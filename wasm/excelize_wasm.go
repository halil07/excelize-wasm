// Package main exposes the excelize library to JavaScript via WebAssembly
// (GOOS=js GOARCH=wasm) for the @halil07/excelize npm package. It binds a
// comprehensive subset of the excelize API to a global `ExcelizeWasm` object
// using syscall/js.
//
// Conventions:
//   - Functions that create a *excelize.File return a numeric handle (id).
//     Handles are managed in an internal registry; call close(id) to release.
//   - Simple arguments are passed as JS primitives. Complex option structs
//     (Style, Table, DataValidation, Chart, ...) are passed as JSON strings or
//     plain JS objects and decoded into the corresponding Go struct.
//   - Byte payloads (XLSX in/out, embedded pictures) use Uint8Array.
//   - Every function returns { ok:true, data:v } on success or
//     { ok:false, err:"message" } on failure. The JS loader turns the latter
//     into thrown errors so callers can use try/catch.
//   - Disk-based APIs (OpenFile/SaveAs) are NOT exposed because browser wasm
//     has no usable filesystem. Use openFile(bytes) / save(id) instead.
//   - Temp-file fallback paths (streaming, >16MB unzip) are disabled by forcing
//     very large in-memory limits in parseOptions; the StreamWriter API is not
//     bound because it requires temp files.
//
// Build:
//
//	GOOS=js GOARCH=wasm go build -o dist/excelize.wasm ./wasm
package main

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"sync"
	"syscall/js"
	"time"

	excelize "github.com/xuri/excelize/v2"
)

// ---------------------------------------------------------------------------
// Argument helpers
// ---------------------------------------------------------------------------

// asStr returns the string at args[i], or "" when unset.
func asStr(args []js.Value, i int) string {
	if i >= len(args) || args[i].IsUndefined() || args[i].IsNull() {
		return ""
	}
	return args[i].String()
}

// asInt returns the int at args[i], or def when unset.
func asInt(args []js.Value, i, def int) int {
	if i >= len(args) || args[i].IsUndefined() || args[i].IsNull() {
		return def
	}
	if args[i].Type() == js.TypeNumber {
		return args[i].Int()
	}
	return def
}

// asFloat returns the float64 at args[i], or def when unset.
func asFloat(args []js.Value, i int, def float64) float64 {
	if i >= len(args) || args[i].IsUndefined() || args[i].IsNull() {
		return def
	}
	if args[i].Type() == js.TypeNumber {
		return args[i].Float()
	}
	return def
}

// asBool returns the bool at args[i], or def when unset.
func asBool(args []js.Value, i int, def bool) bool {
	if i >= len(args) || args[i].IsUndefined() || args[i].IsNull() {
		return def
	}
	if args[i].Type() == js.TypeBoolean {
		return args[i].Bool()
	}
	return def
}

// asOpts returns the raw js.Value carrying options for argument i (or undefined).
func asOpts(args []js.Value, i int) js.Value {
	if i >= len(args) {
		return js.Undefined()
	}
	return args[i]
}

// ---------------------------------------------------------------------------
// Result helpers
// ---------------------------------------------------------------------------

// ok wraps a successful Go value as {ok:true, data:v}.
func ok(v interface{}) js.Value {
	switch v := v.(type) {
	case nil:
		return jsObject("ok", true, "data", nil)
	case js.Value:
		return jsObject("ok", true, "data", v)
	default:
		return jsObject("ok", true, "data", jsonToJS(v))
	}
}

// errResult wraps a Go error as {ok:false, err:msg}; nil errors become ok.
func errResult(err error) js.Value {
	if err == nil {
		return ok(nil)
	}
	return jsObject("ok", false, "err", err.Error())
}

// jsonToJS marshals a Go value to JSON and parses it back as a JS value so
// exported structs become plain JS objects.
func jsonToJS(v interface{}) js.Value {
	b, err := json.Marshal(v)
	if err != nil {
		return js.ValueOf(fmt.Sprintf("%v", v))
	}
	return jsParseJSON(string(b))
}

// jsonFromJS decodes a JS value (object or JSON string) into target.
func jsonFromJS(src js.Value, target any) error {
	if src.IsUndefined() || src.IsNull() {
		return nil
	}
	var raw string
	switch src.Type() {
	case js.TypeString:
		raw = src.String()
		if !json.Valid([]byte(raw)) {
			return fmt.Errorf("invalid json string: %q", raw)
		}
	case js.TypeObject:
		raw = jsStringify(src)
	default:
		return fmt.Errorf("unsupported options type: %v", src.Type())
	}
	return json.Unmarshal([]byte(raw), target)
}

// ---------------------------------------------------------------------------
// JS bridge primitives
// ---------------------------------------------------------------------------

func jsObject(kv ...any) js.Value {
	o := js.Global().Get("Object").New()
	for i := 0; i+1 < len(kv); i += 2 {
		o.Set(fmt.Sprint(kv[i]), js.ValueOf(kv[i+1]))
	}
	return o
}

func jsParseJSON(s string) js.Value {
	return js.Global().Get("JSON").Get("parse").Invoke(s)
}

func jsStringify(v js.Value) string {
	return js.Global().Get("JSON").Get("stringify").Invoke(v).String()
}

// ---------------------------------------------------------------------------
// Byte bridge (Uint8Array <-> []byte)
// ---------------------------------------------------------------------------

// bytesFromJS copies a JS Uint8Array/ArrayBuffer at args[i] into a Go slice.
func bytesFromJS(args []js.Value, i int) ([]byte, error) {
	if i >= len(args) || args[i].IsUndefined() || args[i].IsNull() {
		return nil, nil
	}
	src := args[i]
	if src.Type() != js.TypeObject || src.Get("byteLength").IsUndefined() {
		return nil, fmt.Errorf("argument %d must be a Uint8Array / ArrayBuffer (got %v)", i, src.Type())
	}
	n := src.Get("byteLength").Int()
	if n == 0 {
		return []byte{}, nil
	}
	view := js.Global().Get("Uint8Array").New(src)
	buf := make([]byte, n)
	js.CopyBytesToGo(buf, view)
	return buf, nil
}

// bytesToJS wraps a Go byte slice as a fresh JS Uint8Array (copied).
func bytesToJS(b []byte) js.Value {
	n := len(b)
	arr := js.Global().Get("Uint8Array").New(n)
	if n > 0 {
		js.CopyBytesToJS(arr, b)
	}
	return arr
}

// ---------------------------------------------------------------------------
// Typed value dispatching for SetCellValue
// ---------------------------------------------------------------------------

// valueToInterface converts a JS value into a Go value suitable for
// excelize.SetCellValue (string/float64/bool/nil/time.Time).
func valueToInterface(v js.Value) (interface{}, error) {
	switch v.Type() {
	case js.TypeUndefined, js.TypeNull:
		return nil, nil
	case js.TypeString:
		return v.String(), nil
	case js.TypeNumber:
		return v.Float(), nil
	case js.TypeBoolean:
		return v.Bool(), nil
	case js.TypeObject:
		if ctor := v.Get("constructor"); !ctor.IsUndefined() {
			if name := ctor.Get("name"); !name.IsUndefined() && name.String() == "Date" {
				return time.UnixMilli(int64(v.Call("getTime").Float())), nil
			}
		}
		return jsStringify(v), nil
	default:
		return nil, fmt.Errorf("unsupported cell value type: %v", v.Type())
	}
}

// jsonArrayToInterface decodes a JSON array of mixed primitives for
// SetSheetRow / SetSheetCol.
func jsonArrayToInterface(v js.Value) ([]interface{}, error) {
	if v.IsUndefined() || v.IsNull() {
		return nil, errors.New("missing row/column value array")
	}
	var out []interface{}
	if err := json.Unmarshal([]byte(jsStringify(v)), &out); err != nil {
		return nil, err
	}
	return out, nil
}

// ---------------------------------------------------------------------------
// File handle registry
// ---------------------------------------------------------------------------

var (
	files     = make(map[int]*excelize.File)
	filesMu   sync.Mutex
	filesNext = 1
)

func registerFile(f *excelize.File) int {
	filesMu.Lock()
	defer filesMu.Unlock()
	id := filesNext
	filesNext++
	files[id] = f
	return id
}

func getFile(id int) (*excelize.File, error) {
	filesMu.Lock()
	f, hit := files[id]
	filesMu.Unlock()
	if !hit {
		return nil, fmt.Errorf("unknown file handle: %d (already closed or never created)", id)
	}
	return f, nil
}

func releaseFile(id int) {
	filesMu.Lock()
	if f, hit := files[id]; hit {
		_ = f.Close()
		delete(files, id)
	}
	filesMu.Unlock()
}

// ---------------------------------------------------------------------------
// Options decoding with wasm-safe defaults
// ---------------------------------------------------------------------------

// defaultWasmOptions forces in-memory-only handling to avoid temp-file code
// paths that would call os.CreateTemp/os.Remove (unsupported in browser wasm).
func defaultWasmOptions() excelize.Options {
	return excelize.Options{
		UnzipSizeLimit:    1 << 40, // ~1 TB: prevents fallback to disk extraction
		UnzipXMLSizeLimit: 1 << 40,
	}
}

func parseOptions(src js.Value) (excelize.Options, error) {
	opts := defaultWasmOptions()
	if err := jsonFromJS(src, &opts); err != nil {
		return opts, fmt.Errorf("invalid options: %w", err)
	}
	// Force wasm-safe limits regardless of caller overrides.
	if opts.UnzipXMLSizeLimit <= 0 || opts.UnzipXMLSizeLimit < (1<<30) {
		opts.UnzipXMLSizeLimit = 1 << 40
	}
	if opts.UnzipSizeLimit <= 0 || opts.UnzipSizeLimit < (1<<30) {
		opts.UnzipSizeLimit = 1 << 40
	}
	opts.TmpDir = "" // browser wasm has no writable tmp dir
	return opts, nil
}

// ---------------------------------------------------------------------------
// Binding registry
// ---------------------------------------------------------------------------

var bindings []binding

type binding struct {
	name string
	fn   js.Func
}

func register(name string, cb func(this js.Value, args []js.Value) js.Value) {
	wrapped := func(this js.Value, args []js.Value) any { return cb(this, args) }
	bindings = append(bindings, binding{name, js.FuncOf(wrapped)})
}

// fileOp wires a function whose first argument is a file handle id.
func fileOp(name string, cb func(f *excelize.File, args []js.Value) js.Value) {
	register(name, func(this js.Value, args []js.Value) js.Value {
		if len(args) == 0 || args[0].Type() != js.TypeNumber {
			return errResult(fmt.Errorf("%s: missing file handle id", name))
		}
		f, err := getFile(args[0].Int())
		if err != nil {
			return errResult(err)
		}
		return cb(f, args)
	})
}

func main() {
	registerFileBindings()
	registerSheetBindings()
	registerCellBindings()
	registerStyleBindings()
	registerMergeBindings()
	registerTableBindings()
	registerDataValidationBindings()
	registerPictureBindings()
	registerChartBindings()
	registerHyperLinkBindings()
	registerPageLayoutBindings()
	registerRowColBindings()
	registerOutlineBindings()
	registerDocPropsBindings()
	registerCalcBindings()

	api := js.Global().Get("Object").New()
	for _, b := range bindings {
		api.Set(b.name, b.fn)
	}
	js.Global().Set("ExcelizeWasm", api)

	// Keep the wasm process alive for the lifetime of the page.
	c := make(chan struct{})
	<-c
}

// ---------------------------------------------------------------------------
// File bindings
// ---------------------------------------------------------------------------

func registerFileBindings() {
	register("newFile", func(this js.Value, args []js.Value) js.Value {
		opts, err := parseOptions(asOpts(args, 0))
		if err != nil {
			return errResult(err)
		}
		f := excelize.NewFile(opts)
		return ok(registerFile(f))
	})

	register("openFile", func(this js.Value, args []js.Value) js.Value {
		b, err := bytesFromJS(args, 0)
		if err != nil {
			return errResult(err)
		}
		opts, err := parseOptions(asOpts(args, 1))
		if err != nil {
			return errResult(err)
		}
		f, err := excelize.OpenReader(bytes.NewReader(b), opts)
		if err != nil {
			return errResult(err)
		}
		return ok(registerFile(f))
	})

	fileOp("close", func(f *excelize.File, args []js.Value) js.Value {
		releaseFile(args[0].Int())
		return ok(nil)
	})

	fileOp("save", func(f *excelize.File, args []js.Value) js.Value {
		buf, err := f.WriteToBuffer()
		if err != nil {
			return errResult(err)
		}
		return jsObject("ok", true, "data", bytesToJS(buf.Bytes()))
	})

	fileOp("getSheetList", func(f *excelize.File, args []js.Value) js.Value {
		return ok(f.GetSheetList())
	})

	fileOp("getSheetCount", func(f *excelize.File, args []js.Value) js.Value {
		return ok(f.SheetCount)
	})

	fileOp("defaultSheetName", func(f *excelize.File, args []js.Value) js.Value {
		return ok(f.GetSheetName(0))
	})
}

// ---------------------------------------------------------------------------
// Sheet bindings
// ---------------------------------------------------------------------------

func registerSheetBindings() {
	fileOp("newSheet", func(f *excelize.File, args []js.Value) js.Value {
		idx, err := f.NewSheet(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		return ok(idx)
	})
	fileOp("getSheetIndex", func(f *excelize.File, args []js.Value) js.Value {
		idx, err := f.GetSheetIndex(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		return ok(idx)
	})
	fileOp("getSheetName", func(f *excelize.File, args []js.Value) js.Value {
		return ok(f.GetSheetName(asInt(args, 1, 0)))
	})
	fileOp("setSheetName", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetSheetName(asStr(args, 1), asStr(args, 2)))
	})
	fileOp("deleteSheet", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.DeleteSheet(asStr(args, 1)))
	})
	fileOp("setSheetVisible", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetSheetVisible(asStr(args, 1), asBool(args, 2, true), asBool(args, 3, false)))
	})
	fileOp("getSheetVisible", func(f *excelize.File, args []js.Value) js.Value {
		v, err := f.GetSheetVisible(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		return ok(v)
	})
	fileOp("getSheetMap", func(f *excelize.File, args []js.Value) js.Value {
		return ok(f.GetSheetMap())
	})
	fileOp("setActiveSheet", func(f *excelize.File, args []js.Value) js.Value {
		f.SetActiveSheet(asInt(args, 1, 0))
		return ok(nil)
	})
	fileOp("getActiveSheetIndex", func(f *excelize.File, args []js.Value) js.Value {
		return ok(f.GetActiveSheetIndex())
	})
	fileOp("setSheetBackgroundFromBytes", func(f *excelize.File, args []js.Value) js.Value {
		b, err := bytesFromJS(args, 3)
		if err != nil {
			return errResult(err)
		}
		return errResult(f.SetSheetBackgroundFromBytes(asStr(args, 1), asStr(args, 2), b))
	})
	fileOp("setSheetDimension", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetSheetDimension(asStr(args, 1), asStr(args, 2)))
	})
	fileOp("getSheetDimension", func(f *excelize.File, args []js.Value) js.Value {
		dim, err := f.GetSheetDimension(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		return ok(dim)
	})
}

// ---------------------------------------------------------------------------
// Cell bindings
// ---------------------------------------------------------------------------

func registerCellBindings() {
	fileOp("setCellValue", func(f *excelize.File, args []js.Value) js.Value {
		val, err := valueToInterface(asOpts(args, 3))
		if err != nil {
			return errResult(err)
		}
		return errResult(f.SetCellValue(asStr(args, 1), asStr(args, 2), val))
	})
	fileOp("getCellValue", func(f *excelize.File, args []js.Value) js.Value {
		var opts excelize.Options
		if err := jsonFromJS(asOpts(args, 3), &opts); err != nil {
			return errResult(err)
		}
		v, err := f.GetCellValue(asStr(args, 1), asStr(args, 2), opts)
		if err != nil {
			return errResult(err)
		}
		return ok(v)
	})
	fileOp("setCellFormula", func(f *excelize.File, args []js.Value) js.Value {
		var fo excelize.FormulaOpts
		if err := jsonFromJS(asOpts(args, 4), &fo); err != nil {
			return errResult(err)
		}
		return errResult(f.SetCellFormula(asStr(args, 1), asStr(args, 2), asStr(args, 3), fo))
	})
	fileOp("getCellFormula", func(f *excelize.File, args []js.Value) js.Value {
		v, err := f.GetCellFormula(asStr(args, 1), asStr(args, 2))
		if err != nil {
			return errResult(err)
		}
		return ok(v)
	})
	fileOp("setSheetRow", func(f *excelize.File, args []js.Value) js.Value {
		slice, err := jsonArrayToInterface(asOpts(args, 3))
		if err != nil {
			return errResult(err)
		}
		return errResult(f.SetSheetRow(asStr(args, 1), asStr(args, 2), &slice))
	})
	fileOp("setSheetCol", func(f *excelize.File, args []js.Value) js.Value {
		slice, err := jsonArrayToInterface(asOpts(args, 3))
		if err != nil {
			return errResult(err)
		}
		return errResult(f.SetSheetCol(asStr(args, 1), asStr(args, 2), &slice))
	})
	fileOp("getRows", func(f *excelize.File, args []js.Value) js.Value {
		var opts excelize.Options
		if err := jsonFromJS(asOpts(args, 2), &opts); err != nil {
			return errResult(err)
		}
		rows, err := f.GetRows(asStr(args, 1), opts)
		if err != nil {
			return errResult(err)
		}
		return ok(rows)
	})
	fileOp("getCols", func(f *excelize.File, args []js.Value) js.Value {
		var opts excelize.Options
		if err := jsonFromJS(asOpts(args, 2), &opts); err != nil {
			return errResult(err)
		}
		cols, err := f.GetCols(asStr(args, 1), opts)
		if err != nil {
			return errResult(err)
		}
		return ok(cols)
	})
}

// ---------------------------------------------------------------------------
// Row / column width & visibility
// ---------------------------------------------------------------------------

func registerRowColBindings() {
	fileOp("setColWidth", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetColWidth(asStr(args, 1), asStr(args, 2), asStr(args, 3), asFloat(args, 4, 9.0)))
	})
	fileOp("getColWidth", func(f *excelize.File, args []js.Value) js.Value {
		w, err := f.GetColWidth(asStr(args, 1), asStr(args, 2))
		if err != nil {
			return errResult(err)
		}
		return ok(w)
	})
	fileOp("setColVisible", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetColVisible(asStr(args, 1), asStr(args, 2), asBool(args, 3, true)))
	})
	fileOp("getColVisible", func(f *excelize.File, args []js.Value) js.Value {
		v, err := f.GetColVisible(asStr(args, 1), asStr(args, 2))
		if err != nil {
			return errResult(err)
		}
		return ok(v)
	})
	fileOp("setRowHeight", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetRowHeight(asStr(args, 1), asInt(args, 2, 0), asFloat(args, 3, 15.0)))
	})
	fileOp("getRowHeight", func(f *excelize.File, args []js.Value) js.Value {
		h, err := f.GetRowHeight(asStr(args, 1), asInt(args, 2, 0))
		if err != nil {
			return errResult(err)
		}
		return ok(h)
	})
	fileOp("setRowVisible", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetRowVisible(asStr(args, 1), asInt(args, 2, 0), asBool(args, 3, true)))
	})
	fileOp("getRowVisible", func(f *excelize.File, args []js.Value) js.Value {
		v, err := f.GetRowVisible(asStr(args, 1), asInt(args, 2, 0))
		if err != nil {
			return errResult(err)
		}
		return ok(v)
	})
}

// ---------------------------------------------------------------------------
// Outline level (grouping)
// ---------------------------------------------------------------------------

func registerOutlineBindings() {
	fileOp("setColOutlineLevel", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetColOutlineLevel(asStr(args, 1), asStr(args, 2), uint8(asInt(args, 3, 0))))
	})
	fileOp("getColOutlineLevel", func(f *excelize.File, args []js.Value) js.Value {
		lvl, err := f.GetColOutlineLevel(asStr(args, 1), asStr(args, 2))
		if err != nil {
			return errResult(err)
		}
		return ok(lvl)
	})
	fileOp("getRowOutlineLevel", func(f *excelize.File, args []js.Value) js.Value {
		lvl, err := f.GetRowOutlineLevel(asStr(args, 1), asInt(args, 2, 0))
		if err != nil {
			return errResult(err)
		}
		return ok(lvl)
	})
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

func registerStyleBindings() {
	fileOp("newStyle", func(f *excelize.File, args []js.Value) js.Value {
		var st excelize.Style
		if err := jsonFromJS(asOpts(args, 1), &st); err != nil {
			return errResult(err)
		}
		id, err := f.NewStyle(&st)
		if err != nil {
			return errResult(err)
		}
		return ok(id)
	})
	fileOp("getCellStyle", func(f *excelize.File, args []js.Value) js.Value {
		id, err := f.GetCellStyle(asStr(args, 1), asStr(args, 2))
		if err != nil {
			return errResult(err)
		}
		return ok(id)
	})
	fileOp("setCellStyle", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.SetCellStyle(asStr(args, 1), asStr(args, 2), asStr(args, 3), asInt(args, 4, 0)))
	})
	fileOp("setConditionalFormat", func(f *excelize.File, args []js.Value) js.Value {
		var opts []excelize.ConditionalFormatOptions
		if err := jsonFromJS(asOpts(args, 3), &opts); err != nil {
			return errResult(err)
		}
		return errResult(f.SetConditionalFormat(asStr(args, 1), asStr(args, 2), opts))
	})
	fileOp("getConditionalFormats", func(f *excelize.File, args []js.Value) js.Value {
		m, err := f.GetConditionalFormats(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		return ok(m)
	})
}

// ---------------------------------------------------------------------------
// Merges
// ---------------------------------------------------------------------------

func registerMergeBindings() {
	fileOp("mergeCell", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.MergeCell(asStr(args, 1), asStr(args, 2), asStr(args, 3)))
	})
	fileOp("unmergeCell", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.UnmergeCell(asStr(args, 1), asStr(args, 2), asStr(args, 3)))
	})
	fileOp("getMergeCells", func(f *excelize.File, args []js.Value) js.Value {
		cells, err := f.GetMergeCells(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		if cells == nil {
			cells = []excelize.MergeCell{}
		}
		return ok(cells)
	})
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

func registerTableBindings() {
	fileOp("addTable", func(f *excelize.File, args []js.Value) js.Value {
		var t excelize.Table
		if err := jsonFromJS(asOpts(args, 2), &t); err != nil {
			return errResult(err)
		}
		return errResult(f.AddTable(asStr(args, 1), &t))
	})
	fileOp("getTables", func(f *excelize.File, args []js.Value) js.Value {
		ts, err := f.GetTables(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		if ts == nil {
			ts = []excelize.Table{}
		}
		return ok(ts)
	})
	fileOp("deleteTable", func(f *excelize.File, args []js.Value) js.Value {
		return errResult(f.DeleteTable(asStr(args, 1)))
	})
}

// ---------------------------------------------------------------------------
// Data validation
// ---------------------------------------------------------------------------

func registerDataValidationBindings() {
	fileOp("addDataValidation", func(f *excelize.File, args []js.Value) js.Value {
		var dv excelize.DataValidation
		if err := jsonFromJS(asOpts(args, 2), &dv); err != nil {
			return errResult(err)
		}
		return errResult(f.AddDataValidation(asStr(args, 1), &dv))
	})
}

// ---------------------------------------------------------------------------
// Pictures
// ---------------------------------------------------------------------------

func registerPictureBindings() {
	fileOp("addPictureFromBytes", func(f *excelize.File, args []js.Value) js.Value {
		// args: id, sheet, cell, extension, bytes(, graphicOptionsJSON)
		b, err := bytesFromJS(args, 4)
		if err != nil {
			return errResult(err)
		}
		pic := &excelize.Picture{
			Extension: asStr(args, 3),
			File:      b,
		}
		if opts := asOpts(args, 5); !opts.IsUndefined() && !opts.IsNull() {
			var g excelize.GraphicOptions
			if err := jsonFromJS(opts, &g); err != nil {
				return errResult(err)
			}
			pic.Format = &g
		}
		return errResult(f.AddPictureFromBytes(asStr(args, 1), asStr(args, 2), pic))
	})
	fileOp("getPictures", func(f *excelize.File, args []js.Value) js.Value {
		pics, err := f.GetPictures(asStr(args, 1), asStr(args, 2))
		if err != nil {
			return errResult(err)
		}
		return ok(pics)
	})
}

// ---------------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------------

func registerChartBindings() {
	fileOp("addChart", func(f *excelize.File, args []js.Value) js.Value {
		var c excelize.Chart
		if err := jsonFromJS(asOpts(args, 3), &c); err != nil {
			return errResult(err)
		}
		return errResult(f.AddChart(asStr(args, 1), asStr(args, 2), &c))
	})
	fileOp("addChartSheet", func(f *excelize.File, args []js.Value) js.Value {
		var c excelize.Chart
		if err := jsonFromJS(asOpts(args, 2), &c); err != nil {
			return errResult(err)
		}
		return errResult(f.AddChartSheet(asStr(args, 1), &c))
	})
}

// ---------------------------------------------------------------------------
// Hyperlinks
// ---------------------------------------------------------------------------

func registerHyperLinkBindings() {
	fileOp("setCellHyperLink", func(f *excelize.File, args []js.Value) js.Value {
		var opts []excelize.HyperlinkOpts
		if src := asOpts(args, 5); !src.IsUndefined() && !src.IsNull() {
			var ho excelize.HyperlinkOpts
			if err := jsonFromJS(src, &ho); err != nil {
				return errResult(err)
			}
			opts = append(opts, ho)
		}
		return errResult(f.SetCellHyperLink(asStr(args, 1), asStr(args, 2), asStr(args, 3), asStr(args, 4), opts...))
	})
	fileOp("getCellHyperLink", func(f *excelize.File, args []js.Value) js.Value {
		has, link, err := f.GetCellHyperLink(asStr(args, 1), asStr(args, 2))
		if err != nil {
			return errResult(err)
		}
		return jsObject("ok", true, "data", jsObject("hasLink", has, "link", link))
	})
}

// ---------------------------------------------------------------------------
// Page layout & sheet view / props
// ---------------------------------------------------------------------------

func registerPageLayoutBindings() {
	fileOp("setPageLayout", func(f *excelize.File, args []js.Value) js.Value {
		var po excelize.PageLayoutOptions
		if err := jsonFromJS(asOpts(args, 2), &po); err != nil {
			return errResult(err)
		}
		return errResult(f.SetPageLayout(asStr(args, 1), &po))
	})
	fileOp("getPageLayout", func(f *excelize.File, args []js.Value) js.Value {
		po, err := f.GetPageLayout(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		return ok(po)
	})
	fileOp("setSheetView", func(f *excelize.File, args []js.Value) js.Value {
		var vo excelize.ViewOptions
		if err := jsonFromJS(asOpts(args, 3), &vo); err != nil {
			return errResult(err)
		}
		return errResult(f.SetSheetView(asStr(args, 1), asInt(args, 2, 0), &vo))
	})
	fileOp("getSheetView", func(f *excelize.File, args []js.Value) js.Value {
		vo, err := f.GetSheetView(asStr(args, 1), asInt(args, 2, 0))
		if err != nil {
			return errResult(err)
		}
		return ok(vo)
	})
	fileOp("setSheetProps", func(f *excelize.File, args []js.Value) js.Value {
		var so excelize.SheetPropsOptions
		if err := jsonFromJS(asOpts(args, 2), &so); err != nil {
			return errResult(err)
		}
		return errResult(f.SetSheetProps(asStr(args, 1), &so))
	})
	fileOp("getSheetProps", func(f *excelize.File, args []js.Value) js.Value {
		so, err := f.GetSheetProps(asStr(args, 1))
		if err != nil {
			return errResult(err)
		}
		return ok(so)
	})
}

// ---------------------------------------------------------------------------
// Document properties
// ---------------------------------------------------------------------------

func registerDocPropsBindings() {
	fileOp("setAppProps", func(f *excelize.File, args []js.Value) js.Value {
		var p excelize.AppProperties
		if err := jsonFromJS(asOpts(args, 1), &p); err != nil {
			return errResult(err)
		}
		return errResult(f.SetAppProps(&p))
	})
	fileOp("getAppProps", func(f *excelize.File, args []js.Value) js.Value {
		p, err := f.GetAppProps()
		if err != nil {
			return errResult(err)
		}
		return ok(p)
	})
	fileOp("setDocProps", func(f *excelize.File, args []js.Value) js.Value {
		var p excelize.DocProperties
		if err := jsonFromJS(asOpts(args, 1), &p); err != nil {
			return errResult(err)
		}
		return errResult(f.SetDocProps(&p))
	})
	fileOp("getDocProps", func(f *excelize.File, args []js.Value) js.Value {
		p, err := f.GetDocProps()
		if err != nil {
			return errResult(err)
		}
		return ok(p)
	})
}

// ---------------------------------------------------------------------------
// Formula calculation
// ---------------------------------------------------------------------------

func registerCalcBindings() {
	fileOp("calcCellValue", func(f *excelize.File, args []js.Value) js.Value {
		var opts excelize.Options
		if err := jsonFromJS(asOpts(args, 3), &opts); err != nil {
			return errResult(err)
		}
		result, err := f.CalcCellValue(asStr(args, 1), asStr(args, 2), opts)
		if err != nil {
			return errResult(err)
		}
		return ok(result)
	})
	fileOp("setCalcProps", func(f *excelize.File, args []js.Value) js.Value {
		var p excelize.CalcPropsOptions
		if err := jsonFromJS(asOpts(args, 1), &p); err != nil {
			return errResult(err)
		}
		return errResult(f.SetCalcProps(&p))
	})
	fileOp("getCalcProps", func(f *excelize.File, args []js.Value) js.Value {
		p, err := f.GetCalcProps()
		if err != nil {
			return errResult(err)
		}
		return ok(p)
	})
}
