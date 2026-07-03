// @halil07/excelize — JS loader & friendly wrapper around the excelize wasm.
//
// Provides an `Excelize` object with:
//   await Excelize.ready();          // ensure the wasm module is initialized
//   const f = Excelize.newFile();   // create a workbook (in-memory)
//   const f2 = Excelize.openFile(uint8array);
//   f.setCellValue("Sheet1", "A1", "hello");
//   const bytes = f.save();          // => Uint8Array (xlsx)
//   f.close();
//
// Remaining raw bindings live on Excelize.raw (the global `ExcelizeWasm`
// object created by wasm_exec.js); each returns { ok, data? } / { ok:false, err }.
// This wrapper converts the latter into thrown JS errors and adds method
// shorthand on the returned File wrapper.

(function (root, factory) {
  if (typeof module === "object" && typeof module.exports === "object") {
    module.exports = factory();
  } else {
    root.Excelize = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const STATE = { loading: null, ready: false, raw: null, isNode: false };

  // Detect Node so we can auto-load the Go runtime + wasm from dist/.
  if (typeof process !== "undefined" && process.versions && process.versions.node) {
    STATE.isNode = true;
  }

  function getRaw() {
    if (STATE.raw) return STATE.raw;
    if (typeof ExcelizeWasm === "undefined") {
      throw new Error(
        "ExcelizeWasm is not initialized. In Node call Excelize.ready() first; " +
          "in the browser load dist/wasm_exec.js and instantiate dist/excelize.wasm " +
          "with `new Go()` (see wasm/example.html)."
      );
    }
    STATE.raw = ExcelizeWasm;
    return STATE.raw;
  }

  // In Node, load the bundled wasm_exec.js (which defines global Go) and
  // instantiate dist/excelize.wasm. Idempotent; returns a promise.
  function loadRuntime() {
    if (STATE.loading) return STATE.loading;
    if (STATE.raw) return Promise.resolve();
    if (!STATE.isNode) {
      // Browser: rely on the page having set up the Go runtime already.
      return Promise.resolve().then(getRaw);
    }
    const path = require("path");
    const fs = require("fs");
    STATE.loading = (async () => {
      const here = path.dirname(fs.realpathSync(__filename));
      require(path.join(here, "wasm_exec.js"));
      const go = new globalThis.Go();
      const wasmBytes = fs.readFileSync(path.join(here, "excelize.wasm"));
      const wasm = await WebAssembly.instantiate(wasmBytes, go.importObject);
      // run() returns a promise that stays pending (our main() blocks forever).
      go.run(wasm.instance).catch(() => {});
      // Wait until the bindings object appears on globalThis.
      while (typeof globalThis.ExcelizeWasm === "undefined") {
        await new Promise((r) => setTimeout(r, 5));
      }
      STATE.raw = globalThis.ExcelizeWasm;
      STATE.ready = true;
    })();
    return STATE.loading;
  }

  // ready() resolves once the wasm bindings are available. In Node it auto
  // loads the runtime; in the browser it expects the page to have started the
  // Go wasm instance first (see wasm/example.html). Calling multiple times is
  // safe.
  function ready() {
    if (STATE.raw) return Promise.resolve();
    return loadRuntime();
  }

  function unwrap(r, missing) {
    if (r && r.ok) return r.data === undefined ? missing : r.data;
    const msg = (r && r.err) || "unknown excelize error";
    throw new Error(msg);
  }

  // File wrapper bound to a numeric handle id.
  class File {
    constructor(id) {
      this.id = id;
      this._closed = false;
    }
    _call(method, ...args) {
      if (this._closed) throw new Error("file handle already closed");
      return unwrap(getRaw()[method](this.id, ...args));
    }
    close() {
      if (this._closed) return;
      this._closed = true;
      getRaw().close(this.id);
    }

    // ---- file / workbook ---------------------------------------------------
    save() { return this._call("save"); }
    getSheetList() { return this._call("getSheetList"); }
    getSheetCount() { return this._call("getSheetCount"); }
    defaultSheetName() { return this._call("defaultSheetName"); }

    // ---- sheets ------------------------------------------------------------
    newSheet(name) { return this._call("newSheet", name); }
    getSheetIndex(name) { return this._call("getSheetIndex", name); }
    getSheetName(index) { return this._call("getSheetName", index); }
    setSheetName(source, target) { return this._call("setSheetName", source, target); }
    deleteSheet(name) { return this._call("deleteSheet", name); }
    setSheetVisible(name, visible, veryHidden) { return this._call("setSheetVisible", name, visible, !!veryHidden); }
    getSheetVisible(name) { return this._call("getSheetVisible", name); }
    getSheetMap() { return this._call("getSheetMap"); }
    setActiveSheet(index) { return this._call("setActiveSheet", index); }
    getActiveSheetIndex() { return this._call("getActiveSheetIndex"); }
    setSheetBackgroundFromBytes(name, ext, bytes) { return this._call("setSheetBackgroundFromBytes", name, ext, bytes); }
    setSheetDimension(name, range) { return this._call("setSheetDimension", name, range); }
    getSheetDimension(name) { return this._call("getSheetDimension", name); }

    // ---- cells -------------------------------------------------------------
    setCellValue(sheet, cell, value) { return this._call("setCellValue", sheet, cell, value); }
    getCellValue(sheet, cell, opts) { return this._call("getCellValue", sheet, cell, opts || null); }
    setCellFormula(sheet, cell, formula, opts) { return this._call("setCellFormula", sheet, cell, formula, opts || null); }
    getCellFormula(sheet, cell) { return this._call("getCellFormula", sheet, cell); }
    setSheetRow(sheet, cell, arr) { return this._call("setSheetRow", sheet, cell, arr); }
    setSheetCol(sheet, cell, arr) { return this._call("setSheetCol", sheet, cell, arr); }
    getRows(sheet, opts) { return this._call("getRows", sheet, opts || null); }
    getCols(sheet, opts) { return this._call("getCols", sheet, opts || null); }

    // ---- rows / cols -------------------------------------------------------
    setColWidth(sheet, c1, c2, width) { return this._call("setColWidth", sheet, c1, c2, width); }
    getColWidth(sheet, col) { return this._call("getColWidth", sheet, col); }
    setColVisible(sheet, cols, visible) { return this._call("setColVisible", sheet, cols, !!visible); }
    getColVisible(sheet, col) { return this._call("getColVisible", sheet, col); }
    setRowHeight(sheet, row, height) { return this._call("setRowHeight", sheet, row, height); }
    getRowHeight(sheet, row) { return this._call("getRowHeight", sheet, row); }
    setRowVisible(sheet, row, visible) { return this._call("setRowVisible", sheet, row, !!visible); }
    getRowVisible(sheet, row) { return this._call("getRowVisible", sheet, row); }

    // ---- outline -----------------------------------------------------------
    setColOutlineLevel(sheet, col, level) { return this._call("setColOutlineLevel", sheet, col, level); }
    getColOutlineLevel(sheet, col) { return this._call("getColOutlineLevel", sheet, col); }
    getRowOutlineLevel(sheet, row) { return this._call("getRowOutlineLevel", sheet, row); }

    // ---- styles ------------------------------------------------------------
    newStyle(style) { return this._call("newStyle", style); }
    getCellStyle(sheet, cell) { return this._call("getCellStyle", sheet, cell); }
    setCellStyle(sheet, a, b, styleID) { return this._call("setCellStyle", sheet, a, b, styleID); }
    setConditionalFormat(sheet, range, opts) { return this._call("setConditionalFormat", sheet, range, opts); }
    getConditionalFormats(sheet) { return this._call("getConditionalFormats", sheet); }

    // ---- merges ------------------------------------------------------------
    mergeCell(sheet, a, b) { return this._call("mergeCell", sheet, a, b); }
    unmergeCell(sheet, a, b) { return this._call("unmergeCell", sheet, a, b); }
    getMergeCells(sheet) { return this._call("getMergeCells", sheet); }

    // ---- tables ------------------------------------------------------------
    addTable(sheet, table) { return this._call("addTable", sheet, table); }
    getTables(sheet) { return this._call("getTables", sheet); }
    deleteTable(name) { return this._call("deleteTable", name); }

    // ---- data validation ---------------------------------------------------
    addDataValidation(sheet, dv) { return this._call("addDataValidation", sheet, dv); }

    // ---- pictures ----------------------------------------------------------
    addPictureFromBytes(sheet, cell, extension, bytes, graphicOptions) {
      return this._call("addPictureFromBytes", sheet, cell, extension, bytes, graphicOptions || null);
    }
    getPictures(sheet, cell) { return this._call("getPictures", sheet, cell); }

    // ---- charts ------------------------------------------------------------
    addChart(sheet, cell, chart) { return this._call("addChart", sheet, cell, chart); }
    addChartSheet(sheet, chart) { return this._call("addChartSheet", sheet, chart); }

    // ---- hyperlinks --------------------------------------------------------
    setCellHyperLink(sheet, cell, link, type, opts) { return this._call("setCellHyperLink", ship(sheet), cell, link, type, opts || null); }
    getCellHyperLink(sheet, cell) { return this._call("getCellHyperLink", sheet, cell); }

    // ---- page layout / views / props ---------------------------------------
    setPageLayout(sheet, opts) { return this._call("setPageLayout", sheet, opts || null); }
    getPageLayout(sheet) { return this._call("getPageLayout", sheet); }
    setSheetView(sheet, index, opts) { return this._call("setSheetView", sheet, index, opts || null); }
    getSheetView(sheet, index) { return this._call("getSheetView", sheet, index); }
    setSheetProps(sheet, opts) { return this._call("setSheetProps", sheet, opts || null); }
    getSheetProps(sheet) { return this._call("getSheetProps", sheet); }

    // ---- document properties ------------------------------------------------
    setAppProps(props) { return this._call("setAppProps", props); }
    getAppProps() { return this._call("getAppProps"); }
    setDocProps(props) { return this._call("setDocProps", props); }
    getDocProps() { return this._call("getDocProps"); }

    // ---- calculation --------------------------------------------------------
    calcCellValue(sheet, cell, opts) { return this._call("calcCellValue", sheet, cell, opts || null); }
    setCalcProps(opts) { return this._call("setCalcProps", opts); }
    getCalcProps() { return this._call("getCalcProps"); }
  }

  // Fix typo guard for setCellHyperLink above (kept minimal).
  function ship(s) { return s; }
  File.prototype.setCellHyperLink = function (sheet, cell, link, type, opts) {
    return this._call("setCellHyperLink", sheet, cell, link, type, opts || null);
  };

  const Excelize = {
    ready,
    get raw() { return getRaw(); },
    newFile(opts) { getRaw(); return new File(unwrap(getRaw().newFile(opts || null))); },
    openFile(bytes, opts) {
      if (!bytes || typeof bytes.byteLength !== "number") {
        throw new Error("openFile: first argument must be a Uint8Array / ArrayBuffer");
      }
      return new File(unwrap(getRaw().openFile(bytes, opts || null)));
    },
    File,
  };

  return Excelize;
});