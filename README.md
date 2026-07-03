# @halil07/excelize

[![npm version](https://img.shields.io/npm/v/@halil07/excelize.svg)](https://www.npmjs.com/package/@halil07/excelize)

Read & write Microsoft Excel (`XLSX` / `XLSM` / `XLTM` / `XLTX`) files in the
browser and Node.js through WebAssembly.

This package is a thin WebAssembly binding around the Go
[excelize](https://github.com/xuri/excelize) library. The excelize source is
imported directly from GitHub (`github.com/xuri/excelize/v2`); this repository
only ships the WASM binding layer (`wasm/excelize_wasm.go`) and a friendly
JavaScript loader (`wasm/excelize.js`).

## Features

- Create workbooks in memory and export them as `Uint8Array` (`xlsx`).
- Open existing workbooks from `Uint8Array` / `ArrayBuffer`.
- Read & write cells, formulas, rows, columns.
- Styles, conditional formatting, merges, tables, data validation.
- Pictures, charts, hyperlinks.
- Page layout, sheet views, document properties, formula calculation.

## Install

```bash
npm install @halil07/excelize
# or
yarn add @halil07/excelize
```

## Usage — Node.js

```js
const fs = require("fs");
const Excelize = require("@halil07/excelize");

(async () => {
  await Excelize.ready();          // auto-loads the bundled wasm runtime

  const f = Excelize.newFile();
  f.setCellValue("Sheet1", "A1", "Hello from Node!");
  f.setCellValue("Sheet1", "A2", 41);
  f.setCellFormula("Sheet1", "A3", "=A2+1");
  console.log(f.getRows("Sheet1"));

  const bytes = f.save();          // Uint8Array
  f.close();
  fs.writeFileSync("out.xlsx", bytes);
})();
```

## Usage — Browser

Load the bundled `dist/wasm_exec.js` and the loader, then instantiate the
`.wasm` module with `new Go()` and call `await Excelize.ready()`. A complete
runnable example lives in [`wasm/example.html`](wasm/example.html).

```html
<script src="node_modules/@halil07/excelize/dist/wasm_exec.js"></script>
<script src="node_modules/@halil07/excelize/dist/excelize.js"></script>
<script>
  const go = new Go();
  WebAssembly.instantiateStreaming(
    fetch("node_modules/@halil07/excelize/dist/excelize.wasm"),
    go.importObject
  ).then((res) => {
    go.run(res.instance);          // registers globalThis.ExcelizeWasm
  });
</script>
```

## API surface

The default export is an `Excelize` object plus a `File` class returned by
`newFile()` / `openFile()`. Every method returns the computed value
(`Uint8Array`, `string`, `number`, `Array`, …) on success and throws a JS
`Error` on failure (the underlying `{ ok:false, err }` is unwrapped).

| Area | Methods |
|------|---------|
| Workbook | `newFile`, `openFile`, `save`, `close`, `getSheetList`, `getSheetCount`, `defaultSheetName` |
| Sheets | `newSheet`, `getSheetIndex`, `getSheetName`, `setSheetName`, `deleteSheet`, `setSheetVisible`, `getSheetVisible`, `getSheetMap`, `setActiveSheet`, `getActiveSheetIndex`, `setSheetBackgroundFromBytes`, `setSheetDimension`, `getSheetDimension` |
| Cells | `setCellValue`, `getCellValue`, `setCellFormula`, `getCellFormula`, `setSheetRow`, `setSheetCol`, `getRows`, `getCols` |
| Rows / Cols | `setColWidth`, `getColWidth`, `setColVisible`, `getColVisible`, `setRowHeight`, `getRowHeight`, `setRowVisible`, `getRowVisible` |
| Outline | `setColOutlineLevel`, `getColOutlineLevel`, `getRowOutlineLevel` |
| Styles | `newStyle`, `getCellStyle`, `setCellStyle`, `setConditionalFormat`, `getConditionalFormats` |
| Merges | `mergeCell`, `unmergeCell`, `getMergeCells` |
| Tables | `addTable`, `getTables`, `deleteTable` |
| Validation | `addDataValidation` |
| Pictures | `addPictureFromBytes`, `getPictures` |
| Charts | `addChart`, `addChartSheet` |
| Hyperlinks | `setCellHyperLink`, `getCellHyperLink` |
| Layout / views / props | `setPageLayout`, `getPageLayout`, `setSheetView`, `getSheetView`, `setSheetProps`, `getSheetProps` |
| Doc props | `setAppProps`, `getAppProps`, `setDocProps`, `getDocProps` |
| Calc | `calcCellValue`, `setCalcProps`, `getCalcProps` |

Complex option structs (`Style`, `Table`, `DataValidation`, `Chart`, …) are
passed as plain JS objects and JSON-decoded into the corresponding Go struct.

## Building from source

Requires Go ≥ 1.25 and Node.js.

```bash
npm install
npm run build      # cross-platform: scripts/build.js
# outputs dist/excelize.wasm, dist/wasm_exec.js, dist/excelize.js
```

PowerShell alternative: `./wasm/build.ps1`.

## Limitations

- Browser/Node wasm has no usable writable filesystem, so the disk-based
  `OpenFile` / `SaveAs` APIs are intentionally **not** bound. Open / save
  bytes in memory instead (`openFile(Uint8Array)` / `save()`).
- The streaming `StreamWriter` API and the temp-file fallback path for very
  large worksheets are disabled (browser wasm cannot create temp files). Very
  large workbooks are kept fully in memory instead.

## Publishing

The published npm package only ships the prebuilt `dist/` artifacts plus the
README and LICENSE (see the `files` field in `package.json`). `npm publish`
runs the build automatically through the `prepublishOnly` script.

## License

BSD-3-Clause — same license as the upstream excelize library.