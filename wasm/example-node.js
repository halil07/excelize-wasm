// Node.js example for the @halil07/excelize npm package.
//
//   npm run build
//   node wasm/example-node.js
//
// In Node the loader auto-loads the bundled Go runtime + wasm, so all you
// need is an `await Excelize.ready()` before using the API.
"use strict";

const fs = require("fs");
const path = require("path");
const Excelize = require(path.join(__dirname, "..", "dist", "excelize.js"));

(async function main() {
  await Excelize.ready();

  const f = Excelize.newFile();
  f.setCellValue("Sheet1", "A1", "Hello from Node + @halil07/excelize!");
  f.setCellValue("Sheet1", "A2", 42);
  f.setCellFormula("Sheet1", "A3", "=A2*2");
  console.log("rows =", JSON.stringify(f.getRows("Sheet1")));

  const bytes = f.save();
  f.close();
  fs.writeFileSync(path.join(__dirname, "out.xlsx"), bytes);
  console.log("Wrote", bytes.byteLength, "bytes to wasm/out.xlsx");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});