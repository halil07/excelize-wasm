// Cross-platform build script for @halil07/excelize.
// Produces dist/excelize.wasm, dist/wasm_exec.js, dist/excelize.js (UMD),
// dist/excelize.cjs (CommonJS) and dist/excelize.mjs (ES module).
// Usage: node scripts/build.js
"use strict";

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist");
const wasmExec = path.join(
  execSync("go env GOROOT", { encoding: "utf8" }).trim(),
  "lib",
  "wasm",
  "wasm_exec.js"
);

fs.mkdirSync(dist, { recursive: true });

console.log("Building excelize.wasm (GOOS=js GOARCH=wasm)...");
execSync("go build -trimpath -ldflags=\"-s -w\" -o dist/excelize.wasm ./wasm", {
  stdio: "inherit",
  env: Object.assign({}, process.env, { GOOS: "js", GOARCH: "wasm" }),
  cwd: root,
});

console.log("Bundling wasm_exec.js + loaders...");
fs.copyFileSync(wasmExec, path.join(dist, "wasm_exec.js"));

const loaderSrc = fs.readFileSync(path.join(root, "wasm", "excelize.js"), "utf8");

// UMD build (browser + legacy CJS fallback).
fs.writeFileSync(path.join(dist, "excelize.js"), loaderSrc);

// Explicit CommonJS build.
fs.writeFileSync(path.join(dist, "excelize.cjs"), loaderSrc);

// TypeScript declarations: generic .d.ts plus ESM/CJS-specific variants.
const dtsSrc = path.join(root, "wasm", "excelize.d.ts");
fs.copyFileSync(dtsSrc, path.join(dist, "excelize.d.ts"));
fs.copyFileSync(dtsSrc, path.join(dist, "excelize.d.mts"));
fs.copyFileSync(dtsSrc, path.join(dist, "excelize.d.cts"));

// ESM wrapper that loads the UMD bundle via createRequire and re-exports the API.
// Functions are used for the active bindings so that the wasm runtime is not
// touched at import time (e.g. the `raw` getter must not be evaluated before
// the user has awaited excelizeInit()).
const esmWrapper = `import { createRequire } from "module";
const require = createRequire(import.meta.url);
const Excelize = require("./excelize.js");

export const excelizeInit = (...args) => Excelize.excelizeInit(...args);
export const ready = (...args) => Excelize.ready(...args);
export const newFile = (...args) => Excelize.newFile(...args);
export const openFile = (...args) => Excelize.openFile(...args);
export const File = Excelize.File;
export { Excelize };
export default Excelize;
`;
fs.writeFileSync(path.join(dist, "excelize.mjs"), esmWrapper);

console.log("\nDone. Artifacts in dist/:");
for (const f of fs.readdirSync(dist)) {
  const sz = fs.statSync(path.join(dist, f)).size;
  console.log("  " + f + "  " + (sz / 1024).toFixed(1) + " KB");
}
