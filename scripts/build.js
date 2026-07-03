// Cross-platform build script for @halil07/excelize.
// Produces dist/excelize.wasm, dist/wasm_exec.js and dist/excelize.js.
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

console.log("Bundling wasm_exec.js + loader...");
fs.copyFileSync(wasmExec, path.join(dist, "wasm_exec.js"));
fs.copyFileSync(path.join(root, "wasm", "excelize.js"), path.join(dist, "excelize.js"));

console.log("\nDone. Artifacts in dist/:");
for (const f of fs.readdirSync(dist)) {
  const sz = fs.statSync(path.join(dist, f)).size;
  console.log("  " + f + "  " + (sz / 1024).toFixed(1) + " KB");
}