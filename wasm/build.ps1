# Build @halil07/excelize for Windows / PowerShell (no make required).
# Usage:  pwsh ./wasm/build.ps1   (or  ./wasm/build.ps1)
$ErrorActionPreference = "Stop"

$root  = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$dist  = Join-Path $root "dist"
New-Item -ItemType Directory -Force -Path $dist | Out-Null

$env:GOOS   = "js"
$env:GOARCH = "wasm"

Write-Host "Building excelize.wasm..."
go build -trimpath -ldflags="-s -w" -o (Join-Path $dist "excelize.wasm") (Join-Path $root "wasm")

Write-Host "Bundling wasm_exec.js + loader..."
Copy-Item -Force (Join-Path (go env GOROOT) "lib/wasm/wasm_exec.js") (Join-Path $dist "wasm_exec.js")
Copy-Item -Force (Join-Path $root "wasm/excelize.js") (Join-Path $dist "excelize.js")

Write-Host "Done. Artifacts in: $dist"
Get-ChildItem $dist | Select-Object Name, @{N="KB";E={[int]($_.Length/1KB)}} | Format-Table