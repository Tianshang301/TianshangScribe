# Build .deb package for TianshangScribe on Windows using nfpm
param([string]$Version = "")

$ErrorActionPreference = "Stop"
Push-Location $PSScriptRoot

if (-not $Version) {
    $pyProject = Get-Content "pyproject.toml" -Raw
    if ($pyProject -match 'version\s*=\s*"([^"]+)"') {
        $Version = $matches[1]
    } else {
        $Version = "0.0.0"
    }
}

Write-Host "Building TianshangScribe v$Version .deb package..." -ForegroundColor Cyan

$env:VERSION = $Version

$distDir = "dist\deb"
New-Item -ItemType Directory -Force -Path $distDir | Out-Null

nfpm package --config nfpm.yaml --target $distDir --packager deb
if ($LASTEXITCODE -ne 0) { throw "nfpm failed" }

$debFile = Get-ChildItem "$distDir\tianshang-scribe_*.deb" | Select-Object -First 1
$size = [math]::Round($debFile.Length / 1KB, 1)

Write-Host ""
Write-Host "=== Build Complete ===" -ForegroundColor Green
Write-Host "  File : $($debFile.FullName)" -ForegroundColor Green
Write-Host "  Size : $size KB" -ForegroundColor Green
Write-Host ""
Write-Host "Install on Debian/Ubuntu:" -ForegroundColor Yellow
Write-Host "  sudo dpkg -i $($debFile.Name)" -ForegroundColor White
Write-Host "  tianshang-scribe --help" -ForegroundColor White

Pop-Location
