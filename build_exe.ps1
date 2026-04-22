$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

python -m pip install pyinstaller
python -m PyInstaller `
  --noconfirm `
  --clean `
  ".\InterfaceUpdate.spec"

$distDir = Join-Path $root "dist"
$exePrincipal = Join-Path $distDir "Interface Update.exe"

Get-ChildItem $distDir -Filter "*.exe" -ErrorAction SilentlyContinue |
  Where-Object { $_.FullName -ne $exePrincipal } |
  ForEach-Object {
    try {
      Remove-Item $_.FullName -Force -ErrorAction Stop
    }
    catch {
      Write-Warning "Nao foi possivel remover o executavel extra: $($_.FullName)"
    }
  }

$distCorrigido = Join-Path $root "dist_corrigido"
if (Test-Path $distCorrigido) {
  try {
    Remove-Item $distCorrigido -Recurse -Force -ErrorAction Stop
  }
  catch {
    Write-Warning "Nao foi possivel remover a pasta auxiliar: $distCorrigido"
  }
}

Write-Host ""
Write-Host "EXE gerado em: $root\dist\Interface Update.exe"
