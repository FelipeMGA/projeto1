param(
  [string]$InnoCompilerPath = "",
  [string]$ScriptPath = "installer/windows/MXM_Setup.iss"
)

$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo

if (-not (Test-Path $ScriptPath)) {
  Write-Error "Script Inno Setup não encontrado: $ScriptPath"
  exit 1
}

if ([string]::IsNullOrWhiteSpace($InnoCompilerPath)) {
  $candidates = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
  )
  foreach ($c in $candidates) {
    if (Test-Path $c) { $InnoCompilerPath = $c; break }
  }
}

if (-not (Test-Path $InnoCompilerPath)) {
  Write-Error "ISCC.exe não encontrado. Instale Inno Setup 6 e passe -InnoCompilerPath."
  exit 1
}

& $InnoCompilerPath $ScriptPath
if ($LASTEXITCODE -ne 0) {
  Write-Error "Falha ao compilar instalador Inno Setup"
  exit $LASTEXITCODE
}

Write-Host "Instalador gerado com sucesso." -ForegroundColor Green
