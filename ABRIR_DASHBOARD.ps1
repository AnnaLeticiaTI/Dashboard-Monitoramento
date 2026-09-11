$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path ".env") -and -not (Test-Path ".env.local")) {
    throw "O arquivo .env ainda não foi configurado. Copie .env.example para .env e use as duas chaves do Word administrativo."
}

$python = Join-Path $PSScriptRoot "venv\Scripts\python.exe"

if (-not (Test-Path $python)) {
    throw "O ambiente virtual ainda não foi criado. Execute primeiro instalar_e_executar.ps1."
}

& $python -c "import flask; from app import app"
if ($LASTEXITCODE -ne 0) {
    throw "O ambiente virtual está incompleto. Execute instalar_e_executar.ps1 para corrigi-lo."
}

$env:PYTHONUTF8 = "1"
$env:PYTHONUNBUFFERED = "1"
$env:DASHBOARD_OPEN_BROWSER = "1"

Write-Host "Iniciando o Dashboard Auditoria..." -ForegroundColor Cyan
Write-Host "O navegador será aberto em http://127.0.0.1:5000"
Write-Host "Para encerrar o servidor, pressione Ctrl+C nesta janela."

& $python (Join-Path $PSScriptRoot "app.py")
