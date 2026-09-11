$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path ".env") -and -not (Test-Path ".env.local")) {
    throw "O arquivo .env ainda não foi configurado. Copie .env.example para .env e use as duas chaves do Word administrativo."
}

Write-Host "Dashboard Auditoria - instalação automática" -ForegroundColor Cyan

$python = Join-Path $PSScriptRoot "venv\Scripts\python.exe"

if (Test-Path $python) {
    Write-Host "Verificando o ambiente virtual existente..."
    & $python -c "import flask, numpy, pandas, openpyxl, plotly, reportlab, matplotlib" 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "O venv existente está incompleto ou incompatível. Recriando..." -ForegroundColor Yellow
        Remove-Item ".\venv" -Recurse -Force
    }
}

if (-not (Test-Path $python)) {
    Write-Host "Criando ambiente virtual..."
    $created = $false

    if (Get-Command py -ErrorAction SilentlyContinue) {
        foreach ($version in @("3.13", "3.12", "3.14", "3")) {
            & py "-$version" -m venv venv 2>$null
            if (Test-Path $python) {
                $created = $true
                break
            }
        }
    }

    if (-not $created -and (Get-Command python -ErrorAction SilentlyContinue)) {
        & python -m venv venv
    }
}

if (-not (Test-Path $python)) {
    throw "Não foi possível criar o venv. Instale o Python e marque a opção Add Python to PATH."
}

& $python -m pip install --upgrade pip setuptools wheel
if ($LASTEXITCODE -ne 0) { throw "Não foi possível atualizar o pip." }

& $python -m pip install --upgrade --only-binary=:all: -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "Não foi possível instalar as bibliotecas." }

& $python -c "import flask, numpy, pandas, openpyxl, plotly, reportlab, matplotlib"
if ($LASTEXITCODE -ne 0) { throw "O ambiente virtual foi criado, mas as bibliotecas não carregaram corretamente." }

& $python -c "from app import app; from services.data_service import load_data; from services.paint_service import load_paint_data; from services.carteira_service import load_carteira_data, load_aguardando_pa_data; load_data(); load_paint_data(); load_carteira_data(); load_aguardando_pa_data()"
if ($LASTEXITCODE -ne 0) { throw "As bibliotecas foram instaladas, mas o aplicativo ou uma das planilhas não pôde ser carregado." }

& (Join-Path $PSScriptRoot "ABRIR_DASHBOARD.ps1")
