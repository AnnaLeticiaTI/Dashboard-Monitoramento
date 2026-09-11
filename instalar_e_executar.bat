@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"

if exist "%~dp0.env" goto configuracao_ok
if exist "%~dp0.env.local" goto configuracao_ok
goto erro_configuracao

:configuracao_ok

set "DASH_PYTHON=%~dp0venv\Scripts\python.exe"

echo ==============================================
echo  Dashboard Auditoria - instalacao automatica
echo ==============================================
echo.

if not exist "%DASH_PYTHON%" goto criar_venv

echo Verificando o ambiente virtual existente...
"%DASH_PYTHON%" -c "import flask, numpy, pandas, openpyxl, plotly, reportlab, matplotlib" >nul 2>&1
if not errorlevel 1 goto validar_aplicativo

echo O ambiente virtual esta incompleto ou incompativel.
echo Recriando o ambiente virtual...
rmdir /s /q "%~dp0venv"

:criar_venv
echo Criando ambiente virtual...

where py >nul 2>&1
if errorlevel 1 goto criar_com_python

py -3.13 -m venv "%~dp0venv" >nul 2>&1
if exist "%DASH_PYTHON%" goto instalar_pacotes

py -3.12 -m venv "%~dp0venv" >nul 2>&1
if exist "%DASH_PYTHON%" goto instalar_pacotes

py -3.14 -m venv "%~dp0venv" >nul 2>&1
if exist "%DASH_PYTHON%" goto instalar_pacotes

py -3 -m venv "%~dp0venv" >nul 2>&1
if exist "%DASH_PYTHON%" goto instalar_pacotes

:criar_com_python
where python >nul 2>&1
if errorlevel 1 goto erro_python

python -m venv "%~dp0venv"
if not exist "%DASH_PYTHON%" goto erro_python

:instalar_pacotes
echo Atualizando o instalador de pacotes...
"%DASH_PYTHON%" -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto erro_instalacao

echo Instalando as bibliotecas necessarias...
"%DASH_PYTHON%" -m pip install --upgrade --only-binary=:all: -r "%~dp0requirements.txt"
if errorlevel 1 goto erro_instalacao

:validar_aplicativo
echo Conferindo as bibliotecas, o aplicativo e as planilhas...
"%DASH_PYTHON%" -c "import flask, numpy, pandas, openpyxl, plotly, reportlab, matplotlib; from app import app; from services.data_service import load_data; from services.paint_service import load_paint_data; from services.carteira_service import load_carteira_data, load_aguardando_pa_data; load_data(); load_paint_data(); load_carteira_data(); load_aguardando_pa_data()"
if errorlevel 1 goto erro_aplicativo

echo Ambiente validado com sucesso.
call "%~dp0ABRIR_DASHBOARD.bat"
set "DASH_RESULT=%ERRORLEVEL%"
endlocal & exit /b %DASH_RESULT%

:erro_python
echo.
echo ERRO: nao foi possivel criar o ambiente virtual.
echo Instale o Python e marque a opcao Add Python to PATH.
pause
endlocal & exit /b 1

:erro_configuracao
echo.
echo ERRO: o arquivo .env ainda nao foi configurado.
echo Copie .env.example para .env e use as duas chaves do Word administrativo.
echo Nunca envie o arquivo .env para o GitHub.
pause
endlocal & exit /b 1

:erro_instalacao
echo.
echo ERRO: as bibliotecas nao foram instaladas corretamente.
echo Verifique a conexao com a internet e execute este arquivo novamente.
pause
endlocal & exit /b 1

:erro_aplicativo
echo.
echo ERRO: o aplicativo ou uma das planilhas nao foi carregado.
echo A mensagem exibida acima indica o arquivo que precisa ser verificado.
pause
endlocal & exit /b 1
