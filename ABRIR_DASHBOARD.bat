@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"

if exist "%~dp0.env" goto configuracao_ok
if exist "%~dp0.env.local" goto configuracao_ok
echo ERRO: o arquivo .env ainda nao foi configurado.
echo Copie .env.example para .env e use as duas chaves do Word administrativo.
pause
endlocal & exit /b 1

:configuracao_ok

set "DASH_PYTHON=%~dp0venv\Scripts\python.exe"
set "PYTHONUTF8=1"
set "PYTHONUNBUFFERED=1"
set "DASHBOARD_OPEN_BROWSER=1"

if not exist "%DASH_PYTHON%" goto corrigir_ambiente

"%DASH_PYTHON%" -c "import flask; from app import app" >nul 2>&1
if errorlevel 1 goto corrigir_ambiente

echo Iniciando o Dashboard Auditoria...
echo O navegador sera aberto em http://127.0.0.1:5000
echo Para encerrar o servidor, pressione Ctrl+C nesta janela.
echo.

"%DASH_PYTHON%" "%~dp0app.py"
set "DASH_RESULT=%ERRORLEVEL%"

if "%DASH_RESULT%"=="0" goto finalizar

echo.
echo O dashboard foi encerrado com erro %DASH_RESULT%.
echo Copie a mensagem exibida acima caso precise de suporte.
pause
goto finalizar

:corrigir_ambiente
echo O ambiente virtual precisa ser instalado ou corrigido.
call "%~dp0instalar_e_executar.bat"
set "DASH_RESULT=%ERRORLEVEL%"

:finalizar
endlocal & exit /b %DASH_RESULT%
