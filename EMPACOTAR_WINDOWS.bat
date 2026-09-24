@echo off
setlocal
cd /d "%~dp0"

if exist "Entrega_Carla" (
    echo A pasta Entrega_Carla ja existe. Guarde-a em outro local antes de gerar uma nova entrega.
    exit /b 1
)
if exist "SistemaCobranca_Carla_portatil.zip" (
    echo O arquivo SistemaCobranca_Carla_portatil.zip ja existe. Guarde-o antes de gerar outro.
    exit /b 1
)

py -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto erro

py -m PyInstaller --noconfirm --clean SistemaCobranca_Windows.spec
if errorlevel 1 goto erro
if not exist "dist\SistemaCobranca.exe" goto erro

mkdir "Entrega_Carla"
copy "dist\SistemaCobranca.exe" "Entrega_Carla\SistemaCobranca.exe" >nul
if errorlevel 1 goto erro
copy "MODO_PORTATIL.txt" "Entrega_Carla\MODO_PORTATIL.txt" >nul
if errorlevel 1 goto erro

powershell -NoProfile -Command "Compress-Archive -Path 'Entrega_Carla' -DestinationPath 'SistemaCobranca_Carla_portatil.zip'"
if errorlevel 1 goto erro

echo.
echo Pronto: SistemaCobranca_Carla_portatil.zip
echo Antes de enviar, teste em um pendrive e confirme que a pasta dados foi criada.
pause
exit /b 0

:erro
echo.
echo Nao foi possivel concluir o empacotamento. Veja a mensagem de erro acima.
pause
exit /b 1
