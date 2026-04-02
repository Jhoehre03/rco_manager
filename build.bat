@echo off
echo ============================================
echo  RCO Manager - Build com PyInstaller
echo ============================================
echo.

:: Usa o Python do venv se existir, senão usa o do PATH
if exist "venv\Scripts\python.exe" (
    set PYTHON=venv\Scripts\python.exe
    set PIP=venv\Scripts\pip.exe
) else (
    set PYTHON=python
    set PIP=pip
)

:: Instala PyInstaller se ainda nao estiver instalado
%PYTHON% -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Instalando PyInstaller...
    %PIP% install pyinstaller
    echo.
)

:: Roda o build
echo Iniciando build...
echo.
%PYTHON% -m PyInstaller rco_manager.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo [ERRO] Build falhou. Verifique as mensagens acima.
    pause
    exit /b 1
)

:: Verifica se o executavel foi gerado
if not exist "dist\RCO Manager\RCO Manager.exe" (
    echo.
    echo [ERRO] Executavel nao encontrado em dist\RCO Manager\RCO Manager.exe
    echo        O build pode ter falhado silenciosamente.
    pause
    exit /b 1
)

:: Copia arquivos necessarios para a pasta de distribuicao
echo.
echo Copiando arquivos necessarios...

if exist "oauth_credentials.json" (
    copy /Y "oauth_credentials.json" "dist\RCO Manager\oauth_credentials.json" >nul
    echo   [OK] oauth_credentials.json copiado
) else (
    echo   [AVISO] oauth_credentials.json nao encontrado na pasta do projeto.
    echo           O app nao conseguira autenticar com o Google sem este arquivo.
)

:: Lê a versão do version.py
for /f "tokens=3 delims== " %%v in ('findstr "VERSION" ui\version.py') do set VERSION=%%~v

:: Gera o zip sem pasta raiz (conteúdo direto na raiz do zip)
echo.
echo Gerando zip de release...
set ZIP_NAME=%~dp0rco_manager_v%VERSION%.zip
if exist "%ZIP_NAME%" del /Q "%ZIP_NAME%"
set DIST_DIR=%~dp0dist\RCO Manager
%PYTHON% -c "import zipfile,os; src=r'%DIST_DIR%'; out=r'%ZIP_NAME%'; z=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED); [z.write(os.path.join(r,f),os.path.relpath(os.path.join(r,f),src).replace(os.sep,'/')) for r,_,fs in os.walk(src) for f in fs]; z.close(); print('  [OK]',os.path.basename(out),'gerado')"

echo.
echo ============================================
echo  Build concluido! Pasta: dist\RCO Manager\
echo  Release zip: %ZIP_NAME%
echo  Verifique se oauth_credentials.json esta
echo  na pasta dist\RCO Manager\
echo ============================================
pause
