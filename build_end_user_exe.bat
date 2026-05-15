@echo off
setlocal

pushd "%~dp0"

set "PYTHON_EXE="

if exist ".venv\Scripts\python.exe" set "PYTHON_EXE=.venv\Scripts\python.exe"
if not defined PYTHON_EXE if exist "venv\Scripts\python.exe" set "PYTHON_EXE=venv\Scripts\python.exe"
if not defined PYTHON_EXE set "PYTHON_EXE=py -3"

echo ==========================================
echo BUILDING END-USER EXE
echo ==========================================

echo Using Python: %PYTHON_EXE%
echo.

%PYTHON_EXE% -m pip install --upgrade pip
if errorlevel 1 goto :error

%PYTHON_EXE% -m pip install pyinstaller
if errorlevel 1 goto :error

echo [1/3] Cleaning old build folders...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo [2/3] Building application...
%PYTHON_EXE% -m PyInstaller ReductionAppGUI.spec
if errorlevel 1 goto :error

echo [3/3] Build finished successfully.
echo Output folder:
echo dist\ReductionAppGUI
echo.
pause
popd
exit /b 0

:error
echo.
echo BUILD FAILED.
echo Check the messages above.
echo.
pause
popd
exit /b 1
