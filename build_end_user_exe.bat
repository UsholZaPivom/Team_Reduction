@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ==========================================
echo BUILDING END-USER EXE
echo ==========================================
echo Project folder: %CD%
echo Using Python: .venv\Scripts\python.exe
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: Python virtual environment was not found: .venv\Scripts\python.exe
    echo.
    pause
    exit /b 1
)

if not exist "ui_backend_entry_enduser.py" (
    echo ERROR: ui_backend_entry_enduser.py was not found in project folder.
    echo.
    pause
    exit /b 1
)

if not exist "end_user_app.py" (
    echo ERROR: end_user_app.py was not found in project folder.
    echo.
    pause
    exit /b 1
)

echo [CHECK] Checking GUI source file...
".venv\Scripts\python.exe" -c "from pathlib import Path; p=Path('ui_backend_entry_enduser.py'); t=p.read_text(encoding='utf-8'); assert 'self.analysis_button = ttk.Button' in t, 'new analysis button marker is missing'; assert 'self.process_button = ttk.Button' in t, 'new process button marker is missing'; assert 'self._copy_run_log_next_to_output' in t, 'run log copy marker is missing'; assert 'APP_LOG_PATH' in t, 'app log path marker is missing'; assert 'threading.Thread(target=worker, daemon=True).start()' in t, 'background worker marker is missing'; assert 'command=self._refresh_from_stage3).grid(row=0, column=1)' not in t, 'old refresh button marker was found'; print('OK: ui_backend_entry_enduser.py has new UI and logging markers')"
if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    echo Check ui_backend_entry_enduser.py: it still looks like the old UI or logging patch is missing.
    echo.
    pause
    exit /b 1
)

echo [CHECK] Checking Python syntax...
".venv\Scripts\python.exe" -m py_compile ui_backend_entry_enduser.py end_user_app.py
if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    echo Python syntax check failed.
    echo.
    pause
    exit /b 1
)

echo [CLEAN] Removing build and dist...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [BUILD] Running PyInstaller...
".venv\Scripts\python.exe" -m PyInstaller ReductionAppGUI.spec --clean --noconfirm
if errorlevel 1 (
    echo.
    echo BUILD FAILED.
    echo PyInstaller failed.
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] Build finished successfully.
echo Output folder:
echo dist\ReductionAppGUI
echo.
pause
endlocal
