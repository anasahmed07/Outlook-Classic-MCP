@echo off
REM ---------------------------------------------------------------------
REM outloook-classic-mcp Windows installer (uv-based)
REM
REM 1. Installs uv (Astral) if missing and uses it to fetch Python.
REM 2. Creates .venv with Python 3.11.
REM 3. Installs the package in editable mode (-e .).
REM 4. Pre-warms the pywin32 typelib for Outlook.
REM 5. Launches scripts\install_to_clients.py — interactive menu that
REM    detects your installed MCP clients (Claude Desktop, Claude Code,
REM    Cursor, Cline, Continue, Windsurf) and writes the configs you tick.
REM ---------------------------------------------------------------------

setlocal enabledelayedexpansion

set "INSTALL_DIR=%~dp0"
if "%INSTALL_DIR:~-1%"=="\" set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"

set "VENV_DIR=%INSTALL_DIR%\.venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"

echo.
echo ========================================================================
echo   outloook-classic-mcp installer
echo   Install location: %INSTALL_DIR%
echo ========================================================================
echo.

REM ---- Check for uv, install if missing ----
where uv >nul 2>nul
if %errorlevel% neq 0 (
    echo [1/5] uv not found. Installing uv via the official PowerShell installer ...
    powershell -ExecutionPolicy ByPass -NoProfile -Command "irm https://astral.sh/uv/install.ps1 | iex"
    if !errorlevel! neq 0 (
        echo.
        echo [error] uv install failed. You can install it manually with:
        echo         winget install --id=astral-sh.uv -e
        echo         or visit https://docs.astral.sh/uv/getting-started/installation/
        echo.
        pause
        exit /b 1
    )
    set "PATH=%USERPROFILE%\.local\bin;%PATH%"
    where uv >nul 2>nul
    if !errorlevel! neq 0 (
        echo.
        echo [error] uv was installed but isn't on PATH yet. Open a NEW
        echo         terminal and re-run install.bat.
        echo.
        pause
        exit /b 1
    )
) else (
    echo [1/5] uv is already installed.
)

REM ---- Create venv with Python 3.11 ----
echo [2/5] Creating virtual environment (Python 3.11) ...
cd /d "%INSTALL_DIR%"
uv venv --python 3.11 "%VENV_DIR%"
if %errorlevel% neq 0 (
    echo.
    echo [error] uv venv creation failed.
    pause
    exit /b 1
)

REM ---- Install the package in editable mode ----
echo [3/5] Installing outloook-classic-mcp (editable) and dependencies ...
uv pip install --python "%PYTHON_EXE%" -e "%INSTALL_DIR%"
if %errorlevel% neq 0 (
    echo.
    echo [error] Package install failed.
    pause
    exit /b 1
)

REM ---- Pre-warm pywin32 typelib ----
echo [4/5] Pre-warming pywin32 typelib cache for Outlook ...
> "%TEMP%\_outlook_mcp_warmup.py" echo import win32com.client; win32com.client.gencache.EnsureDispatch('Outlook.Application')
"%PYTHON_EXE%" "%TEMP%\_outlook_mcp_warmup.py" 2>nul
del /q "%TEMP%\_outlook_mcp_warmup.py" 2>nul
if %errorlevel% neq 0 (
    echo [warn] Could not pre-warm typelib. Fine if Outlook isn't running yet
    echo        - the bridge will warm it on first call.
)

REM ---- Smart client installer ----
echo [5/5] Launching client installer ...
echo.
"%PYTHON_EXE%" "%INSTALL_DIR%\scripts\install_to_clients.py"

echo.
echo ========================================================================
echo   INSTALL COMPLETE
echo ========================================================================
echo.
echo   Standalone smoke test:
echo     "%PYTHON_EXE%" -m outlook_mcp
echo.
echo   MCP Inspector (forward-slash paths to avoid Inspector backslash bug):
set "FWD_PY=%PYTHON_EXE:\=/%"
echo     npx @modelcontextprotocol/inspector "%FWD_PY%" -m outlook_mcp
echo.
echo   Re-run client install at any time:
echo     "%PYTHON_EXE%" "%INSTALL_DIR%\scripts\install_to_clients.py"
echo.
echo ========================================================================
echo.
pause

endlocal
