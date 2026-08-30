@echo off
title TicketAudit - Know your queue before you report it

:: Light green on black: the nearest cmd.exe gets to the brand's green-on-chrome
:: (#3fc16e on #0f1518).
::
:: COLOR takes TWO HEX DIGITS and nothing else - background then foreground.
:: Given anything it cannot parse it prints its own 25-line help text, leaves
:: the colour unset, and still returns 0, so the script carries on with the
:: usage block sitting on top of the banner. @echo off does not suppress that,
:: because it is the command's output rather than the command being echoed.
color 0A

:: ---------------------------------------------------------------------------
:: NOTE ON BATCH SYNTAX
:: cmd.exe parses the entire body of an "if (...)" block when it reaches the
:: "if", so a syntax error inside fires even when the condition is false.
:: An unescaped ")" inside a block - including one in a comment or in echo
:: text - closes the block early and the rest of the line is then run as a
:: command. That is what "and was unexpected at this time." came from.
:: Therefore, inside any (...) block:
::   - keep explanatory comments OUT of the block
::   - escape parentheses in echo text as ^( and ^)
::   - use "if errorlevel 1", not "if %errorlevel% neq 0", because %VAR% is
::     expanded once when the block is parsed and would be a stale value
:: ---------------------------------------------------------------------------

:: The wordmark is a FIGlet "ANSI Shadow" banner (Unicode block glyphs), kept in
:: banner.txt and printed with "type" rather than echoed line-by-line. This is
:: deliberate: echoing UTF-8 block characters straight from the .bat corrupts
:: cmd's own line parsing after "chcp 65001" - the echo prefix gets dropped and
:: the art runs as commands ("'████╗' is not recognized"). Reading a separate
:: UTF-8 data file with "type" sidesteps that entirely and keeps THIS file pure
:: ASCII so its parsing is never affected by the codepage switch.
::
:: banner.txt MUST stay UTF-8 without a BOM. The version/tagline/links live in
:: it too; the version is duplicated from show_about_dialog in gui/app_pyside.py
:: (a base64 literal under the integrity hash) - the About dialog is the
:: authority, banner.txt is a convenience, so bump both.
::
:: chcp 65001 is set here, not at the top, because changing the codepage early
:: causes a Windows cmd.exe parsing bug that garbles subsequent :: comment lines.
chcp 65001 >nul
type "%~dp0banner.txt"

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.10+ from https://python.org
    echo.
    pause
    exit /b 1
)

:: Show Python version
echo [OK] Python found:
python --version
echo.

:: Check if we're in the correct directory (look for main.py)
if not exist "%~dp0main.py" (
    echo [ERROR] main.py not found next to this batch file.
    echo Please run this batch file from the TicketAudit folder.
    echo.
    pause
    exit /b 1
)

:: Change to script directory
cd /d "%~dp0"

:: Create virtual environment if it doesn't exist
if not exist ".venv" (
    echo [INFO] Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        echo.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
    echo.
)

:: Activate virtual environment
call "%~dp0.venv\Scripts\activate.bat"

:: Check whether the core dependencies are present (PySide6 stands in for all)
python -c "import PySide6" >nul 2>&1
if errorlevel 1 goto :install_core
goto :core_ready

:install_core
echo [INFO] Installing dependencies in virtual environment...
echo This may take a few minutes on first run...
echo.
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install core dependencies.
    echo Try running: pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)
echo.
echo [OK] Core dependencies installed successfully.
echo.

:core_ready

:: Run the application
echo [INFO] Starting TicketAudit...
echo.
python main.py

:: Capture the exit code before the block: %VAR% inside a (...) block is
:: expanded when the block is parsed, so reading %errorlevel% there would
:: report a stale value.
set "DP_EXIT=%errorlevel%"
if not "%DP_EXIT%"=="0" (
    echo.
    echo [ERROR] Application exited with error code %DP_EXIT%
    pause
)
