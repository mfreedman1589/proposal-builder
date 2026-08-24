@echo off
setlocal
rem ---------------------------------------------------------------------------
rem  Render slide images for any case study that doesn't have them yet.
rem
rem  Double-click this, or pin it to the taskbar (see CLAUDE.md). It renders
rem  through PowerPoint, uploads the images to Supabase, prints a summary and
rem  waits for a keypress so the window doesn't vanish before you've read it.
rem
rem  Runs from its own folder rather than whatever directory the shortcut
rem  happened to start in -- %~dp0 is where this file lives, which is the
rem  project root. Without it, a desktop shortcut lands in C:\Windows\System32
rem  and every relative path in the app is wrong.
rem ---------------------------------------------------------------------------
cd /d "%~dp0"

rem A project venv if there is one, otherwise whatever `python` is on PATH --
rem this project was installed globally via winget, so both are normal.
set "PY=python"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if exist "venv\Scripts\python.exe"  set "PY=venv\Scripts\python.exe"

echo ============================================================
echo  Proposal Builder -- rendering pending case studies
echo  Folder: %CD%
echo  Python: %PY%
echo ============================================================
echo.

"%PY%" render_case_study_images.py --pending
set "CODE=%ERRORLEVEL%"

echo.
if "%CODE%"=="0" (
    echo Finished cleanly.
) else if "%CODE%"=="1" (
    echo At least one case study FAILED -- see the log path above.
) else (
    echo Could not run. Check that PowerPoint is installed and Supabase is reachable.
)

echo.
pause
exit /b %CODE%
