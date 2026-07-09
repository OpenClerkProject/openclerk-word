@echo off
rem Double-clickable wrapper around install-openclerk.ps1.
rem Windows' default action for a .ps1 file is to open it in an editor rather than run it,
rem which makes the installer look like it "does nothing" -- this .cmd actually executes when
rem double-clicked, and the final `pause` keeps the console open so errors are visible.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-openclerk.ps1" %*
echo.
pause
