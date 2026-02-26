@echo off
setlocal

set "SCRIPT=%~dp0time-warp-indicator.ahk"
set "AHK="

for %%P in (
  "%ProgramFiles%\AutoHotkey\v2\AutoHotkey64.exe"
  "%ProgramFiles%\AutoHotkey\v2\AutoHotkey32.exe"
  "%ProgramFiles%\AutoHotkey\v2\AutoHotkey.exe"
) do (
  if exist "%%~P" (
    set "AHK=%%~P"
    goto :run
  )
)

echo AutoHotkey v2 executable not found.
echo Install AutoHotkey v2, then run this file again.
exit /b 1

:run
"%AHK%" "%SCRIPT%"
