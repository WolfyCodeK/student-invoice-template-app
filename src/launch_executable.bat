@echo off

REM Check if the first argument (executable path) is provided
if "%~1" == "" (
    echo Please provide the path to the executable.
    exit /b
)

set EXE_PATH=%~1
set PARAMS=%*

REM Remove the first argument (executable path) from PARAMS
shift

REM Launch the executable with parameters
start "" "%EXE_PATH%" %PARAMS%