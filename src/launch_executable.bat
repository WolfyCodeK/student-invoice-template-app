@echo off

REM Check if the first argument (executable path) is provided
if "%~1" == "" (
    echo Please provide the path to the Python script and its arguments.
    exit /b
)

REM Launch the Python script with arguments excluding the first one (executable path)
python "%~1" %*