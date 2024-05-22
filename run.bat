@echo off

:: List of required packages
set packages=scipy tk openpyxl pandas plotly kaleido

:: Check and install missing packages
for %%p in (%packages%) do (
    pip show %%p >nul 2>nul
    if errorlevel 1 (
        echo Installing %%p...
        pip install %%p
    ) else (
        echo %%p is already installed.
    )
)

:: Python script execution
python -m desktop_ui
cmd /k