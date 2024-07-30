
@echo off

:input_loop
set /p filename="Enter the excel file name : "
if "%filename%"=="q" goto :exit_script

set /p lut="Enter the lut file name : "
if "%lut%"=="q" goto :exit_script

python "F:\RECORDSTOPDF\project\invoice_generation\app.py" -f "%filename%" -l "%lut%"

:pause_option
echo Script execution completed.
set /p pause_option="Enter 'q' to quit or press any other key to continue: "
if "%pause_option%"=="q" goto :exit_script

goto :input_loop

:exit_script
echo Press any key to exit.
pause > nul










