@echo off

REM To set the environemnt variable so that the wrapper can behave different like exit wihout halting the screen etc
set QUEUE=YES

python batch_queue_runner_gui.py

echo.
echo Close the window manually.
:hi
pause > nul
goto :hi
