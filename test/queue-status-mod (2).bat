@echo off
title %0
echo 1.Current Directory: %cd%
echo.
echo 2.Changing CD to script actual path: %~dp0
echo.

pushd %~dp0
echo 3.Current Directory (new): %cd%
echo.

:: Check if QUEUE is set to YES
if /i "%QUEUE%"=="YES" (
    echo 4.QUEUE is set to YES.
) else (
    echo 4.QUEUE is not set or set to NO
)

pause > nul

:eof
endlocal
exit /b 0