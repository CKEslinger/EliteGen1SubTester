@ECHO OFF
REM Exit Codes
REM ----------------------------------------------------------------------------
REM 
REM      0 - All of the selected operations complete successfully
REM      1 - Invalid Command Line Argument
REM      2 - RealICE/ICD3 Communication Failed
REM      3 - Selected Operation Failed
REM      4 - Unknown Runtime Failure
REM      5 - Invalid Device Detected
REM      6 - SQTP Failed
IF ".%1"=="." %0 v
IF ".%2"=="." %0 %1 24FJ256GA106
IF ".%3"=="." %0 %1 %2 C:\PassTimeData\elite02.01.hex

TITLE  %0 %1 %2 %3
PROMPT $n$g

IF EXIST %3 GOTO pPicRun
ECHO. > progPic.txt
ECHO Could not find file  >> progPic.txt
TYPE progPic.txt
CMD
EXIT 1

:pPicRun
IF "%1"=="v" GOTO pPicVerify
IF "%1"=="p" GOTO pPicProgram
GOTO pPicBadParm

:pPicVerify
ECHO Verifying...
ECHO C:^>ICD3CMD.exe -P%2 -F%3 -Y > progPic.txt
ICD3CMD.exe -P%2 -F%3 -Y >> progPic.txt
IF ERRORLEVEL 1 GOTO pPicRunErr
GOTO pPicEnd

:pPicProgram
ECHO Firmware downloading...
ECHO C:^>ICD3CMD.exe -P%2 -F%3 -M > progPic.txt
ICD3CMD.exe -P%2 -F%3 -M >> progPic.txt
IF ERRORLEVEL 1 GOTO pPicRunErr
GOTO pPicEnd

:pPicRunErr
ECHO. >> progPic.txt
ECHO. >> progPic.txt
ECHO ICD3CMD Returned Error Level %ERRORLEVEL% >> progPic.txt
ECHO. >> progPic.txt
ECHO EXIT 1 >> progPic.txt
ECHO. >> progPic.txt
REM The "TYPE" dos command will reset ERRORLEVEL to zero
TYPE progPic.txt
REM CMD
REM PAUSE
EXIT 1

:pPicBadParm
ECHO. > progPic.txt
ECHO Invalid Paramter for batch file %0 >> progPic.txt
ECHO. > progPic.txt
TYPE progPic.txt
REM CMD
EXIT 1

:pPicEnd
ECHO. >> progPic.txt
ECHO. >> progPic.txt
ECHO C:^>ECHO ERRORLEVEL: %%ERRORLEVEL%% >> progPic.txt 
ECHO ERRORLEVEL: %ERRORLEVEL% >> progPic.txt
REM The "TYPE" dos command will reset ERRORLEVEL to zero
TYPE progPic.txt
REM TITLE Batch File %0 Shell
REM CMD
TITLE Batch File %0 Terminated
ECHO.
REM CMD
REM PAUSE
