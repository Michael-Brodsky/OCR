::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
:: PaperShredder setup restart script.
::
:: (c) 2025 Michael Brodsky
:: mbrodskiis@gmail.com
:: All rights reserved. Unauthorized use prohibited.
::
:: Restarts the application after setup is complete so 
:: so that navigation options take effect. 
::
:: Parameters:
::    %1 - Application process id
::    %2 - Application close timeout
::    %3 - Full path to application file w/o extension
::    %4 - Application file extension.
::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

@echo off
SETLOCAL EnableDelayedExpansion
SET pid=%1
SET wait=%2
SET filepath=%~f3.%4
SET /a counter=0

:CHECK_TASKLIST
::
:: Periodically check the Windows tasklist for the given 
:: pid to terminate, then restart the application.
::
PING 127.0.0.1 -n 1 -w 100 > NUL
SET /a counter+=1
IF /I "%counter%" EQU "%wait%" (GOTO TIMEOUT)
TASKLIST /fi "ImageName eq msaccess.exe" /fo csv 2>NUL | FIND /I "%pid%">NUL
IF "%ERRORLEVEL%"=="0" (GOTO CHECK_TASKLIST)
ECHO STARTING %filepath%
START "" "%filepath%"
EXIT /B %ERRORLEVEL%

:TIMEOUT
ECHO ERROR msaccess.exe pid=%pid% QUIT TIMED OUT
EXIT /B 1
