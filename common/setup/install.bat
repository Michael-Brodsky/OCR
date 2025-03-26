::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
:: Windows application dependency checker and installer.
::
:: (c) 2025 Michael Brodsky
:: mbrodskiis@gmail.com
:: All rights reserved. Unauthorized use prohibited.
::
:: Checks the Windows registry for an application install 
:: key and, if necessary, downloads and runs the 
:: application installer. 
::
:: Parameters:
::    $1 - Registry key
::    $2 - Installer URL (without file name)
::    $3 - Local download folder
::    $4 - Installer file name
::    $5 - Installer command line args.
::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

@echo off
SETLOCAL
SET PATH=%PATH%;C:\Windows\System32\
SET recheck=
SET reinstall=

:MAIN  
CALL :CHECK_DEPENDENCY %1 %2 %3 %4 %5 
EXIT /B %ERRORLEVEL%

:CHECK_DEPENDENCY
::
:: Checks the registry for the application key and, if not
:: found, installs the application.
::
REG QUERY %1 2>NUL
IF %ERRORLEVEL% NEQ 0 (
    IF NOT DEFINED recheck (
        SET recheck=1
        CALL :INSTALL_DEPENDENCY %2 %3 %4 %5
        GOTO CHECK_DEPENDENCY
    ) 
) 
EXIT /B %ERRORLEVEL%

:INSTALL_DEPENDENCY
::
:: Checks for and, if necessary, downloads the installer 
:: and runs it with the given command line args.
::
SET folder=%~2
SET file=%~3
SET "installer=%folder%\%file%"
IF NOT EXIST "%installer%" (
    IF NOT DEFINED reinstall (
        SET reinstall=1
        CALL :DOWNLOAD_INSTALLER %1 %2 %3
        GOTO INSTALL_DEPENDENCY
    ) ELSE (
	EXIT /B 1
    ) 
)
CALL :RUN_INSTALLER %2 %3 %4
EXIT /B %ERRORLEVEL%

:DOWNLOAD_INSTALLER
::
:: Downloads an installer from the given URL to the 
:: specified download folder.
::
SET url=%~1
SET folder=%~2
SET file=%~3
SET "download_from=%url%/%file%"
SET "save_to=%folder%\%file%"
CALL :DOWNLOAD_FILE "%download_from%" "%save_to%"
EXIT /B %ERRORLEVEL%

:RUN_INSTALLER
::
:: Runs the installer and returns its exit code.
::
SET "folder=%1"
SET "file=%2"
SET "args=%~3"
ECHO RUNNING %folder%\%file%
CALL START "" /D %folder% /WAIT %file% %args%
EXIT /B %ERRORLEVEL%

:DOWNLOAD_FILE
::
:: This is the main function for downloading files.
::
ECHO DOWNLOADING %1
CALL bitsadmin /transfer mydownloadjob /download /priority FOREGROUND %1 %2
EXIT /B %ERRORLEVEL%

:DOWNLOAD_PROXY_ON
::
:: This function is called before the main function
:: to enable download via proxy server.
::
ECHO PROXY_ON %1
CALL bitsadmin /setproxysettings mydownloadjob OVERRIDE %1
EXIT /B %ERRORLEVEL%

:DOWNLOAD_PROXY_OFF
::
:: This function disables the proxy server.
::
ECHO PROXY_OFF
CALL bitsadmin /setproxysettings mydownloadjob NO_PROXY
EXIT /B %ERRORLEVEL%