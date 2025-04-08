::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
:: PaperShredder setup script.
::
:: (c) 2025 Michael Brodsky
:: mbrodskiis@gmail.com
:: All rights reserved. Unauthorized use prohibited.
::
:: Downloads and installs any required dependencies, and
:: Extracts application files to the root folder. 
::
:: Parameters:
::    %1 - Application root folder (optional).
::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

@echo off
SETLOCAL EnableExtensions EnableDelayedExpansion
SET PATH=%PATH%;C:\Windows\System32\
SET Program=PaperShredder Setup
SET Version=1.0
SET SetupSubDir=setup
SET InstallFile=install.bat
SET AppCompressedFiles=PaperShredderAIO.zip

:: Application dependencies.
SET dependencies[0].Name="GPL Ghostscript"
SET dependencies[0].RegistryKey="HKEY_LOCAL_MACHINE\SOFTWARE\GPL Ghostscript\9.56.1"
SET dependencies[0].InstallerUrl="https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs9561"
SET dependencies[0].FileName="gs9561w64.exe"
SET dependencies[0].Args="/S"
SET dependencies[1].Name="ImageMagick"
SET dependencies[1].RegistryKey="HKEY_LOCAL_MACHINE\SOFTWARE\ImageMagick\Current"
SET dependencies[1].InstallerUrl="https://imagemagick.org/archive/binaries"
SET dependencies[1].FileName="ImageMagick-7.1.1-47-Q16-HDRI-x64-dll.exe"
SET dependencies[1].Args="/VERYSILENT /NORESTART /FORCECLOSEAPPLICATIONS /LOADINF="imsetup.txt""
SET dependencies[2].Name="OCR Tesseract"
SET dependencies[2].RegistryKey="HKEY_LOCAL_MACHINE\SOFTWARE\Tesseract-OCR"
SET dependencies[2].InstallerUrl="https://digi.bib.uni-mannheim.de/tesseract"
SET dependencies[2].FileName="tesseract-ocr-w64-setup-v5.3.0.20221214.exe"
SET dependencies[2].Args="/S"

SET /A sz_dependencies=2

:MAIN
::
:: Download and install any required dependencies.
::
:: The install root is %1 or this script's 
:: parent folder, if omitted. All other locations 
:: are subfolders thereof.
::
IF "%~1"=="" (SET "RootFolder=%~dp0") ELSE (SET "RootFolder=%~1")
SET "SetupFolder=%RootFolder%\%SetupSubDir%"
SET "Installer=%SetupFolder%\%InstallFile%"
ECHO %Program% %Version% 
VER
ECHO Setup will download and install any required dependencies.
PAUSE

:: Loop through each defined dependency and, if neccessary, 
:: install it.
FOR /L %%i IN (0 1 %sz_dependencies%) DO  ( 
    CALL ECHO INSTALLING !dependencies[%%i].Name:"=!
    CALL "%Installer%" %%dependencies[%%i].RegistryKey%% %%dependencies[%%i].InstallerUrl%% "%SetupFolder%" %%dependencies[%%i].FileName%% %%dependencies[%%i].Args%%
    IF NOT ERRORLEVEL 1 (
	CALL ECHO %%dependencies[%%i].Name%% installed successfully
    ) ELSE (
	CALL ECHO %%dependencies[%%i].Name%% install failed
	GOTO MAIN_DONE
    ) 
)
::
:: Extract the application files from the compressed folder.
::
TAR -xf %AppCompressedFiles% -C "%RootFolder%"
IF ERRORLEVEL 1 (GOTO MAIN_DONE)
::
:: Success
::
ECHO Setup completed successfully.
:MAIN_DONE
PAUSE
EXIT /B %ERRORLEVEL%
