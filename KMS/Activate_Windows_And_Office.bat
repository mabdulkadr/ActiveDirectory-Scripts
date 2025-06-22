@echo off
title Activate Windows & Office 2019 - Qassim University
cls

echo ================================================
echo Starting Activation for Windows and Office 2019
echo ================================================

:: Step 1: Activate Office 2019
echo Activating Office 2019...

cd /d "C:\Program Files\Microsoft Office\Office16"

:: (Optional) Set Office 2019 KMS client key (uncomment if needed)
:: cscript //nologo ospp.vbs /inpkey:6MWKP-HQYW6-DF2VR-TTVJX-2BFGV

cscript //nologo ospp.vbs /sethst:kms.qu.edu.sa
cscript //nologo ospp.vbs /act
echo Office activation status:
cscript //nologo ospp.vbs /dstatus

:: Step 2: Activate Windows
echo.
echo Activating Windows...

:: (Optional) Set Windows KMS client key (uncomment if needed)
:: slmgr /ipk WFG99-8FYR6-6HK9W-WXFFM-YYW8H

slmgr /skms kms.qu.edu.sa
slmgr /ato

:: Display Windows activation info
timeout /t 2 >nul
echo Windows activation status:
slmgr /dlv

echo.
echo Activation process completed.
echo Press any key to close this window...
pause >nul
