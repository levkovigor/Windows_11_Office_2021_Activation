@echo off
:: Batch file to upgrade and activate Windows 11 Pro
:: Run this script as Administrator

echo Starting the process to upgrade to Windows 11 Pro...

:: Step 1: Uninstall the current product key
echo Uninstalling the current product key...
cscript //B slmgr.vbs /upk
timeout /t 5 >nul

:: Step 2: Clear the product key from the registry
echo Clearing the product key from the registry...
cscript //B slmgr.vbs /cpky
timeout /t 5 >nul

:: Step 3: Clear the KMS server address
echo Clearing the KMS server address...
cscript //B slmgr.vbs /ckms
timeout /t 5 >nul

:: Step 4: Check if the edition is upgradable to Pro
echo Checking if the edition is upgradable to Pro...
DISM /online /Get-TargetEditions
timeout /t 10 >nul

:: Step 5: Run the Windows Pro installer
echo Running the Windows Pro installer...
sc config LicenseManager start= auto & net start LicenseManager
sc config wuauserv start= auto & net start wuauserv
changepk.exe /productkey VK7JG-NPHTM-C97JM-9MPGT-3V66T
timeout /t 10 >nul

echo The installer will now run. Please wait until it reaches 100%...
echo If you encounter an error, simply exit and reboot your PC.
pause

:: Step 6: Activate Windows 11 Pro
echo Activating Windows 11 Pro...
cscript //B slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
cscript //B slmgr.vbs /skms kms8.msguides.com
cscript //B slmgr.vbs /ato

echo Activation complete! Check your settings to confirm Windows 11 Pro is activated.
pause
exit
