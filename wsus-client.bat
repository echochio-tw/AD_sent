if exist "c:\wsus\chkwsus.bat"  goto end
del /f /q /a "c:\wsus\chkwsus.bat" >nul 2>nul
del /f /q /a "c:\wsus\wsus-client.reg" >nul 2>nul
xcopy \\192.168.1.1\wsus\chkwsus.bat "c:\wsus\"
xcopy \\192.168.1.1\wsus\wsus-client.reg "c:\wsus\"
regedit.exe /s c:\wsus\wsus-client.reg
del /f /q /a "c:\wsus\wsus-client.reg" >nul 2>nul
:end
c:\wsus\chkwsus.bat
exit /b
