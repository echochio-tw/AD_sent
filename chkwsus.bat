@echo off
net stop wuauserv
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v AccountDomainSid /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v PingID /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientId /f
del %SystemRoot%\SoftwareDistribution\*.* /S /Q
net start wuauserv
wuauclt /resetauthorization /detectnow
wuauclt.exe /downloadnow
wuauclt.exe /reportnow
