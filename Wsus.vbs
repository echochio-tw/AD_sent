WScript.Sleep 5*60*1000
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.currentdirectory="c:\"
WshShell.Run "\\192.168.1.1\WSUS\wsus-client.bat", 0
Set WshShell = Nothing
