keystring = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme"
set objShell = WScript.CreateObject("WScript.Shell")
strValue = objShell.RegRead(keystring)

If strValue = "1" Then
    objShell.run "cmd /c regedit ""%cd%\DarkMode.reg""",0,True
Else
    objShell.run "cmd /c regedit ""%cd%\LightMode.reg""",0,True
End If

objShell.Run "taskkill /f /im explorer.exe", 0, true
objShell.Run "explorer"