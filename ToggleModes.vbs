keystring = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme"
set objShell = WScript.CreateObject("WScript.Shell")
strValue = objShell.RegRead(keystring)

If strValue = "1" Then
    WshShell.run "cmd /c regedit %cd%\DarkMode.reg",0,True
Else
    WshShell.run "cmd /c regedit %cd%\LightMode.reg",0,True
End If

Set objScriptShell = CreateObject("Wscript.Shell")
objScriptShell.Run "taskkill /f /im explorer.exe", 0, true
objScriptShell.Run "explorer"