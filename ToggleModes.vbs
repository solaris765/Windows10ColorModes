Option Explicit
'~ On Error Resume Nex
RequireAdmin

Dim objReg, keystring, objshell, Err_Number, strValue
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")


keystring = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme"
set objShell = WScript.CreateObject("WScript.Shell")

On Error Resume Next
strValue = objShell.RegRead(keystring)
Err_Number=err.number
On Error Goto 0

If err_number <> 0 Then
    strValue = 1
End If

If strValue = "1" Then
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "AccentPalette", "REG_BINARY", Array(166,216,255,0,118,185,237,0,66,156,227,0,0,120,215,0,0,90,158,0,0,66,117,0,0,38,66,0,136,23,152,0)
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "StartColorMenu", "REG_DWORD", 4288567808
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "AccentColorMenu", "REG_DWORD", 4284706145
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorizationColorBalance", "REG_DWORD", 89
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorizationColor", "REG_DWORD", 73226840
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorizationAfterglow", "REG_DWORD", 3294452312
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorizationBlurBalance", "REG_DWORD", 1
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "EnableTransparency", "REG_DWORD", 1
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorPrevalence", "REG_DWORD", 1
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "AccentColor", "REG_DWORD", 4283980381
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "EnableWindowColorization", "REG_DWORD", 1
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "AccentColorInactive", "REG_DWORD", 4282927692
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", "REG_DWORD", 0
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "AccentPalette", "REG_BINARY", Array(197,191,185,0,163,158,154,0,135,131,128,0,93,90,88,0,62,60,59,0,43,42,41,0,31,30,29,0,255,67,67,0)
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "StartColorMenu", "REG_DWORD", 4282072126
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "AccentColorMenu", "REG_DWORD", 4283980381
Else
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorizationColor", "REG_DWORD", 3288365271
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorizationColorBalance", "REG_DWORD", 89
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorizationAfterglow", "REG_DWORD", 3288365271
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorizationBlurBalance", "REG_DWORD", 1
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "EnableTransparency", "REG_DWORD", 0
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "ColorPrevalence", "REG_DWORD", 0
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "AccentColor", "REG_DWORD", 4292311040
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "EnableWindowColorization", "REG_DWORD", 0
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\DWM", "AccentColorInactive", "REG_DWORD", 4294967295
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", "REG_DWORD", 1
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "AccentPalette", "REG_BINARY", Array(166,216,255,0,118,185,237,0,66,156,227,0,0,120,215,0,0,90,158,0,0,66,117,0,0,38,66,0,247,99,12,0)
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "StartColorMenu", "REG_DWORD", 4282072126
    RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent", "AccentColorMenu", "REG_DWORD", 4283980381
End If

objShell.Run "taskkill /f /im explorer.exe", 0, true
objShell.Run "explorer"


Function RegWrite(reg_keyname, reg_valuename,reg_type,ByVal reg_value)
	Dim aRegKey, Return
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegWrite = 0
		Exit Function
	End If

	Return = RegWriteKey(aRegKey)
	If Return = 0 Then
		RegWrite = 0
		Exit Function
	End If

	Select Case reg_type
		Case "REG_SZ"
			Return = objReg.SetStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case "REG_EXPAND_SZ"
			Return = objReg.SetExpandedStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case "REG_BINARY"
			If IsArray(reg_value) = 0 Then reg_value = Array()
			Return = objReg.SetBinaryValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		Case "REG_DWORD"
			If IsNumeric(reg_value) = 0 Then reg_value = 0
			Return = objReg.SetDWORDValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		Case "REG_MULTI_SZ"
			If IsArray(reg_value) = 0 Then
				If Len(reg_value) = 0 Then
					reg_value = Array()
				Else
					reg_value = Array(reg_value)
				End If
			End If
			Return = objReg.SetMultiStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		'Case "REG_QWORD"
			'Return = oReg.SetQWORDValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case Else
			RegWrite = 0
			Exit Function
	End Select

	If (Return <> 0) Or (Err.Number <> 0) Then
		RegWrite = 0
		Exit Function
	End If
	RegWrite = 1
End Function

Function RegWriteKey(RegKeyName)
	Dim Return
	If IsArray(RegKeyName) = 0 Then
		RegKeyName = RegSplitKey(RegKeyName)
	End If

	If (IsArray(RegKeyName) = 0) Or (UBound(RegKeyName) <> 1) Then
		RegWriteKey = 0
		Exit Function
	End If

	Return = objReg.CreateKey(RegKeyName(0),RegKeyName(1))
	If (Return <> 0) Or (Err.Number <> 0) Then
		RegWriteKey = 0
		Exit Function
	End If
	RegWriteKey = 1
End Function

Function RegDelete(reg_keyname, reg_valuename)
	Dim Return,aRegKey
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegDelete = 0
		Exit Function
	End If

	Return = objReg.DeleteValue(aRegKey(0),aRegKey(1),reg_valuename)
	If (Return <> 0) And (Err.Number <> 0) Then
		RegDelete = 0
		Exit Function
	End If
	RegDelete = 1
End Function

Function RegDeleteKey(reg_keyname)
	Dim Return,aRegKey
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegDeleteKey = 0
		Exit Function
	End If

	'On Error Resume Next
	Return = RegDeleteSubKey(aRegKey(0),aRegKey(1))
	'On Error Goto 0
	If Return = 0 Then
		RegDeleteKey = 0
		Exit Function
	End If
	RegDeleteKey = 1
End Function

Function RegDeleteSubKey(strRegHive, strKeyPath)
	Dim Return,arrSubkeys,strSubkey
    objReg.EnumKey strRegHive, strKeyPath, arrSubkeys
    If IsArray(arrSubkeys) <> 0 Then
        For Each strSubkey In arrSubkeys
            RegDeleteSubKey strRegHive, strKeyPath & "\" & strSubkey
        Next
    End If

	Return = objReg.DeleteKey(strRegHive, strKeyPath)
	If (Return <> 0) Or (Err.Number <> 0) Then
		RegDeleteSubKey = 0
		Exit Function
	End If
	RegDeleteSubKey = 1
End Function

Function RegSplitKey(RegKeyName)
	Dim strHive, strInstr, strLeft
	strInstr=InStr(RegKeyName,"\")
	If strInstr = 0 Then Exit Function
	strLeft=left(RegKeyName,strInstr-1)

	Select Case strLeft
		Case "HKCR","HKEY_CLASSES_ROOT"	strHive = &H80000000
		Case "HKCU","HKEY_CURRENT_USER"	strHive = &H80000001
		Case "HKLM","HKEY_LOCAL_MACHINE"	strHive = &H80000002
		Case "HKU","HKEY_USERS" 	strHive = &H80000003
		Case "HKCC","HKEY_CURRENT_CONFIG"	strHive = &H80000005
	  Case Else Exit Function
	End Select

    RegSplitKey = Array(strHive,Mid(RegKeyName,strInstr+1))
End Function

Function RequireAdmin()
	Dim reg_valuename, WShell, Cmd, CmdLine, I

	GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")_
	.EnumValues &H80000003, "S-1-5-19\Environment",  reg_valuename
	If IsArray(reg_valuename) <> 0 Then
		RequireAdmin = 1
		Exit Function
	End If

	Set Cmd = WScript.Arguments
	For I = 0 to Cmd.Count - 1
		If Cmd(I) = "/admin" Then
			Wscript.Echo "To script you must have administrator rights!"
			'RequireAdmin = 0
			'Exit Function
			WScript.Quit
		End If
		CmdLine = CmdLine & Chr(32) & Chr(34) & Cmd(I) & Chr(34)
	Next
	CmdLine = CmdLine & Chr(32) & Chr(34) & "/admin" & Chr(34)

	Set WShell= WScript.CreateObject( "WScript.Shell")
	CreateObject("Shell.Application").ShellExecute WShell.ExpandEnvironmentStrings(_
	"%SystemRoot%\System32\WScript.exe"),Chr(34) & WScript.ScriptFullName & Chr(34) & CmdLine, "", "runas"
	WScript.Quit
End Function