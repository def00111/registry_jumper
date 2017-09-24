Set objArgs = Wscript.Arguments
If (objArgs.Count = 0) Then WScript.Quit

Dim arrHives, arrHivesAbbr
arrHives     = Array("HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE", "HKEY_USERS", "HKEY_CURRENT_CONFIG")
arrHivesAbbr = Array("HKCR",              "HKCU",              "HKLM",               "HKU",        "HKCC")

strInput = UnEscape(objArgs(0))

sErrMsg = ""
Do
  hiveCode = GetHiveCode
  If (hiveCode = -1) Then
    strInput = InputBox("Please enter registry key." &_
                        vbcrlf & vbcrlf & sErrMsg, WScript.ScriptName, strInput)
    If (IsEmpty(strInput)) Then WScript.Quit
    If (sErrMsg = "") Then sErrMsg = "Invalid registry key!"
  Else Exit Do
  End If
Loop While True

strInput = Trim(strInput)

strInput = Replace(strInput, "/", "\")

If (Right(strInput, 1) = "]") Then
  strInput = Left(strInput, Len(strInput) - 1)
End If

If (Right(strInput, 1) = "\") Then
  strInput = Left(strInput, Len(strInput) - 1)
End If

Set re = New RegExp
re.Global = True
re.Pattern = "\s+\\+|\\{2,}"
strInput = re.Replace(strInput, "\")
re.Pattern = "\\+\s+|\\{2,}"
strInput = re.Replace(strInput, "\")

re.Pattern = "(\\Current)\s+(Version\\?)"
re.IgnoreCase = True
strInput = re.Replace(strInput, "$1$2")

strKeyPath = ""
strHive = arrHives(hiveCode)

If (InStr(1, strInput, strHive & "\", vbTextCompare) > 0) Then                     
  arrSplit = Split(strInput, "\")
  arrSize = UBound(arrSplit)
  
  If (arrSize >= 1) Then
    Dim rootKeys
    rootKeys = Array(&H80000000, &H80000001, &H80000002, &H80000003, &H80000005)
    rootKey = rootKeys(hiveCode)
    
    sComputer = "."
	  Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
	                                  sComputer & "\root\default:StdRegProv")
	       
	  Dim arrSubKeys()                      
    For i = 1 To arrSize
      ok = (oReg.EnumKey(rootKey, strKeyPath, arrSubKeys) = 0)
      If Not ok Then Exit For
        
      bVerifiedKey = False
      For Each subKey In arrSubKeys
        If (StrComp(subKey, arrSplit(i), vbTextCompare) = 0) Then
          bVerifiedKey = True
        End If
          
        If (bVerifiedKey <> True) Then
          tempKey = arrSplit(i)
          Do
            x = Len(tempKey)
            If (x > 0) Then
              tempKey = Left(tempKey, x - 1)
              If (StrComp(subKey, tempKey, vbTextCompare) = 0) Then
                arrSplit(i) = tempKey
                bVerifiedKey = True
              End If
            Else Exit Do
            End If
          Loop Until (bVerifiedKey = True)
        End If
          
        If (bVerifiedKey = True) Then
          If (i = 1) Then
            strKeyPath = arrSplit(i)
          Else
            strKeyPath = strKeyPath & "\" & arrSplit(i)
          End If
          Exit For
        End If
      Next
      
      If (bVerifiedKey = False) Then Exit For
    Next
  End If
End If

If (strKeyPath = "") Then
  strKeyPath = strHive
Else
  strKeyPath = strHive & "\" & strKeyPath
End If

Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\Lastkey", strKeyPath

On Error Resume Next
WshShell.Run "regedit.exe -m"  
On Error Goto 0 ' canceled by user?

Function GetHiveCode()
  i = -1
  
	For j = 0 To UBound(arrHives)
	  ' replace abbreviations
	  strInput = Replace(strInput, arrHivesAbbr(j), arrHives(j), 1, -1, vbTextCompare)
	  
		x = InStr(1, strInput, arrHives(j), vbTextCompare)
		If (x > 0) Then
		  ' trim garbage before key start and find hive
			strInput = Mid(strInput, x)
			i = j
			Exit For
		End If	
	Next
	GetHiveCode = i
End Function
