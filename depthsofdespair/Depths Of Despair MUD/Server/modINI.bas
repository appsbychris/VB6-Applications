Attribute VB_Name = "modINI"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modINI
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
'for the INI files
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKey As Any, ByVal lpString As String, ByVal lpFileName As String)
Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)

Public Function GetINI(sKey As String) As String
Rem Function to get the user defined settings
Dim strSpace As String, theLength As Long
strSpace = Space(255)
theLength = GetPrivateProfileString("USER", sKey, "Error", strSpace, 255, App.Path & "\userdefined.dat")
strSpace = ReplaceFast(strSpace, Chr$(0), "")
strSpace = TrimIt(strSpace)
GetINI = strSpace
End Function

Public Sub WriteINI(sKey As String, sString As String)
Rem sub to write user defined settings
WritePrivateProfileString "USER", sKey, sString, App.Path & "\userdefined.dat"
End Sub
