Attribute VB_Name = "modSpell"
Public Function GetItemDurFromUnFormattedString(ByVal s As String) As Long
Dim m As Long
Dim n As Long
   On Error GoTo GetItemDurFromUnFormattedString_Error

m = InStr(1, s, ":")
n = InStr(m, s, "/")
m = InStr(n + 1, s, "/")
GetItemDurFromUnFormattedString = CLng(Mid$(s, n + 1, m - n - 1))

   On Error GoTo 0
   Exit Function

GetItemDurFromUnFormattedString_Error:
    GetItemDurFromUnFormattedString = -1
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetItemDurFromUnFormattedString of Module modItemManip"
End Function

Public Function GetItemEnchantsFromUnFormattedString(ByVal s As String) As String
Dim m As Long
Dim n As Long
   On Error GoTo GetItemEnchantsFromUnFormattedString_Error

m = InStr(1, s, "E{") + 1
n = InStr(m, s, "}")
GetItemEnchantsFromUnFormattedString = Mid$(s, m + 1, n - m - 1)

   On Error GoTo 0
   Exit Function

GetItemEnchantsFromUnFormattedString_Error:
End Function

Public Function GetItemFlagsFromUnFormattedString(ByVal s As String) As String
Dim m As Long
Dim n As Long
   On Error GoTo GetItemFlagsFromUnFormattedString_Error

m = InStr(1, s, "F{") + 1
n = InStr(m, s, "}")
GetItemFlagsFromUnFormattedString = Mid$(s, m + 1, n - m - 1)

   On Error GoTo 0
   Exit Function

GetItemFlagsFromUnFormattedString_Error:
End Function

Public Function GetItemUsesFromUnFormattedString(ByVal s As String) As Long
Dim m As Long
Dim n As Long
   On Error GoTo GetItemUsesFromUnFormattedString_Error

m = InStrRev(s, "/")
GetItemUsesFromUnFormattedString = CLng(Mid$(s, m + 1))

   On Error GoTo 0
   Exit Function

GetItemUsesFromUnFormattedString_Error:
    GetItemUsesFromUnFormattedString = -1
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetItemUsesFromUnFormattedString of Module modItemManip"
End Function

Public Function GetItemIDFromUnFormattedString(ByVal s As String) As Long
Dim m As Long
Dim n As Long

   On Error GoTo GetItemIDFromUnFormattedString_Error

m = InStr(1, s, ":")
n = InStr(m, s, "/")
GetItemIDFromUnFormattedString = CLng(Mid$(s, m + 1, n - m - 1))

   On Error GoTo 0
   Exit Function

GetItemIDFromUnFormattedString_Error:

    GetItemIDFromUnFormattedString = -1
End Function
