Attribute VB_Name = "modSC"
Public Function FastStringComp(String1 As String, String2 As String) As Boolean
If LenB(String1) = LenB(String2) Then
    If LenB(String1) = 0 And LenB(String2) = 0 Then
        FastStringComp = True
    Else
        FastStringComp = (InStrB(1, String1, String2, vbBinaryCompare) <> 0)
    End If
End If
End Function
