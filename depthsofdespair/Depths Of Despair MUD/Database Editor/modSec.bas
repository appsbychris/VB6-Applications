Attribute VB_Name = "modSec"
Function uJunkIt(sArg As String) As String
Dim i As Long
Dim j As Long
Dim sBuild As String
For i = 1 To Len(sArg)
    j = j + 1
    sBuild = sBuild & Mid$(sArg, j, 1)
    j = j + i
Next
uJunkIt = sBuild
End Function
Public Sub dB_set_and_load(ByVal sSearch As String, ByVal fn As String, ByVal sReplace As String)
Dim ff As Integer
Dim s As String
Dim ptr As Long
Dim n As Long
Dim iLen As Long
Dim iFound As Long
Dim bDone As Boolean
Const BLOCKSIZE As Long = 500   'number of bytes to read at one time
iLen = Len(sSearch)  'length of string to find
If iLen <> Len(sReplace) Then Exit Sub
If iLen Then
    ff = FreeFile
    ptr = 0  'starting location
    Open fn For Binary As #ff
        n = LOF(ff)   'number of bytes in file
        If n Then
            Do
                If (n - ptr) > BLOCKSIZE Then
                    s = Space$(BLOCKSIZE)
                Else
                    s = Space$(n - ptr)
                    bDone = True    'will exit because end of file
                End If
                Get #ff, ptr + 1, s
                iFound = InStr(s, sSearch)
                If iFound Then
                    'FindInFile = ptr + iFound
                    Put #ff, iFound, sReplace
                    bDone = True   'will exit because a match has been found
                Else
                    ptr = ptr + BLOCKSIZE - iLen
                End If
            Loop While Not bDone
        End If
    Close #ff
End If
End Sub
