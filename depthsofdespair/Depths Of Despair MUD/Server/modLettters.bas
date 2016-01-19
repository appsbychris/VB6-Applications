Attribute VB_Name = "modLetters"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modLetters
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function LetterSubs(Index As Long) As Boolean
If WriteLetter(Index) = True Then LetterSubs = True: Exit Function
If DestroyAll(Index) = True Then LetterSubs = True: Exit Function
If DestroyLetter(Index) = True Then LetterSubs = True: Exit Function
If AppendLetter(Index) = True Then LetterSubs = True: Exit Function
End Function


Public Function WriteLetter(Index As Long) As Boolean
Dim sTitle As String
Dim sMessage As String
Dim lID As Long
Dim m As Long
Dim s As String
Dim dbIndex As Long
If LCaseFast(Left$(X(Index), 6)) = "write " Then
    WriteLetter = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If .lPaper <= 0 Then
            WrapAndSend Index, RED & "You have no paper to write a note on." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    s = X(Index)
    s = Mid$(s, InStr(1, s, " ") + 1)
    m = InStr(1, s, ",")
    If m = 0 Then
        WrapAndSend Index, RED & "You must provide a title and a message." & vbCrLf & "Syntax: write [Title],[Message]" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    sTitle = Mid$(s, 1, m - 1)
    If modLetters.DoesTitleExsist(sTitle) Then
        WrapAndSend Index, RED & "You can't seem to title your note that." & vbCrLf & "Syntax: write [Title],[Message]" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    sMessage = Mid$(s, InStr(1, s, ",") + 1)
    lID = GetNewID
    ReDim Preserve dbLetters(1 To UBound(dbLetters) + 1)
    With dbLetters(UBound(dbLetters))
        .lID = lID
        .sMessage = sMessage
        .sTitle = sTitle
    End With
    With dbPlayers(dbIndex)
        If modSC.FastStringComp(.sLetters, "") Then .sLetters = ""
        .sLetters = .sLetters & ":" & CStr(lID) & ";"
        .lPaper = .lPaper - 1
        SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " writes on a piece of paper." & WHITE & vbCrLf, .lLocation
    End With
    WrapAndSend Index, LIGHTBLUE & "You write a note." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function DestroyLetter(Index As Long) As Boolean
Dim lID As Long
Dim s As String
Dim dbIndex As Long
If LCaseFast(Left$(X(Index), 8)) = "destroy " Then
    DestroyLetter = True
    s = ReplaceFast(X(Index), "destroy ", "", 1, 1)
    s = SmartFind(Index, LCaseFast(s), Inventory_Item)
    
    lID = GetLetterID(ReplaceFast(s, "note: ", ""))
    If lID = 0 Then
        WrapAndSend Index, RED & "You don't seem to have a note titled that." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If InStr(1, .sLetters, ":" & dbLetters(lID).lID & ";") = 0 Then
            WrapAndSend Index, RED & "You don't seem to have a note titled that." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .sLetters = ReplaceFast(.sLetters, ":" & dbLetters(lID).lID & ";", "")
        If modSC.FastStringComp(.sLetters, "") Then .sLetters = "0"
        SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " rips up a piece of paper." & WHITE & vbCrLf, .lLocation
    End With
    RemoveLetter lID
    WrapAndSend Index, LIGHTBLUE & "You rip up your note." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function AppendLetter(Index As Long) As Boolean
Dim lID As Long
Dim s As String
Dim dbIndex As Long
Dim m As Long
Dim sS As String
If LCaseFast(Left$(X(Index), 7)) = "append " Then
    AppendLetter = True
    s = ReplaceFast(X(Index), "append ", "", 1, 1)
    m = InStr(1, s, ",")
    If m = 0 Then
        WrapAndSend Index, RED & "You must provide a title and a message." & vbCrLf & "Syntax: append [Title],[Message]" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    s = Mid$(s, 1, m - 1)
    sS = ReplaceFast(LCaseFast(X(Index)), "append " & s & ",", "", 1, 1)
    s = SmartFind(Index, LCaseFast(s), Inventory_Item)
    lID = GetLetterID(ReplaceFast(s, "note: ", ""))
    If lID = 0 Then
        WrapAndSend Index, RED & "You don't seem to have a note titled that." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If InStr(1, .sLetters, ":" & dbLetters(lID).lID & ";") = 0 Then
            WrapAndSend Index, RED & "You don't seem to have a note titled that." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        With dbLetters(lID)
            .sMessage = .sMessage & vbCrLf & "Added by " & BRIGHTRED & dbPlayers(dbIndex).sPlayerName & GREEN & ":" & vbCrLf & "  " & sS
        End With
        SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " writes on a piece of paper." & WHITE & vbCrLf, .lLocation
    End With
    WrapAndSend Index, LIGHTBLUE & "You append stuff to the note." & WHITE & vbCrLf

    X(Index) = ""
End If
End Function

Public Function DestroyAll(Index As Long) As Boolean
Dim s As String
Dim dbIndex As Long
Dim tArr() As String
If LCaseFast(X(Index)) = "destroy all notes" Then
    dbIndex = GetPlayerIndexNumber(Index)
    DestroyAll = True
    With dbPlayers(dbIndex)
        s = .sLetters
        s = ReplaceFast(s, ":", "")
        If s <> "0" Then
            SplitFast s, tArr, ";"
            For i = LBound(tArr) To UBound(tArr)
                If tArr(i) <> "" Then RemoveLetter GetLetterID(, CLng(tArr(i)))
                If DE Then DoEvents
            Next
            .sLetters = "0"
            WrapAndSend Index, LIGHTBLUE & "You rip up all your letters." & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " rips up all their letters." & WHITE & vbCrLf, .lLocation
            X(Index) = ""
        Else
            WrapAndSend Index, RED & "You have no letters." & WHITE & vbCrLf
            X(Index) = ""
        End If
    End With
End If
End Function

Public Function GetNewID() As Long
Dim i As Long
Dim m As Long
For i = LBound(dbLetters) To UBound(dbLetters)
    With dbLetters(i)
        If .lID > m Then m = .lID
    End With
    If DE Then DoEvents
Next
m = m + 1
GetNewID = m
End Function

Sub RemoveLetter(lID As Long)
On Error GoTo FUCKIMESSEDUP
Dim lngIndex As Long
lngIndex = 1
For lngIndex = lID To UBound(dbLetters) - 1
    dbLetters(lngIndex) = dbLetters(lngIndex + 1)
    If DE Then DoEvents
Next lngIndex
ReDim Preserve dbLetters(1 To UBound(dbLetters) - 1)
FUCKIMESSEDUP:
End Sub

Public Function DoesTitleExsist(ByVal s As String) As Boolean
Dim i As Long
s = LCaseFast(s)
For i = LBound(dbLetters) To UBound(dbLetters)
    With dbLetters(i)
        If LCaseFast(dbLetters(i).sTitle) = s Then
            DoesTitleExsist = True
            Exit Function
            Exit For
        End If
    End With
Next
DoesTitleExsist = False
End Function
