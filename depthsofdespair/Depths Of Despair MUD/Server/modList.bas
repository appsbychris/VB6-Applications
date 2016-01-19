Attribute VB_Name = "modList"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modList
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function ListCommands(Index As Long) As Boolean
If TopList(Index) = True Then ListCommands = True: Exit Function
If MUDHelp(Index) = True Then ListCommands = True: Exit Function 'check for help commands
If ListEmotes(Index) = True Then ListCommands = True: Exit Function 'check for 'emotions' command
If ListSpells(Index) = True Then ListCommands = True: Exit Function 'check for the 'spells' command
If Inventory(Index) = True Then ListCommands = True: Exit Function 'check for the 'inv' command
If EXP(Index) = True Then ListCommands = True: Exit Function 'check for the 'exp' command
If Who(Index) = True Then ListCommands = True: Exit Function 'check for the 'who' command
ListCommands = False
End Function

Public Function TopList(Index As Long) As Boolean
If modSC.FastStringComp(TrimIt(LCaseFast(X(Index))), "top") Then
    TopList = True
    Dim tArr(9) As String
    
    Dim i As Long
    Dim dCurEXP As Double
    Dim iCurID As Long
    Dim iLev As Long
    Dim sAddedAlready As String
    Dim sToSend As String
    Dim LL As Long
    Dim j As Long
    sToSend = YELLOW & " Rank Player's Name" & Space(7) & "EXP" & vbCrLf & "É" & String(61, "Í") & "»" & vbCrLf
    iLev = 0
    dCurEXP = -1
RecurLoop:
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If dCurEXP < .dTotalEXP And InStr(1, sAddedAlready, ":" & i & ";") = 0 Then
                dCurEXP = .dTotalEXP
                iCurID = i
            End If
        End With
        If i = UBound(dbPlayers) Then
            tArr(iLev) = YELLOW & CStr(iLev + 1) & "." & Space(5 - Len(CStr(iLev + 1) & ".")) & LIGHTBLUE & dbPlayers(iCurID).sPlayerName & Space(20 - Len(dbPlayers(iCurID).sPlayerName)) & CStr(dCurEXP) & YELLOW
            iLev = iLev + 1
            dCurEXP = -1
            sAddedAlready = sAddedAlready & ":" & iCurID & ";"
            If iLev = 11 Or iLev = UBound(dbPlayers) - 1 Then
                For j = LBound(tArr) To UBound(tArr)
                    If Len(tArr(j)) > LL Then LL = Len(tArr(j))
                    tArr(j) = "º" & tArr(j)
                    If DE Then DoEvents
                Next
                For j = LBound(tArr) To UBound(tArr)
                    tArr(j) = tArr(j) & Space(IIf(tArr(j) = "º", 61, 65 - Len(tArr(j)))) & "º"
                    If DE Then DoEvents
                Next
                WrapAndSend Index, sToSend & YELLOW & Join(tArr, vbCrLf) & vbCrLf & "È" & String(61, "Í") & "¼" & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            Else
                GoTo RecurLoop
            End If
        End If
        If DE Then DoEvents
    Next
End If
End Function

Public Function TopGuild(Index As Long) As Boolean
Dim sG As String
Dim i As Long, j As Long
Dim tArr1() As String, tArr2() As String
Dim LL As Long
If modSC.FastStringComp(LCaseFast(X(Index)), "top guild") Then
    TopGuild = True
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If InStr(1, sG, .sGuild) = 0 Then
                If Not modSC.FastStringComp(.sGuild, "0") Then
                    sG = sG & .sGuild & "Ø"
                End If
            End If
        End With
        If DE Then DoEvents
    Next
    If Len(sG) = 0 Then
        WrapAndSend Index, RED & "There are currently no established guilds." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    SplitFast Left$(sG, Len(sG) - 1), tArr1, "Ø"
    ReDim tArr2(UBound(tArr1)) As String
    For i = LBound(tArr2) To UBound(tArr2)
        tArr2(i) = "0"
        If DE Then DoEvents
    Next
    For i = LBound(tArr1) To UBound(tArr1)
        For j = LBound(dbPlayers) To UBound(dbPlayers)
            With dbPlayers(j)
                If modSC.FastStringComp(tArr1(i), .sGuild) Then
                    tArr2(i) = CDbl(tArr2(i)) + .dTotalEXP
                End If
            End With
            If DE Then DoEvents
        Next
        If DE Then DoEvents
    Next
    Dim bFlag As Boolean
    bFlag = False
RecurLoop:
    For i = LBound(tArr1) To UBound(tArr1) - 1
        If CDbl(tArr2(i + 1)) > CDbl(tArr2(i)) Then
            sG = tArr2(i)
            tArr2(i) = tArr2(i + 1)
            tArr2(i + 1) = sG
            sG = tArr1(i)
            tArr1(i) = tArr1(i + 1)
            tArr1(i + 1) = sG
            i = LBound(tArr1)
        End If
        If Not bFlag And i = UBound(tArr1) - 1 Then
            bFlag = True
            GoTo RecurLoop
        End If
        If DE Then DoEvents
    Next
    sG = ""
    For i = LBound(tArr1) To UBound(tArr1)
        sG = sG & YELLOW & (i + 1) & "." & Space(5 - Len(CStr(i + 1))) & GREEN & tArr1(i) & Space(20 - Len(tArr1(i))) & BRIGHTBLUE & tArr2(i) & YELLOW & vbCrLf
        If DE Then DoEvents
    Next
    Erase tArr1
    SplitFast sG, tArr1, vbCrLf
    For j = LBound(tArr1) To UBound(tArr1) - 1
        If Len(tArr1(j)) > LL Then LL = Len(tArr1(j))
        tArr1(j) = "º" & tArr1(j)
        If DE Then DoEvents
    Next
    For j = LBound(tArr1) To UBound(tArr1) - 1
        tArr1(j) = tArr1(j) & Space(IIf(tArr1(j) = "º", 61, 66 - Len(tArr1(j)))) & "º"
        If DE Then DoEvents
    Next
    sG = YELLOW & " Rank  Guild's Name" & Space(8) & "EXP" & vbCrLf & "É" & String(61, "Í") & "»" & vbCrLf
    WrapAndSend Index, sG & YELLOW & Join(tArr1, vbCrLf) & "È" & String(61, "Í") & "¼" & WHITE & vbCrLf
    X(Index) = ""
End If
End Function
