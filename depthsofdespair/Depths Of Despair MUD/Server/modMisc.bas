Attribute VB_Name = "modMisc"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modMisc
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'


Public Function DeathCommands(Index As Long) As Boolean
If KillFam(Index) = True Then DeathCommands = True: Exit Function 'check for the 'kill fam' command
If Suicide(Index) = True Then DeathCommands = True: Exit Function 'check for the 'suicide' command
If ReRoll(Index) = True Then DeathCommands = True: Exit Function
DeathCommands = False
End Function

Public Function PlayerHasEcho(Index As Long, dbIndex As Long) As Boolean
On Error GoTo eh1
If dbIndex = 0 Then
    dbIndex = GetPlayerIndexNumber(Index)
    If dbIndex = 0 Then
        If modMain.pLogOnPW(Index) = False Then PlayerHasEcho = True
    Else
        PlayerHasEcho = IIf(dbPlayers(dbIndex).iEcho = 0, False, True)
    End If
Else
    PlayerHasEcho = IIf(dbPlayers(dbIndex).iEcho = 0, False, True)
End If
Exit Function
eh1:
End Function

Public Function pRest(Index As Long) As Boolean
'function to put players in the 'rest' mode
If modSC.FastStringComp(LCaseFast(X(Index)), "rest") Then
    pRest = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        .iResting = 1
        .iMeditating = 0
        X(Index) = ""
        WrapAndSend Index, BRIGHTWHITE & "You sit and rest" & WHITE & vbCrLf
        If .iSneaking = 0 Then SendToAllInRoom Index, WHITE & .sPlayerName & " sits and rest." & vbCrLf, .lLocation
    End With
ElseIf modSC.FastStringComp(LCaseFast(Left(X(Index), 3)), "med") Then
    pRest = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        .iMeditating = 1
        .iResting = 0
        X(Index) = ""
        WrapAndSend Index, BRIGHTWHITE & "You sit and meditate" & WHITE & vbCrLf
        If .iSneaking = 0 Then SendToAllInRoom Index, WHITE & .sPlayerName & " sits and meditates." & vbCrLf, .lLocation
    End With
End If
End Function

Public Function SetEcho(Index As Long) As Boolean
Dim CurEcho As Long
Dim OnOff$
Dim dbIndex As Long
If modSC.FastStringComp(TrimIt(LCaseFast(X(Index))), "=echo") Then
    SetEcho = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        CurEcho = .iEcho
        If CurEcho = 0 Then
            .iEcho = 1
            OnOff$ = "on."
        Else
            .iEcho = 0
            OnOff$ = "off."
        End If
        X(Index) = ""
        WrapAndSend Index, BLUE & "Your echo is now " & OnOff$ & WHITE & vbCrLf
    End With
ElseIf modSC.FastStringComp(TrimIt(LCaseFast(X(Index))), "=a") Then
    SetEcho = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        CurEcho = modMiscFlag.GetStatsPlus(dbIndex, [Pallete Number])
        If CurEcho <> 4 Then
            modMiscFlag.SetStatsPlus dbIndex, [Pallete Number], 4
            OnOff$ = "off."
        Else
            modMiscFlag.SetStatsPlus dbIndex, [Pallete Number], 0
            OnOff$ = "on."
        End If
        X(Index) = ""
        WrapAndSend Index, BLUE & "Your ANSI is now " & OnOff$ & WHITE & vbCrLf
    End With
End If
End Function

Public Function Who(Index As Long) As Boolean
'////////WHO////////
Dim Peeps As String, pLevel As String, pClass As String, pGuild As String
Dim ToSend$, PeepToLevelSpace As Long, ClassToLevelSpace As Long 'spaces in-between the words
Dim iSpaceToGuild As Long
Dim iLongest As Long
Dim pEvil As String
Dim test As String
Dim i As Long
Dim sCol As String

'tempary arrays
Dim tArr1() As String, tArr2() As String, tArr3() As String, tArr4() As String, tArr5() As String
'function to list everyone currently in the game
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "who") Then  'keyword
    Who = True
    sCol = YELLOW & "| "
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If .iIndex <> 0 Then
                Peeps = Peeps & .sPlayerName & ","
                If Len( _
                    .sPlayerName) > PeepToLevelSpace Then PeepToLevelSpace = Len( _
                    .sPlayerName)
                pLevel = pLevel & .iLevel & ","
                If Len( _
                    "Lvl " & .iLevel) > ClassToLevelSpace Then ClassToLevelSpace = _
                    Len("Lvl " & .iLevel)
                pClass = pClass & modGetData.GetClassPointLevel(i, True) & " " & .sClass & ","
                If Len(.sGuild) > iSpaceToGuild Then iSpaceToGuild = Len(.sGuild)
                pGuild = pGuild & .sGuild & ","
                pEvil = pEvil & GetRep(.iIndex) & ","
            End If
        End With
        If DE Then DoEvents
    Next
    SplitFast Left$(Peeps, Len(Peeps) - 1), tArr1, ","
    SplitFast Left$(pLevel, Len(pLevel) - 1), tArr2, ","
    SplitFast Left$(pClass, Len(pClass) - 1), tArr3, ","
    SplitFast Left$(pGuild, Len(pGuild) - 1), tArr4, ","
    SplitFast Left$(pEvil, Len(pEvil) - 1), tArr5, ","
    For i = LBound(tArr3) To UBound(tArr3)
        If Len(tArr3(i)) > iSpaceToGuild Then iSpaceToGuild = Len(tArr3(i))
        If DE Then DoEvents
    Next
    
'    iLongest = iLongest + 6
    For i = LBound(tArr4) To UBound(tArr4)
        If Len(tArr4(i)) > iLongest Then iLongest = Len(tArr4(i))
        If DE Then DoEvents
    Next
    For i = LBound(tArr1) To UBound(tArr1) 'configure the list, and make it nicely formated
        ToSend$ = ToSend$ & tArr5(i) & Space$(18 - Len(tArr5(i))) & sCol & GREEN & tArr1(i) & Space$(( _
            PeepToLevelSpace + 1) - Len(tArr1(i))) & sCol & MAGNETA & "Lvl " & tArr2( _
            i) & Space$((ClassToLevelSpace + 1) - Len("Lvl " & tArr2( _
            i))) & sCol & BLUE & tArr3(i)
            If Not modSC.FastStringComp(tArr4(i), "0") Then
                ToSend$ = ToSend$ & Space$((iSpaceToGuild + 1) - Len(tArr3(i))) & sCol & "of " & tArr4(i) & vbCrLf
                'ToSend$ = ToSend$ & Space$(iLongest + 1 - Len(tArr4(i))) & Space$(5) & "º" & vbCrLf
            Else
                ToSend$ = ToSend$ & Space$((iLongest + 1) + (iSpaceToGuild + 4) - Len(tArr3(i))) & YELLOW & vbCrLf
            End If
        If DE Then DoEvents
    Next
    Dim Templong As Long
    Templong = iLongest
    For i = LBound(tArr1) To UBound(tArr1)
        test = tArr5(i) & Space$(18 - Len(tArr5(i))) & tArr1(i) & Space(( _
            PeepToLevelSpace + 1) - Len(tArr1(i))) & "Lvl " & tArr2( _
            i) & Space((ClassToLevelSpace + 1) - Len("Lvl " & tArr2( _
            i))) & tArr3(i)
        If Not modSC.FastStringComp(tArr4(i), "0") Then
            test = test & Space$((iSpaceToGuild + 1) - Len(tArr3(i))) & "of " & tArr4(i)
            'test = test & Space$(Templong + 1 - Len(tArr4(i))) & "º" & vbCrLf
        Else
            test = test & Space$((Templong + 1) + (iSpaceToGuild + 4) - Len(tArr3(i))) & vbCrLf
        End If
        If Len(test) > iLongest Then iLongest = Len(test)
        If DE Then DoEvents
        Debug.Print Len(test)
    Next
    iLongest = iLongest - 4
    'make a header on the message
    test = ""
    For i = 1 To iLongest Step 4
        test = test & "=-_-"
        If DE Then DoEvents
    Next
    ToSend$ = YELLOW & "Listing villagers..." & vbCrLf & test & vbCrLf & ToSend$
    'close off the message
    ToSend$ = ToSend$ & test & vbCrLf & WHITE
    'send the message
    WrapAndSend Index, ToSend$
    X(Index) = ""
    Exit Function
End If
'////////END////////
End Function
