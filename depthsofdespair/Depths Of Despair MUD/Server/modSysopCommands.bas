Attribute VB_Name = "modSysopCommands"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modSysopCommands
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function SysAddEXP(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 12)), "sys add exp ") Then
    SysAddEXP = True
    Dim Adding As Double
    If IsSysop(Index) = False Then Exit Function
    Adding = Val(Mid$(X(Index), 12, Len(X(Index)) - 11))
    If Adding > 999999999# Then
        WrapAndSend Index, RED & "You are unable to create that much experience!" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(GetPlayerIndexNumber(Index))
        .dEXP = .dEXP + Adding
        .dTotalEXP = .dTotalEXP + Adding
        WrapAndSend Index, BLUE & "You focus your energy into one area..." & vbCrLf & _
            BLUE & "Small " & GREEN & "green " & BLUE & "lights surround you!" & vbCrLf & _
            WHITE & "You gain " & Adding & " exprience!" & WHITE & vbCrLf
        SendToAllInRoom Index, BLUE & .sPlayerName & " focuses energy into one area..." & vbCrLf & _
            BLUE & "Small " & GREEN & "green " & BLUE & "lights surround " & .sPlayerName & "." & vbCrLf & _
            WHITE, .lLocation
    End With
    X(Index) = ""
End If
End Function

Public Function SysBoot(Index As Long) As Boolean
Dim s As String
Dim i As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 9)), "sys boot ") Then
    SysBoot = True
    If IsSysop(Index) = False Then Exit Function
    s = Mid$(X(Index), 10)
    s = LCaseFast(s)
    s = SmartFind(Index, s, All_Players)
    For i = 1 To UBound(dbPlayers)
        With dbPlayers(i)
            If modSC.FastStringComp(LCaseFast(.sPlayerName), s) Then
                If .iIndex <> 0 Then
                    If frmMain.ws(.iIndex).State = sckConnected Then
                        WrapAndSend .iIndex, BRIGHTBLUE & "Sysop discontected you." & WHITE, False
                        frmMain.ws(.iIndex).Close
                    End If
                    frmMain.lstUsers.SetItemText .iIndex, "[Line " & CStr(.iIndex) & " - Open]"
                    If Val(frmMain.Online.Caption) > 0 Then frmMain.Online.Caption = Val(frmMain.Online.Caption) - 1
                    X(.iIndex) = ""
                    PNAME(.iIndex) = ""
                    pPoint(.iIndex) = 0
                    UpdateList "}bLine " & (.iIndex) & " has been booted from the server. }b(}n}i" & Time & "}n}b)"
                    .iIndex = 0
                    X(Index) = ""
                    WrapAndSend Index, BRIGHTRED & "Done. (" & .sPlayerName & " booted from server.)" & WHITE & vbCrLf
                    Exit Function
                End If
            End If
        End With
        If DE Then DoEvents
    Next
    X(Index) = ""
    WrapAndSend Index, RED & "Couldn't find player: " & s & "." & WHITE & vbCrLf
End If
End Function

Public Function SysReName(Index As Long) As Boolean
Dim s As String
Dim t As String
Dim i As Long
Dim j As Long
Dim bFound As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 9)), "sys name ") Then
    SysReName = True
    If IsSysop(Index) = False Then Exit Function
    s = Mid$(X(Index), 10)
    i = InStr(1, s, " as ")
    t = Left$(s, i - 1)
    t = SmartFind(Index, t, All_Players)
    j = GetPlayerIndexNumber(, t)
    If j = 0 Then
        WrapAndSend Index, RED & "You are unable to find that person!" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    t = Mid$(s, i + 4)
    If modValidate.ValidateName(LCaseFast(t)) = True Then
        WrapAndSend Index, MAGNETA & "Invalid name: " & t & WHITE & vbCrLf
    End If
    'name has to be a minimum of 6 characters
    If Len(t) < 5 Then
        'send errormessage
        WrapAndSend Index, MAGNETA & "The name must be at least 5 characters long." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    'max of 12 characters
    ElseIf Len(t) > 12 Then 'send error message
        WrapAndSend Index, MAGNETA & "The name cannot be longer then 12 characters." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If InStr(1, t, " ") Then
        WrapAndSend Index, MAGNETA & "The name can not contain a space." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If modSC.FastStringComp(LCaseFast(Left$(t, 3)), "new") Then t = Right$(t, Len(t) - 3)
    If t = "" Then
        WrapAndSend Index, MAGNETA & "Your name cannot contain be nothing." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(j)
        If .sSeenAs = .sPlayerName Then
            .sSeenAs = t
            .sPlayerName = .sSeenAs
        Else
            .sSeenAs = t
        End If
    End With
    WrapAndSend Index, BRIGHTWHITE & "Done." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function SysMonster(Index As Long) As Boolean
Dim s As String
Dim db As Long
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 12)), "sys gen mon ") Then
    SysMonster = True
    If IsSysop(Index) = False Then Exit Function
    s = ReplaceFast(X(Index), "sys gen mon ", "")
    If IsNumeric(s) Then
        db = GetMonsterID(, CLng(Val(s)))
    Else
        s = SmartFind(Index, s, All_Monsters)
        db = GetMonsterID(s)
    End If
    If db = 0 Then
        WrapAndSend Index, RED & "No monster exsist by " & s & WHITE & vbCrLf
    Else
        dbIndex = GetPlayerIndexNumber(Index)
        modMonsters.GenAMonster CLng(dbPlayers(dbIndex).lLocation), True, CLng(dbMonsters(db).lMobGroup), CLng(dbMonsters(db).lID), dbPlayers(dbIndex).lDBLocation
    End If
    X(Index) = ""
End If
End Function

Public Function IsSysCommand(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "sys") Then
    If SysList(Index) = True Then IsSysCommand = True: Exit Function
    If Conjure(Index) = True Then IsSysCommand = True: Exit Function
    If Teleport(Index) = True Then IsSysCommand = True: Exit Function
    If SysAddEXP(Index) = True Then IsSysCommand = True: Exit Function
    If SysLight(Index) = True Then IsSysCommand = True: Exit Function
    If HellStorm(Index) = True Then IsSysCommand = True: Exit Function
    If PhaseWalls(Index) = True Then IsSysCommand = True: Exit Function
    If SysListUsers(Index) = True Then IsSysCommand = True: Exit Function
    If SysListItemsBy(Index) = True Then IsSysCommand = True: Exit Function
    If SysListItems(Index) = True Then IsSysCommand = True: Exit Function
    If SysLookInv(Index) = True Then IsSysCommand = True: Exit Function
    If SysSpeak(Index) = True Then IsSysCommand = True: Exit Function
    If SysAddEvil(Index) = True Then IsSysCommand = True: Exit Function
    If SysGhost(Index) = True Then IsSysCommand = True: Exit Function
    If SysListLimited(Index) = True Then IsSysCommand = True: Exit Function
    If SysEdit(Index) = True Then IsSysCommand = True: Exit Function
    If SysDebug(Index) = True Then IsSysCommand = True: Exit Function
    If SysRoomInfo(Index) = True Then IsSysCommand = True: Exit Function
    If SysMonster(Index) = True Then IsSysCommand = True: Exit Function
    If SysReName(Index) = True Then IsSysCommand = True: Exit Function
    If SysBoot(Index) = True Then IsSysCommand = True: Exit Function
    If SysHelps(Index) = True Then IsSysCommand = True: Exit Function
    'If Horse(Index) = True Then IsSysCommand = True: Exit Function
End If
If modDebug.ParseDebug(Index) = True Then IsSysCommand = True: Exit Function
End Function

Public Function SysGhost(Index As Long) As Boolean
If Left$(LCaseFast(X(Index)), 10) = "sys ghost " Then
    SysGhost = True
    If IsSysop(Index) = False Then Exit Function
    Select Case CLng(Right$(X(Index), 1))
        Case 0
            dbPlayers(GetPlayerIndexNumber(Index)).iGhostMode = 0
            WrapAndSend Index, LIGHTBLUE & "You return from " & WHITE & "ghost " & LIGHTBLUE & "mode." & WHITE & vbCrLf
        Case 1
            dbPlayers(GetPlayerIndexNumber(Index)).iGhostMode = 1
            WrapAndSend Index, LIGHTBLUE & "You enter " & WHITE & "ghost " & LIGHTBLUE & "mode." & WHITE & vbCrLf
        Case Else
            WrapAndSend Index, RED & "Syntax: sys ghost [0 or 1]" & WHITE & vbCrLf
    End Select
    X(Index) = ""
End If
End Function

Public Function SysRoomInfo(Index As Long) As Boolean
Dim dbIndex As Long
Dim dbMapId As Long
Dim s As String
If LCaseFast(X(Index)) = "sys room info" Then
    SysRoomInfo = True
    If IsSysop(Index) = False Then Exit Function
    dbIndex = GetPlayerIndexNumber(Index)
    dbMapId = dbPlayers(dbIndex).lDBLocation
    With dbMap(dbMapId)
        s = GREEN & "Description     Value" & vbCrLf & YELLOW
        s = s & "Room ID         " & .lRoomID & vbCrLf
        s = s & "Room Title      " & .sRoomTitle & vbCrLf
        s = s & "Mon #'s         " & .sMonsters & vbCrLf
        s = s & "AMONIDS         " & .sAMonIds & vbCrLf
        s = s & "Player #'s      " & modGetData.GetPlayersIDsHere(.lRoomID) & vbCrLf
        s = s & "Items           " & .sItems & vbCrLf
        s = s & "Hidden          " & .sHidden & vbCrLf
        s = s & "Enviroment      " & modGetData.GetMapEnviron(dbMapId) & vbCrLf
        s = s & "Max Regen       " & .iMaxRegen & vbCrLf
        s = s & "Mon Group       " & .iMobGroup & vbCrLf
        s = s & "Special Mon     " & .lSpecialMon & vbCrLf
        s = s & "Special Item    " & .lSpecialItem & vbCrLf
        s = s & "Light Value     " & .lLight & vbCrLf
        s = s & "Safe Room       " & modGetData.GetMapSafe(dbMapId) & vbCrLf
        s = s & "Death Room      " & .lDeathRoom & vbCrLf
        s = s & "Room Type       " & modGetData.GetMapRoomType(dbMapId) & vbCrLf
        s = s & "Train Class     " & .iTrainClass & vbCrLf
        s = s & "Script          " & .sScript & vbCrLf
        WrapAndSend Index, s & WHITE
    End With
    X(Index) = ""
End If

End Function

Public Function Conjure(Index As Long) As Boolean
'////////CONJURE////////
Dim a As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 7)), "sys con") Then
    Conjure = True
    If IsSysop(Index) = False Then Exit Function
    X(Index) = ReplaceFast(X(Index), "sys ", "")
    For a = 1 To InStr(1, X(Index), " ")
        X(Index) = Mid$(X(Index), 2)
    Next a
    If modSC.FastStringComp(TrimIt(X(Index)), "") Then
        WrapAndSend Index, RED & "To create an item, you have to name an item." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    Dim TempItem$
    Dim lItemID As Long
    TempItem$ = SmartFind(Index, X(Index), All_Items)
    lItemID = GetItemID(TempItem$)
    If lItemID = 0 Then
        WrapAndSend Index, RED & "You can't create something that doesn't exsist." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    With dbItems(lItemID)
        If .iInGame >= .iLimit And .iLimit <> 0 Then
            WrapAndSend Index, RED & "You have a problem trying to create this item." & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        ElseIf .iLimit <> 0 Then
            .iInGame = .iInGame + 1
        End If
        TempItem$ = .sItemName
    End With
    Dim dbItemID As Long
    dbItemID = GetItemID(TempItem$)
    With dbPlayers(GetPlayerIndexNumber(Index))
        If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
        .sInventory = .sInventory & ":" & dbItems(dbItemID).iID & "/" & dbItems(dbItemID).lDurability & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(dbItemID).iUses & ";"
        WrapAndSend Index, BLUE & "You wave you hand over your pouch and..." & vbCrLf & _
                    "A " & RED & "red " & BLUE & "sphere hovers over you!" & vbCrLf & _
                    BLUE & "You conjure a(n) " & TempItem$ & "!" & vbCrLf & WHITE
        SendToAllInRoom Index, BLUE & .sPlayerName & " waves their hand over their pouch and..." & vbCrLf & _
                    "A " & RED & "red " & BLUE & "sphere hovers over them!" & vbCrLf & _
                    WHITE, .lLocation
    End With
    X(Index) = ""
End If
'////////END////////
End Function

Public Function Teleport(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 7)), "sys tel") Then
    Teleport = True
    Dim tLocation&, tLocationTeleport&
    If IsSysop(Index) = False Then Exit Function
    Dim tXINDEX As String
    tXINDEX = ReplaceFast(X(Index), "sys ", "")
    For a = 1 To InStr(1, tXINDEX, " ")
        tXINDEX = Mid$(tXINDEX, 2)
    Next a
    If tXINDEX = "" Then
        WrapAndSend Index, RED & "If ye wish to teleport, ye must specify is room." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    On Error GoTo bad
    tLocationTeleport& = CLng(TrimIt(tXINDEX))
    Dim FoundRoom As Boolean
    Dim lRoomIndex As Long
    lRoomIndex = GetMapIndex(CLng(tLocationTeleport&))
    If lRoomIndex = 0 Then
        X(Index) = ""
        WrapAndSend Index, RED & "You are unable to locate the room." & WHITE & vbCrLf
        Exit Function
    End If
    RemoveFromParty Index
    With dbPlayers(GetPlayerIndexNumber(Index))
        SendToAllInRoom Index, _
            BLUE & .sPlayerName & " chants some words..." & vbCrLf & "In a bright " & WHITE & _
            "white " & BLUE & "light, " & .sPlayerName & " vanishes!" & vbCrLf & WHITE, .lLocation
        .lLocation = tLocationTeleport&
        .lDBLocation = GetMapIndex(tLocationTeleport&)
        .lRoomSearched = -1
        X(Index) = ""
        WrapAndSend Index, _
            BLUE & "You chant some words..." & vbCrLf & "In a bright " & WHITE & _
            "white " & BLUE & "light, you are elsewhere!" & vbCrLf & WHITE
        SendToAllInRoom Index, BLUE & "In a bright " & WHITE & "white " & BLUE & "light, " & .sPlayerName & " appears in the room." & WHITE & vbCrLf, .lLocation
    End With
End If
Exit Function
bad:
X(Index) = ""
WrapAndSend Index, RED & "You fucker, treat it right!"
End Function

Public Function IsSysop(Index As Long) As Boolean
If modMiscFlag.GetStatsPlus(GetPlayerIndexNumber(Index), [Is A Sysop]) = 0 Then
    IsSysop = False
    WrapAndSend Index, RED & "You are not a sysop." & WHITE & vbCrLf
    X(Index) = ""
    Exit Function
Else
    IsSysop = True
End If
End Function

Public Function SysLight(Index As Long) As Boolean
Dim sTarget As String
Dim iTargetID As Long
Dim dDam As Double
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 9)), "sys light") Then
    SysLight = True
    If IsSysop(Index) = False Then Exit Function
    X(Index) = ReplaceFast(X(Index), "sys ", "", 1, 1)
    For a = 1 To InStr(1, X(Index), " ")
        X(Index) = Mid$(X(Index), 2)
        If DE Then DoEvents
    Next a
    sTarget = X(Index)
    sTarget = SmartFind(Index, sTarget, All_Players)
    iTargetID = GetPlayerIndexNumber(, sTarget)
    If iTargetID = 0 Then
        WrapAndSend Index, RED & "You cannot locate " & sTarget & "." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If dbPlayers(iTargetID).iIndex = 0 Then
        WrapAndSend Index, RED & "You cannot locate " & sTarget & "." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(iTargetID)
        dDam = RndNumber(CDbl(.lMaxHP / 8), CDbl(.lMaxHP / 2))
        .lHP = .lHP - dDam
        WrapAndSend .iIndex, modANSIConst.BGBLUE & BRIGHTYELLOW & "A bolt of lightning strikes you for " & BRIGHTRED & CStr(dDam) & BRIGHTYELLOW & " damage!" & WHITE & vbCrLf
        SendToAllInRoom .iIndex, BRIGHTYELLOW & "A bolt of lightning stikes " & .sPlayerName & " for " & BRIGHTRED & CStr(dDam) & BRIGHTYELLOW & " damage!" & WHITE & vbCrLf, .lLocation
        WrapAndSend Index, BRIGHTYELLOW & "You strike " & .sPlayerName & " with a " & BRIGHTRED & "lightning bolt " & BRIGHTYELLOW & "for " & BRIGHTRED & CStr(dDam) & BRIGHTYELLOW & " damage!" & WHITE & vbCrLf
        CheckDeath .iIndex
    End With
    X(Index) = ""
End If
End Function

Public Function HellStorm(Index As Long) As Boolean
Dim sTarget As String
Dim iTargetID As Long
Dim dDam As Double
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 8)), "sys hell") Then
    HellStorm = True
    If IsSysop(Index) = False Then Exit Function
    X(Index) = ReplaceFast(X(Index), "sys ", "", 1, 1)
    For a = 1 To InStr(1, X(Index), " ")
        X(Index) = Mid$(X(Index), 2)
        If DE Then DoEvents
    Next a
    sTarget = X(Index)
    sTarget = SmartFind(Index, sTarget, All_Players)
    iTargetID = GetPlayerIndexNumber(, sTarget)
    If iTargetID = 0 Then
        WrapAndSend Index, RED & "You cannot locate " & sTarget & "." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If dbPlayers(iTargetID).iIndex = 0 Then
        WrapAndSend Index, RED & "You cannot locate " & sTarget & "." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(iTargetID)
        dDam = RndNumber(CDbl(.lMaxHP / 6), CDbl(.lMaxHP / 1.5))
        .lHP = .lHP - dDam
        WrapAndSend .iIndex, modANSIConst.BGBLUE & BRIGHTYELLOW & "You are ravaged by a " & BRIGHTRED & "hellstorm " & BRIGHTYELLOW & "for " & BRIGHTRED & CStr(dDam) & BRIGHTYELLOW & " damage!" & WHITE & vbCrLf
        SendToAllInRoom .iIndex, BRIGHTYELLOW & .sPlayerName & " is ravaged by a " & BRIGHTRED & "hellstorm " & BRIGHTYELLOW & "for " & BRIGHTRED & CStr(dDam) & BRIGHTYELLOW & " damage!" & WHITE & vbCrLf, .lLocation
        WrapAndSend Index, BRIGHTYELLOW & "You ravage " & .sPlayerName & " with a " & BRIGHTRED & "hellstorm " & BRIGHTYELLOW & "for " & BRIGHTRED & CStr(dDam) & BRIGHTYELLOW & " damage!" & WHITE & vbCrLf
        CheckDeath .iIndex
    End With
    X(Index) = ""
End If
End Function

Public Function PhaseWalls(Index As Long) As Boolean
Dim Direction As String
Dim CurLoc As Long
Dim GoToLoc As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 9)), "sys phase") Then
    If IsSysop(Index) = False Then Exit Function
    PhaseWalls = True
    Direction = ReplaceFast(LCaseFast(X(Index)), "sys phase ", "")
    With dbPlayers(GetPlayerIndexNumber(Index))
        CurLoc = .lLocation
    End With
    Direction = modGetData.GetLongDir(Direction)
    With dbMap(GetMapIndex(CurLoc))
        Select Case Direction
            Case "northwest"
                If .lDNW <> 0 Then
                    If .lNorthWest <> 0 Then
                        GoToLoc = .lNorthWest
                    End If
                End If
            Case "northeast"
                If .lDNE <> 0 Then
                    If .lNorthEast <> 0 Then
                        GoToLoc = .lNorthEast
                    End If
                End If
            Case "southwest"
                If .lDSW <> 0 Then
                    If .lSouthWest <> 0 Then
                        GoToLoc = .lSouthWest
                    End If
                End If
            Case "southeast"
                If .lDSE <> 0 Then
                    If .lSouthEast <> 0 Then
                        GoToLoc = .lSouthEast
                    End If
                End If
            Case "north"
                If .lDN <> 0 Then
                    If .lNorth <> 0 Then
                        GoToLoc = .lNorth
                    End If
                End If
            Case "south"
                If .lDS <> 0 Then
                    If .lSouth <> 0 Then
                        GoToLoc = .lSouth
                    End If
                End If
            Case "east"
                If .lDE <> 0 Then
                    If .lEast <> 0 Then
                        GoToLoc = .lEast
                    End If
                End If
            Case "west"
                If .lDW <> 0 Then
                    If .lWest <> 0 Then
                        GoToLoc = .lWest
                    End If
                End If
            Case "up"
                If .lDU <> 0 Then
                    If .lUp <> 0 Then
                        GoToLoc = .lUp
                    End If
                End If
            Case "down"
                If .lDD <> 0 Then
                    If .lDown <> 0 Then
                        GoToLoc = .lDown
                    End If
                End If
            Case Else
                WrapAndSend Index, RED & "You cannot phase through the door in that direction." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
        End Select
    End With
    If GoToLoc = 0 Then
        WrapAndSend Index, RED & "You cannot phase through the door in that direction." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(GetPlayerIndexNumber(Index))
        .lLocation = GoToLoc
        .lDBLocation = GetMapIndex(.lLocation)
        .lBackUpLoc = GoToLoc
        RemoveFromParty Index
        SendToAllInRoom Index, BRIGHTBLUE & .sPlayerName & " phases through the door to the " & GREEN & Direction & BRIGHTBLUE & "." & WHITE & vbCrLf, CurLoc
        SendToAllInRoom Index, BRIGHTBLUE & .sPlayerName & " comes out of the door to the " & GREEN & modGetData.GetOppositeDirection(Direction) & BRIGHTBLUE & "." & WHITE & vbCrLf, GoToLoc
        WrapAndSend Index, BRIGHTBLUE & "You phase through the door to the " & GREEN & Direction & BRIGHTBLUE & "." & WHITE & vbCrLf
    End With
    X(Index) = ""
End If
End Function

Public Function SysListUsers(Index As Long) As Boolean
Dim i As Long
Dim s As String
If modSC.FastStringComp(LCaseFast(X(Index)), "sys list users") Then
    If IsSysop(Index) = False Then Exit Function
    SysListUsers = True
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If .iIndex <> 0 Then
                s = s & .sPlayerName & Space$(20 - Len(.sPlayerName)) & .lLocation & Space$(18 - Len(CStr(.lLocation))) & .iIndex & vbCrLf
            End If
        End With
        If DE Then DoEvents
    Next
   s = "Name" & Space(16) & "Location" & Space$(10) & "Line Number" & vbCrLf & s
   WrapAndSend Index, s & WHITE & vbCrLf
   X(Index) = ""
End If
End Function

Public Function SysLookInv(Index As Long) As Boolean
Dim s As String
Dim iPID As Long
Dim ToSend$
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 12)), "sys look inv") Then
    If IsSysop(Index) = False Then Exit Function
    SysLookInv = True
    s = ReplaceFast(LCaseFast(X(Index)), "sys look inv ", "")
    s = SmartFind(Index, s, All_Players)
    iPID = GetPlayerIndexNumber(, s)
    If iPID = 0 Then
        WrapAndSend Index, RED & "You cannot find that person." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(iPID)
        If .iIndex = 0 Then
            WrapAndSend Index, RED & "You cannot find that person." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If (.dGold <> 0) And (.lPaper = 0) Then
            ToSend$ = ToSend$ & YELLOW & "You have " & GREEN & .dGold & YELLOW & " gold,"
        ElseIf (.dGold = 0) And (.lPaper <> 0) Then
            ToSend$ = ToSend$ & YELLOW & "You have " & GREEN & .lPaper & YELLOW & " pieces of paper,"
        ElseIf (.dGold <> 0) And (.lPaper <> 0) Then
            ToSend$ = ToSend$ & YELLOW & "You have " & GREEN & .dGold & YELLOW & " gold," & GREEN & .lPaper & YELLOW & " pieces of paper,"
        End If
        If modSC.FastStringComp(.sInventory, "") Then .sInventory = "0"
        If .sInventory <> "0" Then
            ToSend$ = ToSend$ & GREEN & modGetData.GetPlayersInvFromNums(.iIndex, True, iPID)
        End If
        ToSend$ = ToSend$ & GREEN
        'get all the equipment they are wearing
        ToSend$ = ToSend$ & modGetData.GetPlayersEqFromNums(.iIndex, , iPID) & modItemManip.GetListOfLettersFromInv(iPID)
        If Not modSC.FastStringComp(ToSend$, MAGNETA & "You have upon you:" & vbCrLf & GREEN) Then
            ToSend$ = ReplaceFast(ToSend$, ",", YELLOW & ", " & GREEN) 'format the message
            ToSend$ = ReplaceFast(ToSend, YELLOW & ", " & GREEN, YELLOW & "." & GREEN, Len(ToSend$) - 4, 1)
            'ToSend$ = Left$(ToSend$, Len(ToSend$) - 3) & YELLOW & "." 'finish the message
            ToSend$ = GREEN & ToSend$ & WHITE & vbCrLf 'finish up the message
        Else
            ToSend$ = ToSend$ & "Absolutly nothing." & WHITE & vbCrLf
        End If
        WrapAndSend Index, ToSend$ 'send to the player
        X(Index) = ""
    End With
End If
End Function

Public Function SysSpeak(Index As Long) As Boolean
Dim s As String
Dim iPID As Long
Dim ToSend$
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 9)), "sys speak") Then
    If IsSysop(Index) = False Then Exit Function
    SysSpeak = True
    s = ReplaceFast(LCaseFast(X(Index)), "sys speak ", "")
    ToSend$ = s
    s = Mid$(s, 1, InStr(1, s, " ") - 1)
    s = SmartFind(Index, s, All_Players)
    iPID = GetPlayerIndexNumber(, s)
    s = ReplaceFast(ToSend$, Mid$(ToSend$, 1, InStr(1, ToSend, " ")), "")
    If iPID = 0 Then
        WrapAndSend Index, RED & "You cannot find that person." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(iPID)
        If .iIndex = 0 Then
            WrapAndSend Index, RED & "You cannot find that person." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        ToSend$ = BRIGHTWHITE & "A message from the gods: " & vbCrLf & s & WHITE & vbCrLf
        WrapAndSend .iIndex, ToSend$
        WrapAndSend Index, BRIGHTWHITE & "Message has been sent." & vbCrLf & WHITE
    End With
    X(Index) = ""
End If
End Function

Public Function SysHelps(Index As Long) As Boolean
Dim s As String
If modSC.FastStringComp(LCaseFast(X(Index)), "sys help1") Then
    If IsSysop(Index) = False Then Exit Function
    SysHelps = True
    s = _
    BRIGHTWHITE & "Sysop Help Menu 1:" & vbCrLf & "--------------------------------------------------------------------" & vbCrLf & "sys list items by." & LIGHTBLUE & "<Race>" & BRIGHTWHITE & "." & LIGHTBLUE & "<Class>" & BRIGHTWHITE & "." & LIGHTBLUE & "<Level>" & BRIGHTWHITE & "." & LIGHTBLUE & "<ClassPts>" & BRIGHTWHITE & vbCrLf & "sys edit " & LIGHTBLUE & "<user name>" & BRIGHTWHITE & " " & LIGHTBLUE & "<([------------------------------>" & BRIGHTWHITE & vbCrLf & "                        maxhp " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        maxma " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        str " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        int " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        cha " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        dex " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        agil " & LIGHTBLUE & "<#>" & _
    BRIGHTWHITE & vbCrLf & "                        exp " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        level " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        class " & LIGHTBLUE & "<(classid,class name)>" & BRIGHTWHITE & vbCrLf & "                        hunger " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        stamina " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        sysop " & LIGHTBLUE & "<(0,1,2)>" & BRIGHTWHITE & vbCrLf & "                        lives " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        gold " & LIGHTBLUE & "<#>" & BRIGHTWHITE & vbCrLf & "                        vision " & LIGHTBLUE & "<#(-5 to 10)>" & BRIGHTWHITE & vbCrLf & "                        misc.invisible " & LIGHTBLUE & "<(0,1)>" & BRIGHTWHITE & vbCrLf & "                        misc.seehidden " & LIGHTBLUE & "<(0,1)>" & BRIGHTWHITE & vbCrLf & "                        misc.seeinvisible " & _
    LIGHTBLUE & "<(0,1)>" & _
    BRIGHTWHITE & vbCrLf & "                        misc.gibberish " & LIGHTBLUE & "<#(0-4)>" & BRIGHTWHITE & vbCrLf & "                        evil " & LIGHTBLUE & "<#(1000 or less)>" & BRIGHTWHITE & "])>" & BRIGHTWHITE & WHITE & vbCrLf
    WrapAndSend Index, s
    X(Index) = ""
ElseIf modSC.FastStringComp(LCaseFast(X(Index)), "sys help2") Then
    If IsSysop(Index) = False Then Exit Function
    SysHelps = True
    s = _
    BRIGHTWHITE & "Sysop Help Menu 2" & vbCrLf & "---------------------------------------------------------------------" & vbCrLf & "debug." & vbCrLf & "      view." & vbCrLf & "           roomitems" & vbCrLf & "           pinv" & vbCrLf & "           inv." & LIGHTBLUE & "<user name>" & BRIGHTWHITE & vbCrLf & "           remoteroomitems." & LIGHTBLUE & "<roomid>" & BRIGHTWHITE & vbCrLf & "           pquest" & vbCrLf & "           classpts." & LIGHTBLUE & "<user name>" & BRIGHTWHITE & vbCrLf & "           quest." & LIGHTBLUE & "<user name>" & BRIGHTWHITE & vbCrLf & "           date" & vbCrLf & "           eq" & vbCrLf & "           mons" & vbCrLf & "           roomswithmons" & vbCrLf & "           roomlight" & vbCrLf & "      edit." & vbCrLf & "           time." & LIGHTBLUE & "<# hours>" & BRIGHTWHITE & "." & LIGHTBLUE & "<# minutes>" & BRIGHTWHITE & "." & LIGHTBLUE & "<# seconds>" & BRIGHTWHITE & vbCrLf & "           classpts." & LIGHTBLUE & "<user name>" & _
    BRIGHTWHITE & "." & LIGHTBLUE & _
    "<#to modify by>" & BRIGHTWHITE & vbCrLf & "           limited." & LIGHTBLUE & "<item name>" & BRIGHTWHITE & "." & LIGHTBLUE & "<# to modify by>" & BRIGHTWHITE & vbCrLf & "           ac." & LIGHTBLUE & "<user name>" & BRIGHTWHITE & "." & LIGHTBLUE & "<# to modify by>" & BRIGHTWHITE & vbCrLf & "           destroyitem." & LIGHTBLUE & "<item id #>" & BRIGHTWHITE & vbCrLf & "           roomlight." & LIGHTBLUE & "<-200 to 200>" & BRIGHTWHITE & vbCrLf & "           saferoom." & LIGHTBLUE & "<0 or 1>" & BRIGHTWHITE & vbCrLf & "           monattack." & LIGHTBLUE & "<monster name in room>" & BRIGHTWHITE & "." & LIGHTBLUE & "<0 or 1>" & BRIGHTWHITE & vbCrLf & "           killmon." & LIGHTBLUE & "<monster name in room>" & BRIGHTWHITE & vbCrLf & "      crte." & vbCrLf & "           item." & LIGHTBLUE & "<FORMAT: #/#/E{}F{}A{}B{0|0|0|0}/#>" & BRIGHTWHITE & vbCrLf & "           door." & LIGHTBLUE & "<n,s,e,w,nw,ne,sw,se,u,d>" & BRIGHTWHITE & "." & LIGHTBLUE & "<0 or 1>" & _
    BRIGHTWHITE & vbCrLf & "      runn." & vbCrLf & "           script." & LIGHTBLUE & "<script>" & BRIGHTWHITE & WHITE & vbCrLf
    WrapAndSend Index, s
    X(Index) = ""
End If
End Function

Public Function SysList(Index As Long) As Boolean
Dim s As String
If modSC.FastStringComp(LCaseFast(X(Index)), "sys") Then
    If IsSysop(Index) = False Then Exit Function
    SysList = True
'    s = BRIGHTWHITE & "sys conjure [item]" & vbCrLf & "sys add exp [###]" & _
'        vbCrLf & "sys hellstorm [user]" & vbCrLf & "sys phase [direction]" & vbCrLf _
'        & "sys list users" & vbCrLf & "sys look inv [user]" & vbCrLf & _
'        "sys speak [user] [message]" & WHITE & vbCrLf
    s = _
    BRIGHTWHITE & "Sys Add EXP " & LIGHTBLUE & "[Amount To Add]" & BRIGHTWHITE & vbCrLf & "Sys Name " & LIGHTBLUE & "[Players Name]" & BRIGHTWHITE & " as " & LIGHTBLUE & "[New Name]" & BRIGHTWHITE & vbCrLf & "Sys Gen Mon " & LIGHTBLUE & "[Monster Number or Monster Name]" & BRIGHTWHITE & vbCrLf & "Sys Ghost " & LIGHTBLUE & "[0,1]" & BRIGHTWHITE & vbCrLf & "Sys Con " & LIGHTBLUE & "[Item Name]" & BRIGHTWHITE & vbCrLf & "Sys Tel " & LIGHTBLUE & "[Room Number]" & BRIGHTWHITE & vbCrLf & "Sys Light " & LIGHTBLUE & "[Players Name]" & BRIGHTWHITE & vbCrLf & "Sys Hell " & LIGHTBLUE & "[Players Name]" & BRIGHTWHITE & vbCrLf & "Sys Phase " & LIGHTBLUE & "[n/s/e/w/nw/se/sw/ne/u/d]" & BRIGHTWHITE & vbCrLf & "Sys List Users" & vbCrLf & "Sys Look Inv " & LIGHTBLUE & "[Players Name]" & BRIGHTWHITE & vbCrLf & "Sys Speak " & LIGHTBLUE & "[Players Name]" & BRIGHTWHITE & " " & LIGHTBLUE & "[Message]" & BRIGHTWHITE & vbCrLf & "Sys Help1" & _
    vbCrLf & "Sys Help2" & vbCrLf & "Sys Add Evil " & LIGHTBLUE & "[Amount To Add]" & BRIGHTWHITE & vbCrLf & "Sys Room Info" & vbCrLf & "Sys List Items" & vbCrLf & "Sys List Items By.*.*.*.*" & vbCrLf & "Sys List Limited" & vbCrLf & "Sys Edit " & LIGHTBLUE & "[Player Name]" & BRIGHTWHITE & " " & LIGHTBLUE & "[Edit]" & BRIGHTWHITE & " " & LIGHTBLUE & "[Value]" & BRIGHTWHITE & vbCrLf & "Sys Boot " & LIGHTBLUE & "[Player Name]" & BRIGHTWHITE & vbCrLf & "Sys Debug" & WHITE & vbCrLf
    WrapAndSend Index, s
    X(Index) = ""
End If
End Function

Public Function SysAddEvil(Index As Long) As Boolean
Dim s As String
Dim d As Double
Dim pID As Long
Dim i As Long
'sys add evil master -100
If modSC.FastStringComp(Left$(LCaseFast(X(Index)), 13), "sys add evil ") Then
    If IsSysop(Index) = False Then Exit Function
    SysAddEvil = True
    s = X(Index)
    s = ReplaceFast(s, "sys add evil ", "", 1, 1)
    i = InStrRev(s, " ")
    If i = 0 Then
        WrapAndSend Index, RED & "You cannot find that person." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    d = CDbl(Val(Mid$(s, i)))
    s = ReplaceFast(s, CStr(d), "")
    s = TrimIt(s)
    s = SmartFind(Index, s, All_Players)
    pID = GetPlayerIndexNumber(, s)
    If pID = 0 Then
        WrapAndSend Index, RED & "You cannot find that person." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(pID)
        If .iIndex = 0 Then
            WrapAndSend Index, RED & "You cannot find that person." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .iEvil = .iEvil + d
        WrapAndSend Index, BLUE & "You give " & RED & .sPlayerName & " " & GREEN & CStr(d) & BLUE & " evil points." & WHITE & vbCrLf
    End With
    X(Index) = ""
End If
End Function

Public Function SysListItems(Index As Long) As Boolean
Dim s As String
Dim i As Long
If modSC.FastStringComp(LCaseFast(X(Index)), "sys list items") Then
    If IsSysop(Index) = False Then Exit Function
    SysListItems = True
    For i = LBound(dbItems) To UBound(dbItems)
        s = s & dbItems(i).sItemName & vbCrLf
        If DE Then DoEvents
    Next
    X(Index) = ""
    WrapAndSend Index, s
End If
End Function

Public Function SysListItemsBy(Index As Long) As Boolean
Dim s As String
Dim i As Long
Dim Arr() As String
Dim dbIndex As Long
'sys list items by.race.class.level.classpts
If LCaseFast(X(Index)) Like "sys list items by.*.*.*.*" Then
    If IsSysop(Index) = False Then Exit Function
    SysListItemsBy = True
    X(Index) = Mid$(X(Index), InStr(1, X(Index), ".") + 1)
    dbIndex = GetPlayerIndexNumber(Index)
    For i = LBound(dbItems) To UBound(dbItems)
        SplitFast LCaseFast(X(Index)), Arr, "."
        With dbItems(i)
            If .sRaceRestriction <> "0" Then
                If InStr(1, Arr(0), LCaseFast(.sRaceRestriction)) = 0 Then GoTo nNext
            End If
            If .sClassRestriction <> "0" Then
                If InStr(1, Arr(1), LCaseFast(.sClassRestriction)) = 0 Then GoTo nNext
            End If
            With dbClass(GetClassID(Arr(1)))
                If modWeaponsAndArmor.GenericCanUseArmor(.iArmorType, i) = False Then GoTo nNext
                If dbItems(i).iType <> 0 Then
                    If modWeaponsAndArmor.GenericCanUseWeapon(.iWeapon, i) = False Then GoTo nNext
                End If
            End With
            If dbPlayers(dbIndex).iLevel < Val(Arr(2)) Then GoTo nNext
            If dbPlayers(dbIndex).dClassPoints < Val(Arr(3)) Then GoTo nNext
            s = s & dbItems(i).sItemName & vbCrLf
        End With
nNext:
        If DE Then DoEvents
        
    Next
    X(Index) = ""
    WrapAndSend Index, s
End If
End Function

Public Function SysListLimited(Index As Long) As Boolean
Dim s As String
Dim i As Long
If modSC.FastStringComp(LCaseFast(X(Index)), "sys list limited") Then
    If IsSysop(Index) = False Then Exit Function
    SysListLimited = True
    For i = LBound(dbItems) To UBound(dbItems)
        If dbItems(i).iLimit > 0 Then
            s = s & dbItems(i).sItemName & Space$(35 - Len(dbItems(i).sItemName)) & dbItems(i).iInGame & "/" & dbItems(i).iLimit & vbCrLf
        End If
        If DE Then DoEvents
    Next
    X(Index) = ""
    WrapAndSend Index, s
End If
End Function

Public Function SysEdit(Index As Long) As Boolean
'sys edit [player] [stat] [new value]
Dim s As String
Dim sStat As String
Dim sPlay As String
Dim sNewV As String
Dim dbIndex As Long
Dim i As Long
   On Error GoTo SysEdit_Error

If LCaseFast(Left$(X(Index), 9)) = "sys edit " Then
    If IsSysop(Index) = False Then Exit Function
    SysEdit = True
    s = ReplaceFast(LCaseFast(X(Index)), "sys edit ", "", , 1)
    i = InStr(1, s, " ")
    sPlay = Left$(s, i - 1)
    i = i + 1
    s = Mid$(s, i)
    i = InStr(1, s, " ")
    sStat = Left$(s, i - 1)
    i = i + 1
    s = Mid$(s, i)
    sNewV = s
    sPlay = SmartFind(Index, sPlay, All_Players)
    dbIndex = GetPlayerIndexNumber(, sPlay)
    If dbIndex = 0 Then
        WrapAndSend Index, RED & "There is no player by that name." & WHITE & vbCrLf
        GoTo SysEdit_Error
    End If
    s = ""
    If sStat = "" Then GoTo SysEdit_Error
    With dbPlayers(dbIndex)
        Select Case sStat
            Case "maxhp"
                If Val(sNewV) <= 2000000# Then
                    .lMaxHP = Val(sNewV)
                    If .lMaxHP < 1 Then .lMaxHP = 1
                    If .lHP > .lMaxHP Then .lHP = .lMaxHP
                    s = "max hp"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "maxma"
                If Val(sNewV) <= 2000000# Then
                    .lMaxMana = Val(sNewV)
                    If .lMaxMana < 0 Then .lMaxMana = 0
                    If .lMana > .lMaxMana Then .lMana = .lMaxMana
                    s = "max mana"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "str"
                If Val(sNewV) <= 32000# Then
                    .iStr = Val(sNewV)
                    s = "strength"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "int"
                If Val(sNewV) <= 32000# Then
                    .iInt = Val(sNewV)
                    s = "intellect"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "cha"
                If Val(sNewV) <= 32000# Then
                    .iCha = Val(sNewV)
                    s = "charm"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "dex"
                If Val(sNewV) <= 32000# Then
                    .iDex = Val(sNewV)
                    s = "dexterity"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "agil"
                If Val(sNewV) <= 32000# Then
                    .iAgil = Val(sNewV)
                    s = "agility"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "exp"
                If Val(sNewV) <= 999999999# Then
                    .dEXP = .dEXP + Val(sNewV)
                    .dTotalEXP = .dTotalEXP + Val(sNewV)
                    If .dEXP < 0 Then .dEXP = 0
                    If .dTotalEXP < 0 Then .dTotalEXP = 0
                    s = "exp has increased"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "level"
                If Val(sNewV) <= 32000# Then
                    .iLevel = Val(sNewV)
                    s = "level"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "class"
                For i = LBound(dbClass) To UBound(dbClass)
                    If LCaseFast(dbClass(i).sName) = sNewV Then
                        JustChangeClass CLng(dbIndex), dbClass(i).iID
                        s = "class"
                    ElseIf dbClass(i).iID = Val(sNewV) Then
                        JustChangeClass CLng(dbIndex), dbClass(i).iID
                        s = "class"
                    End If
                    If DE Then DoEvents
                Next
                If s = "" Then GoTo SysEdit_Error
            Case "hunger"
                If Val(sNewV) <= 32000# Then
                    .dHunger = sNewV
                    s = "hunger level"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "stamina"
                If Val(sNewV) <= 32000# Then
                    .dStamina = sNewV
                    s = "stamina level"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "sysop"
                If modMiscFlag.GetMiscFlag(dbIndex, [Can Be De-Sysed]) = 1 Then
                    WrapAndSend Index, BRIGHTRED & "ACCESS DENIED." & WHITE & vbCrLf
                    X(Index) = ""
                    Exit Function
                End If
                Select Case Val(sNewV)
                    Case 0
                        '.iIsSysop = 0
                        modMiscFlag.SetStatsPlus dbIndex, [Is A Sysop], 0
                        s = "sysop powers"
                        sNewV = BRIGHTRED & "DISABLED" & GREEN
                    Case 1
                        modMiscFlag.SetStatsPlus dbIndex, [Is A Sysop], 1
                        s = "sysop powers"
                        sNewV = LIGHTBLUE & "ENABLED" & WHITE
                    Case 2
                        modMiscFlag.SetStatsPlus dbIndex, [Is A Sysop], 1
                        modMiscFlag.SetMiscFlag dbIndex, [Can Be De-Sysed], 1
                        s = "sysop powers"
                        sNewV = LIGHTBLUE & "ENABLED " & BRIGHTRED & "PERMENTLY" & WHITE
                    Case Else
                        SendTooLarge Index
                        Exit Function
                End Select
            Case "lives"
                If Val(sNewV) > 0 And Val(sNewV) < 20 Then
                    .iLives = Val(sNewV)
                    s = "lives"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "gold"
                If Val(sNewV) < 999999999# Then
                    .dGold = Val(sNewV)
                    s = "gold"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
                
            Case "vision"
                If Val(sNewV) < 10 Then
                    .iVision = sNewV
                    s = "vision"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case "misc.invisible"
                Select Case Val(sNewV)
                    Case 0, 1
                    Case Else
                        sNewV = "0"
                End Select
                modMiscFlag.SetMiscFlag dbIndex, Invisible, Val(sNewV)
                s = "invisible"
                sNewV = IIf(sNewV = "0", "FALSE", "TRUE")
            Case "misc.seehidden"
                Select Case Val(sNewV)
                    Case 0, 1
                    Case Else
                        sNewV = "0"
                End Select
                modMiscFlag.SetMiscFlag dbIndex, [See Hidden], Val(sNewV)
                s = "see-hidden"
                sNewV = IIf(sNewV = "0", "FALSE", "TRUE")
            Case "misc.seeinvisible"
                Select Case Val(sNewV)
                    Case 0, 1
                    Case Else
                        sNewV = "0"
                End Select
                modMiscFlag.SetMiscFlag dbIndex, [See Invisible], Val(sNewV)
                s = "see-invisible"
                sNewV = IIf(sNewV = "0", "FALSE", "TRUE")
            Case "misc.gibberish"
                modMiscFlag.SetMiscFlag dbIndex, [Gibberish Talk], Val(sNewV)
                s = "gibberish talk"
            Case "misc.dualwield"
                Select Case Val(sNewV)
                    Case 0, 1
                    Case Else
                        sNewV = "0"
                End Select
                modMiscFlag.SetMiscFlag dbIndex, [Can Dual Wield], Val(sNewV)
                s = "dual wield"
                sNewV = IIf(sNewV = "0", "FALSE", "TRUE")
            Case "evil"
                If Val(sNewV) < 1001 Then
                    .iEvil = .iEvil + Val(sNewV)
                    s = "evil level"
                Else
                    SendTooLarge Index
                    Exit Function
                End If
            Case Else
                GoTo SysEdit_Error
        End Select
'        .iSC = (.iLevel + .iInt + .iAgil + .iDex) \ 3
'        If .iSC >= 100 Then .iSC = 99
'        .iMaxItems = (.iStr \ 2) + (.iDex \ 3) + (.iCha \ 4) + (.iInt \ 6) + (.iAgil \ 5)
'        If .iMaxItems > 20 Then .iMaxItems = 20
'        If .iMaxItems < 4 Then .iMaxItems = 4
'        .iMaxItems = .iMaxItems + (.iStr \ 10)
        modMiscFlag.RedoStatsPlus dbIndex
        If sStat <> "exp" Then
            s = GREEN & "Your " & YELLOW & s & GREEN & " has been changed too " & YELLOW & sNewV & WHITE & vbCrLf
        Else
            s = GREEN & "Your " & YELLOW & s & GREEN & " by " & YELLOW & sNewV & WHITE & vbCrLf
        End If
        WrapAndSend Index, ReplaceFast(s, "Your", .sPlayerName & "'s")
        WrapAndSend .iIndex, s
        X(Index) = ""
    End With
End If

   On Error GoTo 0
   Exit Function

SysEdit_Error:
    X(Index) = ""
    s = BRIGHTWHITE & "Available edits-" & vbCrLf & "maxhp, maxma, str, int, agil, dex, cha, exp"
    WrapAndSend Index, RED & "Syntax: sys edit [player] [stat] [new value]" & vbCrLf & s & WHITE & vbCrLf
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SysEdit of Module modSysopCommands"
End Function

Public Function SendTooLarge(Index As Long)
WrapAndSend Index, RED & "Value too large." & WHITE & vbCrLf
X(Index) = ""
End Function

Public Function SysDebug(Index As Long) As Boolean
If LCaseFast(X(Index)) = "sys debug" Then
    If IsSysop(Index) = False Then Exit Function
    SysDebug = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        Select Case .iDebugMode
            Case 0
                .iDebugMode = 1
                WrapAndSend Index, BRIGHTRED & "DEBUG MODE " & GREEN & "ON" & BRIGHTRED & "." & WHITE & vbCrLf
            Case 1
                .iDebugMode = 0
                WrapAndSend Index, BRIGHTRED & "DEBUG MODE " & RED & "OFF" & BRIGHTRED & "." & whtie & vbCrLf
        End Select
    End With
    X(Index) = ""
End If
End Function
'public function Horse(Index as long) As Boolean
'If lcasefast(Left$(X(Index), 10)) = "sys horse " Then
'    Horse = True
'    If IsSysop(Index) = False Then Exit Function
'    X(Index) = Replace(X(Index), "sys horse ", "")
'    Dim sUser As String, iDamage as long
'    Dim iNewStr as long, iNewAgil as long
'    Dim bFound As Boolean
'    bFound = False
'    sUser =Trimit(X(Index))
'    If sUser = "" Then
'        WrapAndSend Index, RED & "ERROR - Specify a Player." & vbCrLf & WHITE
'        X(Index) = ""
'        Exit Function
'    End If
'    With pRecords(Index).RS
'        .MoveFirst
'        Do
'            If lcasefast(sUser) = lcasefast(!PlayerName) Then
'                bFound = True
'                iDamage = RndNumber(Clng(!MaxHP) / 4, Clng(!MaxHP))
'                iNewStr = Clng(!STR) - RndNumber(1, 3)
'                If iNewStr < 0 Then iNewStr = 0
'                iNewAgil = Clng(!AGIL) - RndNumber(1, 3)
'                If iNewAgil < 0 Then iNewAgil = 0
'                .Edit
'                !HP = Clng(!HP) - iDamage
'                !STR = iNewStr
'                !AGIL = iNewAgil
'                !Horse = 1
'                .Update
'                WrapAndSend Clng(!Index), BRIGHTMAGNETA & "A large horse walks behind you..." _
'                    & vbCrLf & BRIGHTYELLOW & "You are raped in the ass for " & _
'                    iDamage & "!" & vbCrLf & BRIGHTLIGHTBLUE & _
'                    "You feel weak and slow!" & vbCrLf & BRIGHTRED & _
'                    "You strength has dropped to " & iNewStr & "!" & vbCrLf & _
'                    "You agility has droped to " & iNewAgil & "!" & vbCrLf & _
'                    BRIGHTGREEN & "You will never forget this day..." & WHITE & _
'                    vbCrLf
'
'                Exit Do
'            ElseIf Not .EOF Then
'                .MoveNext
'            End If
'            If de then doevents
'        Loop Until .EOF
'    End With
'    If Not bFound Then
'        WrapAndSend Index, RED & "The horse could not find the specified player." & vbCrLf & WHITE
'        X(Index) = ""
'        Exit Function
'    End If
'    WrapAndSend Index, BRIGHTGREEN & "The grotesque deed has been performed on the unlucky soul of " & sUser & "." & WHITE & vbCrLf
'    X(Index) = ""
'End If
'End Function

