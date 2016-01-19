Attribute VB_Name = "modMovement"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modMovement
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Function MoveMentCommands(Index As Long) As Boolean
'If BackUpLoc(Index) = True Then MoveMentCommands = True: Exit Function 'backup the location
If Search(Index) = True Then MoveMentCommands = True: Exit Function 'search the area
'If QuestSay(Index) = True Then MoveMentCommands = True: Exit Function 'check for a 'say' command
If HideItem(Index) = True Then MoveMentCommands = True: Exit Function
'If GoCommand(Index) = True Then MoveMentCommands = True: Exit Function
If Doors(Index) = True Then MoveMentCommands = True: Exit Function 'check door commands
If ShopCommands(Index) = True Then MoveMentCommands = True: Exit Function 'check for shopping commands(buy,sell,list)
If Movement(Index) = True Then MoveMentCommands = True: Exit Function 'check for movement commands
MoveMentCommands = False
End Function
'Public Function BackUpLoc(Index As Long) As Boolean
''This function is to back up the players location
''*********************************************************
''***             Reason for this Function              ***
''*** If the server crashes in the middle of movement   ***
''*** or something is missed in the middle of movement  ***
''*** The player's location becomes !0!, and is stuck   ***
''*** in a "void".  This allows that player to come out ***
''*** of that "void" and get back into the game         ***
''*********************************************************
'If GetPlayerIndexNumber(Index) = 0 Then Exit Function
'With dbPlayers(GetPlayerIndexNumber(Index))
'    If .lLocation = 0 Then
'        .lLocation = .lBackUpLoc
'    Else
'        Exit Function
'    End If
'    BackUpLoc = True 'make the function true
'    X(Index) = ""
'    'send error message to them
'    WrapAndSend Index, BRIGHTRED & "[INTERNAL SERVER ERROR: MOVEMENT, Inform Sysop]" & WHITE & vbCrLf
'    'record error to log or list
'    UpdateList "A movement error occured on line " & Index & ". }b(}n}i" & Time & "}n}b)"
'End With
'End Function

Sub SendRoomError(Index As Long)
With dbPlayers(GetPlayerIndexNumber(Index))
    If .iSneaking <> 0 Then
        .iSneaking = 0
        WrapAndSend Index, BRIGHTRED & "You are no longer sneaking around." & WHITE, False
    End If
End With
WrapAndSend Index, RED & "If you were paying attention, you would notice that the door is closed there." & WHITE & vbCrLf
X(Index) = ""
End Sub

Public Function Movement(Index As Long, Optional LeaderHasMoved As Boolean = False) As Boolean
'////////MOVEMENT////////
'function for moving on the map
'Lots of movement commands...
'*************************************
'***********/////////////////////////*
'**********/////*(N)orth********////**
'*********/////*(S)outh********////***
'********/////*(E)ast*********////****
'*******/////*(W)est*********////*****
'******/////*(U)p***********////******
'*****/////*(D)own*********////*******
'****/////*(NW)Northwest**////********
'***/////*(NE)Northeast**////*********
'**/////*(SW)Southwest**////**********
'*/////*(SE)Southeast**////***********
'/////////////////////////************
'*************************************
Dim sField As String, ToSend$
Dim NewPP As Long
Dim OldPP As Long
Dim dbIndex As Long
Dim lMax As Long
Dim iChance As Long
Dim i As Long
Dim sOth As String
Dim s As String
Dim Arr() As String
Dim dbMapId As Long
Dim sS As String
If IsADirection(X(Index)) = True Then
    Movement = True
    dbIndex = GetPlayerIndexNumber(Index)
    'If modSC.FastStringComp(TempPP, "0") Then BackUpLoc Index
    If modGetData.GetPlayersTotalItems(Index, dbIndex) > modGetData.GetPlayersMaxItems(Index, dbIndex) Then
        WrapAndSend Index, RED & "You are carring so much, you can't even move!" & WHITE & vbCrLf
        RemoveFromParty Index
        X(Index) = ""
        Exit Function
    End If
    dbMapId = dbPlayers(dbIndex).lDBLocation
    With dbMap(dbMapId)
        If Len(X(Index)) > 1 Then 'if dir is NW,NE,SW,SE
            Select Case LCaseFast(Left$(X(Index), 2))
                Case "nw": 'northwest
                    sField = "NorthWest"
                    sS = " to the "
                    If (.lDNW = 1 Or .lDNW = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lNorthWest
                Case "ne": 'northeast
                    sField = "NorthEast"
                    sS = " to the "
                    If (.lDNE = 1 Or .lDNE = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lNorthEast
                Case "sw": 'southwest
                    sField = "SouthWest"
                    sS = " to the "
                    If (.lDSW = 1 Or .lDSW = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lSouthWest
                Case "se" 'southeast
                    sField = "SouthEast"
                    sS = " to the "
                    If (.lDSE = 1 Or .lDSE = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lSouthEast
            End Select
        End If
        If modSC.FastStringComp(sField, "") Then
            Select Case LCaseFast(Left$(X(Index), 1)) 'normal directions
                Case "n": 'north
                    sField = "North"
                    sS = " to the "
                    If (.lDN = 1 Or .lDN = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lNorth
                Case "s": 'south
                    sField = "South"
                    sS = " to the "
                    If (.lDS = 1 Or .lDS = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lSouth
                Case "e": 'east
                    sField = "East"
                    sS = " to the "
                    If (.lDE = 1 Or .lDE = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lEast
                Case "w": 'west
                    sField = "West"
                    sS = " to the "
                    If (.lDW = 1 Or .lDW = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lWest
                Case "u": 'up
                    sField = "Upwards"
                    sS = ""
                    If (.lDU = 1 Or .lDU = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lUp
                Case "d": 'down
                    sField = "Downwards"
                    sS = ""
                    If (.lDD = 1 Or .lDD = 2) And dbPlayers(dbIndex).iGhostMode = 0 Then SendRoomError Index: Exit Function
                    NewPP = .lDown
            End Select
        End If
    End With
    With dbPlayers(dbIndex)
    
        If LeaderHasMoved = False And _
           modSC.FastStringComp(CStr(.iLeadingParty), "0") And _
           Not modSC.FastStringComp(.sParty, "0") Then
           
                RemoveFromParty Index
        End If
        
        X(Index) = ""
        
        If NewPP <> 0 Then
            OldPP = .lLocation
            .lLocation = NewPP
            .iResting = 0
            .iMeditating = 0
            .lRoomSearched = -1
            If RndNumber(0, 1) = 1 Then .dHunger = .dHunger - RndNumber(0, 1)
            If .iHorse > 0 Then
                If sS = "" Then sS = " "
                ToSend$ = LIGHTBLUE & "You ride your " & .sFamName & sS & LCaseFast(sField) & "." & WHITE & vbCrLf
            Else
                If RndNumber(0, 1) = 1 Then .dStamina = .dStamina - RndNumber(0, 1)
                If sS <> "" Then sS = Mid$(sS, 2)
                ToSend$ = LIGHTBLUE & "You walk " & sS & LCaseFast(sField) & "." & vbCrLf & WHITE
            End If
            If Not modSC.FastStringComp(.sFamName, "0") And .iHorse < 1 Then ToSend$ = ToSend$ & LIGHTBLUE & "Your " & .sFamName & " follows you." & WHITE & vbCrLf
            If .iHorse < 1 Then SendDelayMessage dbIndex
            ToSend$ = ToSend$ & modGetData.GetRoomDescription(dbIndex, .lLocation, , True)
            WaitFor modGetData.GetMoveDelay(dbIndex)
            If .iSneaking <> 0 Then
                lMax = modMiscFlag.GetStatsPlusTotal(dbIndex, Steath) + RndNumber(1, 40)
                If lMax > 96 Then lMax = 96
                iChance = RndNumber(1, 100)
                If iChance <= lMax Then
                    ToSend$ = ToSend$ & "Sneaking..." & vbCrLf & WHITE
                Else
                    Select Case RndNumber(0, 4)
                        Case 0
                            ToSend$ = ToSend$ & BRIGHTRED & "You stumble over your own feet!" & WHITE & vbCrLf
                            sOth = LIGHTBLUE & .sPlayerName & " stumbles over their own feet." & WHITE & vbCrLf
                        Case 1
                            ToSend$ = ToSend$ & BRIGHTRED & "You sneeze!" & WHITE & vbCrLf
                            sOth = LIGHTBLUE & .sPlayerName & " sneezes." & WHITE & vbCrLf
                        Case 2
                            ToSend$ = ToSend$ & BRIGHTRED & "You notice something looking at you!" & WHITE & vbCrLf
                            sOth = LIGHTBLUE & .sPlayerName & " looks around as if something is watching them." & WHITE & vbCrLf
                        Case 3
                            If RndNumber(0, 1) = 0 Then
                                If .sWeapon <> "0" Then
                                    If .iAgil < 20 Then
                                        ToSend$ = ToSend$ & BRIGHTRED & "You drop your weapon!" & WHITE & vbCrLf
                                        s = .sWeapon
                                        modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), modItemManip.GetItemIDFromUnFormattedString(s)
                                        modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), modItemManip.GetItemIDFromUnFormattedString(s)
                                        sOth = LIGHTBLUE & .sPlayerName & "'s weapon falls out of their hands." & WHITE & vbCrLf
                                    Else
                                        ToSend$ = ToSend$ & BRIGHTRED & "You almost drop your weapon!" & WHITE & vbCrLf
                                        sOth = LIGHTBLUE & .sPlayerName & " fumbles with their weapon." & WHITE & vbCrLf
                                    End If
                                Else
                                    ToSend$ = ToSend$ & BRIGHTRED & "You hiccup!" & WHITE & vbCrLf
                                    sOth = LIGHTBLUE & .sPlayerName & " hiccups." & WHITE & vbCrLf
                                End If
                            Else
                                ToSend$ = ToSend$ & BRIGHTRED & "You get a chill up your spine!" & WHITE & vbCrLf
                                sOth = LIGHTBLUE & .sPlayerName & " shivers." & WHITE & vbCrLf
                            End If
                        Case 4
                            ToSend$ = ToSend$ & BRIGHTRED & "You cough!" & WHITE & vbCrLf
                            sOth = LIGHTBLUE & .sPlayerName & " coughs." & WHITE & vbCrLf
                    End Select
                    .iSneaking = 0
                End If
            End If
            WrapAndSend Index, ToSend$
            GenAMonster NewPP, , , , dbPlayers(dbIndex).lDBLocation
            If modMiscFlag.GetMiscFlag(dbIndex, Invisible) = 0 Then
                If iChance > lMax Or .iSneaking = 0 Then
                    If .iHorse > 0 Then
                        SendToAllInRoom Index, sOth & LIGHTBLUE & .sPlayerName & " rides " & modGetData.GetGenderPronoun(dbIndex, True) & " " & .sFamName & sS & LCaseFast(sField) & "." & vbCrLf & WHITE, OldPP
                        If sS <> "" And sS <> " " Then sS = "the "
                        If sS = " " Then sS = ""
                        SendToAllInRoom Index, sOth & LIGHTBLUE & .sPlayerName & " just arrived from " & sS & modGetData.GetOppositeDirection(sField, True) & " on " & modGetData.GetGenderPronoun(dbIndex, True) & " " & .sFamName & "." & vbCrLf & WHITE, NewPP
                        MoveMessages dbIndex
                    Else
                        If sS = "" Then sS = " "
                        SendToAllInRoom Index, sOth & LIGHTBLUE & .sPlayerName & " just left " & sS & LCaseFast(sField) & "." & vbCrLf & WHITE, OldPP
                        If sS <> "" Then sS = "the "
                        SendToAllInRoom Index, sOth & LIGHTBLUE & .sPlayerName & " just arrived from " & modGetData.GetOppositeDirection(sField, True) & "." & vbCrLf & WHITE, NewPP
                        MoveMessages dbIndex
                    End If
                End If
            Else
                s = modGetData.GetPlayersDBIndexesHere(.lLocation)
                SplitFast s, Arr, ";"
                For i = LBound(Arr) To UBound(Arr)
                    If Arr(i) <> "" And Arr(i) <> "0" Then
                        With dbPlayers(CLng(Arr(i)))
                            If CLng(Arr(i)) <> dbIndex Then
                                If modMiscFlag.GetMiscFlag(CLng(Arr(i)), [See Invisible]) = 1 Or modMiscFlag.GetMiscFlag(CLng(Arr(i)), [See Hidden]) = 1 Then
                                    If dbPlayers(dbIndex).iHorse > 0 Then
                                        If sS <> "" Then sS = "the "
                                        WrapAndSend .iIndex, sOth & LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " just arrived from " & modGetData.GetOppositeDirection(sField, True) & " on " & modGetData.GetGenderPronoun(dbIndex, True) & " " & dbPlayers(dbIndex).sFamName & "." & vbCrLf & WHITE
                                    Else
                                        If sS <> "" Then sS = "the "
                                        WrapAndSend .iIndex, sOth & LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " just arrived from " & modGetData.GetOppositeDirection(sField, True) & "." & vbCrLf & WHITE
                                    End If
                                ElseIf (RndNumber(1, 100) < modMiscFlag.GetStatsPlusTotal(CLng(Arr(i)), Perception)) Then
                                    If sS <> "" Then sS = "the "
                                    WrapAndSend .iIndex, sOth & LIGHTBLUE & "You hear movement coming from " & modGetData.GetOppositeDirection(sField, True) & "." & WHITE & vbCrLf
                                End If
                            End If
                        End With
                    End If
                    If DE Then DoEvents
                Next
                s = ""
                s = modGetData.GetPlayersDBIndexesHere(OldPP)
                SplitFast s, Arr, ";"
                For i = LBound(Arr) To UBound(Arr)
                    If Arr(i) <> "" And Arr(i) <> "0" Then
                        With dbPlayers(CLng(Arr(i)))
                            If CLng(Arr(i)) <> dbIndex Then
                                If modMiscFlag.GetMiscFlag(CLng(Arr(i)), [See Invisible]) = 1 Or modMiscFlag.GetMiscFlag(CLng(Arr(i)), [See Hidden]) = 1 Then
                                    If dbPlayers(dbIndex).iHorse > 0 Then
                                        If sS = "" Then sS = " "
                                        WrapAndSend .iIndex, sOth & LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " rides " & modGetData.GetGenderPronoun(dbIndex, True) & " " & dbPlayers(dbIndex).sFamName & sS & LCaseFast(sField) & "." & vbCrLf & WHITE
                                    Else
                                        If sS = "" Then sS = " "
                                        WrapAndSend .iIndex, sOth & LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " just left" & sS & LCaseFast(sField) & "." & vbCrLf & WHITE
                                    End If
                                ElseIf RndNumber(1, 100) < modMiscFlag.GetStatsPlusTotal(CLng(Arr(i)), Perception) Then
                                    If sS = "" Then sS = " "
                                    WrapAndSend .iIndex, sOth & LIGHTBLUE & "You hear movement going away" & sS & LCaseFast(sField) & "." & WHITE & vbCrLf
                                End If
                            End If
                        End With
                    End If
                    If DE Then DoEvents
                Next
            End If
            PartyMovement dbIndex, modGetData.GetShortDir(sField)
            X(Index) = ""
        Else
            WrapAndSend Index, RED & "You walk right into a wall." & vbCrLf & WHITE
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " walks straight into a wall." & WHITE & vbCrLf, .lLocation
            If .iSneaking <> 0 Then
                .iSneaking = 0
                WrapAndSend Index, BRIGHTRED & "You are no longer sneaking around." & WHITE, False
            End If
            X(Index) = ""
            Exit Function
        End If
        If RndNumber(1, 100) > 68 Then GenAMonster .lLocation, , , , .lDBLocation
        With dbMap(dbMapId)
            If .lNorth <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lNorth
            If .lSouth <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lSouth
            If .lEast <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lEast
            If .lWest <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lWest
            If .lNorthEast <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lNorthEast
            If .lNorthWest <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lNorthWest
            If .lSouthEast <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lSouthEast
            If .lSouthWest <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lSouthWest
            If .lUp <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lUp
            If .lDown <> 0 And RndNumber(1, 100) > 68 Then GenAMonster .lDown
        End With
    End With
End If
'////////END////////
End Function

Sub PartyMovement(dbIndex As Long, Direction As String)
'sub for haveing the party follow the leader when they move
Dim Party As String
Dim sLeadingParty As String
Dim tArr() As String 'tempary array
With dbPlayers(dbIndex)
    sLeadingParty = CStr(.iLeadingParty)
    Party = .sParty
End With
If Not modSC.FastStringComp(sLeadingParty, "1") Then Exit Sub  'make sure they are leading the party
If modSC.FastStringComp(Party, "0") Then Exit Sub 'if they don't have a party, exit the sub
Party = ReplaceFast(Party, ":", "")
If DCount(Party, ";") > 1 Then 'get the amount of people
    SplitFast Party, tArr, ";"
Else
    ReDim tArr(0) As String
    tArr(0) = Left$(Party, Len(Party) - 1)
End If
For i = 0 To UBound(tArr) 'loop the array
    If Not modSC.FastStringComp(tArr(i), "") Then
        MovePartyMember CLng(tArr(i)), Direction
    End If
Next
End Sub

Sub MovePartyMember(Index As Long, Direction As String)
Dim sField As String, ToSend$
Dim OldPP As Long, NewPP As Long
Dim dbIndex As Long
Dim lMax As Long
Dim iChance As Long
Dim sOth As String
dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    OldPP = .lLocation
End With
'If modSC.FastStringComp(TempPP, "0") Then BackUpLoc Index
If modGetData.GetPlayersTotalItems(Index, dbIndex) > modGetData.GetPlayersMaxItems(Index, dbIndex) Then
    WrapAndSend Index, RED & "You are carring so much, you can't even move!" & WHITE & vbCrLf
    RemoveFromParty Index
    X(Index) = ""
    Exit Sub
End If
With dbMap(dbPlayers(dbIndex).lDBLocation)
    If Len(Direction) > 1 Then 'if dir is NW,NE,SW,SE
        Select Case LCaseFast(Left$(Direction, 2))
            Case "nw": 'northwest
                sField = "NorthWest"
                If .lDNW = 1 Or .lDNW = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lNorthWest
            Case "ne": 'northeast
                sField = "NorthEast"
                If .lDNE = 1 Or .lDNE = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lNorthEast
            Case "sw": 'southwest
                sField = "SouthWest"
                If .lDSW = 1 Or .lDSW = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lSouthWest
            Case "se" 'southeast
                sField = "SouthEast"
                If .lDSW = 1 Or .lDSW = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lSouthEast
        End Select
    End If
    If modSC.FastStringComp(sField, "") Then
        Select Case LCaseFast(Left$(Direction, 1)) 'normal directions
            Case "n": 'north
                sField = "North"
                If .lDN = 1 Or .lDN = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lNorth
            Case "s": 'south
                sField = "South"
                If .lDS = 1 Or .lDS = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lSouth
            Case "e": 'east
                sField = "East"
                If .lDE = 1 Or .lDE = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lEast
            Case "w": 'west
                sField = "West"
                If .lDW = 1 Or .lDW = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lWest
            Case "u": 'up
                sField = "Up"
                If .lDU = 1 Or .lDU = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lUp
            Case "d": 'down
                sField = "Down"
                If .lDD = 1 Or .lDD = 2 Then SendRoomError Index: Exit Sub
                NewPP = .lDown
        End Select
    End If
End With
With dbPlayers(dbIndex)

    'If LeaderHasMoved = False And modSC.FastStringComp(CStr(.iLeadingParty), "0") And Not modSC.FastStringComp(.sParty, "0") Then RemoveFromParty Index
    'X(Index) = ""
    
    If Not modSC.FastStringComp(CStr(NewPP), "0") Then
        .lLocation = NewPP
        .iResting = 0
        .iMeditating = 0
        .lRoomSearched = -1
        .dStamina = .dStamina - RndNumber(0, 1)
        .dHunger = .dHunger - RndNumber(0, 1)
        ToSend$ = BRIGHTWHITE & "Following your party leader to the " & sField & "." & whtie & vbCrLf
        If Not modSC.FastStringComp(.sFamName, "0") And .iHorse = 0 Then ToSend$ = ToSend$ & LIGHTBLUE & "Your " & .sFamName & " follows you." & WHITE & vbCrLf
        ToSend$ = ToSend$ & modGetData.GetRoomDescription(CLng(dbIndex), .lLocation, , True)
        If .iSneaking <> 0 Then
            lMax = .iAgil + (.iDex \ 2) + (.iCha \ 3)
            If lMax > 96 Then lMax = 96
            iChance = RndNumber(1, 100)
            If iChance <= lMax Then
                ToSend$ = ToSend$ & "Sneaking..." & vbCrLf & WHITE
            Else
                Select Case RndNumber(0, 4)
                    Case 0
                        ToSend$ = ToSend$ & BRIGHTRED & "You stumble over your own feet!" & WHITE & vbCrLf
                        sOth = LIGHTBLUE & .sPlayerName & " stumbles over their own feet." & WHITE & vbCrLf
                    Case 1
                        ToSend$ = ToSend$ & BRIGHTRED & "You sneeze!" & WHITE & vbCrLf
                        sOth = LIGHTBLUE & .sPlayerName & " sneezes." & WHITE & vbCrLf
                    Case 2
                        ToSend$ = ToSend$ & BRIGHTRED & "You notice something looking at you!" & WHITE & vbCrLf
                        sOth = LIGHTBLUE & .sPlayerName & " looks around as if something is watching them." & WHITE & vbCrLf
                    Case 3
                        If RndNumber(0, 1) = 0 Then
                            If .sWeapon <> "0" Then
                                If .iAgil < 20 Then
                                    ToSend$ = ToSend$ & BRIGHTRED & "You drop your weapon!" & WHITE & vbCrLf
                                    modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), modItemManip.GetItemIDFromUnFormattedString(.sWeapon)
                                    sOth = LIGHTBLUE & .sPlayerName & "'s weapon falls out of their hands." & WHITE & vbCrLf
                                Else
                                    ToSend$ = ToSend$ & BRIGHTRED & "You almost drop your weapon!" & WHITE & vbCrLf
                                    sOth = LIGHTBLUE & .sPlayerName & " fumbles with their weapon." & WHITE & vbCrLf
                                End If
                            Else
                                ToSend$ = ToSend$ & BRIGHTRED & "You hiccup!" & WHITE & vbCrLf
                                sOth = LIGHTBLUE & .sPlayerName & " hiccups." & WHITE & vbCrLf
                            End If
                        Else
                            ToSend$ = ToSend$ & BRIGHTRED & "You get a chill up your spine!" & WHITE & vbCrLf
                            sOth = LIGHTBLUE & .sPlayerName & " shivers." & WHITE & vbCrLf
                        End If
                    Case 4
                        ToSend$ = ToSend$ & BRIGHTRED & "You cough!" & WHITE & vbCrLf
                        sOth = LIGHTBLUE & .sPlayerName & " coughs." & WHITE & vbCrLf
                End Select
                .iSneaking = 0
            End If
        End If
        WrapAndSend Index, ToSend$
        If iChance > lMax And .iSneaking = 0 Then
            SendToAllInRoom Index, sOth & LIGHTBLUE & .sPlayerName & " just left to the " & LCaseFast(sField) & "." & vbCrLf & WHITE, OldPP
            SendToAllInRoom Index, sOth & LIGHTBLUE & .sPlayerName & " just arrived from the " & modGetData.GetOppositeDirection(sField) & "." & vbCrLf & WHITE, NewPP
        End If
        'X(Index) = ""
    Else
        If .iSneaking <> 0 Then
            .iSneaking = 0
            WrapAndSend Index, BRIGHTRED & "You are no longer sneaking around." & WHITE, False
        End If
        WrapAndSend Index, RED & "You walk right into a wall." & vbCrLf & WHITE
        SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " walks straight into a wall." & WHITE & vbCrLf, .lLocation
        'X(Index) = ""
        Exit Sub
    End If
End With
End Sub

Sub SendDelayMessage(dbIndex As Long)
Select Case dbPlayers(dbIndex).dStamina
    Case Is <= 0
        WrapAndSend dbPlayers(dbIndex).iIndex, BRIGHTRED & "You slowly get up and move." & WHITE & vbCrLf, False
    Case Is < 30
        WrapAndSend dbPlayers(dbIndex).iIndex, RED & "You eventually get up and move." & WHITE & vbCrLf, False
    
    Case Is < 50
        WrapAndSend dbPlayers(dbIndex).iIndex, BRIGHTYELLOW & "You sluggishly move on." & WHITE & vbCrLf, False
    
End Select
End Sub

Public Sub MoveMessages(dbIndex As Long)
Dim dbMapId As Long
Dim ToSend As String
With dbPlayers(dbIndex)
    dbMapId = GetMapIndex(.lLocation)
End With
With dbMap(dbMapId)
    If .lNorth <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement to the south." & WHITE & vbCrLf, .lNorth
    If .lSouth <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement to the north." & WHITE & vbCrLf, .lSouth
    If .lEast <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement to the west." & WHITE & vbCrLf, .lEast
    If .lWest <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement to the east." & WHITE & vbCrLf, .lWest
    If .lUp <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement below you." & WHITE & vbCrLf, .lUp
    If .lDown <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement above you." & WHITE & vbCrLf, .lDown
    If .lNorthWest <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement to the southeast." & WHITE & vbCrLf, .lNorthWest
    If .lNorthEast <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement to the southwest." & WHITE & vbCrLf, .lNorthEast
    If .lSouthEast <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement to the northwest." & WHITE & vbCrLf, .lSouthEast
    If .lSouthWest <> 0 Then SendToAllInRoom 0, LIGHTBLUE & "You hear movement to the northeast." & WHITE & vbCrLf, .lSouthWest
End With
End Sub
