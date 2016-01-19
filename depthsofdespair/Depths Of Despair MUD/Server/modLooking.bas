Attribute VB_Name = "modLooking"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modLooking
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function IsMonsterHostile(sMon As String) As Boolean
IsMonsterHostile = False
With dbMonsters(GetMonsterID(sMon))
    If .iHostile = 1 Then IsMonsterHostile = True
End With
End Function

Public Function LookAround(Index As Long, Optional CustomPP As Long = 0) As Boolean
'////////LOOK AROUND////////
'function to look at things (room, players, monsters, items, familiars, etc)
Dim TempItemHere As String, TempLoc As String, ToSend$
Dim TempPeeps As String, TempInv As String, TempGold$
Dim Mons$, bFound As Boolean
Dim MonsterName As String, TempPName As String
Dim dbMapId As Long
Dim tItem$
Dim s As String
Dim t As String
Dim ByRefString As String
Dim dbIndex As Long
Dim iIn As Long
Dim i As Long
Dim MonID As Long
Dim Arr() As String
Dim sTemp As String
'check the syntax
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 1)), "l") And Not modSC.FastStringComp(LCaseFast(Left$(X(Index), 4)), "list") Then
    X(Index) = TrimIt(X(Index))
    If Len(X(Index)) > 1 Then
        If Not modSC.FastStringComp(LCaseFast(Mid$(X(Index), 2, 1)), " ") And Not LCaseFast(Mid$(X(Index), 2, 1)) Like "[o]" Then
            Exit Function
        ElseIf Len(X(Index)) > 2 Then
            If Not modSC.FastStringComp(LCaseFast(Mid$(X(Index), 2, 1)), " ") Then
                If Not LCaseFast(Mid$(X(Index), 3, 1)) Like "[o]" Then
                    Exit Function
                End If
            End If
        End If
    End If
    LookAround = True
    dbIndex = GetPlayerIndexNumber(Index)
    If CustomPP = 0 Then 'check the optional parameter
        dbMapId = dbPlayers(dbIndex).lDBLocation
    Else
        dbMapId = GetMapIndex(CustomPP)
    End If
    With dbPlayers(dbIndex)
        If Not modSC.FastStringComp(.sInventory, "0") Then TempInv$ = TempInv$ & GREEN & modGetData.GetPlayersInvFromNums(Index, , dbIndex)
        'get all the equipment they are wearing
        TempInv$ = TempInv$ & modGetData.GetPlayersEqFromNums(Index, , dbIndex)
        'get whoever is in the room
        TempPeeps = modGetData.GetPlayersHere(.lLocation, dbIndex)
    End With
    'add the familiars to that
    'TempPeeps = TempPeeps & modgetdata.GetFamiliarsHere(CLng(TempLoc))
    'get what gold is in the room
    TempGold$ = modGetData.GetGoldHere(dbPlayers(dbIndex).lLocation, dbMapId)
    'get the items in the room
    TempItemHere = modGetData.GetRoomItemsFromNums(0, True, True, 0, dbMapId)
    'get the monsters in the room
    Mons$ = modGetData.GetMonsHere(dbPlayers(dbIndex).lLocation, True, dbIndex, dbMapId)
    If modSC.FastStringComp(TempItemHere, "") Then  'if there are no items
        'if there is some gold, set the message
        If Not modSC.FastStringComp(TempGold$, "") Then TempItemHere = Left$(TempGold$, Len(TempGold$) - 1)
    Else
        'set the message of gold and items
        TempItemHere = TempGold$ & TempItemHere
    End If
    'format the items
    TempItemHere = ReplaceFast(TempItemHere, ";", YELLOW & ", " & GREEN)
    'start the message with thr room title and description
    ToSend$ = BRIGHTLIGHTBLUE & modGetData.GetRoomTitle(dbMap(dbMapId).lRoomID, dbMapId) & _
              vbCrLf & BRIGHTWHITE & modGetData.GetRoomDesc(dbMap(dbMapId).lRoomID, dbMapId) & vbCrLf
    '********************************************
    'Set the messages for monsters and people here
    
    If Not modSC.FastStringComp(TempPeeps, "") And Not modSC.FastStringComp(Mons$, "") Then
        TempPeeps = TempPeeps & Mons$
    ElseIf modSC.FastStringComp(TempPeeps, "") And Not modSC.FastStringComp(Mons$, "") Then
        TempPeeps = Mons$
    ElseIf Not modSC.FastStringComp(TempPeeps, "") And modSC.FastStringComp(Mons$, "") Then
        TempPeeps = Left$(TempPeeps, Len(TempPeeps) - 2)
    ElseIf modSC.FastStringComp(TempPeeps, "") And modSC.FastStringComp(Mons$, "") Then
        TempPeeps = ""
    End If
    
    '*******************************************
    'make sure there is someone there before adding a Also here: tag
    If Not modSC.FastStringComp(TempPeeps, "") Then ToSend$ = ToSend$ & MAGNETA & "Also here: " & TempPeeps
    If Not modSC.FastStringComp(TempItemHere, "") Then 'see if there are items
        'set the items message with room exits
        ToSend$ = ToSend$ & vbCrLf & YELLOW & "You notice " _
            & GREEN & TempItemHere & YELLOW & " here." & vbCrLf _
            & BRIGHTMAGNETA & "Visible Exits: " & GREEN & _
            modGetData.GetRoomExits(dbMap(dbMapId).lRoomID, dbMapId) _
            & vbCrLf & WHITE
    Else
        'just send the exits
        ToSend$ = ToSend$ & vbCrLf _
            & BRIGHTMAGNETA & "Visible Exits: " & GREEN & _
            modGetData.GetRoomExits(dbMap(dbMapId).lRoomID, dbMapId) _
            & vbCrLf & WHITE
    End If
    'now checking what they are looking at
    If InStr(1, X(Index), " ") = 0 Then 'if no target, show the room
        'set the message
        If Not modSC.FastStringComp(TempPeeps, "") Then TempPeeps = Left$(TempPeeps, Len(TempPeeps) - 1) & "."
        'send the message
        With dbPlayers(dbIndex)
            Select Case .iVision + modGetData.GetRoomLight(dbMap(dbMapId).lRoomID, dbMapId)
                Case Is < -3
                    ToSend$ = WHITE & "This room is too dark, you can't see a thing!" & WHITE & vbCrLf
                Case -3 To -1
                    ToSend$ = ToSend$ & WHITE & "This room is barely visible." & WHITE & vbCrLf
                Case 0 To 2
                    ToSend$ = ToSend$ & WHITE & "This room has little light in it." & WHITE & vbCrLf
            End Select
        End With
        WrapAndSend Index, ToSend$
        X(Index) = ""
        Exit Function
    Else
        X(Index) = LCaseFast(Mid$(X(Index), InStr(1, X(Index), " ") + 1))
        TempPeeps = TempPeeps & ", " & dbPlayers(dbIndex).sPlayerName
        If LookDirs(Index) = True Then
            Exit Function
        ElseIf InStr(1, LCaseFast(Mons$), SmartFind(Index, X(Index), Monster_In_Room)) Or _
               InStr(1, LCaseFast(modGetData.GetFamiliarsHere(dbPlayers(dbIndex).lLocation)), _
                        SmartFind(Index, X(Index), Monster_In_Room)) Then
            MonsterName = SmartFind(Index, LCaseFast(X(Index)), Monster_In_Room)
            If InStr(1, MonsterName, Chr$(0)) > 0 Then MonsterName = Mid$(MonsterName, InStr(1, MonsterName, Chr$(0)) + 1)
            bFound = False 'flag
            sTemp = modGetData.GetAllMonstersInRoom(dbPlayers(dbIndex).lLocation, dbPlayers(dbIndex).lDBLocation)
            If sTemp <> "" Then
                SplitFast sTemp, Arr, ";"
                For i = LBound(Arr) To UBound(Arr)
                    If Arr(i) <> "" Then
                        If modSC.FastStringComp(LCaseFast(aMons(Val(Arr(i))).mName), MonsterName) Then
                            With aMons(Val(Arr(i)))
                                bFound = True
                                ToSend$ = BRIGHTLIGHTBLUE & .mName & vbCrLf
                                ToSend$ = ToSend$ & GREEN & .mDesc & vbCrLf
                                ToSend$ = ToSend$ & BRIGHTRED & .mName & modGetData.GetMonDamDesc(.mHP, .mMaxHP) & WHITE & vbCrLf
                                WrapAndSend Index, ToSend$
                                X(Index) = ""
                            End With
                            Exit For
                        End If
                    End If
                    If DE Then DoEvents
                Next
            End If
            If bFound = True Then Exit Function 'if they were looking at a monster, exit the sub
            bFound = False
            MonID = CStr(GetFamID(, MonsterName))
            If MonID <> 0 Then
                bFound = True
                With dbFamiliars(MonID)
                    'if it finds the familiar,
                    'set the message
                    ToSend$ = BRIGHTLIGHTBLUE & .sFamName & vbCrLf
                    ToSend$ = ToSend$ & GREEN & .sDescription & vbCrLf & WHITE
                    'send the message
                    WrapAndSend Index, ToSend$
                    X(Index) = ""
                End With
            End If
            If Not bFound Then
                'if it didn't find a monster or familiar, then send an error message
                WrapAndSend Index, RED & "You can't seem to find the " & X(Index) & " anywhere." & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
            End If
        'check the people at the current room
        ElseIf InStr(1, LCaseFast(TempPeeps), SmartFind(Index, X(Index), Player_In_Room)) And Not modSC.FastStringComp(TempPeeps, "") Then
            'get the name
            TempPName = SmartFind(Index, LCaseFast(TrimIt(X(Index))), Player_In_Room)
            TempInv$ = "" 'clear the inventory variable
            If modSC.FastStringComp(TempPName, "") Then  'if no name
                'they must specify a player
                WrapAndSend Index, RED & "Don't be silly, you can't look at nothing." & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
            End If
            bFound = False
            For i = LBound(dbPlayers) To UBound(dbPlayers)
                With dbPlayers(i)
                    If modSC.FastStringComp(LCaseFast(.sPlayerName), LCaseFast(TempPName)) Then
                        If .lLocation = dbPlayers(dbIndex).lLocation Then
                            bFound = True
                            If modSC.FastStringComp(.sOverrideDesc, "0") Then
                                ToSend$ = .sPlayerName & " has equiped upon themselves: " & vbCrLf & GREEN
                                ToSend$ = ToSend$ & modGetData.GetPlayersEqFromNums(.iIndex, , i)
                                ToSend$ = BRIGHTBLUE & ToSend$
                                ToSend$ = ReplaceFast(ToSend$, ",", vbCrLf & GREEN)
                                ToSend$ = ReplaceFast(ToSend$, "(", YELLOW & "(")
                                ToSend$ = BRIGHTWHITE & modDesc.sDescription(.iIndex) & vbCrLf & ToSend$
                            Else
                                ToSend$ = .sOverrideDesc
                            End If
                            'If .iHorse = 1 Then ToSend$ = ToSend$ & BRIGHTYELLOW & .sPlayerName & " appears to have been raped by a horse..." & vbCrLf
                            If .iHorse > 0 Then ToSend$ = ToSend$ & BRIGHTYELLOW & .sPlayerName & " is riding their " & .sFamName & "." & vbCrLf
                            ToSend$ = ToSend$ & WHITE
                            WrapAndSend .iIndex, BRIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " looked you over." & WHITE & vbCrLf
                            WrapAndSend Index, BRIGHTBLUE & "You look " & .sPlayerName & " over..." & WHITE & vbCrLf & ToSend$ & vbCrLf
                            X(Index) = ""
                            Exit Function
                        End If
                    End If
                End With
                If DE Then DoEvents
            Next
            If Not bFound Then
                'send error message
                WrapAndSend Index, RED & "You can't seem to locate " & TempPName & "." & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
            End If
        'search through the inventory
        
        ElseIf InStr(1, modGetData.GetPlayersEqFromNums(Index, , dbIndex), SmartFind(Index, X(Index), Equiped_Item)) Then
            'get the item name
            tItem$ = SmartFind(Index, LCaseFast(X(Index)), Equiped_Item, True, ByRefString)
            If InStr(1, tItem$, Chr$(0)) > 0 Then tItem$ = Mid$(tItem$, InStr(1, tItem$, Chr$(0)) + 1)
            iIn = GetItemID(tItem)
            If iIn <> 0 Then
                With dbItems(iIn)
                    'start the message
                    ToSend$ = LIGHTBLUE & "You look at the " & .sItemName & vbCrLf & BRIGHTLIGHTBLUE & .sItemName & vbCrLf
                    'finish the message
                    ToSend$ = ToSend$ & GREEN & .sDesc & vbCrLf
                    If modItemManip.GetItemBulletsID(ByRefString) <> 0 Then
                        With dbItems(GetItemID(, modItemManip.GetItemBulletsID(ByRefString)))
                            s = .sItemName
                            t = .iUses
                        End With
                        ToSend$ = ToSend$ & GREEN & "AMMO: " & LIGHTBLUE & s & GREEN & " (" & LIGHTBLUE & modItemManip.GetItemBulletsLeft(ByRefString) & GREEN & "/" & LIGHTBLUE & t & GREEN & ")" & vbCrLf
                    End If
                    ToSend = ToSend & modGetData.GetItemsDurPercent(CDbl(modItemManip.GetItemDurFromUnFormattedString(ByRefString)), CDbl(.lDurability)) & WHITE
                    WrapAndSend Index, ToSend$ 'send the message
                    X(Index) = ""
                    Exit Function
                End With
            Else
                'send error message
                WrapAndSend Index, RED & "You have no idea what a " & tItem$ & " is." & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
            End If
        ElseIf InStr(1, LCaseFast(TempInv & modItemManip.GetListOfLettersFromInv(dbIndex)), SmartFind(Index, X(Index), Inventory_Item)) Then
            'get the item name
            tItem$ = SmartFind(Index, LCaseFast(X(Index)), Inventory_Item, True, ByRefString)
            If InStr(1, tItem$, Chr$(0)) > 0 Then tItem$ = Mid$(tItem$, InStr(1, tItem$, Chr$(0)) + 1)
            iIn = GetItemID(tItem)
            If iIn = 0 Or ByRefString = "" Then
                tItem = SmartFind(Index, tItem, Equiped_Item, True, ByRefString)
                If InStr(1, tItem$, Chr$(0)) > 0 Then tItem$ = Mid$(tItem$, InStr(1, tItem$, Chr$(0)) + 1)
                iIn = GetItemID(tItem)
            End If
            If iIn <> 0 And ByRefString <> "" Then
                With dbItems(iIn)
                    'start the message
                    ToSend$ = LIGHTBLUE & "You look at the " & .sItemName & vbCrLf & BRIGHTLIGHTBLUE & .sItemName & vbCrLf
                    'finish the message
                    ToSend$ = ToSend$ & GREEN & .sDesc & vbCrLf
                    If modItemManip.GetItemBulletsID(ByRefString) <> 0 Then
                        With dbItems(GetItemID(, modItemManip.GetItemBulletsID(ByRefString)))
                            s = .sItemName
                            t = .iUses
                        End With
                        ToSend$ = ToSend$ & GREEN & "AMMO: " & LIGHTBLUE & s & GREEN & " (" & LIGHTBLUE & modItemManip.GetItemBulletsLeft(ByRefString) & GREEN & "/" & LIGHTBLUE & t & GREEN & ")" & vbCrLf
                    End If
                    ToSend$ = ToSend$ & modGetData.GetItemsDurPercent(CDbl(modItemManip.GetItemDurFromUnFormattedString(ByRefString)), CDbl(.lDurability)) & WHITE
                    WrapAndSend Index, ToSend$ 'send the message
                    X(Index) = ""
                    Exit Function
                End With
            ElseIf GetLetterID(ReplaceFast(tItem$, "note: ", "", 1, 1)) <> 0 Then
                With dbLetters(GetLetterID(ReplaceFast(tItem$, "note: ", "", 1, 1)))
                    ToSend$ = LIGHTBLUE & "You look at the note" & vbCrLf & BRIGHTLIGHTBLUE & .sTitle & vbCrLf
                    'finish the message
                    ToSend$ = ToSend$ & GREEN & .sMessage & vbCrLf & WHITE
                    WrapAndSend Index, ToSend$ 'send the message
                    X(Index) = ""
                    Exit Function
                End With
            Else
            
                'send error message
                WrapAndSend Index, RED & "You have no idea what a " & tItem$ & " is." & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
            End If
        'Letters
        
        ElseIf InStr(1, LCaseFast(TempItemHere), SmartFind(Index, X(Index), Item_In_Room)) Then
            tItem$ = SmartFind(Index, LCaseFast(X(Index)), Item_In_Room, True, ByRefString)
            If InStr(1, tItem$, Chr$(0)) > 0 Then tItem$ = Mid$(tItem$, InStr(1, tItem$, Chr$(0)) + 1)
            iIn = GetItemID(tItem)
            If iIn <> 0 Then
                With dbItems(iIn)
                    'start the message
                    ToSend$ = LIGHTBLUE & "You look at the " & .sItemName & vbCrLf & BRIGHTLIGHTBLUE & .sItemName & vbCrLf
                    'finish the message
                    ToSend$ = ToSend$ & GREEN & .sDesc & vbCrLf & modGetData.GetItemsDurPercent(CDbl(modItemManip.GetItemDurFromUnFormattedString(ByRefString)), CDbl(.lDurability)) & WHITE
                    WrapAndSend Index, ToSend$ 'send the message
                    X(Index) = ""
                    Exit Function
                End With
            ElseIf GetLetterID(ReplaceFast(tItem, "note: ", "", 1, 1)) <> 0 Then
                With dbLetters(GetLetterID(ReplaceFast(tItem$, "note: ", "", 1, 1)))
                    ToSend$ = LIGHTBLUE & "You look at the note" & vbCrLf & BRIGHTLIGHTBLUE & .sTitle & vbCrLf
                    'finish the message
                    ToSend$ = ToSend$ & GREEN & .sMessage & vbCrLf & WHITE
                    WrapAndSend Index, ToSend$ 'send the message
                    X(Index) = ""
                    Exit Function
                End With
            Else
                'send error message
                WrapAndSend Index, RED & "You have no idea what a " & tItem$ & " is." & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
            End If
        Else
            'send error message
            WrapAndSend Index, RED & "You have no idea what a " & X(Index) & " is." & vbCrLf & WHITE
            X(Index) = ""
        End If
    End If
End If
'////////END////////
End Function

Sub SendBadMessageLOOKING(Index As Long)
WrapAndSend Index, RED & "You look at a closed door." & WHITE & vbCrLf
X(Index) = ""
End Sub

Public Function LookDirs(Index As Long, Optional dbIndex As Long, Optional dbMapId As Long) As Boolean
'function to look out of the room
'check the syntax
Dim NewPP As Long
Dim sField As String
Dim dd As Long
If IsADirection(X(Index)) = True Then
    LookDirs = True
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    If dbMapId = 0 Then
        dd = GetMapIndex(dbPlayers(dbIndex).lLocation)
    Else
        dd = dbMapId
    End If
    With dbMap(dd)
        If Len(X(Index)) > 1 Then 'if dir is NW,NE,SW,SE
            Select Case LCaseFast(Left$(X(Index), 2))
                Case "nw": 'northwest
                    If (.lDNW = 1 Or .lDNW = 2) And modMapFlags.GetMapFlag(dd, mapGate, NorthWest) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "NorthWest"
                    NewPP = .lNorthWest
                Case "ne": 'northeast
                    If (.lDNE = 1 Or .lDNE = 2) And modMapFlags.GetMapFlag(dd, mapGate, NorthEast) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "NorthEast"
                    NewPP = .lNorthEast
                Case "sw": 'southwest
                    If (.lDSW = 1 Or .lDSW = 2) And modMapFlags.GetMapFlag(dd, mapGate, SouthWest) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "SouthWest"
                    NewPP = .lSouthWest
                Case "se" 'southeast
                    If (.lDSE = 1 Or .lDSE = 2) And modMapFlags.GetMapFlag(dd, mapGate, SouthEast) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "SouthEast"
                    NewPP = .lSouthEast
            End Select
        End If
        If modSC.FastStringComp(sField, "") Then
            Select Case LCaseFast(Left$(X(Index), 1)) 'normal directions
                Case "n": 'north
                    If (.lDN = 1 Or .lDN = 2) And modMapFlags.GetMapFlag(dd, mapGate, North) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "North"
                    NewPP = .lNorth
                Case "s": 'south
                    If (.lDS = 1 Or .lDS = 2) And modMapFlags.GetMapFlag(dd, mapGate, South) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "South"
                    NewPP = .lSouth
                Case "e": 'east
                    If (.lDE = 1 Or .lDE = 2) And modMapFlags.GetMapFlag(dd, mapGate, East) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "East"
                    NewPP = .lEast
                Case "w": 'west
                    If (.lDW = 1 Or .lDW = 2) And modMapFlags.GetMapFlag(dd, mapGate, West) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "West"
                    NewPP = .lWest
                Case "u": 'up
                    If (.lDU = 1 Or .lDU = 2) And modMapFlags.GetMapFlag(dd, mapGate, Up) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "above"
                    NewPP = .lUp
                Case "d": 'down
                    If (.lDD = 1 Or .lDD = 2) And modMapFlags.GetMapFlag(dd, mapGate, Down) <> "1" Then SendBadMessageLOOKING Index: Exit Function
                    sField = "below"
                    NewPP = .lDown
            End Select
        End If
    End With
    With dbPlayers(dbIndex)
        If NewPP <> 0 Then
            
            SendToAllInRoom Index, BRIGHTLIGHTBLUE & .sPlayerName & " looks to the " & LCaseFast(sField) & "." & WHITE & vbCrLf, .lLocation
            'send to people in the other room
            SendToAllInRoom Index, BRIGHTLIGHTBLUE & .sPlayerName & " peeks into the room from the " & modGetData.GetOppositeDirection(sField) & "." & WHITE & vbCrLf, NewPP
            'get the room description
            WrapAndSend Index, modGetData.GetRoomDescription(dbIndex, NewPP)
            X(Index) = ""
        Else
            LookDirs = False
        End If
    End With
End If
End Function

Public Function WhichDirIsIt(sDir As String) As String
If Len(sDir) > 1 Then 'if dir is NW,NE,SW,SE
    Select Case LCaseFast(Left$(sDir, 2))
        Case "nw": 'northwest
            If Len(sDir) = 2 Then: WhichDirIsIt = "northwest": Exit Function
        Case "ne": 'northeast
            If Len(sDir) = 2 Then: WhichDirIsIt = "northeast": Exit Function
        Case "sw": 'southwest
            If Len(sDir) = 2 Then: WhichDirIsIt = "southwest": Exit Function
        Case "se" 'southeast
            If Len(sDir) = 2 Then: WhichDirIsIt = "southeast": Exit Function
    End Select
End If
Select Case LCaseFast(Left$(sDir, 1)) 'normal directions
    Case "n": 'north
        If Len(sDir) <= 5 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[o]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[r]" Then
                            WhichDirIsIt = "north"
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                WhichDirIsIt = "north"
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "s": 'south
        If Len(sDir) <= 5 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[o]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[u]" Then
                            WhichDirIsIt = "south"
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                WhichDirIsIt = "south"
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "e": 'east
        If Len(sDir) <= 4 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[a]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[s]" Then
                            WhichDirIsIt = "east"
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                WhichDirIsIt = "east"
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "w": 'west
        If Len(sDir) <= 4 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[e]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[s]" Then
                            WhichDirIsIt = "west"
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                WhichDirIsIt = "west"
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "u": 'up
        If Len(sDir) <= 2 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[p]" Then
                    WhichDirIsIt = "up"
                    Exit Function
                Else
                    Exit Function
                End If
            Else
                WhichDirIsIt = "up"
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "d": 'down
        If Len(sDir) <= 4 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[o]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[w]" Then
                            WhichDirIsIt = "down"
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                WhichDirIsIt = "down"
                Exit Function
            End If
        Else
            Exit Function
        End If
End Select
End Function

Public Function IsADirection(sDir As String) As Boolean
IsADirection = False
If Len(sDir) > 1 Then 'if dir is NW,NE,SW,SE
    Select Case LCaseFast(Left$(sDir, 2))
        Case "nw": 'northwest
            If Len(sDir) = 2 Then IsADirection = True: Exit Function
        Case "ne": 'northeast
            If Len(sDir) = 2 Then IsADirection = True: Exit Function
        Case "sw": 'southwest
            If Len(sDir) = 2 Then IsADirection = True: Exit Function
        Case "se" 'southeast
            If Len(sDir) = 2 Then IsADirection = True: Exit Function
    End Select
End If
Select Case LCaseFast(Left$(sDir, 1)) 'normal directions
    Case "n": 'north
        If Len(sDir) <= 5 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[o]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[r]" Then
                            IsADirection = True
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                IsADirection = True
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "s": 'south
        If Len(sDir) <= 5 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[o]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[u]" Then
                            IsADirection = True
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                IsADirection = True
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "e": 'east
        If Len(sDir) <= 4 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[a]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[s]" Then
                            IsADirection = True
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                IsADirection = True
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "w": 'west
        If Len(sDir) <= 4 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[e]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[s]" Then
                            IsADirection = True
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                IsADirection = True
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "u": 'up
        If Len(sDir) <= 2 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[p]" Then
                    IsADirection = True
                    Exit Function
                Else
                    Exit Function
                End If
            Else
                IsADirection = True
                Exit Function
            End If
        Else
            Exit Function
        End If
    Case "d": 'down
        If Len(sDir) <= 4 Then
            If Len(sDir) > 1 Then
                If Mid$(sDir, 2, 1) Like "[o]" Then
                    If Len(sDir) > 2 Then
                        If Mid$(sDir, 3, 1) Like "[w]" Then
                            IsADirection = True
                            Exit Function
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                IsADirection = True
                Exit Function
            End If
        Else
            Exit Function
        End If
End Select
End Function
