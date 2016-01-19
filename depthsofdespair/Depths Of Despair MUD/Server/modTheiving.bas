Attribute VB_Name = "modTheiving"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modTheiving
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function Steal(Index As Long) As Boolean
'rob [player] {of [item]/(g)old}
Dim i As Long
Dim Arr() As String
Dim s As String
Dim sI As String
Dim sP As String
Dim b As Boolean
Dim dbVIndex As Long
Dim dbIndex As Long
Dim sU As String
Dim l As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "rob") Then
    Steal = True
    s = LCaseFast(X(Index))
    If s = "rob" Then Steal = False: Exit Function
    s = Mid$(s, 5)
    If InStr(1, s, " of ") <> 0 Then
        i = InStr(1, s, " of ") + 4
        sI = Mid$(s, i)
        If Len(sI) > 2 Then
            If Mid$(LCaseFast(sI), 3, 1) = "l" Then
                If Len(sI) > 3 Then
                    If Mid$(LCaseFast(sI), 4, 1) = "d" Then
                        If Len(sI) > 4 Then
                            b = False
                        Else
                            b = True
                        End If
                    Else
                        b = True
                    End If
                Else
                    b = True
                End If
            End If
        Else
            b = True
        End If
        sP = Left$(s, i - 5)
        sP = SmartFind(Index, sP, Player_In_Room)
        dbVIndex = GetPlayerIndexNumber(, sP)
        If dbVIndex = 0 Then
            WrapAndSend Index, RED & "You can't find " & sP & " in this room." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If Not b Then sI = SmartFind(dbPlayers(dbVIndex).iIndex, sI, Inventory_Item, True, sU)
    Else
        sP = SmartFind(Index, s, Player_In_Room)
        dbVIndex = GetPlayerIndexNumber(, sP)
        If dbVIndex = 0 Then
            WrapAndSend Index, RED & "You can't find " & sP & " in this room." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sI = dbPlayers(dbVIndex).sInventory
        SplitFast sI, Arr, ";"
        If UBound(Arr) = 0 Then
            WrapAndSend Index, RED & "You can't find " & sP & " in this room." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        i = RndNumber(LBound(Arr), UBound(Arr) - 1)
        
        sU = Arr(i)
        i = modItemManip.GetItemIDFromUnFormattedString(sU)
        i = GetItemID(, i)
        If i = 0 Then
            WrapAndSend Index, RED & "You can't find anything on " & sP & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sI = SmartFind(dbPlayers(dbVIndex).iIndex, dbItems(i).sItemName, Inventory_Item, True, sU)
        If sU = "" Then
            WrapAndSend Index, RED & "You can't find anything on " & sP & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End If
    If InStr(1, dbPlayers(dbVIndex).sInventory, sU) = 0 Or sU = "" Then
        WrapAndSend Index, RED & "You can't find " & sI & " on " & sP & "." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    If modMiscFlag.GetMiscFlag(dbIndex, [Can Steal]) <> 1 Then
        WrapAndSend Index, RED & "You can't figure out how to steal it without getting caught." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    l = modMiscFlag.GetStatsPlusTotal(dbIndex, Thieving)
    If RndNumber(1, 100) < l Then
        'success
        If b = False Then
            With dbPlayers(dbVIndex)
                .sInventory = ReplaceFast(.sInventory, sU & ";", "")
                If .sInventory = "" Then .sInventory = "0"
            End With
            With dbPlayers(dbIndex)
                If .sInventory = "0" Then .sInventory = ""
                .sInventory = .sInventory & sU & ";"
            End With
            WrapAndSend Index, BRIGHTYELLOW & "You steal " & sI & " from " & sP & "." & WHITE & vbCrLf
        Else
            i = RndNumber(0, dbPlayers(dbVIndex).dGold)
            If i > 0 Then
                With dbPlayers(dbVIndex)
                    .dGold = .dGold - i
                End With
                With dbPlayers(dbIndex)
                    .dGold = .dGold + i
                    l = modGetData.GetPlayersMaxGold(Index, dbIndex)
                    If l > .dGold Then
                        i = .dGold - l
                        .dGold = l
                        With dbPlayers(dbVIndex)
                            .dGold = .dGold + i
                        End With
                    End If
                End With
                WrapAndSend Index, BRIGHTYELLOW & "You steal " & i & " gold from " & sP & "." & WHITE & vbCrLf
            Else
                WrapAndSend Index, BRIGHTRED & "You bump into " & sP & "!" & WHITE & vbCrLf
                WrapAndSend dbPlayers(dbVIndex).iIndex, BRIGHTRED & dbPlayers(dbIndex).sPlayerName & " bumps into you!" & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bumps into " & dbPlayers(dbVIndex).sPlayerName & "." & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation, dbPlayers(dbVIndex).iIndex
            End If
        End If
    Else
        WrapAndSend Index, BRIGHTRED & "You bump into " & sP & "!" & WHITE & vbCrLf
        WrapAndSend dbPlayers(dbVIndex).iIndex, BRIGHTRED & dbPlayers(dbIndex).sPlayerName & " bumps into you!" & WHITE & vbCrLf
        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bumps into " & dbPlayers(dbVIndex).sPlayerName & "." & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation, dbPlayers(dbVIndex).iIndex
    End If
    X(Index) = ""
End If
End Function

Public Function Mug(Index As Long) As Boolean
'mug [player] {of [item]/(g)old}
Dim i As Long
Dim Arr() As String
Dim s As String
Dim sI As String
Dim sP As String
Dim b As Boolean
Dim c As Boolean
Dim a&, m&, u&
Dim dbVIndex As Long
Dim dbIndex As Long
Dim sU As String
Dim l As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "mug") Then
    Mug = True
    If lIsPvP <> 1 Then
        WrapAndSend Index, RED & "You are unable to intiate combat here." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    s = LCaseFast(X(Index))
    If s = "mug" Then Mug = False: Exit Function
    s = Mid$(s, 5)
    If InStr(1, s, " of ") <> 0 Then
        i = InStr(1, s, " of ") + 4
        sI = Mid$(s, i)
        If Len(sI) > 2 Then
            If Mid$(LCaseFast(sI), 3, 1) = "l" Then
                If Len(sI) > 3 Then
                    If Mid$(LCaseFast(sI), 4, 1) = "d" Then
                        If Len(sI) > 4 Then
                            b = False
                        Else
                            b = True
                        End If
                    Else
                        b = True
                    End If
                Else
                    b = True
                End If
            End If
        Else
            b = True
        End If
        sP = Left$(s, i - 5)
        sP = SmartFind(Index, sP, Player_In_Room)
        dbVIndex = GetPlayerIndexNumber(, sP)
        If dbVIndex = 0 Then
            WrapAndSend Index, RED & "You can't find " & sP & " in this room." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If Not b Then sI = SmartFind(dbPlayers(dbVIndex).iIndex, sI, Inventory_Item, True, sU): c = False
        If sU = "" Then sI = SmartFind(dbPlayers(dbVIndex).iIndex, sI, Equiped_Item, True, sU): c = True
    Else
        sP = SmartFind(Index, s, Player_In_Room)
        dbVIndex = GetPlayerIndexNumber(, sP)
        If dbVIndex = 0 Then
            WrapAndSend Index, RED & "You can't find " & sP & " in this room." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        c = False
        If RndNumber(0, 1) = 0 Then
            sI = dbPlayers(dbVIndex).sInventory
        Else
            c = True
            sI = modGetData.GetPlayersEq(dbPlayers(dbVIndex).iIndex)
            sI = ReplaceFast(sI, ":0;", "")
        End If
        SplitFast sI, Arr, ";"
        If UBound(Arr) = 0 Then
            WrapAndSend Index, RED & "You can't find " & sP & " in this room." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        i = RndNumber(LBound(Arr), UBound(Arr) - 1)
        sU = Arr(i)
        i = modItemManip.GetItemIDFromUnFormattedString(sU)
        i = GetItemID(, i)
        If i = 0 Then
            WrapAndSend Index, RED & "You can't find anything on " & sP & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If c Then
            sI = SmartFind(dbPlayers(dbVIndex).iIndex, dbItems(i).sItemName, Equiped_Item, True, sU)
        Else
            sI = SmartFind(dbPlayers(dbVIndex).iIndex, dbItems(i).sItemName, Inventory_Item, True, sU)
        End If
        If sU = "" Then
            WrapAndSend Index, RED & "You can't find anything on " & sP & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End If
    If InStr(1, dbPlayers(dbVIndex).sInventory, sU) = 0 And c = False Then
        WrapAndSend Index, RED & "You can't find " & sI & " on " & sP & "." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    If modMiscFlag.GetMiscFlag(dbIndex, [Can Steal]) <> 1 Then
        WrapAndSend Index, RED & "You can't figure out how to steal it without getting caught." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    a& = dbPlayers(dbIndex).iLevel
    m& = dbPlayers(dbVIndex).iLevel - lPvPLevel
    u& = dbPlayers(dbVIndex).iLevel + lPvPLevel
    If (a& < m&) Or (a& > u&) Then
        WrapAndSend Index, RED & "You are about to mug " & dbPlayers(dbVIndex).sPlayerName & ", but decide not to." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If dbMap(GetMapIndex(dbPlayers(dbIndex).lLocation)).iSafeRoom = 1 Then
        WrapAndSend Index, RED & "You are unable to intiate combat here." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    l = modMiscFlag.GetStatsPlusTotal(dbIndex, Thieving) + dbPlayers(dbIndex).iStr
    If RndNumber(1, 175 + dbPlayers(dbVIndex).iStr) < l Then
        'success
        dbPlayers(dbIndex).iSneaking = 0
        dbPlayers(dbIndex).iResting = 0
        dbPlayers(dbIndex).iMeditating = 0
        If b = False Then
            If c = False Then
                modItemManip.TakeFromYourInvAndPutInAnothersInv dbVIndex, dbIndex, modItemManip.GetItemIDFromUnFormattedString(sU)
            Else
                modItemManip.TakeEqItemAndPlaceInInv dbVIndex, modItemManip.GetItemIDFromUnFormattedString(sU)
                modItemManip.TakeFromYourInvAndPutInAnothersInv dbVIndex, dbIndex, modItemManip.GetItemIDFromUnFormattedString(sU)
                WrapAndSend dbPlayers(dbVIndex).iIndex, BRIGHTYELLOW & dbPlayers(dbIndex).sPlayerName & " sneaks up and grabs your " & BRIGHTRED & sI & BRIGHTYELLOW & "!" & WHITE & vbCrLf
                
            End If
            WrapAndSend Index, BRIGHTYELLOW & "You mug " & sP & " of their " & sI & "!" & WHITE & vbCrLf
            bResult = CheckPlayerAttack(Index, dbIndex, dbPlayers(dbVIndex).sPlayerName)
        Else
            i = RndNumber(0, dbPlayers(dbVIndex).dGold)
            If i > 0 Then
                With dbPlayers(dbVIndex)
                    .dGold = .dGold - i
                End With
                With dbPlayers(dbIndex)
                    .dGold = .dGold + i
                    l = modGetData.GetPlayersMaxGold(Index, dbIndex)
                    If l > .dGold Then
                        i = .dGold - l
                        .dGold = l
                        With dbPlayers(dbVIndex)
                            .dGold = .dGold + i
                        End With
                    End If
                End With
                WrapAndSend Index, BRIGHTYELLOW & "You mug " & sP & " of " & i & " gold!" & WHITE & vbCrLf
                bResult = CheckPlayerAttack(Index, dbIndex, dbPlayers(dbVIndex).sPlayerName)
            Else
                WrapAndSend Index, BRIGHTRED & "You bump into " & sP & "!" & WHITE & vbCrLf
                WrapAndSend dbPlayers(dbVIndex).iIndex, BRIGHTRED & dbPlayers(dbIndex).sPlayerName & " bumps into you!" & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bumps into " & dbPlayers(dbVIndex).sPlayerName & "." & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation, dbPlayers(dbVIndex).iIndex
            End If
        End If
    Else
        WrapAndSend Index, BRIGHTRED & "You bump into " & sP & "!" & WHITE & vbCrLf
        WrapAndSend dbPlayers(dbVIndex).iIndex, BRIGHTRED & dbPlayers(dbIndex).sPlayerName & " bumps into you!" & WHITE & vbCrLf
        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bumps into " & dbPlayers(dbVIndex).sPlayerName & "." & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation, dbPlayers(dbVIndex).iIndex
    End If
    X(Index) = ""
End If
End Function
