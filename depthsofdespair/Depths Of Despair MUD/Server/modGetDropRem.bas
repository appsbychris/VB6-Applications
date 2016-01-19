Attribute VB_Name = "modGetDropRem"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modGetDropRem
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Function RemDropGetEQ(Index As Long) As Boolean
'function to check these 5 items
If GiveItems(Index) = True Then RemDropGetEQ = True: Exit Function
If EqAll(Index) = True Then RemDropGetEQ = True: Exit Function
If EquipItems(Index) = True Then RemDropGetEQ = True: Exit Function
If RemAll(Index) = True Then RemDropGetEQ = True: Exit Function
If RemoveItems(Index) = True Then RemDropGetEQ = True: Exit Function
If DropAllItems(Index) = True Then RemDropGetEQ = True: Exit Function
If DropStuff(Index) = True Then RemDropGetEQ = True: Exit Function
If GetAllItems(Index) = True Then RemDropGetEQ = True: Exit Function
If iGetAnItem(Index) = True Then RemDropGetEQ = True: Exit Function
End Function

Public Function IsANumber(ByVal KeyAscii As Long) As Boolean
Select Case KeyAscii
    Case 48 To 57
        IsANumber = True
    Case Else
        IsANumber = False
End Select
End Function

Public Function EqAll(Index As Long) As Boolean
Dim dbIndex As Long
Dim s As String
Dim Arr() As String
Dim i As Long
Dim Message1 As String
Dim Message2 As String
If modSC.FastStringComp(LCaseFast(X(Index)), "eq all") Then
    EqAll = True
    dbIndex = GetPlayerIndexNumber(Index)
    s = dbPlayers(dbIndex).sInventory
    SplitFast s, Arr, ";"
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) <> "" Then
            modGetDropRem.EqItemNoFuss dbIndex, GetItemID(, modItemManip.GetItemIDFromUnFormattedString(Arr(i))), Message1, Message2
        End If
        If DE Then DoEvents
    Next
    WrapAndSend Index, Message1
    SendToAllInRoom Index, Message2, dbPlayers(dbIndex).lLocation
    X(Index) = ""
End If
End Function

Public Function RemAll(Index As Long) As Boolean
Dim dbIndex As Long
Dim s As String
Dim Arr() As String
Dim i As Long
Dim Message1 As String
Dim Message2 As String
If modSC.FastStringComp(LCaseFast(X(Index)), "rem all") Then
    RemAll = True
    dbIndex = GetPlayerIndexNumber(Index)
    s = modGetData.GetPlayersEq(Index)
    SplitFast s, Arr, ";"
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) <> "" Then
            modGetDropRem.RemoveItemNoFuss dbIndex, GetItemID(, modItemManip.GetItemIDFromUnFormattedString(Arr(i))), Message1, Message2
        End If
        If DE Then DoEvents
    Next
    WrapAndSend Index, Message1
    SendToAllInRoom Index, Message2, dbPlayers(dbIndex).lLocation
    X(Index) = ""
End If
End Function

Public Sub IsGold(ByVal s As String, ByRef lAmount As Long, ByRef bIsGold As Boolean)
s = LCaseFast(s)
If Left$(s, 1) = "g" Then
    If Len(s) > 1 Then
        If Mid$(s, 2, 1) <> "o" Then
            bIsGold = False
        Else
            If Len(s) > 2 Then
                If Mid$(s, 3, 1) = "l" Then
                    If Len(s) > 3 Then
                        If Mid$(s, 4, 1) = "d" Then
                            If Len(s) > 4 Then
                                bIsGold = False
                            Else
                                If lAmount > 1 Then
                                    bIsGold = True
                                Else
                                    lAmount = -1
                                    bIsGold = True
                                End If
                            End If
                        Else
                            If lAmount > 1 Then
                                bIsGold = True
                            Else
                                lAmount = -1
                                bIsGold = True
                            End If
                        End If
                    Else
                        If lAmount > 1 Then
                            bIsGold = True
                        Else
                            lAmount = -1
                            bIsGold = True
                        End If
                    End If
                End If
            Else
                If lAmount > 1 Then
                    bIsGold = True
                Else
                    lAmount = -1
                    bIsGold = True
                End If
            End If
        End If
    Else
        If lAmount > 1 Then
            bIsGold = True
        Else
            lAmount = -1
            bIsGold = True
        End If
    End If
End If
End Sub


Public Function DropStuff(Index As Long) As Boolean
'////////DROP STUFF////////
Dim sEQ As String
Dim bGold As Boolean
Dim dbIndex As Long
Dim dbItemID As Long
Dim lAmount As Long
Dim i As Long
Dim j As Long
Dim bIsLetter As Boolean


s = LCaseFast(X(Index))
If modSC.FastStringComp(Left$(s, 4), "drop") Then   'if the command
    DropStuff = True
    If Len(s) < 6 Then
        DropStuff = False
        Exit Function
    End If
    If s Like "drop #* *" Then
        i = InStr(1, s, " ")
        j = InStr(i + 1, s, " ")
        lAmount = Val(Mid$(s, i + 1, j - i - 1))
        If lAmount = 0 Then
            WrapAndSend Index, RED & "Why would you try to drop 0 of something?" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sEQ = Mid$(s, j + 1)
        s = Mid$(s, j + 1)
    Else
        sEQ = Mid$(s, 6)
        s = Mid$(s, 6)
        lAmount = 1
    End If
    sEQ = SmartFind(Index, sEQ, Inventory_Item)
    If InStr(1, sEQ, Chr$(0)) > 0 Then sEQ = Mid$(sEQ, InStr(1, sEQ, Chr$(0)) + 1)
    dbIndex = GetPlayerIndexNumber(Index)
    IsGold s, lAmount, bGold
    If bGold = False Then
        dbItemID = GetItemID(sEQ)
        If dbItemID = 0 Then
            dbItemID = GetLetterID(ReplaceFast(sEQ, "note: ", "", 1, 1))
            If dbItemID = 0 Then
                WrapAndSend Index, RED & "You do not have that." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            Else
                bIsLetter = True
            End If
        End If
    End If
    If bIsLetter Then
        If InStr(1, LCaseFast(modItemManip.GetListOfLettersFromInv(dbIndex)), sEQ & ",") <> 0 Then
            modItemManip.TakeLetterFromInvAndDropIt dbIndex, dbLetters(dbItemID).lID
        Else
            WrapAndSend Index, RED & "You do not have that." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    ElseIf Not bGold And Not bIsLetter Then
        If InStr(1, dbPlayers(dbIndex).sInventory, ":" & dbItems(dbItemID).iID & "/") = 0 Then
            WrapAndSend Index, RED & "You do not have that." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        j = 0
        For i = 1 To lAmount
            If InStr(1, dbPlayers(dbIndex).sInventory, ":" & dbItems(dbItemID).iID & "/") <> 0 Then
                modItemManip.TakeItemFromInvAndPutOnGround dbIndex, dbItems(dbItemID).iID
                j = j + 1
            Else
                Exit For
            End If
            If DE Then DoEvents
        Next
        lAmount = j
    ElseIf bGold Then
        If dbPlayers(dbIndex).dGold < lAmount Then
            WrapAndSend Index, RED & "You do not have that much gold." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        With dbPlayers(dbIndex)
            If lAmount = -1 Then lAmount = .dGold
            If lAmount = 0 Then
                WrapAndSend Index, RED & "You do not have any gold." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
            .dGold = .dGold - lAmount
        End With
        With dbMap(dbPlayers(dbIndex).lDBLocation)
            .dGold = .dGold + lAmount
        End With
        sEQ = lAmount & " gold"
    End If
    If Not bGold Then
        If lAmount > 1 Then
            WrapAndSend Index, LIGHTBLUE & "You drop " & lAmount & " " & modGetData.GetItemsNameAddS(dbItemID) & "." & vbCrLf & WHITE
            SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " drops " & lAmount & " " & modGetData.GetItemsNameAddS(dbItemID) & "." & vbCrLf & WHITE, dbPlayers(dbIndex).lLocation
        Else
            WrapAndSend Index, LIGHTBLUE & "You drop your " & sEQ & "." & vbCrLf & WHITE
            SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " drops their " & sEQ & "." & vbCrLf & WHITE, dbPlayers(dbIndex).lLocation
        End If
    Else
        WrapAndSend Index, LIGHTBLUE & "You drop " & sEQ & "." & vbCrLf & WHITE
        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " drops " & sEQ & "." & vbCrLf & WHITE, dbPlayers(dbIndex).lLocation
    End If
    X(Index) = ""
End If
'////////END////////
End Function



Public Function ClassCanWear(ItemID As Long, ClassCheck As String, Optional dbItemID As Long) As Boolean
Dim i As Long, tArr() As String
If dbItemID = 0 Then dbItemID = GetItemID(, ItemID)
With dbItems(dbItemID)
    If .sClassRestriction = "0" Then ClassCanWear = True: Exit Function
    'tArr() = Split(Left$(.sClassRestriction, Len(.sClassRestriction) - 1), ";")
    SplitFast Left$(.sClassRestriction, Len(.sClassRestriction) - 1), tArr, ";"
    For i = 0 To UBound(tArr)
        If modSC.FastStringComp(LCaseFast(tArr(i)), LCaseFast(ClassCheck)) Then
            ClassCanWear = True
            Exit Function
        End If
        If DE Then DoEvents
    Next
End With
ClassCanWear = False
End Function

Public Function ClassCanuseMagical(PlayersClass As String, Magical As Long) As Boolean
ClassCanuseMagical = False
If Magical <= dbClass(GetClassID(PlayersClass)).iUseMagical Then ClassCanuseMagical = True
End Function

Public Function RaceCanWear(ItemID As Long, RaceCheck As String, Optional dbItemID As Long) As Boolean
Dim i As Long, tArr() As String
If dbItemID = 0 Then dbItemID = GetItemID(, ItemID)
With dbItems(dbItemID)
    If .sRaceRestriction = "0" Then RaceCanWear = True: Exit Function
    'tArr() = Split(Left$(.sRaceRestriction, Len(.sRaceRestriction) - 1), ";")
    SplitFast Left$(.sRaceRestriction, Len(.sRaceRestriction) - 1), tArr, ";"
    For i = 0 To UBound(tArr)
        If modSC.FastStringComp(LCaseFast(tArr(i)), LCaseFast(RaceCheck)) Then
            RaceCanWear = True
            Exit Function
        End If
        If DE Then DoEvents
    Next
End With
End Function

Public Function IsLedgendary(dbItemName As String) As Boolean
If GetItemID(dbItemName) = 0 Then IsLedgendary = False: Exit Function
With dbItems(GetItemID(dbItemName))
    IsLedgendary = IIf(.iIsLedgenary = 1, True, False)
End With
End Function

Public Function MaxLedgendary(dbPlayerIndexNumber As Long) As Long
Dim iAmount As Long
Dim i As Long
With dbPlayers(dbPlayerIndexNumber)
    If IsLedgendary(.sArms) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sBody) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sFeet) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sHands) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sHead) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sLegs) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sWaist) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sWeapon) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sShield) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sBack) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sEars) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sFace) = True Then iAmount = iAmount + 1
    If IsLedgendary(.sNeck) = True Then iAmount = iAmount + 1
    For i = 0 To 5
        If IsLedgendary(.sRings(i)) = True Then iAmount = iAmount + 1
        If DE Then DoEvents
    Next
End With
MaxLedgendary = iAmount
End Function

Public Sub EqItemNoFuss(dbIndex As Long, dbItemID As Long, ByRef Message1 As String, ByRef Message2 As String)
Dim iRealItemID As Long
Dim i As Long
Dim bDual As Boolean
With dbPlayers(dbIndex)
    If ClassCanuseMagical(.sClass, dbItems(dbItemID).iMagical) = False Then Exit Sub
    If ClassCanWear(CLng(dbItems(dbItemID).iID), .sClass) = False Then Exit Sub
    If IsLedgendary(dbItems(dbItemID).sItemName) = True Then
        If MaxLedgendary(dbIndex) >= 2 Then Exit Sub
    End If
    If RaceCanWear(CLng(dbItems(dbItemID).iID), .sRace) = False Then Exit Sub
    If .iLevel < dbItems(dbItemID).lLevel Then Exit Sub
    If modWeaponsAndArmor.PlayerCanUseWeapon(dbIndex, dbItemID) = False And dbItems(dbItemID).iType <> 0 Then Exit Sub
    If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then Exit Sub
    For i = 0 To 5
        If Not modSC.FastStringComp(.sRings(i), "0") Then
            If dbItems(dbItemID).iID = modItemManip.GetItemIDFromUnFormattedString(.sRings(i)) Then Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
If modSC.FastStringComp(dbItems(iItemID).sWorn, "shield") Then
    If dbPlayers(dbpID).sWeapon <> "0" Then
        Select Case dbItems(GetItemID(, modItemManip.GetItemIDFromUnFormattedString(dbPlayers(dbpID).sWeapon))).iType
            Case 1, 2, 4, 5
            Case Else
                Exit Sub
        End Select
    End If
End If
bDual = False
If modSC.FastStringComp(dbItems(dbItemID).sWorn, "weapon") Then
    If dbPlayers(dbIndex).sWeapon <> "0" Then
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Dual Wield]) = 1 Then
            Select Case dbItems(GetItemID(, modItemManip.GetItemIDFromUnFormattedString(dbPlayers(dbIndex).sWeapon))).iType
                Case 1, 2, 4, 5
                    Select Case dbItems(dbItemID).iType
                        Case 1, 2, 4, 5
                            If dbPlayers(dbIndex).iDualWield = 0 Then
                                dbPlayers(dbIndex).iDualWield = 1
                                bDual = True
                            End If
                    End Select
            End Select
        End If
    End If
End If
With dbItems(dbItemID)
    iRealItemID = .iID
    If modSC.FastStringComp(.sWorn, "item") Or modSC.FastStringComp(.sWorn, _
        "scroll") Or modSC.FastStringComp(.sWorn, _
        "corpse") Or modSC.FastStringComp(.sWorn, _
        "food") Or modSC.FastStringComp(.sWorn, _
        "ofood") Or modSC.FastStringComp(.sWorn, "key") Or modSC.FastStringComp(.sWorn, "projectile") Then Exit Sub
End With
With dbPlayers(dbIndex)
    modItemManip.TakeItemFromInvAndEqIt dbIndex, iRealItemID, bDual
    Message1 = Message1 & LIGHTBLUE & "You are now wearing " & dbItems(dbItemID).sItemName & " on your " & dbItems(dbItemID).sWorn & "." & vbCrLf & WHITE
    Message2 = Message2 & LIGHTBLUE & .sPlayerName & " equips " & dbItems(dbItemID).sItemName & "." & WHITE & vbCrLf
End With
End Sub

Public Function EquipItems(Index As Long) As Boolean
'////////EQUIP ITEMS////////
Dim TempEq As String
Dim iItemID As Long, iRealItemID As Long
Dim dbpID As Long, i As Long
Dim a As Long
Dim lI As Long
Dim bDual As Boolean
'function to equip items
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "eq ") Then   'if the command
    EquipItems = True
    For a = 1 To InStr(1, X(Index), " ") 'trim off the command
        X(Index) = Mid$(X(Index), 2)
    Next a
    TempEq = TrimIt(X(Index)) 'get the name of the equipemnt
    'and then make sure it is the full name
    TempEq = SmartFind(Index, TempEq, Inventory_Item)
    If InStr(1, TempEq, Chr$(0)) > 0 Then TempEq = Mid$(TempEq, InStr(1, TempEq, Chr$(0)) + 1)
    iItemID = GetItemID(TempEq)
    dbpID = GetPlayerIndexNumber(Index)
    With dbPlayers(dbpID)
        'check restrictions
        If iItemID = 0 Then
            WrapAndSend Index, RED & "You have no idea what that is." & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        If ClassCanuseMagical(.sClass, dbItems(iItemID).iMagical) = False Then   'magic level
            WrapAndSend Index, RED & "You may not use that!" & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        If ClassCanWear(CLng(dbItems(iItemID).iID), .sClass) = False Then      'class
            WrapAndSend Index, RED & "You may not use that!" & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        If IsLedgendary(TempEq) = True Then
            If MaxLedgendary(dbpID) >= 2 Then
                WrapAndSend Index, RED & "This item, combined with the ones you are currently wearing, will cause great confliction." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        End If
        If RaceCanWear(CLng(dbItems(iItemID).iID), .sRace) = False Then   'race
            WrapAndSend Index, RED & "You may not use that!" & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        If .iLevel < dbItems(iItemID).lLevel Then   'level
            WrapAndSend Index, RED & "You may not use that!" & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        If modWeaponsAndArmor.PlayerCanUseWeapon(CLng(dbpID), CLng(iItemID)) = False And dbItems(iItemID).iType <> 0 Then    'weapon type
            WrapAndSend Index, RED & "You may not weild that weapon!" & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        If modWeaponsAndArmor.PlayerCanUseArmor(CLng(dbpID), CLng(iItemID)) = False Then   'armour type
            WrapAndSend Index, RED & "You may not wear that armor!" & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        For i = 0 To 5
            If Not modSC.FastStringComp(.sRings(i), "0") Then
                If dbItems(iItemID).iID = modItemManip.GetItemIDFromUnFormattedString(.sRings(i)) Then
                    WrapAndSend Index, RED & "You are already wearing one of those!" & vbCrLf & WHITE
                    X(Index) = ""
                    Exit Function
                End If
            End If
            If DE Then DoEvents
        Next
        If InStr(1, dbItems(iItemID).sWorn, "ring") > 0 Then
            lI = 0
            For i = 0 To 5
                If modSC.FastStringComp(.sRings(i), "0") Then
                    Select Case i
                        Case 0
                            If modMiscFlag.GetMiscFlag(dbpID, [Can Eq Ring 0]) = 1 Then lI = lI + 1
                        Case 1
                            If modMiscFlag.GetMiscFlag(dbpID, [Can Eq Ring 1]) = 1 Then lI = lI + 1
                        Case 2
                            If modMiscFlag.GetMiscFlag(dbpID, [Can Eq Ring 2]) = 1 Then lI = lI + 1
                        Case 3
                            If modMiscFlag.GetMiscFlag(dbpID, [Can Eq Ring 3]) = 1 Then lI = lI + 1
                        Case 4
                            If modMiscFlag.GetMiscFlag(dbpID, [Can Eq Ring 4]) = 1 Then lI = lI + 1
                        Case 5
                            If modMiscFlag.GetMiscFlag(dbpID, [Can Eq Ring 5]) = 1 Then lI = lI + 1
                    End Select
                Else
                    lI = lI + 1
                End If
                If DE Then DoEvents
            Next
            If lI = 6 Then
                WrapAndSend Index, RED & "You don't have anymore room on your fingers for that!" & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        End If
    End With
    If iItemID = 0 Then
        'send error message
        WrapAndSend Index, RED & "You just remembered, you don't know what a " & TempEq & " is." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    If modSC.FastStringComp(dbItems(iItemID).sWorn, "shield") Then
        If dbPlayers(dbpID).sWeapon <> "0" Then
            Select Case dbItems(GetItemID(, modItemManip.GetItemIDFromUnFormattedString(dbPlayers(dbpID).sWeapon))).iType
                Case 1, 2, 4, 5
                Case Else
                    WrapAndSend Index, RED & "Your two hands are busy with your weapon!" & WHITE & vbCrLf
                    X(Index) = ""
                    Exit Function
            End Select
        End If
    End If
    With dbItems(iItemID)
        iRealItemID = .iID
        If modSC.FastStringComp(.sWorn, "item") Or modSC.FastStringComp(.sWorn, _
            "scroll") Or modSC.FastStringComp(.sWorn, _
            "corpse") Or modSC.FastStringComp(.sWorn, _
            "food") Or modSC.FastStringComp(.sWorn, _
            "ofood") Or modSC.FastStringComp(.sWorn, "key") Or modSC.FastStringComp(.sWorn, "projectile") Then
            
                WrapAndSend Index, RED & "You may not equip that!" & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
        End If
    End With
    If InStr(1, dbPlayers(dbpID).sInventory, ":" & iRealItemID & "/") = 0 Then
        WrapAndSend Index, RED & "You do not have that!" & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    bDual = False
    If modSC.FastStringComp(dbItems(iItemID).sWorn, "weapon") Then
        If dbPlayers(dbpID).sWeapon <> "0" Then
            If modMiscFlag.GetMiscFlag(dbpID, [Can Dual Wield]) = 1 Then
                Select Case dbItems(GetItemID(, modItemManip.GetItemIDFromUnFormattedString(dbPlayers(dbpID).sWeapon))).iType
                    Case 1, 2, 4, 5
                        Select Case dbItems(iItemID).iType
                            Case 1, 2, 4, 5
                                If dbPlayers(dbpID).iDualWield = 0 Then
                                    dbPlayers(dbpID).iDualWield = 1
                                    bDual = True
                                End If
                            Case Else
                                modItemManip.TakeEqItemAndPlaceInInv dbpID, dbItems(iItemID).iID
                        End Select
                End Select
            End If
        End If
    End If
    With dbPlayers(dbpID)
        modItemManip.TakeItemFromInvAndEqIt dbpID, iRealItemID, bDual
        If bDual Then
            WrapAndSend Index, LIGHTBLUE & "You are now wearing " & TempEq & " on your off-hand." & vbCrLf & WHITE
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " equips " & TempEq & "." & WHITE & vbCrLf, .lLocation
        Else
            WrapAndSend Index, LIGHTBLUE & "You are now wearing " & TempEq & " on your " & dbItems(iItemID).sWorn & "." & vbCrLf & WHITE
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " equips " & TempEq & "." & WHITE & vbCrLf, .lLocation
        End If
    End With
    X(Index) = ""
End If
'////////END////////
End Function

Public Function IsGetable(ItemName As String) As Boolean
If modSC.FastStringComp(LCaseFast(ItemName), "gold") Then IsGetable = True: Exit Function
If GetItemID(ItemName) = 0 Then
    IsGetable = False
    Exit Function
End If
With dbItems(GetItemID(ItemName))
    If .iMoveable = 0 Then
        IsGetable = False
        Exit Function
    Else
        IsGetable = True
    End If
End With
End Function

Public Function iGetAnItem(Index As Long) As Boolean
'////////GET AN ITEM////////
Dim sEQ As String
Dim bGold As Boolean
Dim dbIndex As Long
Dim dbItemID As Long
Dim lAmount As Long
Dim i As Long
Dim j As Long
Dim bIsLetter As Boolean
Dim bIsHidden As Boolean

s = LCaseFast(X(Index))
If modSC.FastStringComp(Left$(s, 1), "g") Then
    iGetAnItem = True
    If s Like "g* #* *" Then
        i = InStr(1, s, " ")
        j = InStr(i + 1, s, " ")
        lAmount = Val(Mid$(s, i + 1, j - i - 1))
        If lAmount = 0 Then
            WrapAndSend Index, RED & "Why would you try to pick up 0 of something?" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sEQ = Mid$(s, j + 1)
        s = Mid$(s, j + 1)
    Else
        If Left$(s, 2) <> "gu" Then
            i = InStr(1, s, " ")
            s = Mid$(s, i + 1)
            sEQ = s
            lAmount = 1
        Else
            iGetAnItem = False
            Exit Function
        End If
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    IsGold s, lAmount, bGold
    If bGold = False Then
        sEQ = SmartFind(Index, sEQ, Item_In_Room)
        If InStr(1, sEQ, Chr$(0)) > 0 Then sEQ = Mid$(sEQ, InStr(1, sEQ, Chr$(0)) + 1)
        With dbMap(dbPlayers(dbIndex).lDBLocation)
            dbItemID = GetItemID(sEQ)
            If dbItemID = 0 Then
                dbItemID = GetLetterID(ReplaceFast(sEQ, "note: ", "", 1, 1))
                If dbItemID = 0 Then
                    If dbPlayers(dbIndex).lRoomSearched = dbPlayers(dbIndex).lLocation Then
                        sEQ = SmartFind(Index, sEQ, Hidden_Item)
                        If InStr(1, sEQ, Chr$(0)) > 0 Then sEQ = Mid$(sEQ, InStr(1, sEQ, Chr$(0)) + 1)
                        dbItemID = GetItemID(sEQ)
                        If dbItemID = 0 Then
                            dbItemID = GetLetterID(ReplaceFast(sEQ, "note: ", "", 1, 1))
                            If dbItemID = 0 Then
                                WrapAndSend Index, RED & "You can't find that here." & WHITE & vbCrLf
                                X(Index) = ""
                                Exit Function
                            End If
                            If InStr(1, .sHLetters, ":" & dbLetters(dbItemID).lID & ";") <> 0 Then
                                bIsHidden = True
                                bIsLetter = True
                            Else
                                WrapAndSend Index, RED & "You can't find that here." & WHITE & vbCrLf
                                X(Index) = ""
                                Exit Function
                            End If
                        Else
                            If InStr(1, .sHidden, ":" & dbItems(dbItemID).iID & "/") = 0 Then
                                WrapAndSend Index, RED & "You can't find that here." & WHITE & vbCrLf
                                X(Index) = ""
                                Exit Function
                            End If
                            bIsHidden = True
                        End If
                    Else
                        WrapAndSend Index, RED & "You can't find that here." & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Function
                    End If
                Else
                    If InStr(1, .sLetters, ":" & dbLetters(dbItemID).lID & ";") <> 0 Then
                        bIsLetter = True
                    ElseIf InStr(1, .sHLetters, ":" & dbLetters(dbItemID).lID & ";") <> 0 Then
                        bIsHidden = True
                        bIsLetter = True
                    Else
                        WrapAndSend Index, RED & "You can't find that here." & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Function
                    End If
                End If
            End If
            If Not bIsHidden And Not bIsLetter Then
                If InStr(1, .sItems, ":" & dbItems(dbItemID).iID & "/") = 0 Then
                    WrapAndSend Index, RED & "You can't find that here." & WHITE & vbCrLf
                    X(Index) = ""
                    Exit Function
                End If
            End If
        End With
        If bIsLetter = False Then
            With dbItems(dbItemID)
                If .iMoveable = 0 Then
                    WrapAndSend Index, RED & "You struggle with all your might, and you are unable to lift the item." & WHITE & vbCrLf
                    X(Index) = ""
                    Exit Function
                End If
            End With
            If modGetData.GetPlayersTotalItems(Index, dbIndex) + lAmount > modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) Then
                WrapAndSend Index, RED & "You don't have any more room on you for that!" & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        End If
    Else
        With dbMap(dbPlayers(dbIndex).lDBLocation)
            If lAmount = -1 Then lAmount = .dGold
            If dbPlayers(dbIndex).dGold + lAmount > modGetData.GetPlayersMaxGold(Index, dbIndex) Then
                lAmount = modGetData.GetPlayersMaxGold(Index, dbIndex) - dbPlayers(dbIndex).dGold
            End If
            If lAmount > .dGold Then
                WrapAndSend Index, RED & "You can't find that much gold here." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        End With
    End If
    If bGold Then
        With dbMap(dbPlayers(dbIndex).lDBLocation)
            .dGold = .dGold - lAmount
            dbPlayers(dbIndex).dGold = dbPlayers(dbIndex).dGold + lAmount
        End With
        If lAmount = 1 Then
            sEQ = "gold piece"
        Else
            sEQ = lAmount & " gold pieces"
        End If
    ElseIf bIsLetter And Not bIsHidden Then
        j = 0
        For i = 1 To lAmount
            If InStr(1, dbMap(dbPlayers(dbIndex).lDBLocation).sLetters, ":" & dbLetters(dbItemID).lID & ";") <> 0 Then
                modItemManip.TakeLetterFromGroundAndPutItInInv dbIndex, dbLetters(dbItemID).lID
                j = j + 1
            Else
                Exit For
            End If
            If DE Then DoEvents
        Next
        lAmount = j
        If InStr(1, sEQ, "note: ") = 0 Then sEQ = "note: " & sEQ
    ElseIf bIsLetter And bIsHidden Then
        j = 0
        For i = 1 To lAmount
            If InStr(1, dbMap(dbPlayers(dbIndex).lDBLocation).sHLetters, ":" & dbLetters(dbItemID).lID & ";") <> 0 Then
                modItemManip.TakeLetterFromHiddenAndPutItInInv dbIndex, dbLetters(dbItemID).lID
                j = j + 1
            Else
                Exit For
            End If
            If DE Then DoEvents
        Next
        lAmount = j
        If InStr(1, sEQ, "note: ") = 0 Then sEQ = "note: " & sEQ
    ElseIf Not bIsLetter And bIsHidden Then
        j = 0
        For i = 1 To lAmount
            If InStr(1, dbMap(dbPlayers(dbIndex).lDBLocation).sHidden, ":" & dbItems(dbItemID).iID & "/") <> 0 Then
                modItemManip.TakeHiddenItemAndPutInInv dbIndex, dbItems(dbItemID).iID
                j = j + 1
            End If
            If DE Then DoEvents
        Next
        lAmount = j
    ElseIf Not bGold And Not bIsLetter And Not bIsHidden Then
        j = 0
        For i = 1 To lAmount
            If InStr(1, dbMap(dbPlayers(dbIndex).lDBLocation).sItems, ":" & dbItems(dbItemID).iID & "/") <> 0 Then
                modItemManip.TakeItemFromGroundAndPutInInv dbIndex, dbItems(dbItemID).iID
                j = j + 1
            End If
            If DE Then DoEvents
        Next
        lAmount = j
    End If
    If lAmount > 1 And Not bGold Then
        WrapAndSend Index, LIGHTBLUE & "You pick up " & lAmount & " " & modGetData.GetItemsNameAddS(dbItemID) & "." & vbCrLf & WHITE
        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks up " & lAmount & " " & modGetData.GetItemsNameAddS(dbItemID) & "." & vbCrLf & WHITE, dbPlayers(dbIndex).lLocation
    Else
        WrapAndSend Index, LIGHTBLUE & "You pick up " & sEQ & "." & vbCrLf & WHITE
        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks up " & sEQ & "." & vbCrLf & WHITE, dbPlayers(dbIndex).lLocation
    End If
    X(Index) = ""
End If
'////////END////////
End Function

Public Function RemoveItems(Index As Long) As Boolean
'////////REMOVE ITEMS////////

Dim theEq As String
Dim iItemID As Long, dbpID As Long
'Removeing equiped items from yourself
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "rem") Then    'if the command
    RemoveItems = True
    For a = 1 To InStr(1, X(Index), " ") 'remove the command
        X(Index) = Mid$(X(Index), 2)
    Next a
    theEq = TrimIt(X(Index)) 'get the selected equipment
    'make sure it is the full name
    theEq = SmartFind(Index, theEq, Equiped_Item)
    If InStr(1, theEq, Chr$(0)) > 0 Then theEq = Mid$(theEq, InStr(1, theEq, Chr$(0)) + 1)
    iItemID = GetItemID(theEq)
    If iItemID = 0 Then
        WrapAndSend Index, RED & "You can't seem to find " & theEq & " anywhere on here." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    dbpID = GetPlayerIndexNumber(Index)
    'theEq = dbItems(iItemID).iID
    modItemManip.TakeEqItemAndPlaceInInv dbpID, dbItems(iItemID).iID
    With dbPlayers(dbpID)
        WrapAndSend Index, LIGHTBLUE & "You remove your " & theEq & " from your " & dbItems(iItemID).sWorn & "." & vbCrLf & WHITE
        SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " removes " & theEq & "." & WHITE & vbCrLf, .lLocation
        .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(iItemID).sWorn & "/" & dbItems(iItemID).iOnEquipKillDur & ";", "", 1, 1)
        If modSC.FastStringComp(.sKillDurItems, "") Then .sKillDurItems = "0"
    End With
    
    X(Index) = ""
End If
'////////END////////
End Function

Public Sub RemoveItemNoFuss(dbIndex As Long, dbItemID As Long, ByRef Message1 As String, ByRef Message2 As String)
If dbItemID = 0 Then Exit Sub
modItemManip.TakeEqItemAndPlaceInInv dbIndex, dbItems(dbItemID).iID
With dbPlayers(dbIndex)
    Message1 = Message1 & LIGHTBLUE & "You remove your " & dbItems(dbItemID).sItemName & " from your " & dbItems(dbItemID).sWorn & "." & vbCrLf & WHITE
    Message2 = Message2 & LIGHTBLUE & .sPlayerName & " removes " & dbItems(dbItemID).sItemName & "." & WHITE & vbCrLf
    .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(dbItemID).sWorn & "/" & dbItems(dbItemID).iOnEquipKillDur & ";", "", 1, 1)
    If modSC.FastStringComp(.sKillDurItems, "") Then .sKillDurItems = "0"
End With
End Sub

Public Function GetAllItems(Index As Long) As Boolean
Dim s As String
Dim t As String
Dim v As String
Dim tArr() As String
Dim dbIndex As Long
Dim l As Long
Dim m As Long
Dim i As Long
If LCaseFast(X(Index)) = "get all items" Then
    GetAllItems = True
    dbIndex = GetPlayerIndexNumber(Index)
    s = modGetData.GetItemsHere(dbPlayers(dbIndex).lLocation, dbPlayers(dbIndex).lDBLocation)
    If s = "0" Then
        WrapAndSend Index, RED & "There are no items here." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    SplitFast s, tArr, ";"
    l = modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items])
    m = modGetData.GetPlayersTotalItems(Index, dbIndex)
    If m >= l Then
        WrapAndSend Index, RED & "You don't have room for anymore items." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    s = ""
    l = l - m
    If l > UBound(tArr) Then l = UBound(tArr)
    For i = 0 To l
        If tArr(i) <> "" And tArr(i) <> "0" Then
            v = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
            If IsGetable(dbItems(v).sItemName) = True Then
                modItemManip.TakeItemFromGroundAndPutInInv CLng(dbIndex), modItemManip.GetItemIDFromUnFormattedString(tArr(i))
                s = s & LIGHTBLUE & "You pick up " & dbItems(v).sItemName & "." & WHITE & vbCrLf
                t = s & LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks up " & dbItems(v).sItemName & "." & WHITE & vbCrLf
            End If
        End If
        If DE Then DoEvents
    Next
    WrapAndSend Index, s
    SendToAllInRoom Index, t, dbPlayers(dbIndex).lLocation
    X(Index) = ""
End If
End Function

Public Function DropAllItems(Index As Long) As Boolean
Dim s As String
Dim dbIndex As Long
Dim dbMapIndex As Long
If LCaseFast(X(Index)) = "drop all items" Then
    DropAllItems = True
    dbIndex = GetPlayerIndexNumber(Index)
    dbMapIndex = GetMapIndex(dbPlayers(dbIndex).lLocation)
    With dbPlayers(dbIndex)
        s = .sInventory
        .sInventory = "0"
    End With
    With dbMap(dbMapIndex)
        If .sItems = "0" Then .sItems = ""
        .sItems = .sItems & s
    End With
    WrapAndSend Index, LIGHTBLUE & "You drop all of your items." & WHITE & vbCrLf
    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " drops all of their items." & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
    X(Index) = ""
End If
End Function
