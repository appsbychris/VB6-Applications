Attribute VB_Name = "modSmartFind"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modSmartFind
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Enum SmartFindChoices
    Inventory_Item = 0
    Equiped_Item = 1
    Item_In_Room = 2
    Monster_In_Room = 3
    Item_In_Shop = 4
    Player_In_Room = 5
    Hidden_Item = 6
    All_Players = 7
    All_Items = 8
    All_Monsters = 9
End Enum

Public Function SmartFind(Index As Long, SmartFindWhat As String, Which As SmartFindChoices, Optional bPutUnFormattedInByRef As Boolean = False, Optional ByRef UnFormatted As String) As String
Dim i As Long
If Index = 0 Or modSC.FastStringComp(SmartFindWhat, "") Then Exit Function
SmartFindWhat = LCaseFast(SmartFindWhat)
Dim pINV$, tArr() As String, Floor$, Monsters$, ShopItems$, ShopItemsIDs$, _
    Players$, ShopDBIndex As Long, tArr2() As String
Dim s As String
Dim j As Long
Dim k As Long
Dim t As String
Dim dbIndex As Long
Select Case Which
    Case 0
        dbIndex = GetPlayerIndexNumber(Index)
        With dbPlayers(dbIndex)
            pINV$ = modGetData.GetPlayersInvFromNums(Index, , dbIndex) & modItemManip.GetListOfLettersFromInv(dbIndex)
            pINV$ = LCaseFast(ReplaceFast(pINV$, "0", ""))
            If modSC.FastStringComp(pINV$, "") Then SmartFind = SmartFindWhat: Exit Function
            SplitFast Left$(pINV$, Len(pINV$) - 1), tArr, ","
            For i = 0 To UBound(tArr)
                tArr(i) = LCaseFast(tArr(i))
                If InStr(1, ReplaceFast(tArr(i), Chr$(0), ""), SmartFindWhat) Then
                    s = s & LCaseFast(tArr(i)) & ";"
                    If bPutUnFormattedInByRef And InStr(1, tArr(i), "note: ") = 0 Then
                        SplitFast .sInventory, tArr2, ";"
                        t = t & tArr2(i) & ";"
                    End If
                End If
                If DE Then DoEvents
            Next
            If DCount(s, ";") = 1 Then
                SmartFind = Left$(s, Len(s) - 1)
                If t <> "" Then
                    UnFormatted = Left$(t, Len(t) - 1)
                End If
                Exit Function
            ElseIf s <> "" Then
                Erase tArr
                SplitFast s, tArr, ";"
                SplitFast t, tArr2, ";"
                j = Len(tArr(LBound(tArr)))
                For i = LBound(tArr) To UBound(tArr)
                    If tArr(i) <> "" Then
                        If Len(tArr(i)) < j Then
                            j = Len(tArr(i))
                            k = i
                        End If
                    End If
                    If DE Then DoEvents
                Next
                SmartFind = tArr(k)
                If k < UBound(tArr2) Then
                    If tArr2(k) <> "" Then
                        UnFormatted = tArr2(k)
                    End If
                End If
                Exit Function
            End If
            SmartFind = "----1"
            UnFormatted = ""
        End With
    Case 1
        dbIndex = GetPlayerIndexNumber(Index)
        With dbPlayers(dbIndex)
            pINV$ = modGetData.GetPlayersEqFromNums(Index, True, dbIndex)
            SplitFast LCaseFast(Left$(pINV$, Len(pINV$) - 1)), tArr, ","
            pINV = .sArms & ";" & .sBack & ";" & .sBody & ";" & .sEars & ";" & .sFace & _
                ";" & .sFeet & ";" & .sHands & ";" & .sHead & ";" & .sLegs & ";" & .sNeck & _
                ";" & .sShield & ";" & .sWaist & ";" & .sWeapon & ";" & .sRings( _
                0) & ";" & .sRings(1) & ";" & .sRings(2) & ";" & .sRings(3) & ";" & .sRings( _
                4) & ";" & .sRings(5)
            SplitFast pINV$, tArr2, ";"
            For i = 0 To UBound(tArr)
                If InStr(1, ReplaceFast(tArr(i), Chr$(0), ""), SmartFindWhat) Then
                    s = s & tArr(i) & ";"
                    If bPutUnFormattedInByRef Then
                        t = t & tArr2(i) & ";"
                    End If
                End If
                If DE Then DoEvents
            Next
            If DCount(s, ";") = 1 Then
                SmartFind = Left$(s, Len(s) - 1)
                If t <> "" Then
                    UnFormatted = Left$(t, Len(t) - 1)
                End If
                Exit Function
            ElseIf s <> "" Then
                Erase tArr
                SplitFast s, tArr, ";"
                SplitFast t, tArr2, ";"
                j = Len(tArr(LBound(tArr)))
                For i = LBound(tArr) To UBound(tArr)
                    If tArr(i) <> "" Then
                        If Len(tArr(i)) < j Then
                            j = Len(tArr(i))
                            k = i
                        End If
                    End If
                    If DE Then DoEvents
                Next
                SmartFind = tArr(k)
                If k < UBound(tArr2) Then
                    If tArr2(k) <> "" Then
                        UnFormatted = tArr2(k)
                    End If
                End If
                Exit Function
            End If
            SmartFind = "----1"
        End With
    Case 2
        dbIndex = GetPlayerIndexNumber(Index)
        Floor$ = modGetData.GetRoomItemsFromNums(Index, , , dbIndex)
        If modSC.FastStringComp(Floor$, "") Then SmartFind = SmartFindWhat: Exit Function
        SplitFast LCaseFast(Left$(Floor$, Len(Floor$) - 1)), tArr, ","
        For i = 0 To UBound(tArr)
            If InStr(1, ReplaceFast(tArr(i), Chr$(0), ""), SmartFindWhat) Then
                SmartFind = LCaseFast(tArr(i))
                If bPutUnFormattedInByRef And InStr(1, tArr(i), "note: ") = 0 Then
                    With dbMap(dbPlayers(dbIndex).lDBLocation)
                       SplitFast Left$(.sItems, Len(.sItems) - 1), tArr2, ";"
                    End With
                    UnFormatted = tArr2(i)
                End If
                Exit Function
            End If
            If DE Then DoEvents
        Next
        SmartFind = "----1"
    Case 3
        dbIndex = GetPlayerIndexNumber(Index)
        Monsters$ = modGetData.GetMonsHere(dbPlayers(dbIndex).lLocation, , dbIndex, dbPlayers(dbIndex).lDBLocation) & ";"
        If modSC.FastStringComp(Monsters$, ";") Then Monsters$ = ""
        Monsters$ = Monsters$ & modGetData.GetFamiliarsHere(dbPlayers(dbIndex).lLocation)
        Monsters$ = ReplaceFast(Monsters$, ", ", ";")
        Monsters$ = ReplaceFast(Monsters$, BRIGHTMAGNETA, "")
        Monsters$ = ReplaceFast(Monsters$, YELLOW, "")
        Monsters$ = ReplaceFast(Monsters$, BRIGHTBLUE, "")
        Monsters$ = ReplaceFast(Monsters$, LIGHTBLUE, "")
        If modSC.FastStringComp(Monsters$, "") Then SmartFind = SmartFindWhat: Exit Function
        SplitFast Left$(Monsters$, Len(Monsters$) - 1), tArr, ";"
        For i = 0 To UBound(tArr)
            If InStr(1, LCaseFast(tArr(i)), SmartFindWhat) Then
                SmartFind = LCaseFast(tArr(i))
                Exit Function
            End If
            If DE Then DoEvents
        Next
        SmartFind = "----1"
    Case 4
        dbIndex = GetPlayerIndexNumber(Index)
        With dbMap(dbPlayers(dbIndex).lDBLocation)
            ShopDBIndex = GetShopIndex(CLng(.sShopItems))
        End With
        If ShopDBIndex = 0 Then
            SmartFind = SmartFindWhat
            Exit Function
        End If
        For i = 0 To 14
            If dbShops(ShopDBIndex).iItems(i) <> 0 Then
                With dbItems(GetItemID(, CLng(dbShops(ShopDBIndex).iItems(i))))
                    ShopItems$ = ShopItems$ & .sItemName & ","
                End With
            End If
            If DE Then DoEvents
        Next
        If modSC.FastStringComp(ShopItems, "") Then SmartFind = SmartFindWhat: Exit Function
        SplitFast Left$(ShopItems$, Len(ShopItems$) - 1), tArr, ","
        For i = LBound(tArr) To UBound(tArr)
            If InStr(1, LCaseFast(tArr(i)), SmartFindWhat) Then
                SmartFind = LCaseFast(tArr(i))
                Exit Function
            End If
            If DE Then DoEvents
        Next
        SmartFind = "----1"
    Case 5
        dbIndex = GetPlayerIndexNumber(Index)
        Players$ = modGetData.GetPlayersHereWithoutRiding(dbPlayers(dbIndex).lLocation, dbIndex)
        If modSC.FastStringComp(Players$, "") Then SmartFind = SmartFindWhat: Exit Function
        SplitFast Left$(Players$, Len(Players$) - 1), tArr, ";"
        For i = 0 To UBound(tArr)
            tArr(i) = LCaseFast(tArr(i))
            If InStr(1, tArr(i), SmartFindWhat) <> 0 Then
                SmartFind = tArr(i)
                Exit Function
            End If
            If DE Then DoEvents
        Next
        SmartFind = "----1"
    Case 6
        dbIndex = GetPlayerIndexNumber(Index)
        Floor$ = modGetData.GetRoomHiddenItemsFromNums(Index, , , dbIndex) & modItemManip.GetListOfLettersFromHidden(dbPlayers(dbIndex).lDBLocation)
        If modSC.FastStringComp(Floor$, "") Or modSC.FastStringComp(Floor$, "0") Then SmartFind = SmartFindWhat: Exit Function
        SplitFast LCaseFast(Left$(Floor$, Len(Floor$) - 1)), tArr, ","
        For i = 0 To UBound(tArr)
            If InStr(1, ReplaceFast(tArr(i), Chr$(0), ""), SmartFindWhat) Then
                SmartFind = LCaseFast(tArr(i))
                If bPutUnFormattedInByRef And InStr(1, tArr(i), "note: ") = 0 Then
                    With dbMap(dbPlayers(dbIndex).lLocation)
                       SplitFast Left$(.sHidden, Len(.sHidden) - 1), tArr2, ";"
                    End With
                    UnFormatted = tArr2(i)
                End If
                Exit Function
            End If
            If DE Then DoEvents
        Next
        SmartFind = "----1"
    Case 7
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            With dbPlayers(i)
                If modSC.FastStringComp(LCaseFast(Left$(.sPlayerName, Len(SmartFindWhat))), LCaseFast(SmartFindWhat)) Or modSC.FastStringComp(LCaseFast(Left$(.sSeenAs, Len(SmartFindWhat))), LCaseFast(SmartFindWhat)) Then
                    SmartFindWhat = .sPlayerName
                    Exit For
                End If
            End With
            If DE Then DoEvents
        Next
        SmartFind = "----1"
    Case 8
        For i = LBound(dbItems) To UBound(dbItems)
            With dbItems(i)
                If modSC.FastStringComp(LCaseFast(Left$(.sItemName, Len(SmartFindWhat))), LCaseFast(SmartFindWhat)) Then
                    SmartFindWhat = .sItemName
                    Exit For
                End If
            End With
            If DE Then DoEvents
        Next
        SmartFind = "----1"
    Case 9
        For i = LBound(dbMonsters) To UBound(dbMonsters)
            If InStr(1, LCaseFast(dbMonsters(i).sMonsterName), SmartFindWhat) Then
                SmartFind = LCaseFast(dbMonsters(i).sMonsterName)
                Exit Function
            End If
            If DE Then DoEvents
        Next
        SmartFind = "----1"
End Select
SmartFind = LCaseFast(SmartFindWhat)
End Function
