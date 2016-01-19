Attribute VB_Name = "modShop"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modShop
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function ShopCommands(Index As Long) As Boolean
'Function to check the shop commands
'selling
If SellItem(Index) = True Then ShopCommands = True: Exit Function
'list items in store
If ListItems(Index) = True Then ShopCommands = True: Exit Function
'buy items
If BuyAllFree(Index) = True Then ShopCommands = True: Exit Function
If PurchaseItem(Index) = True Then ShopCommands = True: Exit Function
If AppraiseItem(Index) = True Then ShopCommands = True: Exit Function
End Function

Public Function ListItems(Index As Long) As Boolean
'Function to list all the items in a store
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 4)), "list") Then  'keyword
    ListItems = True
    Dim Loc%
    Dim aryItems(14) As String 'tempary arrays
    Dim ToSend$
    Dim iMostLen As Long 'space between words
    Dim Templen As Long
    Dim aryCost(14) As String 'array for the cost of theitem
    Dim aryQuant(14) As String
    Dim Cost As String 'temp string
    Dim dbShopIndex As Long
    With dbMap(dbPlayers(GetPlayerIndexNumber(Index)).lDBLocation)
        If .iType <> 1 Then
            WrapAndSend Index, RED & "After looking around, you notice this is not a shop." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        Else
            dbShopIndex = GetShopIndex(CLng(.sShopItems))
        End If
    End With
    If dbShopIndex = 0 Then
        WrapAndSend Index, RED & "After looking around, you notice this is not a shop." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    For i = 0 To 14
        With dbShops(dbShopIndex)
            If .iQ(i) <> 0 Then
                aryItems(i) = .iItems(i)
                aryQuant(i) = .iQ(i)
            End If
        End With
        If DE Then DoEvents
    Next
    For i = 0 To 14
        If Not modSC.FastStringComp(aryItems(i), "") Then
            With dbItems(GetItemID(, CLng(aryItems(i))))
                aryItems(i) = YELLOW & "º " & GREEN & .sItemName
                Cost = RoundFast(.dCost + (.dCost * dbShops(dbShopIndex).iMarkUp / 100), 0)
                If modSC.FastStringComp(Cost, "0") Then
                    Cost = "Free"
                    aryCost(i) = BRIGHTYELLOW & Cost & GREEN & Space(6 - Len(Cost)) & "    "
                Else
                    aryCost(i) = BRIGHTYELLOW & Cost & GREEN & Space(6 - Len(Cost)) & "Gold"
                End If
                aryCost(i) = aryCost(i) & PlayerUseable(Index, .sItemName)
            End With
        End If
        If DE Then DoEvents
    Next
    For i = 0 To 14
        If Not modSC.FastStringComp(aryItems(i), "") Then
            Templen = Len(aryItems(i) & Space(35 - Len(aryItems(i))) & aryQuant(i) & Space(5 - Len(aryQuant(i))) & aryCost(i))
            If Templen - 6 > iMostLen Then iMostLen = Templen - 6
        End If
        If DE Then DoEvents
    Next
    For i = 0 To 14
        If Not modSC.FastStringComp(aryItems(i), "") Then
            ToSend$ = ToSend$ & aryItems(i) & Space(35 - Len(aryItems(i))) & BRIGHTBLUE & aryQuant(i) & Space(5 - Len(aryQuant(i))) & aryCost(i) & vbCrLf
        End If
        If DE Then DoEvents
    Next
    ToSend$ = YELLOW & "É" & String$(iMostLen - 2, "Í") & "»" & vbCrLf & ToSend$ & YELLOW & "È" & String$(iMostLen - 2, "Í") & "¼" & WHITE & vbCrLf
    ToSend$ = LIGHTBLUE & "Listing items..." & vbCrLf & YELLOW & "  Item Name" & Space(33 - Len("  Item Name")) & "#" & Space(4) & "Cost" & vbCrLf & ToSend$
    
    WrapAndSend Index, ToSend$
    X(Index) = ""
End If
End Function

Public Function PlayerUseable(Index As Long, Item As String) As String
'function to check if a player can use or wear the item
Dim Temp As Long
Dim dbIndex As Long
Dim dbItemIndex As Long
Dim lItemID As Long
dbIndex = GetPlayerIndexNumber(Index)
dbItemIndex = GetItemID(Item)
With dbItems(dbItemIndex)
    If modSC.FastStringComp(.sWorn, "scroll") Then
        Temp = CanUseSpell(dbIndex, dbItemIndex)
        If Temp = 0 Then
            If dbPlayers(dbIndex).iLevel < .lLevel Or dbPlayers(dbIndex).dClassPoints < .dClassPoints Then
                PlayerUseable = RED & " (Too Powerfull)  " & YELLOW & "º"
            Else
                PlayerUseable = RED & YELLOW & "                  º"
                Exit Function
            End If
        ElseIf Temp = 2 Then
            PlayerUseable = MAGNETA & " (Mastered)       " & YELLOW & "º"
        Else
            PlayerUseable = BRIGHTRED & " (You cannot use) " & YELLOW & "º"
        End If
        Exit Function
    End If
    lItemID = CLng(.iID)
End With
With dbPlayers(dbIndex)
    If .dClassPoints < dbItems(dbItemIndex).dClassPoints Then
        PlayerUseable = BRIGHTRED & " (You cannot use) " & YELLOW & "º"
        Exit Function
    End If
    If ClassCanWear(lItemID, .sClass) = False Then  'class
        PlayerUseable = BRIGHTRED & " (You cannot use) " & YELLOW & "º"
        Exit Function
    End If
    If RaceCanWear(lItemID, .sRace) = False Then 'race
        PlayerUseable = BRIGHTRED & " (You cannot use) " & YELLOW & "º"
        Exit Function
    End If
    If .iLevel < dbItems(dbItemIndex).lLevel Then  'level
        PlayerUseable = BRIGHTRED & " (You cannot use) " & YELLOW & "º"
        Exit Function
    End If
    If modWeaponsAndArmor.PlayerCanUseWeapon(CLng(dbIndex), CLng(dbItemIndex)) = False And dbItems(dbItemIndex).iType <> 0 Then  'weapon type
        PlayerUseable = BRIGHTRED & " (You cannot use) " & YELLOW & "º"
        Exit Function
    End If
    If modWeaponsAndArmor.PlayerCanUseArmor(CLng(dbIndex), CLng(dbItemIndex)) = False And dbItems(dbItemIndex).iArmorType <> 0 Then   'armour type
        PlayerUseable = BRIGHTRED & " (You cannot use) " & YELLOW & "º"
        Exit Function
    End If
    'if none of the checks failed, set this to false
    PlayerUseable = RED & YELLOW & "                  º"
End With
End Function


Public Function PurchaseItem(Index As Long) As Boolean
'function to buy items from the shop
Dim sItem$, sItemID$, sCost$, sItems$
Dim dbShopIndex As Long, lAmount As Long
Dim iDur As Long, iUses As Long
Dim i As Long
Dim s As String
Dim j As Long
Dim dbItemID As Long
Dim dbIndex As Long
lAmount = 1
s = LCaseFast(X(Index))
If s Like "buy #* *" Or X(Index) Like "bu #* *" Then
    PurchaseItem = True
    i = InStr(1, s, " ")
    j = InStr(i + 1, s, " ")
    lAmount = Val(Mid$(s, i + 1, j - i - 1))
    If lAmount = 0 Then
        WrapAndSend Index, RED & "Why would you want to buy 0 of something?" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    sItem = Mid$(s, j + 1)
    sItem = SmartFind(Index, sItem, Item_In_Shop)
    dbItemID = GetItemID(sItem$)
    If dbItemID = 0 And Not modSC.FastStringComp(sItem$, "paper") Then
        WrapAndSend Index, RED & "The owner has no idea what a " & sItem$ & " is." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    GoTo BuyTheItem
End If
If modSC.FastStringComp(Left$(s, 2), "bu") Then
    If Mid$(s, 3, 1) <> "y" Then
        If Mid$(s, 3, 1) <> " " Then
            Exit Function
        End If
    End If
        
    PurchaseItem = True
    If s = "buy" Then PurchaseItem = False: Exit Function
    sItem$ = Mid$(X(Index), 5, Len(s) - 4)
    sItem$ = SmartFind(Index, sItem$, Item_In_Shop)
    dbItemID = GetItemID(sItem$)
    If dbItemID = 0 And Not modSC.FastStringComp(sItem$, "paper") Then
        WrapAndSend Index, RED & "The owner has no idea what a " & sItem$ & " is." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
BuyTheItem:
    If modSC.FastStringComp(sItem$, "paper") Then GoTo CheckIt
    With dbItems(dbItemID)
        sItemID = .iID
        sCost$ = .dCost
        iUses = .iUses
        iDur = .lDurability
        If .iInGame >= .iLimit Then
            If dbMap(dbPlayers(dbIndex).lDBLocation).iType = 1 Then
                dbShopIndex = GetShopIndex(CLng(dbMap(dbPlayers(dbIndex).lDBLocation).sShopItems))
                For i = 0 To 14
                    If CLng(sItemID) = dbShops(dbShopIndex).iItems(i) Then
                        dbShops(dbShopIndex).iQ(i) = 0
                        Exit For
                    End If
                    If DE Then DoEvents
                Next
            End If
            WrapAndSend Index, RED & "The shop owner relizes he is out of stock on this item." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
CheckIt:
    With dbMap(dbPlayers(dbIndex).lDBLocation)
        If .iType = 1 Then
            dbShopIndex = GetShopIndex(CLng(.sShopItems))
            
            If modSC.FastStringComp(sItem, "paper") Then
                BuyPaper Index
                Exit Function
            End If
            For i = 0 To 14
                If dbShops(dbShopIndex).iItems(i) <> "0" Then
                    sItems$ = sItems$ & ":" & dbShops(dbShopIndex).iItems(i) & ";"
                End If
                If DE Then DoEvents
            Next
            If InStr(1, sItems$, ":" & sItemID & ";") = 0 Then
                WrapAndSend Index, RED & "The owner has no idea what a " & sItem$ & " is." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        Else
            WrapAndSend Index, RED & "After looking around, you notice this is not a shop." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    With dbShops(dbShopIndex)
        sCost = RoundFast(CDbl(sCost$) + (CDbl(sCost) * (.iMarkUp / 100)), 0) * lAmount
        For i = 0 To 14
            If CLng(sItemID) = .iItems(i) Then
                If .iQ(i) < lAmount Then
                    If .iQ(i) = 0 Then
                        WrapAndSend Index, RED & "That item is out of stock at the moment." & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Function
                    Else
                        WrapAndSend Index, RED & "The owner does not have " & lAmount & " " & modGetData.GetItemsNameAddS(dbItemID) & "." & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Function
                    End If
                End If
                Exit For
            End If
            If DE Then DoEvents
        Next
    End With
    With dbPlayers(dbIndex)
        If Not modSC.FastStringComp(sCost$, "0") Then
            sCost$ = RoundFast(CDbl(sCost$) - (.iCha / 10), 0)
            If CDbl(sCost$) < 1 Then sCost$ = CStr(lAmount)
        End If
        If .dGold >= CDbl(sCost$) Then
            If modGetData.GetPlayersTotalItems(Index, dbIndex) + lAmount > modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) Then
                WrapAndSend Index, RED & "You can't carry that much!" & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
            If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
            For i = 1 To lAmount
                .sInventory = .sInventory & ":" & dbItems(dbItemID).iID & "/" & iDur & "/E{}F{}A{}B{0|0|0|0}/" & iUses & ";"
                If dbItems(dbItemID).iLimit > 0 Then dbItems(dbItemID).iInGame = dbItems(dbItemID).iInGame + 1
                If DE Then DoEvents
            Next
            .dGold = .dGold - CDbl(sCost$)
            If lAmount = 1 Then
                WrapAndSend Index, LIGHTBLUE & "You purchase a " & sItem$ & " from the shop for " & sCost$ & " gold." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " purchases a " & sItem$ & ".", .lLocation
            Else
                WrapAndSend Index, LIGHTBLUE & "You purchase " & lAmount & " " & modGetData.GetItemsNameAddS(dbItemID) & " from the shop for " & sCost$ & " gold." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " purchases " & lAmount & " " & modGetData.GetItemsNameAddS(dbItemID) & ".", .lLocation
            End If
            X(Index) = ""
            For i = 0 To 14
                If CLng(sItemID) = dbShops(dbShopIndex).iItems(i) Then
                    dbShops(dbShopIndex).iQ(i) = dbShops(dbShopIndex).iQ(i) - lAmount
                    Exit For
                End If
                If DE Then DoEvents
            Next
            Exit Function
        Else
            If lAmount = 1 Then
                WrapAndSend Index, RED & "You cannot afford this item."
            Else
                WrapAndSend Index, RED & "You cannot afford " & lAmount & " of this item."
            End If
            X(Index) = ""
            Exit Function
        End If
    End With
End If
End Function

Public Function BuyAllFree(Index As Long) As Boolean
'function to buy items from the shop
Dim sItems$
Dim dbShopIndex As Long
Dim i As Long
Dim j As Long
Dim dbItemID As Long
Dim dbIndex As Long
Dim Arr() As String
Dim Message1 As String
Dim Message2 As String
If modSC.FastStringComp(LCaseFast(X(Index)), "buy all free") Then
    BuyAllFree = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbMap(GetMapIndex(dbPlayers(dbIndex).lLocation))
        If .iType = 1 Then
            dbShopIndex = GetShopIndex(CLng(.sShopItems))
        Else
            WrapAndSend Index, RED & "After looking around, you notice this is not a shop." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    With dbShops(dbShopIndex)
        For i = 0 To 14
            If .iQ(i) > 0 Then
                dbItemID = GetItemID(, .iItems(i))
                If dbItems(dbItemID).dCost = 0 Then sItems = sItems & dbItemID & ";"
            End If
            If DE Then DoEvents
        Next
    End With
    With dbPlayers(dbIndex)
        SplitFast sItems, Arr, ";"
        For i = LBound(Arr) To UBound(Arr)
            If Not modSC.FastStringComp(Arr(i), "") Then
                If modGetData.GetPlayersTotalItems(Index, dbIndex) + 1 > modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) Then
                    If Message1 <> "" Then
                        WrapAndSend Index, Message1 & RED & "You can't carry any more!" & WHITE & vbCrLf
                        SendToAllInRoom Index, Message2, .lLocation
                    Else
                        WrapAndSend Index, RED & "You can't carry that much!" & WHITE & vbCrLf
                    End If
                    X(Index) = ""
                    Exit Function
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                dbItemID = CLng(Arr(i))
                .sInventory = .sInventory & ":" & dbItems(dbItemID).iID & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(dbItemID).lDurability & "/" & dbItems(dbItemID).iUses & ";"
                For j = 0 To 14
                    If dbItems(dbItemID).iID = dbShops(dbShopIndex).iItems(i) Then
                        dbShops(dbShopIndex).iQ(i) = dbShops(dbShopIndex).iQ(i) - 1
                        Exit For
                    End If
                    If DE Then DoEvents
                Next
                Message1 = Message1 & LIGHTBLUE & "You purchase a " & dbItems(dbItemID).sItemName & " from the shop for nothing." & WHITE & vbCrLf
                Message2 = Message2 & LIGHTBLUE & .sPlayerName & " purchases a " & dbItems(dbItemID).sItemName & "." & WHITE & vbCrLf
            End If
        Next
        WrapAndSend Index, Message1
        SendToAllInRoom Index, Message2, .lLocation
        X(Index) = ""
    End With
End If
End Function

Public Function SellItem(Index As Long) As Boolean
Dim sItem$, sItemID$, sCost$, sLoc$
Dim dbpID As Long
Dim dbShopIndex As Long
Dim dSell As Double
Dim lAmount As Long
Dim s As String
Dim i As Long
Dim j As Long
Dim dbItemID As Long
Dim lSold As Long
s = LCaseFast(X(Index))
lAmount = 1
If s Like "sell #* *" Then
    SellItem = True
    i = InStr(1, s, " ")
    j = InStr(i + 1, s, " ")
    lAmount = Val(Mid$(s, i + 1, j - i - 1))
    If lAmount = 0 Then
        WrapAndSend Index, RED & "You can't sell 0 of something!" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    sItem = Mid$(s, j + 1)
    sItem = SmartFind(Index, sItem, Inventory_Item)
    If InStr(1, sItem$, Chr$(0)) > 0 Then sItem$ = Mid$(sItem$, InStr(1, sItem$, Chr$(0)) + 1)
    dbItemID = GetItemID(sItem$)
    If dbItemID = 0 Then
        WrapAndSend Index, RED & "The owner has no idea what a " & sItem$ & " is." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    GoTo SellIt
End If
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 4)), "sell") Then
    SellItem = True
    If LCaseFast(X(Index)) = "sell" Then SellItem = False: Exit Function
    sItem$ = Mid$(X(Index), 6, Len(X(Index)) - 5)
    sItem$ = SmartFind(Index, sItem$, Inventory_Item)
    If InStr(1, sItem$, Chr$(0)) > 0 Then sItem$ = Mid$(sItem$, InStr(1, sItem$, Chr$(0)) + 1)
    dbItemID = GetItemID(sItem$)
    If dbItemID = 0 Then
        WrapAndSend Index, RED & "You look all over you for your " & sItem$ & ", yet you cannot find it." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
SellIt:
    dbpID = GetPlayerIndexNumber(Index)
    With dbItems(dbItemID)
        sItemID = .iID
        sCost = .dCost
        sItem$ = .sItemName
    End With
    With dbPlayers(dbpID)
        If InStr(1, .sInventory, ":" & sItemID & "/") = 0 Then
            WrapAndSend Index, RED & "You look all over you for your " & sItem$ & ", yet you cannot find it." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sLoc$ = .lLocation
    End With
    With dbMap(GetMapIndex(CLng(sLoc$)))
        If .iType = 1 Then
            dbShopIndex = GetShopIndex(CLng(.sShopItems))
            For i = 0 To 14
                If Not modSC.FastStringComp(CStr(dbShops(dbShopIndex).iItems(i)), "0") Then
                    sItems$ = sItems$ & ":" & dbShops(dbShopIndex).iItems(i) & ";"
                End If
                If DE Then DoEvents
            Next
            If InStr(1, sItems$, ":" & sItemID & ";") = 0 Then
                WrapAndSend Index, RED & "The owner does not the sell that item, so they cannot purchase it from you." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        Else
            WrapAndSend Index, RED & "After looking around, you notice this is not a shop." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    With dbShops(dbShopIndex)
        sCost$ = (CDbl(sCost$) + (CDbl(sCost$) * (.iMarkUp / 100))) - (dbPlayers(dbpID).iCha \ 10)
    End With
    With dbPlayers(dbpID)
        For i = 1 To lAmount
            If InStr(1, .sInventory, ":" & sItemID & "/") <> 0 Then
                modItemManip.RemoveItemFromInv dbpID, CLng(sItemID)
                lSold = lSold + 1
            End If
            If DE Then DoEvents
        Next
        If Not modSC.FastStringComp(sCost$, "0") Then
            dSell = (CDbl(sCost$) / 3) + (.iCha / 10)
            dSell = RoundFast(dSell, 0)
            If dSell > CDbl(sCost$) Then dSell = CDbl(sCost$) - 1
            If dSell < 0 Then dSell = 0
            dSell = dSell * lSold
            .dGold = .dGold + dSell
        End If
        With dbItems(dbItemID)
            If .iLimit > 0 Then
                If .iInGame > 0 Then
                    .iInGame = .iInGame - 1
                End If
            End If
        End With
        If lSold > 1 Then
            WrapAndSend Index, LIGHTBLUE & "You sell " & lSold & " " & modGetData.GetItemsNameAddS(dbItemID) & " to the shop owner for " & CStr(dSell) & " gold." & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " sells " & lSold & " " & modGetData.GetItemsNameAddS(dbItemID) & " to the shop owner." & WHITE & vbCrLf, .lLocation
        Else
            WrapAndSend Index, LIGHTBLUE & "You sell your " & sItem$ & " to the shop owner for " & CStr(dSell) & " gold." & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " sells thier " & sItem$ & " to the shop owner." & WHITE & vbCrLf, .lLocation
        End If
        X(Index) = ""
        For i = 0 To 14
            If CLng(sItemID) = dbShops(dbShopIndex).iItems(i) Then
                dbShops(dbShopIndex).iQ(i) = dbShops(dbShopIndex).iQ(i) + lSold
                Exit For
            End If
            If DE Then DoEvents
        Next
        Exit Function
    End With
End If
End Function

Public Function CanUseSpell(dbIndex As Long, dbItemID As Long) As Long
Dim aFlgs() As String
Dim WasUsed As Long
Dim dVal As Double
WasUsed = 1
If InStr(1, dbItems(dbItemID).sFlags, "gsp") <> 0 Then
    SplitFast dbItems(dbItemID).sFlags, aFlgs, ";"
    For i = LBound(aFlgs) To UBound(aFlgs)
        dVal = CDbl(Val(Mid$(aFlgs(i), 4)))
        Select Case Left$(aFlgs(i), 3)
            Case "gsp"
                iSpellID = GetSpellID(, CLng(dVal))
                If iSpellID = 0 Then GoTo tNext
                If dbPlayers(dbIndex).iSpellLevel >= dbSpells(iSpellID).iLevel Then
                    If dbPlayers(dbIndex).iSpellType = dbSpells(iSpellID).iType Then
                        WasUsed = 0
                        If InStr(1, dbPlayers(dbIndex).sSpells, ":" & dbSpells(iSpellID).lID & ";") Then
                            
                            WasUsed = 2
                        End If
                        
                    Else
                        
                        WasUsed = 1
                    End If
                Else
                    
                    WasUsed = 1
                End If
        End Select
tNext:
        If DE Then DoEvents
    Next
End If
CanUseSpell = WasUsed
'With dbSpells(GetSpellID(, SpellID))
'    If InStr(1, dbPlayers(dbpID).sSpells, ":" & SpellID & ";") Then
'        CanUseSpell = 2
'        Exit Function
'    End If
'    If dbPlayers(dbpID).iSpellLevel >= .iLevel Then
'        If .iType = dbPlayers(dbpID).iSpellType Then
'            CanUseSpell = 0
'            Exit Function
'        Else
'            CanUseSpell = 1
'            Exit Function
'        End If
'    Else
'        CanUseSpell = 1
'    End If
'End With
End Function

Sub BuyPaper(Index As Long)
Dim dbIndex As Long
Dim dPrice As Double
dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    dPrice = 1 + (dbShops(GetShopIndex(CLng(dbMap(GetMapIndex(.lLocation)).sShopItems))).iMarkUp / 100) - (.iCha / 10)
    dPrice = RoundFast(dPrice, 0)
    If dPrice < 1 Then dPrice = 1
    If .dGold >= dPrice Then
        .dGold = .dGold - dPrice
        .lPaper = .lPaper + 10
        WrapAndSend Index, LIGHTBLUE & "You buy 10 pieces of paper for " & CStr(dPrice) & " gold." & WHITE & vbCrLf
    Else
        WrapAndSend Index, RED & "You cannot afford the price of " & CStr(dPrice) & " gold." & WHITE & vbCrLf
    End If
End With
X(Index) = ""
End Sub

Public Function AppraiseItem(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 9)), "appraise ") Then
    AppraiseItem = True
    Dim sItem$, sItemID$, sCost$, sInv$, sLoc$
    Dim dbShopItem As Long
    Dim dbpID As Long
    Dim dbShopIndex As Long
    dbpID = GetPlayerIndexNumber(Index)
    sItem$ = Mid$(X(Index), 10, Len(X(Index)) - 5)
    sItem$ = SmartFind(Index, sItem$, Inventory_Item)
    If InStr(1, sItem$, Chr$(0)) > 0 Then sItem$ = Mid$(sItem$, InStr(1, sItem$, Chr$(0)) + 1)
    If GetItemID(sItem$) = 0 Then
        WrapAndSend Index, RED & "You look all over you for your " & sItem$ & ", yet you cannot find it." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbItems(GetItemID(sItem$))
        sItemID = .iID
        sCost = .dCost
        sItem$ = .sItemName
    End With
    With dbPlayers(dbpID)
        If InStr(1, .sInventory, ":" & sItemID & "/") = 0 Then
            WrapAndSend Index, RED & "You look all over you for your " & sItem$ & ", yet you cannot find it." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sLoc$ = .lLocation
    End With
    With dbMap(GetMapIndex(CLng(sLoc$)))
        If .iType = 1 Then
            dbShopIndex = GetShopIndex(CLng(.sShopItems))
            For i = 0 To 14
                If Not modSC.FastStringComp(CStr(dbShops(dbShopIndex).iItems(i)), "0") Then
                    sItems$ = sItems$ & ":" & dbShops(dbShopIndex).iItems(i) & ";"
                End If
                If DE Then DoEvents
            Next
            If InStr(1, sItems$, ":" & sItemID & ";") = 0 Then
                WrapAndSend Index, RED & "The owner does not the sell that item, so they cannot appraise it for you." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        Else
            WrapAndSend Index, RED & "After looking around, you notice this is not a shop." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    With dbShops(dbShopIndex)
        sCost$ = CDbl(sCost$) + (CDbl(sCost$) * (.iMarkUp / 100))
    End With
    With dbPlayers(dbpID)
        If Not modSC.FastStringComp(sCost$, "0") Then
            sCost$ = (CDbl(sCost$) / 3) + (.iCha / 10)
            sCost$ = RoundFast(CDbl(sCost$), 0)
        End If
'        Dim FuzzBy As Double
'        FuzzBy = CDbl(RndNumber(1, 100 - .iInt))
'        If FuzzBy <= 0 Then FuzzBy = 1
'        FuzzBy = FuzzBy / 100
'        sCost$ = (CDbl(sCost$) * FuzzBy) * 100
'        sCost$ = RoundFast(CDbl(sCost$), 0)
        WrapAndSend Index, LIGHTBLUE & "You can sell your " & sItem$ & " to the shop owner for " & sCost$ & " gold." & WHITE & vbCrLf
        X(Index) = ""
    End With
End If
End Function
