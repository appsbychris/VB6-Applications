Attribute VB_Name = "modGiveItems"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modGiveItems
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function GiveItems(Index As Long) As Boolean
Dim sReciever As String
Dim sRecIndex As Long
Dim sItem As String
Dim iItemID As Long
Dim Eq As String, bGold As Boolean
Dim a As Long, iID As Long
Dim dbpID As Long
Dim bL As Boolean
Dim lAmount As Long
Dim dbMapId As Long
Dim bTriedItem As Boolean
Dim i As Long
Dim j As Long
Dim lamD As Long
Dim lDropped As Long
Dim s As String
Dim dTemp As Long
s = LCaseFast(X(Index))
'X(Index) = LCaseFast(X(Index))
lamD = 1
If s Like "give #* * to *" Then
    GiveItems = True
    i = InStr(1, s, " ")
    j = InStr(i + 1, s, " ")
    lamD = Val(Mid$(s, i + 1, j - i - 1))
    If lamD = 0 Then
        WrapAndSend Index, RED & "Why would you try to give 0 of something?" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    X(Index) = s
    X(Index) = Mid$(X(Index), 6)
    i = InStr(1, X(Index), " ") + 1
    sItem = TrimIt(Mid$(X(Index), i, InStr(1, X(Index), " to") - 3))
    sReciever = TrimIt(Mid$(X(Index), InStr(i, X(Index), " to") + 3, Len(X(Index)) - InStr(i, X(Index), " to") + 3))
    GoTo GiveItemNow
End If
If modSC.FastStringComp(Left$(s, 5), "give ") Then
    If InStr(6, X(Index), " to") = 0 Then Exit Function
    GiveItems = True
    X(Index) = s
    sItem = TrimIt(Mid$(X(Index), 5, InStr(6, X(Index), " to") - 4))
    sReciever = TrimIt(Mid$(X(Index), InStr(6, X(Index), " to") + 3, Len(X(Index)) - InStr(6, X(Index), " to") + 3))
    lamD = 1
GiveItemNow:
    sReciever = SmartFind(Index, sReciever, Player_In_Room)
    sItem = SmartFind(Index, sItem, Inventory_Item)
    If InStr(1, sItem, Chr$(0)) > 0 Then sItem = Mid$(sItem, InStr(1, sItem, Chr$(0)) + 1)
    sRecIndex = GetPlayerIndexNumber(, sReciever)
    dbpID = GetPlayerIndexNumber(Index)
    If sRecIndex = 0 Then
        WrapAndSend Index, RED & sReciever & " has seemed to vanished, as you cannot find them." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    Eq = sItem
    If Left$(LCaseFast(Eq), 1) = "g" Then
        If Len(Eq) > 1 Then
            If Mid$(LCaseFast(Eq), 2, 1) <> "o" Then
                bGold = False
            Else
                If Len(Eq) > 2 Then
                    If Mid$(LCaseFast(Eq), 3, 1) = "l" Then
                        If Len(Eq) > 3 Then
                            If Mid$(LCaseFast(Eq), 4, 1) = "d" Then
                                If Len(Eq) > 4 Then
                                    bGold = False
                                Else
                                    If lamD > 1 Then
                                        lAmount = lamD
                                        bGold = True
                                    Else
                                        lAmount = -1
                                        bGold = True
                                    End If
                                End If
                            Else
                                If lamD > 1 Then
                                    lAmount = lamD
                                    bGold = True
                                Else
                                    lAmount = -1
                                    bGold = True
                                End If
                            End If
                        Else
                            If lamD > 1 Then
                                lAmount = lamD
                                bGold = True
                            Else
                                lAmount = -1
                                bGold = True
                            End If
                        End If
                    End If
                Else
                    If lamD > 1 Then
                        lAmount = lamD
                        bGold = True
                    Else
                        lAmount = -1
                        bGold = True
                    End If
                End If
            End If
        Else
            If lamD > 1 Then
                lAmount = lamD
                bGold = True
            Else
                lAmount = -1
                bGold = True
            End If
        End If
    ElseIf InStr(1, Eq, " ") <> 0 Then
        lAmount = Val(Mid$(Eq, 1, InStr(1, Eq, " ") - 1))
        bGold = True
        For i = 1 To Len(CStr(lAmount))
            If IsANumber(Asc(Mid$(Eq, i, 1))) = False Then
                bGold = False
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
    
    If bGold = False Then
TryItem:
        iID = GetItemID(Eq)
        If iID = 0 Then
            iID = GetLetterID(ReplaceFast(Eq, "note: ", "", 1, 1))
            If iID = 0 Then
                WrapAndSend Index, RED & "You do not have that." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            Else
                bL = True
                GoTo GiveLetter
            End If
        End If
        iID = dbItems(iID).iID
    End If
    With dbPlayers(dbpID)
        If InStr(1, .sInventory, ":" & iID & "/") <> 0 Then
            For i = 1 To lamD
                If InStr(1, .sInventory, ":" & iID & "/") <> 0 Then
                    If InStr(1, LCaseFast(modGetData.GetPlayersHereWithoutRiding(CLng(.lLocation), dbpID)), LCaseFast(sReciever)) = 0 Then
                        WrapAndSend Index, RED & sReciever & " has seemed to vanished, as you cannot find them." & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Function
                    End If
                    If modGetData.GetPlayersTotalItems(dbPlayers(sRecIndex).iIndex, sRecIndex) + 1 > modGetData.GetPlayersMaxItems(dbPlayers(sRecIndex).iIndex, sRecIndex) Then
                        WrapAndSend Index, RED & "They are carring too many things to accept your offer." & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Function
                    End If
                    modItemManip.TakeFromYourInvAndPutInAnothersInv dbpID, sRecIndex, iID
                    lDropped = lDropped + 1
                End If
                If DE Then DoEvents
            Next
        ElseIf bGold = True Then
            If .dGold = 0 And bTriedItem = False Then
                bTriedItem = True
                GoTo TryItem
            ElseIf (lAmount > .dGold) Or bTriedItem Then
                WrapAndSend Index, RED & "You do not have that much gold." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            Else
                If lAmount > 0 Then
                    If dbPlayers(sRecIndex).dGold + lAmount > modGetData.GetPlayersMaxGold(dbPlayers(sRecIndex).iIndex, sRecIndex) Then
                        dTemp = dbPlayers(sRecIndex).dGold
                        dTemp = modGetData.GetPlayersMaxGold(dbPlayers(sRecIndex).iIndex, sRecIndex) - dTemp
                        With dbPlayers(sRecIndex)
                            .dGold = .dGold + dTemp
                        End With
                        .dGold = .dGold - dTemp
                    Else
                        dTemp = lAmount
                        With dbPlayers(sRecIndex)
                            .dGold = .dGold + dTemp
                        End With
                        .dGold = .dGold - dTemp
                    End If
                ElseIf lAmount = -1 Then
                    If dbPlayers(sRecIndex).dGold + .dGold > modGetData.GetPlayersMaxGold(dbPlayers(sRecIndex).iIndex, sRecIndex) Then
                        dTemp = dbPlayers(sRecIndex).dGold
                        dTemp = modGetData.GetPlayersMaxGold(dbPlayers(sRecIndex).iIndex, sRecIndex) - dTemp
                        With dbPlayers(sRecIndex)
                            .dGold = .dGold + dTemp
                        End With
                        .dGold = .dGold - dTemp
                    Else
                        With dbPlayers(sRecIndex)
                            .dGold = .dGold + dbPlayers(dbpID).dGold
                            lAmount = dbPlayers(dbpID).dGold
                        End With
                        .dGold = 0
                    End If
                Else
                    WrapAndSend Index, RED & "You can't give 0 of something." & WHITE & vbCrLf
                    X(Index) = ""
                    Exit Function
                End If
                Eq = CStr(lAmount) & " gold"
            End If
        Else
            'send error message
            WrapAndSend Index, RED & "You do not have that." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
GiveLetter:
        If (iID <> 0) And bL Then
            modItemManip.TakeLetterFromInvAndPutInAnotherInv dbpID, sRecIndex, dbLetters(iID).lID
        End If
        If Not bGold Then
            If lDropped > 1 Then
                iID = GetItemID(, iID)
                WrapAndSend Index, LIGHTBLUE & "You give " & lDropped & " " & modGetData.GetItemsNameAddS(iID) & " to " & dbPlayers(sRecIndex).sPlayerName & "." & vbCrLf & WHITE
                WrapAndSend dbPlayers(sRecIndex).iIndex, LIGHTBLUE & .sPlayerName & " just handed you " & lDropped & " " & modGetData.GetItemsNameAddS(iID) & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbpID).sPlayerName & " gives " & dbPlayers(sRecIndex).sPlayerName & " something." & vbCrLf & WHITE, dbPlayers(dbpID).lLocation, dbPlayers(sRecIndex).iIndex
            Else
                WrapAndSend Index, LIGHTBLUE & "You give away your " & Eq & " to " & sReciever & "." & WHITE & vbCrLf
                WrapAndSend dbPlayers(sRecIndex).iIndex, LIGHTBLUE & dbPlayers(dbpID).sPlayerName & " just handed you their " & Eq & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbpID).sPlayerName & " just handed " & sReciever & " something." & WHITE & vbCrLf, dbPlayers(dbpID).lLocation, dbPlayers(sRecIndex).iIndex
            End If
        Else
            WrapAndSend Index, LIGHTBLUE & "You give away your " & Eq & " to " & sReciever & "." & WHITE & vbCrLf
            WrapAndSend dbPlayers(sRecIndex).iIndex, LIGHTBLUE & dbPlayers(dbpID).sPlayerName & " just handed you their " & Eq & "." & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " just handed " & sReciever & " something." & WHITE & vbCrLf, dbPlayers(dbpID).lLocation, dbPlayers(sRecIndex).iIndex
        End If
        X(Index) = ""
    End With
End If
End Function
