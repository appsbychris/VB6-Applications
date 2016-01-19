Attribute VB_Name = "modHiddenItems"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modHiddenItems
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function Search(Index As Long) As Boolean
Dim ToSend As String
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "sea") Then
    If Len(X(Index)) > 3 Then
        If Not Mid$(LCaseFast(X(Index)), 4, 1) Like "[r]" Then
            Exit Function
        End If
    End If
    Search = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If .iHorse > 0 Then
            WrapAndSend Index, RED & "You can't search this area while atop your " & .sFamName & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .lRoomSearched = .lLocation
    End With
    With dbMap(dbPlayers(dbIndex).lDBLocation)
        ToSend = LIGHTBLUE & "You search the area..." & WHITE & vbCrLf
        If Not modSC.FastStringComp(.sHidden, "0") Or Not modSC.FastStringComp(.sHLetters, "0") Then
            ToSend = ToSend & GREEN & modGetData.GetRoomHiddenItemsFromNums(Index, True, True, dbIndex)
            ToSend = Left$(ToSend, Len(ToSend) - 2) & YELLOW & "." & WHITE & vbCrLf
        Else
            ToSend = ToSend & LIGHTBLUE & "Your search comes up empty." & WHITE & vbCrLf
        End If
        WrapAndSend Index, ToSend
        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " searches the area." & WHITE & vbCrLf, .lRoomID
        X(Index) = ""
    End With
End If
End Function

Public Function HideItem(Index As Long) As Boolean
Dim Eq As String
Dim a As Long
Dim iItemID As Long
Dim dbIndex As Long
'function to drop items from the inventory
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "hide ") Then  'if the command
    HideItem = True
    For a = 1 To InStr(1, X(Index), " ") 'trim off the command
        X(Index) = Mid$(X(Index), 2)
    Next
    Eq = TrimIt(X(Index)) 'get the eq they want to drop
    Eq = SmartFind(Index, Eq, Inventory_Item)
    If InStr(1, Eq, Chr$(0)) > 0 Then Eq = Mid$(Eq, InStr(1, Eq, Chr$(0)) + 1)
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If .iHorse > 0 Then
            WrapAndSend Index, RED & "You can't an item while atop your " & .sFamName & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        iItemID = GetItemID(Eq)
        If iItemID = 0 Then
            iItemID = GetLetterID(ReplaceFast(Eq, "note: ", "", 1, 1))
            If iItemID <> 0 Then
                modItemManip.TakeLetterFromInvAndHideIt dbIndex, dbLetters(iItemID).lID
                GoTo SendMsg
            Else
                'send error message
                WrapAndSend Index, RED & "You do not have that." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        End If
        iItemID = dbItems(iItemID).iID
        If InStr(1, .sInventory, ":" & iItemID & "/") Then
        
            modItemManip.TakeItemFromInvAndHideIt dbIndex, iItemID
            
        Else
            'send error message
            WrapAndSend Index, RED & "You do not have that." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    'send out messages
SendMsg:
    WrapAndSend Index, LIGHTBLUE & "Your hide your " & Eq & " in a very secret spot." & vbCrLf & WHITE
    X(Index) = ""
End If
End Function
