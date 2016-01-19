Attribute VB_Name = "modInventory"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modInventory
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Function Inventory(Index As Long) As Boolean
'////////INVENTORY////////
'Function to get the players inventory
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 1)), "i") Then   ''i' is the inventory command
    X(Index) = TrimIt(X(Index))
    If Len(X(Index)) > 1 Then
        If Not LCaseFast(Mid$(X(Index), 2, 1)) Like "[n]" Then
            Exit Function
        ElseIf Len(X(Index)) > 2 Then
            If Not LCaseFast(Mid$(X(Index), 3, 1)) Like "[v]" Then
                Exit Function
            ElseIf Len(X(Index)) > 3 Then
                If Not LCaseFast(Mid$(X(Index), 3, 1)) Like "[e]" Then
                    Exit Function
                End If
            End If
        End If
    End If
    Inventory = True
    Dim ToSend$ 'variables for values
    ToSend$ = MAGNETA & "You have upon you:" & vbCrLf 'start the message
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If (.dGold <> 0) And (.lPaper = 0) Then
            ToSend$ = ToSend$ & YELLOW & "You have " & GREEN & .dGold & YELLOW & " gold,"
        ElseIf (.dGold = 0) And (.lPaper <> 0) Then
            ToSend$ = ToSend$ & YELLOW & "You have " & GREEN & .lPaper & YELLOW & " pieces of paper,"
        ElseIf (.dGold <> 0) And (.lPaper <> 0) Then
            ToSend$ = ToSend$ & YELLOW & "You have " & GREEN & .dGold & YELLOW & " gold," & GREEN & .lPaper & YELLOW & " pieces of paper,"
        End If
        If modSC.FastStringComp(.sInventory, "") Then .sInventory = "0"
        If .sInventory <> "0" Then
            ToSend$ = ToSend$ & GREEN & modGetData.GetPlayersInvFromNums(Index, True, dbIndex)
        End If
        ToSend$ = ToSend$ & GREEN
        'get all the equipment they are wearing
        ToSend$ = ToSend$ & modGetData.GetPlayersEqFromNums(Index, , dbIndex) & modItemManip.GetListOfLettersFromInv(dbIndex)
        If Not modSC.FastStringComp(ToSend$, MAGNETA & "You have upon you:" & vbCrLf & GREEN) Then
            ToSend$ = ReplaceFast(ToSend$, ",", YELLOW & ", " & GREEN) 'format the message
            ToSend$ = ReplaceFast(ToSend, YELLOW & ", " & GREEN, YELLOW & "." & GREEN, Len(ToSend$) - 4, 1)
            'ToSend$ = Left$(ToSend$, Len(ToSend$) - 3) & YELLOW & "." 'finish the message
            ToSend$ = GREEN & ToSend$ & WHITE & vbCrLf 'finish up the message
        Else
            ToSend$ = ToSend$ & "Absolutly nothing." & WHITE & vbCrLf
        End If
        ToSend = ToSend & YELLOW & "Enc: " & LIGHTBLUE & modGetData.GetPlayersTotalItems(Index, _
            dbIndex) & YELLOW & "/" & LIGHTBLUE & modGetData.GetPlayersMaxItems(Index, dbIndex) & vbCrLf
        ToSend = ToSend & YELLOW & "You can carry a maximum of " & LIGHTBLUE & modGetData.GetPlayersMaxGold(Index, dbIndex) & YELLOW & " gold pieces." & WHITE & vbCrLf
        WrapAndSend Index, ToSend$ 'send to the player
        X(Index) = ""
    End With
End If
'////////END////////
End Function


