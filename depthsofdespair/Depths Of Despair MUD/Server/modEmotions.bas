Attribute VB_Name = "modEmotions"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modEmotions
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Function Emotes(Index As Long, Optional dbIndex As Long) As Boolean
'Function for emotions
Dim TempEmote As String, TempPlayer As String 'temp vars to hold values
Dim bFound As Boolean 'flag
Dim ieID As Long
Dim i As Long

Dim dbTIndex As Long
Dim dbGen As Long
If InStr(1, X(Index), " ") Then 'if there is a target
    TempEmote = Left$(X(Index), InStr(1, X(Index), " ") - 1)
    TempPlayer = SmartFind(Index, Mid$(X(Index), InStr(1, X(Index), " ") + 1, Len(X(Index)) - InStr(1, X(Index), " ")), Player_In_Room)
Else
    TempEmote = X(Index) 'get the emotions
End If
For i = LBound(dbEmotions) To UBound(dbEmotions) 'see if its a emotion
    If modSC.FastStringComp(LCaseFast(TempEmote), LCaseFast(dbEmotions(i).sSyntax)) Then
        bFound = True 'set the flag
        ieID = i
        Exit For
    End If
    If DE Then DoEvents
Next
If bFound Then 'if there is the emotion
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    Emotes = True
    If modSC.FastStringComp(TempPlayer, "") Then
        'send stuff out
        WrapAndSend Index, GREEN & dbEmotions(ieID).sPhraseYou & WHITE & vbCrLf
        SendToAllInRoom Index, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseOthers, "<player>", dbPlayers(dbIndex).sPlayerName) & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
        X(Index) = ""
    Else
        'if there was a target
        If GetPlayerIndexNumber(, TempPlayer) = 0 Then
            TempPlayer = SmartFind(Index, TempPlayer, Monster_In_Room)
            If InStr(1, TempPlayer, Chr$(0)) > 0 Then TempPlayer = Mid$(TempPlayer, InStr(1, TempPlayer, Chr$(0)) + 1)
            dbGen = GetMonsterID(TempPlayer)
            If dbGen = 0 Then
                dbGen = GetFamID(, TempPlayer)
                If dbGen = 0 Then
                    TempPlayer = SmartFind(Index, TempPlayer, Inventory_Item)
                    dbGen = GetItemID(TempPlayer)
                    If dbGen = 0 Then
TryOther:
                        TempPlayer = SmartFind(Index, TempPlayer, Equiped_Item)
                        dbGen = GetItemID(TempPlayer)
                        If dbGen = 0 Then
TryAgain:
                            TempPlayer = SmartFind(Index, TempPlayer, Item_In_Room)
                            dbGen = GetItemID(TempPlayer)
                            If dbGen = 0 Then
                                WrapAndSend Index, RED & "That player is not in this room!" & WHITE & vbCrLf
                                X(Index) = ""
                                Exit Function
                            Else
                                If InStr(1, modGetData.GetItemsHere(dbPlayers(dbIndex).lLocation, dbPlayers(dbIndex).lDBLocation), ":" & dbItems(dbGen).iID & "/") <> 0 Then
                                    WrapAndSend Index, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseYouToOther, "<victim>", TempPlayer) & WHITE & vbCrLf
                                    SendToAllInRoom Index, GREEN & ReplaceFast(ReplaceFast(dbEmotions(ieID).sPhraseOthers2, "<player>", dbPlayers(dbIndex).sPlayerName), "<victim>", TempPlayer) & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
                                    X(Index) = ""
                                    Exit Function
                                Else
                                    WrapAndSend Index, RED & "That player is not in this room!" & WHITE & vbCrLf
                                    X(Index) = ""
                                    Exit Function
                                End If
                            End If
                        Else
                            If InStr(1, modGetData.GetPlayersEq(Index, dbIndex), ":" & dbItems(dbGen).iID & "/") <> 0 Then
                                WrapAndSend Index, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseYouToOther, "<victim>", "your " & TempPlayer) & WHITE & vbCrLf
                                SendToAllInRoom Index, GREEN & ReplaceFast(ReplaceFast(dbEmotions(ieID).sPhraseOthers2, "<player>", dbPlayers(dbIndex).sPlayerName), "<victim>", "their " & TempPlayer) & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
                                X(Index) = ""
                                Exit Function
                            Else
                                GoTo TryAgain
                            End If
                        End If
                    Else
                        If InStr(1, dbPlayers(dbIndex).sInventory, ":" & dbItems(dbGen).iID & "/") <> 0 Then
                            WrapAndSend Index, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseYouToOther, "<victim>", "your " & TempPlayer) & WHITE & vbCrLf
                            SendToAllInRoom Index, GREEN & ReplaceFast(ReplaceFast(dbEmotions(ieID).sPhraseOthers2, "<player>", dbPlayers(dbIndex).sPlayerName), "<victim>", "their " & TempPlayer) & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
                            X(Index) = ""
                            Exit Function
                        Else
                            GoTo TryOther
                        End If
                    End If
                Else
                    If InStr(1, modGetData.GetFamiliarsHere(dbPlayers(dbIndex).lLocation), dbFamiliars(dbGen).sFamName & YELLOW & ", ") <> 0 Then
                        WrapAndSend Index, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseYouToOther, "<victim>", TempPlayer) & WHITE & vbCrLf
                        SendToAllInRoom Index, GREEN & ReplaceFast(ReplaceFast(dbEmotions(ieID).sPhraseOthers2, "<player>", dbPlayers(dbIndex).sPlayerName), "<victim>", TempPlayer) & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
                        X(Index) = ""
                        Exit Function
                    Else
                        WrapAndSend Index, RED & "That player is not in this room!" & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Function
                    End If
                End If
            Else
                If InStr(1, modGetData.GetMonsHere(dbPlayers(dbIndex).lLocation, , dbIndex, dbPlayers(dbIndex).lDBLocation), dbMonsters(dbGen).sMonsterName) <> 0 Then
                    WrapAndSend Index, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseYouToOther, "<victim>", TempPlayer) & WHITE & vbCrLf
                    SendToAllInRoom Index, GREEN & ReplaceFast(ReplaceFast(dbEmotions(ieID).sPhraseOthers2, "<player>", dbPlayers(dbIndex).sPlayerName), "<victim>", TempPlayer) & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
                    X(Index) = ""
                    Exit Function
                Else
                    WrapAndSend Index, RED & "That player is not in this room!" & WHITE & vbCrLf
                    X(Index) = ""
                    Exit Function
                End If
            End If
            WrapAndSend Index, RED & "That player is not in this room!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        dbTIndex = GetPlayerIndexNumber(, TempPlayer)
        With dbPlayers(dbTIndex)
            If .iIndex = 0 Then
                WrapAndSend Index, RED & "That player is not online!" & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            ElseIf modMiscFlag.GetMiscFlag(dbTIndex, Invisible) = 1 Then
                WrapAndSend Index, RED & "That player is not here!" & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            ElseIf dbPlayers(dbTIndex).iGhostMode = 1 Then
                WrapAndSend Index, RED & "That player is not here!" & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            ElseIf .iSneaking <> 0 Then
                WrapAndSend Index, RED & "That player is not here!" & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        End With
        If dbPlayers(dbIndex).lLocation <> dbPlayers(dbTIndex).lLocation Then  'check if the player is in the room
            'send error message
            WrapAndSend Index, RED & "That player is not in this room!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        'send out messages to everyone
        If dbIndex <> dbTIndex Then
            WrapAndSend Index, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseYouToOther, "<victim>", dbPlayers(dbTIndex).sPlayerName) & WHITE & vbCrLf
            WrapAndSend dbPlayers(dbTIndex).iIndex, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseToYou, "<player>", dbPlayers(dbIndex).sPlayerName) & WHITE & vbCrLf
            SendToAllInRoom Index, GREEN & ReplaceFast(ReplaceFast(dbEmotions(ieID).sPhraseOthers2, "<player>", dbPlayers(dbIndex).sPlayerName), "<victim>", dbPlayers(dbTIndex).sPlayerName) & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation, dbPlayers(dbTIndex).iIndex
        Else
            WrapAndSend Index, GREEN & ReplaceFast(dbEmotions(ieID).sPhraseYouToOther, "<victim>", "yourself") & WHITE & vbCrLf
            SendToAllInRoom Index, GREEN & ReplaceFast(ReplaceFast(dbEmotions(ieID).sPhraseOthers2, "<player>", dbPlayers(dbIndex).sPlayerName), "<victim>", modGetData.GetGenderPronoun(dbIndex, True) & "self") & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
        End If
        X(Index) = ""
    End If
Else
    Emotes = False 'set this falue if its not an emotion
End If
End Function

Public Function ListEmotes(Index As Long) As Boolean
'function to list all available emotions
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 8)), "emotions") Then   'if the keyword
    ListEmotes = True
    Dim ToSend$
    ToSend$ = YELLOW & "List of emotions:" & vbCrLf 'set the message
    For i = LBound(dbEmotions) To UBound(dbEmotions)
        ToSend$ = ToSend$ & LIGHTBLUE & dbEmotions(i).sSyntax & YELLOW & ", "  'get the emotions
        If DE Then DoEvents
    Next
    WrapAndSend Index, Left$(ToSend$, Len(ToSend$) - 2) & WHITE & vbCrLf 'send the data
    X(Index) = ""
End If
End Function


