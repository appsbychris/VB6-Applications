Attribute VB_Name = "modAid"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modAid
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function AidPlayer(Index As Long) As Boolean
Dim s As String
Dim dbIndex2 As Long
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 4)), "aid ") Then
    AidPlayer = True
    s = Mid$(X(Index), 5)
    s = LCaseFast(s)
    s = SmartFind(Index, s, Player_In_Room)
    dbIndex2 = GetPlayerIndexNumber(, s)
    If dbIndex2 = 0 Then
        WrapAndSend Index, RED & "You can't find " & s & "." & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    If InStr(1, modgetdata.GetPlayersHereWithoutRiding(dbPlayers(dbIndex).lLocation, dbIndex), dbPlayers(dbIndex2).sPlayerName) = 0 Then
        WrapAndSend Index, RED & "You can't find " & s & "." & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbPlayers(dbIndex2).lRegain = 1
    WrapAndSend Index, LIGHTBLUE & "You aid " & dbPlayers(dbIndex2).sPlayerName & "'s wounds." & WHITE & vbCrLf
    WrapAndSend dbPlayers(dbIndex2).iIndex, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " aids your wounds." & WHITE & vbCrLf
    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " aids " & dbPlayers(dbIndex2).sPlayerName & "'s wounds." & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation, dbPlayers(dbIndex2).iIndex
    X(Index) = ""
End If
End Function
