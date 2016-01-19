Attribute VB_Name = "modDoors"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modDoors
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Sub CloseDoorsTimer()
Dim i As Long
Dim sS As String
Dim dd As Long
For i = 1 To UBound(dbDoor)
    With dbDoor(i)
        dd = .ldbDoorsMapID
        .lDD = dbMap(dd).lDD
        .lDU = dbMap(dd).lDU
        .lDN = dbMap(dd).lDN
        .lDS = dbMap(dd).lDS
        .lDE = dbMap(dd).lDE
        .lDW = dbMap(dd).lDW
        .lDNW = dbMap(dd).lDNW
        .lDNE = dbMap(dd).lDNE
        .lDSW = dbMap(dd).lDSW
        .lDSE = dbMap(dd).lDSE
        If .lRoomID <> 0 Then
            If .lDD <> 0 And .lDD > 1 Then
                sS = modgetdata.DoorOrGate(dd, 9)
                dbMap(dd).lDD = 1
                SendToAllInRoom 0, LIGHTBLUE & "The trap door below closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDU <> 0 And .lDU > 1 Then
                sS = modgetdata.DoorOrGate(dd, 8)
                dbMap(dd).lDU = 1
                SendToAllInRoom 0, LIGHTBLUE & "The hatch above closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDN <> 0 And .lDN > 1 Then
                sS = modgetdata.DoorOrGate(dd, 0)
                dbMap(dd).lDN = 1
                SendToAllInRoom 0, LIGHTBLUE & "The " & sS & " to the north closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDS <> 0 And .lDS > 1 Then
                sS = modgetdata.DoorOrGate(dd, 1)
                dbMap(dd).lDS = 1
                SendToAllInRoom 0, LIGHTBLUE & "The " & sS & " to the south closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDE <> 0 And .lDE > 1 Then
                sS = modgetdata.DoorOrGate(dd, 2)
                dbMap(dd).lDE = 1
                SendToAllInRoom 0, LIGHTBLUE & "The " & sS & " to the east closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDW <> 0 And .lDW > 1 Then
                sS = modgetdata.DoorOrGate(dd, 3)
                dbMap(dd).lDW = 1
                SendToAllInRoom 0, LIGHTBLUE & "The " & sS & " to the west closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDNW <> 0 And .lDNW > 1 Then
                sS = modgetdata.DoorOrGate(dd, 4)
                dbMap(dd).lDNW = 1
                SendToAllInRoom 0, LIGHTBLUE & "The " & sS & " to the northwest closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDNE <> 0 And .lDNE > 1 Then
                sS = modgetdata.DoorOrGate(dd, 5)
                dbMap(dd).lDNE = 1
                SendToAllInRoom 0, LIGHTBLUE & "The " & sS & " to the northeast closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDSW <> 0 And .lDSW > 1 Then
                sS = modgetdata.DoorOrGate(dd, 6)
                dbMap(dd).lDSW = 1
                SendToAllInRoom 0, LIGHTBLUE & "The " & sS & " to the southwest closes." & WHITE & vbCrLf, .lRoomID
            End If
            If .lDSE <> 0 And .lDSE > 1 Then
                sS = modgetdata.DoorOrGate(dd, 7)
                dbMap(dd).lDSE = 1
                SendToAllInRoom 0, LIGHTBLUE & "The " & sS & " to the southeast closes." & WHITE & vbCrLf, .lRoomID
            End If
        End If
    End With
Next
End Sub

Public Function CloseDoor(Index As Long) As Boolean
Dim dbIndex As Long
Dim sS As String
Dim dd As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 6)), "close ") Then
    CloseDoor = True
    dbIndex = GetPlayerIndexNumber(Index)
    dd = GetMapIndex(dbPlayers(dbIndex).lLocation)
    With dbMap(dd)
        Dim sCloseDir As String
        sCloseDir = Mid$(X(Index), InStr(1, X(Index), " ") + 1, 2)
        sCloseDir = TrimIt(sCloseDir)
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort(sCloseDir))
        sCloseDir = modgetdata.GetLongDir(sCloseDir)
        Select Case LCaseFast(sCloseDir)
            Case "northwest"
                If .lDNW = 3 Then
                    .lDNW = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lNorthWest <> 0 Then
                        With dbMap(GetMapIndex(.lNorthWest))
                            If .lDSE <> 0 Or .lDSE <> 1 Then
                                .lDSE = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sCloseDir) & " closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDNW = 2 Or .lDNW = 1 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "northeast"
                If .lDNE = 3 Then
                    .lDNE = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lNorthEast <> 0 Then
                        With dbMap(GetMapIndex(.lNorthEast))
                            If .lDSW <> 0 Or .lDSW <> 1 Then
                                .lDSW = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sCloseDir) & " closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDNE = 2 Or .lDNE = 1 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "southwest"
                If .lDSW = 3 Then
                    .lDSW = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lSouthWest <> 0 Then
                        With dbMap(GetMapIndex(.lSouthWest))
                            If .lDNE <> 0 Or .lDNE <> 1 Then
                                .lDNE = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sCloseDir) & " closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDSW = 2 Or .lDSW = 1 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "southeast"
                If .lDSE = 3 Then
                    .lDSE = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lSouthEast <> 0 Then
                        With dbMap(GetMapIndex(.lSouthEast))
                            If .lDNW <> 0 Or .lDNW <> 1 Then
                                .lDNW = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sCloseDir) & " closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDSE = 2 Or .lDSE = 1 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "north"
                If .lDN = 3 Then
                    .lDN = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lNorth <> 0 Then
                        With dbMap(GetMapIndex(.lNorth))
                            If .lDS <> 0 Or .lDS <> 1 Then
                                .lDS = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sCloseDir) & " closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDN = 2 Or .lDN = 1 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "south"
                If .lDS = 3 Then
                    .lDS = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lSouth <> 0 Then
                        With dbMap(GetMapIndex(.lSouth))
                            If .lDN <> 0 Or .lDN <> 1 Then
                                .lDN = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sCloseDir) & " closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDS = 2 Or .lDS = 1 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "east"
                If .lDE = 3 Then
                    .lDE = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lEast <> 0 Then
                        With dbMap(GetMapIndex(.lEast))
                            If .lDW <> 0 Or .lDW <> 1 Then
                                .lDW = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sCloseDir) & " closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDE = 2 Or .lDE = 1 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "west"
                If .lDW = 3 Then
                    .lDW = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the " & sS & " to the " & sCloseDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lWest <> 0 Then
                        With dbMap(GetMapIndex(.lWest))
                            If .lDE <> 0 Or .lDE <> 1 Then
                                .lDE = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sCloseDir) & " closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDW = 2 Or .lDW = 1 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "up"
                If .lDU = 3 Then
                    .lDU = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the hatch above." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the hatch above." & WHITE & vbCrLf, .lRoomID
                    If .lUp <> 0 Then
                        With dbMap(GetMapIndex(.lUp))
                            If .lDD <> 0 Or .lDD <> 1 Then
                                .lDD = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The trap door to below closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDU = 2 Or .lDU = 1 Then
                    WrapAndSend Index, RED & "You notice the hatch is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a hatch there." & WHITE & vbCrLf
                End If
            Case "down"
                If .lDD = 3 Then
                    .lDD = 1
                    WrapAndSend Index, LIGHTBLUE & "You close the trap door below." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " closes the trap door below." & WHITE & vbCrLf, .lRoomID
                    If .lDown <> 0 Then
                        With dbMap(GetMapIndex(.lDown))
                            If .lDU <> 0 Or .lDU <> 1 Then
                                .lDU = 1
                                SendToAllInRoom Index, LIGHTBLUE & "The hatch above closes." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDD = 2 Or .lDD = 1 Then
                    WrapAndSend Index, RED & "You notice the trap door is already closed." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a trap door there." & WHITE & vbCrLf
                End If
            Case Else
                WrapAndSend Index, RED & "There is nothing to close there." & WHITE & vbCrLf
        End Select
        X(Index) = ""
    End With
End If
End Function

Public Function BashDoor(Index As Long) As Boolean
Dim dbIndex As Long
Dim sS As String
Dim dd As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "bash ") Then
    BashDoor = True
    dbIndex = GetPlayerIndexNumber(Index)
    dd = GetMapIndex(dbPlayers(dbIndex).lLocation)
    With dbMap(dd)
        Dim sBashDir As String
        sBashDir = Mid$(X(Index), InStr(1, X(Index), " "), 2)
        sBashDir = TrimIt(sBashDir)
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort(sBashDir))
        sBashDir = modgetdata.GetLongDir(sBashDir)
        Select Case LCaseFast(sBashDir)
            Case "northwest"
                If .lDNW = 1 Or .lDNW = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBNW And .lBNW <> -1 Then
                        .lDNW = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf, .lRoomID
                        If .lNorthWest <> 0 Then
                            With dbMap(GetMapIndex(.lNorthWest))
                                If .lDSE <> 0 And .lDSE <> 3 Then
                                    .lDSE = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sBashDir) & " opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the " & sS & " to the " & sBashDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the " & sS & " to the " & sBashDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDNW = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "northeast"
                If .lDNE = 1 Or .lDNE = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBNE And .lBNE <> -1 Then
                        .lDNE = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf, .lRoomID
                        If .lNorthEast <> 0 Then
                            With dbMap(GetMapIndex(.lNorthEast))
                                If .lDSW <> 0 And .lDSW <> 3 Then
                                    .lDSW = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sBashDir) & " opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the " & sS & " to the " & sBashDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the " & sS & " to the " & sBashDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDNE = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "southwest"
                If .lDSW = 1 Or .lDSW = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBSW And .lBSW <> -1 Then
                        .lDSW = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf, .lRoomID
                        If .lSouthWest <> 0 Then
                            With dbMap(GetMapIndex(.lSouthWest))
                                If .lDNE <> 0 And .lDNE <> 3 Then
                                    .lDNE = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sBashDir) & " opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the " & sS & " to the " & sBashDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the " & sS & " to the " & sBashDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDSW = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "southeast"
                If .lDSE = 1 Or .lDSE = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBSE And .lBSE <> -1 Then
                        .lDSE = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf, .lRoomID
                        If .lSouthEast <> 0 Then
                            With dbMap(GetMapIndex(.lSouthEast))
                                If .lDNW <> 0 And .lDNW <> 3 Then
                                    .lDNW = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sBashDir) & " opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the " & sS & " to the " & sBashDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the " & sS & " to the " & sBashDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDSE = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "north"
                If .lDN = 1 Or .lDN = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBN And .lBN <> -1 Then
                        .lDN = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf, .lRoomID
                        If .lNorth <> 0 Then
                            With dbMap(GetMapIndex(.lNorth))
                                If .lDS <> 0 And .lDS <> 3 Then
                                    .lDS = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sBashDir) & " opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the " & sS & " to the " & sBashDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the " & sS & " to the " & sBashDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDN = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "south"
                If .lDS = 1 Or .lDS = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBS And .lBS <> -1 Then
                        .lDS = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf, .lRoomID
                        If .lSouth <> 0 Then
                            With dbMap(GetMapIndex(.lSouth))
                                If .lDN <> 0 And .lDN <> 3 Then
                                    .lDN = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sBashDir) & " opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the " & sS & " to the " & sBashDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the " & sS & " to the " & sBashDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDS = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "east"
                If .lDE = 1 Or .lDE = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBE And .lBE <> -1 Then
                        .lDE = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf, .lRoomID
                        If .lEast <> 0 Then
                            With dbMap(GetMapIndex(.lEast))
                                If .lDW <> 0 And .lDW <> 3 Then
                                    .lDW = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sBashDir) & " opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the " & sS & " to the " & sBashDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the " & sS & " to the " & sBashDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDE = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "west"
                If .lDW = 1 Or .lDW = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBW And .lBW <> -1 Then
                        .lDW = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the " & sS & " to the " & sBashDir & "." & WHITE & vbCrLf, .lRoomID
                        If .lWest <> 0 Then
                            With dbMap(GetMapIndex(.lWest))
                                If .lDE <> 0 And .lDE <> 3 Then
                                    .lDE = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sBashDir) & " opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the " & sS & " to the " & sBashDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the " & sS & " to the " & sBashDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDW = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "up"
                If .lDU = 1 Or .lDU = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBU And .lBU <> -1 Then
                        .lDU = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the hatch above." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the hatch above." & WHITE & vbCrLf, .lRoomID
                        If .lUp <> 0 Then
                            With dbMap(GetMapIndex(.lUp))
                                If .lDD <> 0 And .lDD <> 3 Then
                                    .lDD = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The trap door below opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the hatch above, but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the hatch above, but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDU = 3 Then
                    WrapAndSend Index, RED & "You notice the hatch is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a hatch there." & WHITE & vbCrLf
                End If
            Case "down"
                If .lDD = 1 Or .lDD = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iStr) * 1.5) >= .lBD And .lBD <> -1 Then
                        .lDD = 3
                        WrapAndSend Index, LIGHTBLUE & "You bash the trap door below." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " bashes the trap door below." & WHITE & vbCrLf, .lRoomID
                        If .lDown <> 0 Then
                            With dbMap(GetMapIndex(.lDown))
                                If .lDU <> 0 And .lDU <> 3 Then
                                    .lDU = 3
                                    SendToAllInRoom Index, LIGHTBLUE & "The hatch above opens." & WHITE & vbCrLf, .lRoomID
                                End If
                            End With
                        End If
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to bash the trap door below, but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to bash the trap door below, but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDD = 3 Then
                    WrapAndSend Index, RED & "You notice the trap door is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a trap door there." & WHITE & vbCrLf
                End If
            Case Else
                WrapAndSend Index, RED & "There is nothing to bash there." & WHITE & vbCrLf
        End Select
        X(Index) = ""
    End With
End If
End Function

Public Function PickDoor(Index As Long) As Boolean
Dim dbIndex As Long
Dim dd As Long
Dim sS As String
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "pick ") Then
    PickDoor = True
    dbIndex = GetPlayerIndexNumber(Index)
    dd = GetMapIndex(dbPlayers(dbIndex).lLocation)
    With dbMap(dd)
        Dim sPickDir As String
        sPickDir = Mid$(X(Index), InStr(1, X(Index), " "), 2)
        sPickDir = TrimIt(sPickDir)
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort(sPickDir))
        sPickDir = modgetdata.GetLongDir(sPickDir)
        Select Case LCaseFast(sPickDir)
            Case "northwest"
                If .lDNW = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPNW And .lPNW <> -1 Then
                        .lDNW = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the " & sS & " to the " & sPickDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the " & sS & " to the " & sPickDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDNW = 1 Or .lDNW = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "northeast"
                If .lDNE = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPNE And .lPNE <> -1 Then
                        .lDNE = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the " & sS & " to the " & sPickDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the " & sS & " to the " & sPickDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDNE = 1 Or .lDNE = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "southwest"
                If .lDSW = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPSW And .lPSW <> -1 Then
                        .lDSW = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the " & sS & " to the " & sPickDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the " & sS & " to the " & sPickDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDSW = 1 Or .lDSW = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "southeast"
                If .lDSE = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPSE And .lPSE <> -1 Then
                        .lDSE = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the " & sS & " to the " & sPickDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the " & sS & " to the " & sPickDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDSE = 1 Or .lDSE = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "north"
                If .lDN = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPN And .lPN <> -1 Then
                        .lDN = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the " & sS & " to the " & sPickDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the " & sS & " to the " & sPickDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDN = 1 Or .lDN = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "south"
                If .lDS = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPS And .lPS <> -1 Then
                        .lDS = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the " & sS & " to the " & sPickDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the " & sS & " to the " & sPickDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDS = 1 Or .lDS = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "east"
                If .lDE = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPE And .lPE <> -1 Then
                        .lDE = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the " & sS & " to the " & sPickDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the " & sS & " to the " & sPickDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDE = 1 Or .lDE = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "west"
                If .lDW = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPW And .lPW <> -1 Then
                        .lDW = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the " & sS & " to the " & sPickDir & ", but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the " & sS & " to the " & sPickDir & ", but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDW = 1 Or .lDW = 3 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "up"
                If .lDU = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPU And .lPU <> -1 Then
                        .lDU = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the hatch above." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the " & sS & " to the " & sPickDir & "." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the hatch, but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the hatch above, but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDU = 1 Or .lDU = 3 Then
                    WrapAndSend Index, RED & "You notice the hatch is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a hatch there." & WHITE & vbCrLf
                End If
            Case "down"
                If .lDD = 2 Then
                    If RndNumber(0, CDbl(dbPlayers(dbIndex).iInt) * 1.5) >= .lPD And .lPD <> -1 Then
                        .lDD = 1
                        WrapAndSend Index, LIGHTBLUE & "You pick the trap door below." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " picks the trap door below." & WHITE & vbCrLf, .lRoomID
                    Else
                        WrapAndSend Index, LIGHTBLUE & "You attempt to pick the trap door below, but fail." & WHITE & vbCrLf
                        SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to pick the trap door below, but fails." & WHITE & vbCrLf, .lRoomID
                    End If
                ElseIf .lDD = 1 Or .lDD = 3 Then
                    WrapAndSend Index, RED & "You notice the trap door is already open." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a trap door there." & WHITE & vbCrLf
                End If
            Case Else
                WrapAndSend Index, RED & "There is nothing to pick there." & WHITE & vbCrLf
        End Select
        X(Index) = ""
    End With
End If
End Function

Public Function OpenDoor(Index As Long) As Boolean
Dim dbIndex As Long
Dim dd As Long
Dim sS As String
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "open ") Then
    OpenDoor = True
    dbIndex = GetPlayerIndexNumber(Index)
    dd = GetMapIndex(dbPlayers(dbIndex).lLocation)
    With dbMap(dd)
        Dim sOpenDir As String
        sOpenDir = Mid$(X(Index), InStr(1, X(Index), " ") + 1, 2)
        sOpenDir = TrimIt(sOpenDir)
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort(sOpenDir))
        sOpenDir = modgetdata.GetLongDir(sOpenDir)
        Select Case LCaseFast(sOpenDir)
            Case "northwest"
                If .lDNW = 1 Then
                    .lDNW = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lNorthWest <> 0 Then
                        With dbMap(GetMapIndex(.lNorthWest))
                            If .lDSE <> 0 And .lDSE <> 3 Then
                                .lDSE = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sOpenDir) & " opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDNW = 2 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "northeast"
                If .lDNE = 1 Then
                    .lDNE = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lNorthEast <> 0 Then
                        With dbMap(GetMapIndex(.lNorthEast))
                            If .lDSW <> 0 And .lDSW <> 3 Then
                                .lDSW = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sOpenDir) & " opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDNE = 2 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "southwest"
                If .lDSW = 1 Then
                    .lDSW = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lSouthWest <> 0 Then
                        With dbMap(GetMapIndex(.lSouthWest))
                            If .lDNE <> 0 And .lDNE <> 3 Then
                                .lDNE = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sOpenDir) & " opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDSW = 2 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "southeast"
                If .lDSE = 1 Then
                    .lDSE = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lSouthEast <> 0 Then
                        With dbMap(GetMapIndex(.lSouthEast))
                            If .lDNW <> 0 And .lDNW <> 3 Then
                                .lDNW = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sOpenDir) & " opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDSE = 2 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "north"
                If .lDN = 1 Then
                    .lDN = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lNorth <> 0 Then
                        With dbMap(GetMapIndex(.lNorth))
                            If .lDS <> 0 And .lDS <> 3 Then
                                .lDS = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sOpenDir) & " opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDN = 2 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "south"
                If .lDS = 1 Then
                    .lDS = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lSouth <> 0 Then
                        With dbMap(GetMapIndex(.lSouth))
                            If .lDN <> 0 And .lDN <> 3 Then
                                .lDN = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sOpenDir) & " opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDS = 2 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "east"
                If .lDE = 1 Then
                    .lDE = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lEast <> 0 Then
                        With dbMap(GetMapIndex(.lEast))
                            If .lDW <> 0 And .lDW <> 3 Then
                                .lDW = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sOpenDir) & " opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDE = 2 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "west"
                If .lDW = 1 Then
                    .lDW = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the " & sS & " to the " & sOpenDir & "." & WHITE & vbCrLf, .lRoomID
                    If .lWest <> 0 Then
                        With dbMap(GetMapIndex(.lWest))
                            If .lDE <> 0 And .lDE <> 3 Then
                                .lDE = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The " & sS & " to the " & modgetdata.GetOppositeDirection(sOpenDir) & " opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDW = 2 Then
                    WrapAndSend Index, RED & "You notice the " & sS & " is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
                End If
            Case "up"
                If .lDU = 1 Then
                    .lDU = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the hatch above." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the hatch above." & WHITE & vbCrLf, .lRoomID
                    If .lUp <> 0 Then
                        With dbMap(GetMapIndex(.lUp))
                            If .lDD <> 0 And .lDD <> 3 Then
                                .lDD = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The hatch above opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDU = 2 Then
                    WrapAndSend Index, RED & "You notice the hatch above is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a hatch there." & WHITE & vbCrLf
                End If
            Case "down"
                If .lDD = 1 Then
                    .lDD = 3
                    WrapAndSend Index, LIGHTBLUE & "You open the trap door below." & WHITE & vbCrLf
                    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " opens the trap door below." & WHITE & vbCrLf, .lRoomID
                    If .lDown <> 0 Then
                        With dbMap(GetMapIndex(.lDown))
                            If .lDU <> 0 And .lDU <> 3 Then
                                .lDU = 3
                                SendToAllInRoom Index, LIGHTBLUE & "The hatch above opens." & WHITE & vbCrLf, .lRoomID
                            End If
                        End With
                    End If
                ElseIf .lDD = 2 Then
                    WrapAndSend Index, RED & "You notice the trap door is locked." & WHITE & vbCrLf
                Else
                    WrapAndSend Index, RED & "You notice there isn't a trap door there." & WHITE & vbCrLf
                End If
            Case Else
                WrapAndSend Index, RED & "There is nothing to open there." & WHITE & vbCrLf
        End Select
        X(Index) = ""
    End With
End If
End Function

Sub UnLockDoor(Index As Long, lLocation As Long, lPlayersKey As Long, sDirection As String)
Dim dd As Long
Dim sS As String
dd = GetMapIndex(lLocation)
With dbMap(dd)
    sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort(modgetdata.GetShortDir(sDirection)))
    Select Case sDirection
        Case "north"
            If lPlayersKey = .lKN Then
                If .lDN = 2 Then
                    .lDN = 1
                ElseIf .lDN = 3 Then
                    UnLockBadDoor Index, "You notice that the " & sS & " is already open."
                    Exit Sub
                ElseIf .lDN = 1 Then
                    UnLockBadDoor Index, "You notice that this " & sS & " does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "south"
            If lPlayersKey = .lKS Then
                If .lDS = 2 Then
                    .lDS = 1
                ElseIf .lDS = 3 Then
                    UnLockBadDoor Index, "You notice that the " & sS & " is already open."
                    Exit Sub
                ElseIf .lDS = 1 Then
                    UnLockBadDoor Index, "You notice that this " & sS & " does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "east"
            If lPlayersKey = .lKE Then
                If .lDE = 2 Then
                    .lDE = 1
                ElseIf .lDE = 3 Then
                    UnLockBadDoor Index, "You notice that the " & sS & " is already open."
                    Exit Sub
                ElseIf .lDE = 1 Then
                    UnLockBadDoor Index, "You notice that this " & sS & " does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "west"
            If lPlayersKey = .lKW Then
                If .lDW = 2 Then
                    .lDW = 1
                ElseIf .lDW = 3 Then
                    UnLockBadDoor Index, "You notice that the " & sS & " is already open."
                    Exit Sub
                ElseIf .lDW = 1 Then
                    UnLockBadDoor Index, "You notice that this " & sS & " does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "northwest"
            If lPlayersKey = .lKNW Then
                If .lDNW = 2 Then
                    .lDNW = 1
                ElseIf .lDNW = 3 Then
                    UnLockBadDoor Index, "You notice that the " & sS & " is already open."
                    Exit Sub
                ElseIf .lDNW = 1 Then
                    UnLockBadDoor Index, "You notice that this " & sS & " does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "northeast"
            If lPlayersKey = .lKNE Then
                If .lDNE = 2 Then
                    .lDNE = 1
                ElseIf .lDNE = 3 Then
                    UnLockBadDoor Index, "You notice that the " & sS & " is already open."
                    Exit Sub
                ElseIf .lDNE = 1 Then
                    UnLockBadDoor Index, "You notice that this " & sS & " does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "southwest"
            If lPlayersKey = .lKSW Then
                If .lDSW = 2 Then
                    .lDSW = 1
                ElseIf .lDSW = 3 Then
                    UnLockBadDoor Index, "You notice that the " & sS & " is already open."
                    Exit Sub
                ElseIf .lDSW = 1 Then
                    UnLockBadDoor Index, "You notice that this " & sS & " does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "southeast"
            If lPlayersKey = .lKSE Then
                If .lDSE = 2 Then
                    .lDSE = 1
                ElseIf .lDSE = 3 Then
                    UnLockBadDoor Index, "You notice that the " & sS & " is already open."
                    Exit Sub
                ElseIf .lDSE = 1 Then
                    UnLockBadDoor Index, "You notice that this " & sS & " does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "up"
            If lPlayersKey = .lKU Then
                If .lDU = 2 Then
                    sS = "hatch"
                    .lDU = 1
                ElseIf .lDU = 3 Then
                    UnLockBadDoor Index, "You notice that the hatch is already open."
                    Exit Sub
                ElseIf .lDU = 1 Then
                    UnLockBadDoor Index, "You notice that this hatch does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case "down"
            If lPlayersKey = .lKD Then
                If .lDD = 2 Then
                    .lDD = 1
                    sS = "trap door"
                ElseIf .lDD = 3 Then
                    UnLockBadDoor Index, "You notice that the trap door is already open."
                    Exit Sub
                ElseIf .lDD = 1 Then
                    UnLockBadDoor Index, "You notice that this trap door does not require a key."
                    Exit Sub
                End If
            Else
                UnLockBadDoor Index, "You can't seem to get this key in."
                Exit Sub
            End If
        Case Else
                WrapAndSend Index, RED & "There is nothing to unlock there." & WHITE & vbCrLf
    End Select
    WrapAndSend Index, LIGHTBLUE & "You unlock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
    SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " unlocks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, lLocation
    X(Index) = ""
End With
End Sub

Sub UnLockBadDoor(Index As Long, sWhat As String)
WrapAndSend Index, RED & sWhat & WHITE & vbCrLf
X(Index) = ""
End Sub

Sub LockDoor(Index As Long, lLocation As Long, sDirection As String)
Dim dd As Long
Dim sS As String
dd = GetMapIndex(lLocation)
With dbMap(dd)
    sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort(modgetdata.GetShortDir(sDirection)))
    Select Case LCaseFast(sDirection)
        Case "northwest"
            If (.lDNW = 3 Or .lDNW = 1) And (.lKNW <> 0 Or .lPNW <> -1) Then
                .lDNW = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDNW = 2 Or .lDNW = 1 Then
                WrapAndSend Index, RED & "You notice the " & sS & " is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
            End If
        Case "northeast"
            If (.lDNE = 3 Or .lDNE = 1) And (.lKNE <> 0 Or .lPNE <> -1) Then
                .lDNE = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDNE = 2 Or .lDNE = 1 Then
                WrapAndSend Index, RED & "You notice the " & sS & " is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
            End If
        Case "southwest"
            If (.lDSW = 3 Or .lDSW = 1) And (.lKSW <> 0 Or .lPSW <> -1) Then
                .lDSW = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDSW = 2 Or .lDSW = 1 Then
                WrapAndSend Index, RED & "You notice the " & sS & " is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
            End If
        Case "southeast"
            If (.lDSE = 3 Or .lDSE = 1) And (.lKSE <> 0 Or .lPSE <> -1) Then
                .lDSE = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDSE = 2 Or .lDSE = 1 Then
                WrapAndSend Index, RED & "You notice the " & sS & " is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
            End If
        Case "north"
            If (.lDN = 3 Or .lDN = 1) And (.lKN <> 0 Or .lPN <> -1) Then
                .lDN = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDN = 2 Or .lDN = 1 Then
                WrapAndSend Index, RED & "You notice the " & sS & " is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
            End If
        Case "south"
            If (.lDS = 3 Or .lDS = 1) And (.lKS <> 0 Or .lPS <> -1) Then
                .lDS = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDS = 2 Or .lDS = 1 Then
                WrapAndSend Index, RED & "You notice the " & sS & " is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
            End If
        Case "east"
            If (.lDE = 3 Or .lDE = 1) And (.lKE <> 0 Or .lPE <> -1) Then
                .lDE = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDE = 2 Or .lDE = 1 Then
                WrapAndSend Index, RED & "You notice the " & sS & " is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
            End If
        Case "west"
            If (.lDW = 3 Or .lDW = 1) And (.lKW <> 0 Or .lPW <> -1) Then
                .lDW = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the " & sS & " to the " & sDirection & "." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDW = 2 Or .lDW = 1 Then
                WrapAndSend Index, RED & "You notice the " & sS & " is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a " & sS & " there." & WHITE & vbCrLf
            End If
        Case "up"
            If (.lDU = 3 Or .lDU = 1) And (.lKU <> 0 Or .lPU <> -1) Then
                .lDU = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the hatch above." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the hatch above." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDU = 2 Or .lDU = 1 Then
                WrapAndSend Index, RED & "You notice the hatch is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a hatch there." & WHITE & vbCrLf
            End If
        Case "down"
            If (.lDD = 3 Or .lDD = 1) And (.lKD <> 0 Or .lPD <> -1) Then
                .lDD = 2
                WrapAndSend Index, LIGHTBLUE & "You lock the trap door below." & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " locks the trap door below." & WHITE & vbCrLf, .lRoomID
            ElseIf .lDD = 2 Or .lDD = 1 Then
                WrapAndSend Index, RED & "You notice the trap door is already locked." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You notice there isn't a trap door there." & WHITE & vbCrLf
            End If
        Case Else
                WrapAndSend Index, RED & "There is nothing to lock there." & WHITE & vbCrLf
    End Select
    X(Index) = ""
End With
End Sub

Public Function UnlockADoor(Index As Long) As Boolean
Dim TempItem As String
Dim s As String
Dim i As Long
Dim sDir As String
Dim sC As String
Dim dbItemID As Long
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 2)), "un") Then
    s = LCaseFast(X(Index))
    If modSC.FastStringComp(Mid$(s, 3, 1), "l") Then
        UnlockADoor = True
    ElseIf modSC.FastStringComp(Mid$(s, 3, 1), " ") Then
        UnlockADoor = True
    Else
        Exit Function
    End If
    i = InStr(1, s, " ")
    If i = 0 Then
        UnlockADoor = False
        Exit Function
    End If
    s = Mid$(s, i + 1)
    i = InStr(1, s, " ")
    If i = 0 Then
        UnlockADoor = False
        Exit Function
    End If
    sDir = Mid$(s, 1, i - 1)
    s = Mid$(s, i + 1)
    If Len(s) > 1 Then
        sDir = Mid$(sDir, 1, 2)
    Else
        sDir = Left$(s, 1)
    End If
    sDir = modgetdata.GetLongDir(sDir)
    If sDir = "-1" Then
        WrapAndSend Index, RED & "You can't find that door." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If modSC.FastStringComp(Left$(s, 1), "w") Then
        i = InStr(1, s, " ")
        If i = 0 Then
            UnlockADoor = False
            Exit Function
        End If
        If i < 6 Then
            sC = Mid$(s, 1, i - 1)
            If Len(sC) > 1 Then
                Select Case Len(sC)
                    Case 2
                        If Not modSC.FastStringComp(sC, "wi") Then
                            TempItem = s
                        Else
                            s = Mid$(s, i + 1)
                            TempItem = s
                        End If
                    Case 3
                        If Not modSC.FastStringComp(sC, "wit") Then
                            TempItem = s
                        Else
                            s = Mid$(s, i + 1)
                            TempItem = s
                        End If
                    Case 4
                        If Not modSC.FastStringComp(sC, "with") Then
                            TempItem = s
                        Else
                            s = Mid$(s, i + 1)
                            TempItem = s
                        End If
                    Case Else
                        
                        TempItem = s
                End Select
            Else
                If modSC.FastStringComp(sC, "w") Then
                    s = Mid$(s, i + 1)
                    TempItem = s
                Else
                    TempItem = s
                End If
            End If
        Else
            TempItem = s
        End If
    Else
        TempItem = s
    End If
    TempItem = SmartFind(Index, TempItem, Inventory_Item)
    If InStr(1, TempItem, Chr$(0)) > 0 Then TempItem = Mid$(TempItem, InStr(1, TempItem, Chr$(0)) + 1)
    dbItemID = GetItemID(TempItem)
    If dbItemID = 0 Then
        UnLockBadDoor Index, "You don't seem to have that key."
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If InStr(1, .sInventory, ":" & dbItems(dbItemID).iID & "/") = 0 Then
            UnLockBadDoor Index, "You don't seem to have that key."
            Exit Function
        End If
    End With
    With dbItems(dbItemID)
        If .sWorn = "key" Then
            UnLockDoor Index, dbPlayers(dbIndex).lLocation, CLng(.iID), sDir
        Else
            WrapAndSend Index, RED & "That doesn't appear to be a key." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
End If
End Function

Public Function LockADoor(Index As Long) As Boolean

If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "lock ") Then
    LockADoor = True
    X(Index) = ReplaceFast(X(Index), "lock ", "", 1, 1)
    LockDoor Index, dbPlayers(GetPlayerIndexNumber(Index)).lLocation, modgetdata.GetLongDir(X(Index))
    X(Index) = ""
End If
End Function

Public Function Doors(Index As Long) As Boolean
If OpenDoor(Index) = True Then Doors = True: Exit Function
If CloseDoor(Index) = True Then Doors = True: Exit Function
If BashDoor(Index) = True Then Doors = True: Exit Function
If PickDoor(Index) = True Then Doors = True: Exit Function
If UnlockADoor(Index) = True Then Doors = True: Exit Function
If LockADoor(Index) = True Then Doors = True: Exit Function
Doors = False
End Function
