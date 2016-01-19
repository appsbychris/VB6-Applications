Attribute VB_Name = "modRoaming"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modRoaming
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Function RoamMonsters(laMonsID As Long, lLocationToRoamTo As Long, sShortDir As String, Optional dbMapId As Long = 0) As Boolean
On Error GoTo eh1
Dim iMonID As Long
Dim dd As Long
If aMons(laMonsID).mRoams <> 0 And aMons(laMonsID).mLoc <> -1 Then
    If lLocationToRoamTo <> 0 Then
        iMonID = aMons(laMonsID).mdbMonID
        If dbMapId = 0 Then dbMapId = GetMapIndex(lLocationToRoamTo)
        With dbMap(dbMapId)
            If CLng(.iMobGroup) <> dbMonsters(iMonID).lMobGroup Then
                Exit Function
            End If
            If DCount(.sMonsters, ";") + 1 > .iMaxRegen Then
                Exit Function
            End If
        End With
        dd = aMons(laMonsID).mdbMapID
        If RndNumber(0, 100) > 75 Then DropOutDoorFood dbMapId
        With dbMap(dd)
            If .lNorth = dbMap(dbMapId).lRoomID And (.lDN = 1 Or .lDN = 2) Then Exit Function
            If .lSouth = dbMap(dbMapId).lRoomID And (.lDS = 1 Or .lDS = 2) Then Exit Function
            If .lEast = dbMap(dbMapId).lRoomID And (.lDE = 1 Or .lDE = 2) Then Exit Function
            If .lWest = dbMap(dbMapId).lRoomID And (.lDW = 1 Or .lDW = 2) Then Exit Function
            If .lNorthWest = dbMap(dbMapId).lRoomID And (.lDNW = 1 Or .lDNW = 2) Then Exit Function
            If .lNorthEast = dbMap(dbMapId).lRoomID And (.lDNE = 1 Or .lDNE = 2) Then Exit Function
            If .lSouthWest = dbMap(dbMapId).lRoomID And (.lDSW = 1 Or .lDSW = 2) Then Exit Function
            If .lSouthEast = dbMap(dbMapId).lRoomID And (.lDSE = 1 Or .lDSE = 2) Then Exit Function
            If .lUp = dbMap(dbMapId).lRoomID And (.lDU = 1 Or .lDU = 2) Then Exit Function
            If .lDown = dbMap(dbMapId).lRoomID And (.lDD = 1 Or .lDD = 2) Then Exit Function
            .sMonsters = ReplaceFast(.sMonsters, ":" & dbMonsters(iMonID).lID & ";", "", 1, 1)
            .sAMonIds = ReplaceFast(.sAMonIds, laMonsID & ";", "")
            If modSC.FastStringComp(.sMonsters, "") Then .sMonsters = "0"
            SendToAllInRoom 0, BRIGHTRED & aMons(laMonsID).mName & LIGHTBLUE & " leaves to the " & modgetdata.GetLongDir(sShortDir) & "." & WHITE & vbCrLf, .lRoomID
        End With
        With dbMap(dbMapId)
            If modSC.FastStringComp(.sMonsters, "0") Then .sMonsters = ""
            .sMonsters = .sMonsters & ":" & dbMonsters(iMonID).lID & ";"
            .sAMonIds = .sAMonIds & laMonsID & ";"
            aMons(laMonsID).mLoc = .lRoomID
            aMons(laMonsID).mdbMapID = dbMapId
            SendToAllInRoom 0, BRIGHTRED & aMons(laMonsID).mName & LIGHTBLUE & " walks in from the " & modgetdata.GetOppositeDirection(modgetdata.GetLongDir(sShortDir)) & "." & WHITE & vbCrLf, .lRoomID
        End With
        RoamMonsters = True
    End If
End If
eh1:
End Function
