Attribute VB_Name = "modMapFlags"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modMapFlags
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Enum MapFlag
    mapType = 0
    mapShop = 1
    mapMaxRegen = 2
    mapMobGroup = 3
    mapGold = 4
    mapLight = 5
    mapSafeRoom = 6
    mapDeathRoom = 7
    mapIndoor = 8
    mapTrainClass = 9
    mapGate = 10
    mapOutDoorFood = 11
End Enum

Public Enum mapDir
    North = 0
    South = 1
    East = 2
    West = 3
    NorthWest = 4
    NorthEast = 5
    SouthWest = 6
    SouthEast = 7
    Up = 8
    Down = 9
    GETALL = 10
End Enum

Public Function GetMapFlag(dbMapId As Long, w As MapFlag, Optional GATEDIR As mapDir) As String
Dim Arr() As String
Dim s As String
With dbMap(dbMapId)
'    If .sMapFlags = "" Then UpdateMapFlags dbMapId
    SplitFast .sMapFlags, Arr, "/"
    If w <> mapGate Or GATEDIR = GETALL Then
        GetMapFlag = Arr(w)
    Else
        s = Arr(w)
        Erase Arr
        SplitFast s, Arr, ";"
        GetMapFlag = Arr(GATEDIR)
    End If
End With
End Function

Public Sub SetMapFlag(dbMapId As Long, w As MapFlag, sSet As String, Optional GATEDIR As mapDir)
Dim Arr() As String
Dim Arr2() As String
Dim s As String
With dbMap(dbMapId)
    SplitFast .sMapFlags, Arr, "/"
    If w <> mapGate Then
        Arr(w) = sSet
    Else
        s = Arr(w)
        SplitFast s, Arr2, ";"
        Arr2(GATEDIR) = sSet
        s = Join(Arr2, ";")
        Arr(w) = s
    End If
    .sMapFlags = Join(Arr, "/")
End With
End Sub

Public Sub LoadMapFlags(dbMapId As Long)
If dbMapId = 0 Then Exit Sub
With dbMap(dbMapId)
    .iType = CLng(Val(modMapFlags.GetMapFlag(dbMapId, mapType)))
    .sShopItems = modMapFlags.GetMapFlag(dbMapId, mapShop)
    .iMaxRegen = CLng(Val(modMapFlags.GetMapFlag(dbMapId, mapMaxRegen)))
    .dGold = Val(modMapFlags.GetMapFlag(dbMapId, mapGold))
    .lLight = CLng(Val(modMapFlags.GetMapFlag(dbMapId, mapLight)))
    .iSafeRoom = CLng(Val(modMapFlags.GetMapFlag(dbMapId, mapSafeRoom)))
    .lDeathRoom = CLng(Val(modMapFlags.GetMapFlag(dbMapId, mapDeathRoom)))
    .iInDoor = CLng(Val(modMapFlags.GetMapFlag(dbMapId, mapIndoor)))
    .iTrainClass = CLng(Val(modMapFlags.GetMapFlag(dbMapId, mapTrainClass)))
    .iMobGroup = CLng(Val(modMapFlags.GetMapFlag(dbMapId, mapMobGroup)))
    .sOutDoorFood = modMapFlags.GetMapFlag(dbMapId, mapOutDoorFood)
End With
End Sub

'Public Enum MapFlag
'    mapType = 0
'    mapShop = 1
'    mapMaxRegen = 2
'    mapMobGroup = 3
'    mapGold = 4
'    mapLight = 5
'    mapSafeRoom = 6
'    mapDeathRoom = 7
'    mapIndoor = 8
'    mapTrainClass = 9
'    mapGate = 10
'    mapOutDoorFood = 11
'End Enum

Public Sub UpdateMapFlags(dbMapId As Long)
Dim s As String
With dbMap(dbMapId)
    s = CStr(.iType) & "/" & .sShopItems & "/" & CStr(.iMaxRegen) & "/" & CStr(.iMobGroup) & "/" & CStr(.dGold) & "/" & CStr(.lLight) & "/" & _
        CStr(.iSafeRoom) & "/" & CStr(.lDeathRoom) & "/" & CStr(.iInDoor) & "/" & CStr(.iTrainClass) & "/" & _
        modMapFlags.GetMapFlag(dbMapId, mapGate, GETALL) & "/" & .sOutDoorFood
    .sMapFlags = s
End With
End Sub


