VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : clsGet
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function GetAllMonstersInRoom(lLocation As Long, Optional dbMapId As Long) As String
Dim Arr() As String
Dim i As Long
Dim s As String
If dbMapId = 0 Then dbMapId = GetMapIndex(lLocation)
With dbMap(dbMapId)
    s = .sAMonIds
    If s <> "" Then
        SplitFast s, Arr, ";"
        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> "" Then
                If Val(Arr(i)) <= UBound(aMons) Then
                    If aMons(Val(Arr(i))).mLoc <> .lRoomID Then
                        .sAMonIds = ReplaceFast(.sAMonIds, Arr(i) & ";", "")
                    End If
                Else
                    .sAMonIds = ReplaceFast(.sAMonIds, Arr(i) & ";", "")
                End If
            End If
            If DE Then DoEvents
        Next
    End If
    GetAllMonstersInRoom = .sAMonIds
End With
End Function

Public Function GetAllMonstersInRoomATTACKABLE(lLocation As Long, Optional dbMapId As Long) As String
Dim i As Long
Dim Arr() As String
Dim s As String
If dbMapId = 0 Then dbMapId = GetMapIndex(lLocation)
With dbMap(dbMapId)
    s = .sAMonIds
    If s <> "" Then
        SplitFast s, Arr, ";"
        s = ""
        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> "" Then
                If aMons(Val(Arr(i))).mAttackable = True Then
                    s = s & Arr(i) & ";"
                End If
            End If
            If DE Then DoEvents
        Next
    End If
End With
GetAllMonstersInRoomATTACKABLE = s
End Function

Public Function GetAllMonstersInRoomMONID(lLocation As Long) As String
Dim i As Long
Dim s As String
Dim j As Long
For i = LBound(aMons) To UBound(aMons)
    If j > AmountMons Then Exit For
    If aMons(i).mLoc = lLocation Then
        s = s & ":" & aMons(i).miID & ";"
        j = j + 1
    End If
    If DE Then DoEvents
Next
GetAllMonstersInRoomMONID = s
End Function

Public Function GetMonDamDesc(lCHP As Long, lMHP As Long) As String
Dim d As Double
d = lCHP / lMHP
d = d * 100
d = RoundFast(d, 0)
Select Case d
    Case Is >= 100
        GetMonDamDesc = BRIGHTWHITE & " is completly unwounded."
    Case Is >= 90
        GetMonDamDesc = WHITE & " has a few scrapes and scratches."
    Case Is >= 80
        GetMonDamDesc = BRIGHTBLUE & " has a bruise or two, and some cuts."
    Case Is >= 70
        GetMonDamDesc = BLUE & " has many cuts, with some still bleeding."
    Case Is >= 60
        GetMonDamDesc = BRIGHTLIGHTBLUE & " has some bleeding wounds, and quite a few bruises."
    Case Is >= 50
        GetMonDamDesc = LIGHTBLUE & " has many bleeding wounds, and looks like they have taken a beating."
    Case Is >= 40
        GetMonDamDesc = BRIGHTGREEN & " looks very unhealthy and unstable. They are bleeding heavily."
    Case Is >= 30
        GetMonDamDesc = GREEN & " is bleeding a lot, has many bruises, and looks unstable."
    Case Is >= 20
        GetMonDamDesc = BRIGHTRED & " is heavily wounded, and is barely able to stand up."
    Case Is >= 10
        GetMonDamDesc = RED & " is near death, breathing heavily, and can barely stand."
    Case Else
        GetMonDamDesc = BGRED & " is barely able to breath and stand up. They are bleeding heavily from many wounds."
End Select
End Function

Public Function GetANSIColorChanges(sText As String) As Long
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, RED)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, GREEN)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, YELLOW)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BLUE)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, MAGNETA)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, LIGHTBLUE)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, WHITE)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BGRED)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BGGREEN)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BGYELLOW)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BGBLUE)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BGPURPLE)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BGLIGHTBLUE)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BRIGHTYELLOW)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BRIGHTGREEN)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BRIGHTRED)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BRIGHTBLUE)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BRIGHTMAGNETA)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BRIGHTLIGHTBLUE)
GetANSIColorChanges = GetANSIColorChanges + DCount(sText, BRIGHTWHITE)
End Function

Public Function GetClassFromNum(ClassNum As Long) As String
Dim i As Long
i = GetClassID(, ClassNum)
If i = 0 Then GetClassFromNum = "-1": Exit Function
GetClassFromNum = LCaseFast(dbClass(i).sName)
End Function

Public Function GetRaceFromNum(RaceNum As Long) As String
Dim i As Long
i = GetRaceID(, RaceNum)
If i = 0 Then GetRaceFromNum = "-1": Exit Function
GetRaceFromNum = LCaseFast(dbRaces(i).sName)
End Function

Public Function GetItemNumFromName(ItemName As String) As String
Dim i As Long
i = GetItemID(ItemName)
If i = 0 Then GetItemNumFromName = "(-1)": Exit Function
GetItemNumFromName = dbItems(i).iID
End Function

Public Function GetMonstersFromNums(MonIDs As String) As String
Dim s As String
Dim tArr() As String
Dim i As Long, j As Long
If MonIDs = "0" Then Exit Function
s = MonIDs
s = ReplaceFast(s, ":", "")
SplitFast s, tArr, ";"
s = ""
For i = LBound(tArr) To UBound(tArr)
    If Not clsSC.FastStringComp(tArr(i), "") Then
        For j = LBound(dbMonsters) To UBound(dbMonsters)
            With dbMonsters(j)
                If .lID = CLng(tArr(i)) Then
                    s = s & .sMonsterName & ","
                    Exit For
                End If
            End With
            If DE Then DoEvents
        Next j
    End If
    If DE Then DoEvents
Next i
GetMonstersFromNums = s
End Function

Public Function GetPlayersInvFromNums(Index As Long, Optional GroupIt As Boolean = False, Optional dbIndex As Long = 0) As String
Dim s As String
Dim tArr() As String
Dim i As Long, j As Long
Dim m As Long
Dim n As Long
If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    s = .sInventory
End With
If clsSC.FastStringComp(s, "0") Then GetPlayersInvFromNums = "": Exit Function
SplitFast s, tArr, ";"
s = ""

For i = LBound(tArr) To UBound(tArr)
    m = InStr(1, tArr(i), ":")
    
    If m = 0 Then GoTo nNext
    
    n = InStr(m, tArr(i), "/")
    
    If n = 0 Then GoTo nNext
    
    m = CLng(Mid$(tArr(i), m + 1, n - m - 1))
    
    If m = 0 Then GoTo nNext
    
    For j = LBound(dbItems) To UBound(dbItems)
        With dbItems(j)
            If .iID = m Then
                s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & ","
                Exit For
            End If
        End With
        If DE Then DoEvents
    Next j
nNext:
    If DE Then DoEvents
Next i

If GroupIt And Not clsSC.FastStringComp(s, "0") Then
    Erase tArr
    SplitFast s, tArr, ","
    s = ""
    n = 1
    For i = LBound(tArr) To UBound(tArr)
        If Not clsSC.FastStringComp(tArr(i), "") Then
            For j = LBound(tArr) To UBound(tArr)
                If i <> j Then
                    If Not clsSC.FastStringComp(tArr(i), "") Then
                        If clsSC.FastStringComp(tArr(i), tArr(j)) Then
                            n = n + 1
                            tArr(j) = ""
                        End If
                    End If
                End If
                If DE Then DoEvents
            Next
            If n > 1 Then
                s = s & CStr(n) & " " & tArr(i) & ","
            Else
                s = s & tArr(i) & ","
            End If
            n = 1
        End If
        If DE Then DoEvents
    Next
End If

GetPlayersInvFromNums = s
End Function

Public Function GetPlayersEq(Index As Long, Optional dbIndex As Long) As String
Dim s As String
If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    s = .sArms & ";" & .sBack & ";" & .sBody & ";" & .sEars & ";" & .sFace & _
        ";" & .sFeet & ";" & .sHands & ";" & .sHead & ";" & .sLegs & ";" & .sNeck & _
        ";" & .sShield & ";" & .sWaist & ";" & .sWeapon & ";" & .sRings( _
        0) & ";" & .sRings(1) & ";" & .sRings(2) & ";" & .sRings(3) & ";" & .sRings( _
        4) & ";" & .sRings(5)
End With
If clsSC.FastStringComp(s, "0") Then GetPlayersEq = "": Exit Function
GetPlayersEq = s
End Function

Public Function GetPlayersEqDBINDEX(dbIndex As Long) As String
Dim s As String
With dbPlayers(dbIndex)
    s = .sArms & ";" & .sBack & ";" & .sBody & ";" & .sEars & ";" & .sFace & _
        ";" & .sFeet & ";" & .sHands & ";" & .sHead & ";" & .sLegs & ";" & .sNeck & _
        ";" & .sShield & ";" & .sWaist & ";" & .sWeapon & ";" & .sRings( _
        0) & ";" & .sRings(1) & ";" & .sRings(2) & ";" & .sRings(3) & ";" & .sRings( _
        4) & ";" & .sRings(5)
End With
If clsSC.FastStringComp(s, "0") Then GetPlayersEqDBINDEX = "": Exit Function
GetPlayersEqDBINDEX = s
End Function

Public Function GetPlayersEqFromNums(Index As Long, Optional NoWorn As Boolean = False) As String
Dim s As String
Dim tArr() As String
Dim i As Long
Dim m As Long
With dbPlayers(GetPlayerIndexNumber(Index))
    s = .sArms & ";" & .sBack & ";" & .sBody & ";" & .sEars & ";" & .sFace & _
        ";" & .sFeet & ";" & .sHands & ";" & .sHead & ";" & .sLegs & ";" & .sNeck & _
        ";" & .sShield & ";" & .sWaist & ";" & .sWeapon & ";" & .sRings( _
        0) & ";" & .sRings(1) & ";" & .sRings(2) & ";" & .sRings(3) & ";" & .sRings( _
        4) & ";" & .sRings(5)
End With
SplitFast Left$(s, Len(s) - 1), tArr, ";"
s = ""

For i = LBound(tArr) To UBound(tArr)
    If tArr(i) <> "0" Then
        m = modItemManip.GetItemIDFromUnFormattedString(tArr(i))
        If m = 0 Then GoTo nNext
        m = GetItemID(, CLng(m))
        If m = 0 Then GoTo nNext
        With dbItems(m)
            If NoWorn Then
                s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & ","
            Else
                s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & Me.GetWornLocation(i) & ","
            End If
        End With
    Else
        If NoWorn Then s = s & "0,"
    End If
nNext:
    If DE Then DoEvents
Next i
GetPlayersEqFromNums = s
End Function

Public Function GetPlayersEqFromNumsDBINDEX(dbIndex As Long, Optional NoWorn As Boolean = False) As String
Dim s As String
Dim tArr() As String
Dim i As Long
Dim m As Long
With dbPlayers(dbIndex)
    s = .sArms & ";" & .sBack & ";" & .sBody & ";" & .sEars & ";" & .sFace & _
        ";" & .sFeet & ";" & .sHands & ";" & .sHead & ";" & .sLegs & ";" & .sNeck & _
        ";" & .sShield & ";" & .sWaist & ";" & .sWeapon & ";" & .sRings( _
        0) & ";" & .sRings(1) & ";" & .sRings(2) & ";" & .sRings(3) & ";" & .sRings( _
        4) & ";" & .sRings(5)
End With
SplitFast Left$(s, Len(s) - 1), tArr, ";"
s = ""

For i = LBound(tArr) To UBound(tArr)
    If tArr(i) <> "0" Then
        m = modItemManip.GetItemIDFromUnFormattedString(tArr(i))
        If m = 0 Then GoTo nNext
        m = GetItemID(, CLng(m))
        If m = 0 Then GoTo nNext
        With dbItems(m)
            If NoWorn Then
                s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & ","
            Else
                s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & Me.GetWornLocation(i) & ","
            End If
        End With
    Else
        If NoWorn Then s = s & "0,"
    End If
nNext:
    If DE Then DoEvents
Next i
GetPlayersEqFromNumsDBINDEX = s
End Function

Public Function GetRoomItemsFromNums(Index As Long, Optional Colorize As Boolean = False, Optional GroupIt As Boolean = False, Optional dbIndex As Long) As String
Dim s As String
Dim tArr() As String
Dim i As Long, j As Long
Dim dbLoc As Long
Dim m As Long
Dim n As Long
If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
If dbIndex = 0 Then Exit Function
dbLoc = dbPlayers(dbIndex).lDBLocation
With dbMap(dbLoc)
    s = .sItems
End With
If clsSC.FastStringComp(s, "0") Then GetRoomItemsFromNums = "": GoTo Notes
SplitFast s, tArr, ";"
s = ""
For i = LBound(tArr) To UBound(tArr)
    m = modItemManip.GetItemIDFromUnFormattedString(tArr(i))
    If m <= 0 Then GoTo nNext
    With dbItems(GetItemID(, m))
        s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & ","
    End With
nNext:
    If DE Then DoEvents
Next i
If GroupIt And Not clsSC.FastStringComp(s, "0") Then
    Erase tArr
    SplitFast s, tArr, ","
    s = ""
    n = 1
    For i = LBound(tArr) To UBound(tArr)
        If Not clsSC.FastStringComp(tArr(i), "") Then
            For j = LBound(tArr) To UBound(tArr)
                If i <> j Then
                    If Not clsSC.FastStringComp(tArr(i), "") Then
                        If clsSC.FastStringComp(tArr(i), tArr(j)) Then
                            n = n + 1
                            tArr(j) = ""
                        End If
                    End If
                End If
                If DE Then DoEvents
            Next
            If n > 1 Then
                s = s & CStr(n) & " " & tArr(i) & ","
            Else
                s = s & tArr(i) & ","
            End If
            n = 1
        End If
        If DE Then DoEvents
    Next
End If
Notes:
If clsSC.FastStringComp(s, "0") Then s = ""
s = s & modItemManip.GetListOfLettersFromGround(dbLoc)
If Colorize Then
    If s <> "" Then
        s = Left$(s, Len(s) - 1)
        s = ReplaceFast(s, ",", YELLOW & ", " & GREEN)
    End If
End If
GetRoomItemsFromNums = s
End Function

Public Function GetRoomItemsFromNumsFromRoomID(lRoomID As Long, Optional Colorize As Boolean = False, Optional GroupIt As Boolean = False, Optional dbMapId As Long = 0) As String
Dim s As String
Dim tArr() As String
Dim i As Long, j As Long
Dim dbLoc As Long
Dim m As Long
Dim n As Long

If dbMapId = 0 Then
    dbLoc = GetMapIndex(lRoomID)
Else
    dbLoc = dbMapId
End If
With dbMap(dbLoc)
    s = .sItems
End With
If clsSC.FastStringComp(s, "0") Then GetRoomItemsFromNumsFromRoomID = "": GoTo Notes
's = ReplaceFast(s, ":", "")
SplitFast s, tArr, ";"
s = ""
For i = LBound(tArr) To UBound(tArr)
    m = modItemManip.GetItemIDFromUnFormattedString(tArr(i))
    If m <= 0 Then GoTo nNext
    With dbItems(GetItemID(, m))
        s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & ","
    End With
nNext:
    If DE Then DoEvents
Next i
If GroupIt And Not clsSC.FastStringComp(s, "0") Then
    Erase tArr
    SplitFast s, tArr, ","
    s = ""
    n = 1
    For i = LBound(tArr) To UBound(tArr)
        If Not clsSC.FastStringComp(tArr(i), "") Then
            For j = LBound(tArr) To UBound(tArr)
                If i <> j Then
                    If Not clsSC.FastStringComp(tArr(i), "") Then
                        If clsSC.FastStringComp(tArr(i), tArr(j)) Then
                            n = n + 1
                            tArr(j) = ""
                        End If
                    End If
                End If
                If DE Then DoEvents
            Next
            If n > 1 Then
                s = s & CStr(n) & " " & tArr(i) & ","
            Else
                s = s & tArr(i) & ","
            End If
            n = 1
        End If
        If DE Then DoEvents
    Next
End If
Notes:
If clsSC.FastStringComp(s, "0") Then s = ""
s = s & modItemManip.GetListOfLettersFromGround(CLng(dbLoc))
If Colorize Then
    If s <> "" Then
        s = Left$(s, Len(s) - 1)
        s = ReplaceFast(s, ",", YELLOW & ", " & GREEN)
    End If
End If
GetRoomItemsFromNumsFromRoomID = s
End Function

Public Function GetRoomHiddenItemsFromNums(Index As Long, Optional Colorize As Boolean = False, Optional GroupIt As Boolean = False, Optional dbIndex As Long) As String
Dim s As String
Dim tArr() As String
Dim i As Long, j As Long
Dim dbLoc As Long
Dim m As Long
Dim n As Long
If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
dbLoc = dbPlayers(dbIndex).lDBLocation
With dbMap(dbLoc)
    s = .sHidden
End With
If clsSC.FastStringComp(s, "0") Then GetRoomHiddenItemsFromNums = "": GoTo Notes

SplitFast s, tArr, ";"
s = ""
For i = LBound(tArr) To UBound(tArr)
    m = modItemManip.GetItemIDFromUnFormattedString(tArr(i))
    If m <= 0 Then GoTo nNext
        With dbItems(m)
            s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & ","
        End With
nNext:
    If DE Then DoEvents
Next i
If GroupIt And Not clsSC.FastStringComp(s, "0") Then
    Erase tArr
    SplitFast s, tArr, ","
    s = ""
    n = 1
    For i = LBound(tArr) To UBound(tArr)
        If Not clsSC.FastStringComp(tArr(i), "") Then
            For j = LBound(tArr) To UBound(tArr)
                If i <> j Then
                    If Not clsSC.FastStringComp(tArr(i), "") Then
                        If clsSC.FastStringComp(tArr(i), tArr(j)) Then
                            n = n + 1
                            tArr(j) = ""
                        End If
                    End If
                End If
                If DE Then DoEvents
            Next
            If n > 1 Then
                s = s & CStr(n) & " " & tArr(i) & ","
            Else
                s = s & tArr(i) & ","
            End If
            n = 1
        End If
        If DE Then DoEvents
    Next
End If
Notes:
If clsSC.FastStringComp(s, "0") Then s = ""
s = s & modItemManip.GetListOfLettersFromHidden(dbLoc)
If Colorize Then
    s = ReplaceFast(s, ",", YELLOW & ", " & GREEN)
    s = Left$(s, Len(s) - 1)
End If
GetRoomHiddenItemsFromNums = s
End Function

Public Function GetEquipedItemsFromNums(Index As Long) As String
Dim s As String
Dim tArr() As String
Dim i As Long, j As Long
Dim m As Long
Dim n As Long
With dbPlayers(GetPlayerIndexNumber(Index))
    s = IIf(.sArms = "0", "", .sArms)
    s = s & IIf(.sBack = "0", "-1", .sBack)
    s = s & IIf(.sBody = "0", "-1", .sBody)
    s = s & IIf(.sEars = "0", "-1", .sEars)
    s = s & IIf(.sFace = "0", "-1", .sFace)
    s = s & IIf(.sFeet = "0", "-1", .sFeet)
    s = s & IIf(.sHands = "0", "-1", .sHands)
    s = s & IIf(.sHead = "0", "-1", .sHead)
    s = s & IIf(.sLegs = "0", "-1", .sLegs)
    s = s & IIf(.sNeck = "0", "-1", .sNeck)
    s = s & IIf(.sShield = "0", "-1", .sShield)
    s = s & IIf(.sWaist = "0", "-1", .sWaist)
    s = s & IIf(.sWeapon = "0", "-1", .sWeapon)
End With
SplitFast s, tArr, ";"
s = ""

For i = LBound(tArr) To UBound(tArr)
    If CLng(tArr(i)) = -1 Then GoTo nNext
    m = InStr(1, tArr(i), ":")
    
    If m = 0 Then GoTo nNext
    
    n = InStr(m, tArr(i), "/")
    
    If n = 0 Then GoTo nNext
    
    m = CLng(Mid$(tArr(i), m + 1, n - m - 1))
    
    If m = 0 Then GoTo nNext
    
    For j = LBound(dbItems) To UBound(dbItems)
        With dbItems(j)
            If .iID = m Then
                s = s & ReplaceFast(modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)), "|", " ") & Chr$(0) & .sItemName & Me.GetWornLocation(i) & vbCrLf
                Exit For
            End If
        End With
        If DE Then DoEvents
    Next j
nNext:
    If DE Then DoEvents
Next i
GetEquipedItemsFromNums = s

End Function

Public Function GetWornLocation(iNum As Long) As String
Select Case iNum
    Case 0
        GetWornLocation = " (Arms)"
    Case 1
        GetWornLocation = " (Back)"
    Case 2
        GetWornLocation = " (Body)"
    Case 3
        GetWornLocation = " (Ears)"
    Case 4
        GetWornLocation = " (Face)"
    Case 5
        GetWornLocation = " (Feet)"
    Case 6
        GetWornLocation = " (Hands)"
    Case 7
        GetWornLocation = " (Head)"
    Case 8
        GetWornLocation = " (Legs)"
    Case 9
        GetWornLocation = " (Neck)"
    Case 10
        GetWornLocation = " (Off-Hand)"
    Case 11
        GetWornLocation = " (Waist)"
    Case 12
        GetWornLocation = " (Weapon)"
    Case 13 To 18
        GetWornLocation = " (Ring)"
End Select
    
End Function
Public Function GetPlayersMaxGold(Index As Long) As Double
With dbPlayers(GetPlayerIndexNumber(Index))
    GetPlayersMaxGold = CDbl(.iStr) * 100
End With
End Function

Public Function GetPlayersMaxGoldDBINDEX(dbIndex As Long) As Double
With dbPlayers(dbIndex)
    GetPlayersMaxGoldDBINDEX = CDbl(.iStr) * 100
End With
End Function

Public Function GetGoldHere(Location As Long) As String
'function to get the amount of gold in a room
Dim TempGold As Double
If Location = 0 Then Exit Function
TempGold = dbMap(GetMapIndex(CLng(Location))).dGold
If TempGold <> 0 Then 'if there is gold there
    'set the message
    GetGoldHere = YELLOW & "there are " & GREEN & TempGold & YELLOW & " gold;"
Else
    'if no gold, set the variable to nothing
    GetGoldHere = ""
End If
End Function

Public Function GetGoldHereDBINDEX(dbLocation As Long) As String
'function to get the amount of gold in a room
Dim TempGold As Double
'If Location = 0 Then Exit Function
TempGold = dbMap(dbLocation).dGold
If TempGold <> 0 Then 'if there is gold there
    'set the message
    GetGoldHereDBINDEX = YELLOW & "there are " & GREEN & TempGold & YELLOW & " gold;"
Else
    'if no gold, set the variable to nothing
    GetGoldHereDBINDEX = ""
End If
End Function

Public Function GetPlayersMaxItems(Index As Long) As Long
Dim dbIndex As Long
dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    GetPlayersMaxItems = modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items])
End With
End Function

Public Function GetPlayersTotalItems(Index As Long) As Long
With dbPlayers(GetPlayerIndexNumber(Index))
    If .sInventory = "0" Then GetPlayersTotalItems = 0: Exit Function
    GetPlayersTotalItems = DCount(.sInventory, ";")
End With
End Function

Public Function GetPlayersMaxItemsDBINDEX(dbIndex As Long) As Long
With dbPlayers(dbIndex)
    GetPlayersMaxItemsDBINDEX = modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items])
End With
End Function

Public Function GetPlayersTotalItemsDBINDEX(dbIndex As Long) As Long
With dbPlayers(dbIndex)
    If .sInventory = "0" Then GetPlayersTotalItemsDBINDEX = 0: Exit Function
    GetPlayersTotalItemsDBINDEX = DCount(.sInventory, ";")
End With
End Function

Public Function GetItemsHere(Location As Long, Optional dbMapId As Long) As String
'function to get the amount of items in a room
Dim TempItemHere As String
If dbMapId = 0 Then dbMapId = GetMapIndex(Location)
TempItemHere = dbMap(dbMapId).sItems
GetItemsHere = TempItemHere
End Function

Public Function GetHiddenItemsHere(Location As Long) As String
'function to get the amount of items in a room
Dim TempItemHere As String
TempItemHere = dbMap(GetMapIndex(CLng(Location))).sHidden
'trim off the last ";"
If clsSC.FastStringComp(TempItemHere, "") Then
    TempItemHere = "0"
    dbMap(GetMapIndex(CLng(Location))).sHidden = TempItemHere
End If
TempItemHere = Left$(TempItemHere, Len(TempItemHere) - 1)
'make sure there is something there
If clsSC.FastStringComp(TempItemHere, "0") Then TempItemHere = ""
'set the message
TempItemHere = ReplaceFast(TempItemHere, ";", YELLOW & ", " & GREEN)
GetHiddenItemsHere = TempItemHere 'set the function
End Function

Public Function GetRoomLight(Location As Long, Optional dbMapId As Long) As Long
If dbMapId = 0 Then dbMapId = GetMapIndex(Location)
With dbMap(dbMapId)
    Select Case .lLight
        Case -25 To 25
            GetRoomLight = 0
        Case -75 To -26
            GetRoomLight = -1
        Case -125 To -76
            GetRoomLight = -2
        Case -175 To -126
            GetRoomLight = -3
        Case Is < -175
            GetRoomLight = -4
        Case 26 To 75
            GetRoomLight = 1
        Case 76 To 125
            GetRoomLight = 2
        Case 126 To 175
            GetRoomLight = 3
        Case Is > 175
            GetRoomLight = 4
    End Select
    If .iInDoor = 0 Then
        Select Case CLng(Mid$(modTime.TimeOfDay, 1, 2))
            Case 21 To 24, 1 To 5
                GetRoomLight = GetRoomLight - 3
            Case 6 To 9, 18 To 20
                GetRoomLight = GetRoomLight - 2
            Case 11 To 14
                GetRoomLight = GetRoomLight + 3
        End Select
    End If
End With
End Function

Public Function GetMonsHere(Location As Long, Optional Colorize As Boolean = False, Optional dbIndex As Long = 0, Optional dbMapId As Long) As String
Dim Mons$
Dim tArr() As String
Dim i As Long
Dim j As Long
If Location = 0 Then Exit Function
If dbMapId = 0 Then dbMapId = GetMapIndex(Location)
If Colorize Then
    Mons$ = Me.GetAllMonstersInRoom(Location, dbMapId)
Else
    Mons$ = Me.GetAllMonsNamesFromRoom(Location, dbMapId)
End If
If clsSC.FastStringComp(Mons$, "") Then Exit Function
If Not clsSC.FastStringComp(Mons$, "0") Then
    SplitFast Left$(Mons$, Len(Mons$) - 1), tArr, ";"
    Mons$ = ""
    If Colorize Then
        For i = LBound(tArr) To UBound(tArr)
            If tArr(i) <> "" Then
                j = Val(tArr(i))
                If j <= UBound(aMons) Then
                Select Case dbPlayers(dbIndex).iEvil
                    Case Is >= 40
                        If aMons(j).mHostile = True Then
                            Mons$ = Mons$ & BRIGHTMAGNETA & aMons(j).mName & YELLOW & ", "
                        Else
                            Select Case aMons(j).mEvil
                                Case Is < -41
                                    Mons$ = Mons$ & BRIGHTMAGNETA & aMons(j).mName & YELLOW & ", "
                                Case Else
                                    Mons$ = Mons$ & LIGHTBLUE & aMons(j).mName & YELLOW & ", "
                            End Select
                        End If
                    Case Is <= -41
                        If aMons(j).mHostile = True Then
                            Mons$ = Mons$ & BRIGHTMAGNETA & aMons(j).mName & YELLOW & ", "
                        Else
                            Select Case aMons(Val(tArr(i))).mEvil
                                Case Is > 40
                                    Mons$ = Mons$ & BRIGHTMAGNETA & aMons(j).mName & YELLOW & ", "
                                Case Else
                                    Mons$ = Mons$ & LIGHTBLUE & aMons(j).mName & YELLOW & ", "
                            End Select
                        End If
                    Case Else
                        If aMons(j).mHostile = True Then
                            Mons$ = Mons$ & BRIGHTMAGNETA & aMons(j).mName & YELLOW & ", "
                        Else
                            Mons$ = Mons$ & LIGHTBLUE & aMons(j).mName & YELLOW & ", "
                        End If
                End Select
                End If
            End If
        Next
    Else
        For i = LBound(tArr) To UBound(tArr)
            Mons$ = Mons$ & tArr(i) & ", "
        Next
    End If
    If Mons <> "" Then Mons$ = Left$(Mons$, Len(Mons$) - 2)
Else
    Mons$ = ""
End If
GetMonsHere = Mons$
End Function

Public Function GetAllMonsNamesFromRoom(lLocation As Long, Optional dbMapId As Long) As String
Dim i As Long
Dim s() As String
Dim t As String
If dbMapId = 0 Then dbMapId = GetMapIndex(lLocation)
With dbMap(dbMapId)
    If .sAMonIds <> "" Then
        SplitFast .sAMonIds, s, ";"
        For i = LBound(s) To UBound(s)
            If s(i) <> "" Then
                If Val(s(i)) <= UBound(aMons) Then
                    If aMons(Val(s(i))).mLoc = .lRoomID Then
                        t = t & aMons(Val(s(i))).mName & ";"
                    Else
                        .sAMonIds = ReplaceFast(.sAMonIds, s(i) & ";", "")
                    End If
                End If
            End If
            If DE Then DoEvents
        Next
    End If
End With
GetAllMonsNamesFromRoom = t
End Function

Public Function GetFamiliarsHere(Location As Long) As String
'function to get what players familiars are here
Dim Fams$
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iIndex <> 0 Then
            If .lLocation = Location Then
                If .lFamID <> 0 Then
                    If .sFamCustom <> "0" Then
                        Fams$ = Fams$ & BRIGHTBLUE & .sFamCustom & " the " & Chr$(0) & .sFamName & YELLOW & ", "
                    Else
                        Fams$ = Fams$ & BRIGHTBLUE & .sFamName & YELLOW & ", "
                    End If
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next
'if no one has familiars, set the function to nothing
If clsSC.FastStringComp(Fams$, "") Then GetFamiliarsHere = "": Exit Function
GetFamiliarsHere = Fams$ 'if there is, set the function to the findings
End Function

Public Function GetPlayersHere(Location As Long, Optional dbIndex As Long = -1) As String
'function to get all the players that are here
Dim TempPeeps As String
Dim i As Long
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iIndex <> 0 Then
            If i <> dbIndex Then
                If .lLocation = Location Then
                    If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False And .iGhostMode = 0 Then
                        If .iSneaking = 0 Then
                            If modMiscFlag.GetMiscFlag(i, Invisible) = 0 Then
                                If .lFamID <> 0 Then
                                    If .iHorse > 0 Then
                                        If .sFamCustom <> "0" Then
                                            TempPeeps = TempPeeps & GREEN & .sPlayerName & MAGNETA & " riding " & BRIGHTBLUE & .sFamCustom & " the " & .sFamName & MAGNETA & ", "
                                        Else
                                            TempPeeps = TempPeeps & GREEN & .sPlayerName & MAGNETA & " riding " & BRIGHTBLUE & .sFamName & MAGNETA & ", "
                                        End If
                                    Else
                                        If .sFamCustom <> "0" Then
                                            TempPeeps = TempPeeps & GREEN & .sPlayerName & MAGNETA & ", " & BRIGHTBLUE & .sFamCustom & " the " & .sFamName & MAGNETA & ", "
                                        Else
                                            TempPeeps = TempPeeps & GREEN & .sPlayerName & MAGNETA & ", " & BRIGHTBLUE & .sFamName & MAGNETA & ", "
                                        End If
                                    End If
                                Else
                                    TempPeeps = TempPeeps & GREEN & .sPlayerName & MAGNETA & ", "
                                End If
                            ElseIf dbIndex <> 0 Then
                                If modMiscFlag.GetMiscFlag(dbIndex, [See Invisible]) = 1 Then
                                    If .lFamID <> 0 Then
                                        If .iHorse > 0 Then
                                            If .sFamCustom <> "0" Then
                                                TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Invisible)" & MAGNETA & " riding " & BRIGHTBLUE & .sFamCustom & " the " & .sFamName & MAGNETA & ", "
                                            Else
                                                TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Invisible)" & MAGNETA & " riding " & BRIGHTBLUE & .sFamName & MAGNETA & ", "
                                            End If
                                        Else
                                            If .sFamCustom <> "0" Then
                                                TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Invisible)" & MAGNETA & ", " & BRIGHTBLUE & .sFamCustom & " the " & .sFamName & MAGNETA & ", "
                                            Else
                                                TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Invisible)" & MAGNETA & ", " & BRIGHTBLUE & .sFamName & MAGNETA & ", "
                                            End If
                                        End If
                                    Else
                                        TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Invisible)" & MAGNETA & ", "
                                    End If
                                End If
                            End If
                        ElseIf dbIndex <> 0 Then
                            If modMiscFlag.GetMiscFlag(dbIndex, [See Hidden]) = 1 Then
                                If .lFamID <> 0 Then
                                    If .iHorse > 0 Then
                                        If .sFamCustom <> "0" Then
                                            TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Hidden)" & MAGNETA & " riding " & BRIGHTBLUE & .sFamCustom & " the " & .sFamName & MAGNETA & ", "
                                        Else
                                            TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Hidden)" & MAGNETA & " riding " & BRIGHTBLUE & .sFamName & MAGNETA & ", "
                                        End If
                                    Else
                                        If .sFamCustom <> "0" Then
                                            TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Hidden)" & MAGNETA & ", " & BRIGHTBLUE & .sFamCustom & " the " & .sFamName & MAGNETA & ", "
                                        Else
                                            TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Hidden)" & MAGNETA & ", " & BRIGHTBLUE & .sFamName & MAGNETA & ", "
                                        End If
                                    End If
                                Else
                                    TempPeeps = TempPeeps & GREEN & .sPlayerName & BRIGHTRED & " (Hidden)" & MAGNETA & ", "
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If .lLocation = Location Then
                    If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False And .iGhostMode = 0 Then
                        If .lFamID <> 0 Then
                            If .sFamCustom <> "0" Then
                                TempPeeps = TempPeeps & BRIGHTBLUE & .sFamCustom & " the " & .sFamName & MAGNETA & ", "
                            Else
                                TempPeeps = TempPeeps & BRIGHTBLUE & .sFamName & MAGNETA & ", "
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next
GetPlayersHere = TempPeeps 'set the function
End Function

Public Function GetPlayersHereWithoutRiding(Location As Long, Optional dbIndex As Long) As String
'function to get all the players that are here
Dim TempPeeps As String
Dim i As Long
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iIndex <> 0 Then
            If .lLocation = Location Then
                If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False And .iGhostMode = 0 Then
                    If .iSneaking = 0 Then
                        If modMiscFlag.GetMiscFlag(i, Invisible) = 0 Then
                            TempPeeps = TempPeeps & .sPlayerName & ";"
                        ElseIf dbIndex <> 0 Then
                            If modMiscFlag.GetMiscFlag(dbIndex, [See Invisible]) = 1 Then
                                TempPeeps = TempPeeps & .sPlayerName & ";"
                            End If
                        End If
                    ElseIf dbIndex <> 0 Then
                        If modMiscFlag.GetMiscFlag(dbIndex, [See Hidden]) = 1 Then
                            TempPeeps = TempPeeps & .sPlayerName & ";"
                        End If
                    End If
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next
GetPlayersHereWithoutRiding = TempPeeps 'set the function
End Function

Public Function GetPlayersIDsHere(Location As Long) As String
'function to get all the players that are here

Dim TempPeeps As String
Dim i As Long
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iIndex <> 0 Then
            If .lLocation = Location Then
                If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False And .iGhostMode = 0 Then
                    If .iSneaking = 0 And modMiscFlag.GetMiscFlag(i, Invisible) = 0 Then
                        TempPeeps = TempPeeps & .lPlayerID & ";"
                    End If
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next
GetPlayersIDsHere = TempPeeps 'set the function
End Function

Public Function GetMapRoomType(dbMapId As Long) As String
Select Case dbMap(dbMapId).iType
    Case 0
        GetMapRoomType = "Normal"
    Case 1
        GetMapRoomType = "Shop"
    Case 2
        GetMapRoomType = "Level Trainer"
    Case 3
        GetMapRoomType = "Arena"
    Case 4
        GetMapRoomType = "Boss"
    Case 5
        GetMapRoomType = "Bank"
    Case 6
        GetMapRoomType = "Class Trainer"
End Select
End Function

Public Function GetMapSafe(dbMapId As Long) As String
Select Case dbMap(dbMapId).iSafeRoom
    Case 0
        GetMapSafe = "No"
    Case 1
        GetMapSafe = "Yes"
End Select
End Function

Public Function GetMapEnviron(dbMapId As Long) As String
Select Case dbMap(dbMapId).iInDoor
    Case 0
        GetMapEnviron = "Outdoor"
    Case 1
        GetMapEnviron = "Indoor"
    Case 2
        GetMapEnviron = "Underground"
End Select
End Function

Public Function GetPlayersDBIndexesHere(Location As Long) As String
'function to get all the players that are here
Dim TempPeeps As String
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iIndex <> 0 Then
            If .lLocation = Location Then
                If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False And .iGhostMode = 0 Then
                    TempPeeps = TempPeeps & i & ";"
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next
GetPlayersDBIndexesHere = TempPeeps 'set the function
End Function

Public Function GetPlayersDBIndexesHereNOTINPARTY(dbIndex As Long, Location As Long) As String
'function to get all the players that are here
Dim TempPeeps As String
Dim R As String
Dim i As Long
Dim s As String
Dim l As Long
Dim Arr() As String
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iIndex <> 0 And i <> dbIndex Then
            If .lLocation = Location Then
                If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False And .iGhostMode = 0 Then
                    TempPeeps = TempPeeps & i & ";"
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next
R = dbPlayers(dbIndex).sParty
s = TempPeeps
If Not clsSC.FastStringComp(R, "0") Then
    R = ReplaceFast(R, ":", "")
    SplitFast R, Arr, ";"
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) <> "" Then
            l = GetPlayerIndexNumber(CLng(Val(Arr(i))))
            s = ReplaceFast(s, CStr(l) & ";", "")
        End If
        If DE Then DoEvents
    Next
End If
GetPlayersDBIndexesHereNOTINPARTY = s 'set the function
End Function


Public Function GetRoomDesc(Location As Long, Optional dbMapId As Long) As String
'function to get the room description
If dbMapId = 0 Then GetMapIndex (Location)
GetRoomDesc = dbMap(dbMapId).sRoomDesc  'set the function
End Function

Public Function GetRoomExits(Location As Long, Optional dd As Long) As String
'function to get all the exits in the room
Dim tVal As String
Dim sS As String
'Dim dd As Long
If Location = 0 Then Exit Function
If dd = 0 Then dd = GetMapIndex(Location)
With dbMap(dd)
    If .lDN <> 0 Then
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort("n"))
        If .lDN = 1 Or .lDN = 2 Then
            tVal = tVal & "closed " & sS & " north,"
        Else
            tVal = tVal & "open " & sS & " north,"
        End If
    ElseIf .lNorth <> 0 Then
        tVal = tVal & "north,"
    End If
    If .lDS <> 0 Then
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort("s"))
        If .lDS = 1 Or .lDS = 2 Then
            tVal = tVal & "closed " & sS & " south,"
        Else
            tVal = tVal & "open " & sS & " south,"
        End If
    ElseIf .lSouth <> 0 Then
        tVal = tVal & "south,"
    End If
    If .lDE <> 0 Then
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort("e"))
        If .lDE = 1 Or .lDE = 2 Then
            tVal = tVal & "closed " & sS & " east,"
        Else
            tVal = tVal & "open " & sS & " east,"
        End If
    ElseIf .lEast <> 0 Then
        tVal = tVal & "east,"
    End If
    If .lDW <> 0 Then
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort("w"))
        If .lDW = 1 Or .lDW = 2 Then
            tVal = tVal & "closed " & sS & " west,"
        Else
            tVal = tVal & "open " & sS & " west,"
        End If
    ElseIf .lWest <> 0 Then
        tVal = tVal & "west,"
    End If
    If .lDNE <> 0 Then
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort("ne"))
        If .lDNE = 1 Or .lDNE = 2 Then
            tVal = tVal & "closed " & sS & " northeast,"
        Else
            tVal = tVal & "open " & sS & " northeast,"
        End If
    ElseIf .lNorthEast <> 0 Then
        tVal = tVal & "northeast,"
    End If
    If .lDNW <> 0 Then
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort("nw"))
        If .lDNW = 1 Or .lDNW = 2 Then
            tVal = tVal & "closed " & sS & " northwest,"
        Else
            tVal = tVal & "open " & sS & " northwest,"
        End If
    ElseIf .lNorthWest <> 0 Then
        tVal = tVal & "northwest,"
    End If
    If .lDSE <> 0 Then
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort("se"))
        If .lDSE = 1 Or .lDSW = 2 Then
            tVal = tVal & "closed " & sS & " southeast,"
        Else
            tVal = tVal & "open " & sS & " southeast,"
        End If
    ElseIf .lSouthEast <> 0 Then
        tVal = tVal & "southeast,"
    End If
    If .lDSW <> 0 Then
        sS = modgetdata.DoorOrGate(dd, modgetdata.GetDirIndexFromShort("sw"))
        If .lDSW = 1 Or .lDSW = 2 Then
            tVal = tVal & "closed " & sS & " southwest,"
        Else
            tVal = tVal & "open " & sS & " southwest,"
        End If
    ElseIf .lSouthWest <> 0 Then
        tVal = tVal & "southwest,"
    End If
    If .lDD <> 0 Then
        If .lDD = 1 Or .lDD = 2 Then
            tVal = tVal & "closed trap door down,"
        Else
            tVal = tVal & "open trap door down,"
        End If
    ElseIf .lDown <> 0 Then
        tVal = tVal & "down,"
    End If
    If .lDU <> 0 Then
        If .lDU = 1 Or .lDU = 2 Then
            tVal = tVal & "closed hatch up,"
        Else
            tVal = tVal & "open hatch up,"
        End If
    ElseIf .lUp <> 0 Then
        tVal = tVal & "up,"
    End If
    If clsSC.FastStringComp(tVal, "") Then tVal = "None."
    tVal = Left$(tVal, Len(tVal) - 1) & "."
    tVal = ReplaceFast(tVal, ",", YELLOW & ", " & GREEN)
    GetRoomExits = tVal  'set the function
End With
End Function

Public Function GetRoomTitle(Location As Long, Optional dbMapId As Long) As String
'function to get the title of the room
If Location = 0 Then Exit Function
If dbMapId = 0 Then dbMapId = GetMapIndex(Location)
GetRoomTitle = dbMap(dbMapId).sRoomTitle  'set the function
End Function

Public Function GetOppositeDirection(Current As String, Optional ABOVEBELOW As Boolean = False) As String
'get the opposite direction of a supplyed direction
'pretty self explanitory
If ABOVEBELOW Then
    Select Case LCaseFast(Current)
        Case "downwards"
            GetOppositeDirection = "above"
        Case "upwards"
            GetOppositeDirection = "below"
    End Select
    If GetOppositeDirection <> "" Then Exit Function
End If
Select Case LCaseFast(Current)
    Case "north"
        GetOppositeDirection = "south"
    Case "south"
        GetOppositeDirection = "north"
    Case "east"
        GetOppositeDirection = "west"
    Case "west"
        GetOppositeDirection = "east"
    Case "up"
        GetOppositeDirection = "down"
    Case "down"
        GetOppositeDirection = "up"
    Case "northwest"
        GetOppositeDirection = "southeast"
    Case "northeast"
        GetOppositeDirection = "southwest"
    Case "southwest"
        GetOppositeDirection = "northeast"
    Case "southeast"
        GetOppositeDirection = "northwest"
    Case "above"
        GetOppositeDirection = "below"
    Case "below"
        GetOppositeDirection = "above"
    Case "upwards"
        GetOppositeDirection = "downwards"
    Case "downwards"
        GetOppositeDirection = "upwards"
    
End Select
End Function

Public Function GetmDodge(Index As Long, MonsterID As Long) As Long
'Gets the chance of a monster dodging the player's swings
Dim d As Double
d = (aMons(MonsterID).mAc) '* 10)
If d > 98 Then d = d / 5
If d <= 0 Then d = 2
d = d / 100
d = RoundFast(d, 2)
d = d * 100
If d > 98 Then d = 98
GetmDodge = CLng(d) 'set the function
End Function

Public Function GetmMaxHit(dbIndex As Long, MonsterID As Long) As Long
'get the monsters max number
Dim d As Double
d = (((aMons(MonsterID).mAc + 1) * 2) + RndNumber(1, 5) + RndNumber(1, aMons(MonsterID).mAc + 5)) / 100
d = RoundFast(d, 2)
d = d * 100
If d > 98 Then d = 98
If d < 10 Then d = 15
GetmMaxHit = CLng(d)
End Function

Public Function GetMonID(Index As Long) As Long
'get the monsters ID number according to the aMons array
GetMonID = dbPlayers(GetPlayerIndexNumber(Index)).dMonsterID
End Function

Public Function GetpDodge(dbIndex As Long) As Long
'get the players dodge
Dim d As Double
With dbPlayers(dbIndex)
    d = (.iAgil + .iDodge + .iAC) / 100
    d = RoundFast(d, 2)
    d = d * 100
    If d < 0 Then d = 0
    If d > 62 Then d = 62
    If .iPartyRank = 2 Then d = d + 10
    If d > 70 Then d = 70
End With
GetpDodge = CLng(d)
End Function

Public Function GetpMR(dbIndex As Long) As Long
'get the players dodge
Dim d As Double
With dbPlayers(dbIndex)
    d = (.iInt + .iDex + .iAC + .iLevel) / 100
    d = RoundFast(d, 2)
    d = d * 100
    If d < 0 Then d = 0
    If d > 62 Then d = 62
    If .iPartyRank = 2 Then d = d + 10
    If d > 70 Then d = 70
End With
GetpMR = CLng(d)
End Function

Public Function GetPlayersMRvsMon(dbIndex As Long) As Long
i = modMiscFlag.GetStatsPlusTotal(dbIndex, [Magic Resistance])
If i > 75 Then i = 75
GetPlayersMRvsMon = i
End Function

Public Function GetpMaxHit(Index As Long) As Long
'get the players max hit number
Dim d As Double
With dbPlayers(GetPlayerIndexNumber(CLng(Index)))
    d = (((.iAC + 1) * 2) + .iDex + .iStr + (RndNumber(1, .iCha + .iAgil))) / 100
    d = RoundFast(d, 2) '+ RoundFast((RndNumber(1, .iCha + .iAgil) / 100), 2)
    d = d * 100
    If d > 98 Then d = 98
    Select Case d
        Case Is < 5
            d = 75
        Case Is < 10
            d = 70
        Case Is < 15
            d = 60
        Case Is < 20
            d = 45
        Case Is < 25
            d = 35
    End Select
    d = d + GetpSwings(Index)
    GetpMaxHit = CLng(d)
End With
End Function

Public Function GetpSwings(Index As Long) As Long
'get the amount of swings a player gets with the weapon
Dim Swings As Long
Dim dbIndex As Long
dbIndex = GetPlayerIndexNumber(CLng(Index))
With dbPlayers(dbIndex)
    Swings = (.iAgil + .iDex + (.iInt \ 2) + (.iCha \ 3) + (.iStr \ 3) - pWeapon(Index).wSpeed) / 12
    Swings = CLng(RoundFast(CDbl(Swings), 0))
    Select Case .dStamina
        Case Is < 10
            Swings = Swings - 3
        Case Is < 40
            Swings = Swings - 2
        Case Is < 60
            Swings = Swings - 1
    End Select
    Select Case .dHunger
        Case Is < 10
            Swings = Swings - 2
        Case Is < 40
            Swings = Swings - 1
    End Select
    Select Case .iLevel
        Case Is < 10
        Case Is < 20
        Case Is < 30
            Swings = Swings + 1
        Case Is < 40
        Case Is < 50
            Swings = Swings + 2
        Case Is < 60
            Swings = Swings + 3
        Case Else
            Swings = Swings + 4
    End Select
End With
Swings = Swings + pWeapon(Index).wSB
If Swings < 1 Then Swings = 1
'and max swings is 6
If Swings > 6 Then Swings = 6
GetpSwings = Swings
End Function

Public Function GetSpellChanceFromdbSpell(dbIndex As Long, dbSpellID As Long) As Long
'get the players chance of casting a spell max
Dim Chance&
Dim lBonus&

With dbPlayers(dbIndex)
    Chance& = modMiscFlag.GetStatsPlusTotal(dbIndex, [Spell Casting]) + dbSpells(dbSpellID).iDifficulty
    Select Case Chance&
        Case Is > 99, Is < 150
            lBonus& = 1
        Case Is < 170
            lBonus& = 2
        Case Is < 200
            lBonus& = 3
        Case Is >= 200
            lBonus& = 4
    End Select
End With
If Chance& > 96 Then Chance& = 96
Chance& = Chance& + lBonus&
GetSpellChanceFromdbSpell = Chance&
End Function

Public Function GetSpellChance(Index As Long) As Long
'get the players chance of casting a spell max
Dim Chance&, dbIndex&
Dim lBonus&
dbIndex& = GetPlayerIndexNumber(CLng(Index))

With dbPlayers(dbIndex&)
    Chance& = modMiscFlag.GetStatsPlusTotal(dbIndex, [Spell Casting]) + dbSpells(GetSpellID(pWeapon(Index).wSpellName)).iDifficulty
    Select Case Chance&
        Case Is > 99, Is < 150
            lBonus& = 1
        Case Is < 170
            lBonus& = 2
        Case Is < 200
            lBonus& = 3
        Case Is >= 200
            lBonus& = 4
    End Select
End With
If Chance& > 96 Then Chance& = 96
Chance& = Chance& + lBonus&
GetSpellChance = Chance&
End Function

Public Function GetLongDir(sDir As String) As String
Select Case LCaseFast(sDir)
    Case "nw"
        GetLongDir = "northwest"
    Case "ne"
        GetLongDir = "northeast"
    Case "sw"
        GetLongDir = "southwest"
    Case "se"
        GetLongDir = "southeast"
    Case "n"
        GetLongDir = "north"
    Case "s"
        GetLongDir = "south"
    Case "e"
        GetLongDir = "east"
    Case "w"
        GetLongDir = "west"
    Case "u"
        GetLongDir = "up"
    Case "d"
        GetLongDir = "down"
    Case Else
        GetLongDir = "-1"
End Select
End Function

Public Function DoorOrGate(dbMapId As Long, DirIndex As Long) As String
Select Case modMapFlags.GetMapFlag(dbMapId, mapGate, DirIndex)
    Case 0
        DoorOrGate = "door"
    Case 1
        DoorOrGate = "gate"
    Case Else
        DoorOrGate = "door"
End Select
End Function


Public Function GetDirIndexFromShort(sDir As String) As Long
Select Case LCaseFast(sDir)
    Case "nw"
        GetDirIndexFromShort = 4
    Case "ne"
        GetDirIndexFromShort = 5
    Case "sw"
        GetDirIndexFromShort = 6
    Case "se"
        GetDirIndexFromShort = 7
    Case "n"
        GetDirIndexFromShort = 0
    Case "s"
        GetDirIndexFromShort = 1
    Case "e"
        GetDirIndexFromShort = 2
    Case "w"
        GetDirIndexFromShort = 3
    Case "u"
        GetDirIndexFromShort = 8
    Case "d"
        GetDirIndexFromShort = 9
    Case Else
        GetDirIndexFromShort = 0
End Select
End Function

Public Function GetDragon() As String
Dim Dragon$
Rem sub to make the starting image
Dragon$ = Dragon$ & GREEN & "                   ___====-_  _-====___" & vbCrLf
Dragon$ = Dragon$ & GREEN & "              --^^^" & BLUE & "#####" & GREEN & "//      \\" & BLUE & "#####" & GREEN & "^^^--_" & vbCrLf
Dragon$ = Dragon$ & GREEN & "          _-^" & BLUE & "##########" & GREEN & "//" & YELLOW & " (    )" & GREEN & " \\" & BLUE & "##########" & GREEN & "^-_" & vbCrLf
Dragon$ = Dragon$ & GREEN & "         -" & BLUE & "############" & GREEN & "//" & YELLOW & "  |\^^/|  " & GREEN & "\\" & BLUE & "############" & GREEN & "-" & vbCrLf
Dragon$ = Dragon$ & GREEN & "       _/" & BLUE & "############" & GREEN & "//" & YELLOW & "   (" & RED & "@::@" & YELLOW & ")" & GREEN & "   \\" & BLUE & "############" & GREEN & "\_" & vbCrLf
Dragon$ = Dragon$ & GREEN & "      /" & BLUE & "#############" & GREEN & "((     " & YELLOW & "\\//     " & GREEN & "))" & BLUE & "#############" & GREEN & "\" & vbCrLf
Dragon$ = Dragon$ & GREEN & "     -" & BLUE & "###############" & GREEN & "\\    " & YELLOW & "(" & LIGHTBLUE & "oo" & YELLOW & ")    " & GREEN & "//" & BLUE & "###############" & GREEN & "-" & vbCrLf
Dragon$ = Dragon$ & GREEN & "    -" & BLUE & "#################" & GREEN & "\\  " & YELLOW & "/ " & RED & "VV" & YELLOW & " \  " & GREEN & "//" & BLUE & "#################" & GREEN & "-" & vbCrLf
Dragon$ = Dragon$ & GREEN & "   -" & BLUE & "###################" & GREEN & "\\" & YELLOW & "/      \" & GREEN & "//" & BLUE & "###################" & GREEN & "-" & vbCrLf
Dragon$ = Dragon$ & GREEN & "  _" & BLUE & "#" & GREEN & "/|" & BLUE & "##########" & GREEN & "/\" & BLUE & "######" & GREEN & "(   " & YELLOW & "/\" & GREEN & "   )" & BLUE & "######" & GREEN & "/\" & BLUE & "##########" & GREEN & "|\" & BLUE & "#" & GREEN & "_" & vbCrLf
Dragon$ = Dragon$ & GREEN & "  |/ |" & BLUE & "#" & GREEN & "/\" & BLUE & "#" & GREEN & "/\" & BLUE & "#" & GREEN & "/\/  \" & BLUE & "#" & GREEN & "/\" & BLUE & "##" & GREEN & "\  " & YELLOW & "|  |  " & GREEN & "/" & BLUE & "##" & GREEN & "/\" & BLUE & "#" & GREEN & "/  \/\" & BLUE & "#" & GREEN & "/\" & BLUE & "#" & GREEN & "/\" & BLUE & "#" & GREEN & "| \|" & vbCrLf
Dragon$ = Dragon$ & GREEN & "  `  |/  V  V  `   V  \" & BLUE & "#" & GREEN & "\" & YELLOW & "| |  | |" & GREEN & "/" & BLUE & "#" & GREEN & "/  V   '  V  V  \|  '" & vbCrLf
Dragon$ = Dragon$ & GREEN & "     `   `  `      `   " & YELLOW & "/ | |  | | \" & GREEN & "   '      '  '   '" & vbCrLf
Dragon$ = Dragon$ & YELLOW & "                      (  | |  | |  )" & vbCrLf
Dragon$ = Dragon$ & YELLOW & "                     __\ | |  | | /__" & vbCrLf
Dragon$ = Dragon$ & YELLOW & "                    (" & MAGNETA & "vvv" & YELLOW & "(" & MAGNETA & "VVV" & YELLOW & ")(" & MAGNETA & "VVV" & YELLOW & ")" & MAGNETA & "vvv" & YELLOW & ")" & vbCrLf
GetDragon = Dragon$
End Function

Public Function GetShortDir(sDir As String) As String
Select Case LCaseFast(sDir)
    Case "northwest"
        GetShortDir = "nw"
    Case "northeast"
        GetShortDir = "ne"
    Case "southwest"
        GetShortDir = "sw"
    Case "southeast"
        GetShortDir = "se"
    Case "north"
        GetShortDir = "n"
    Case "south"
        GetShortDir = "s"
    Case "east"
        GetShortDir = "e"
    Case "west"
        GetShortDir = "w"
    Case "up"
        GetShortDir = "u"
    Case "down"
        GetShortDir = "d"
End Select
End Function

Public Function GetStatLine(dbIndex As Long) As String
Dim ToSend$
Dim s As String
Dim d As Double
Dim m As String
Dim n As String
Dim a As String
Dim clsNtW As clsNumsToWords
If dbIndex = 0 Then Exit Function
With dbPlayers(dbIndex)
    s = .sStatline
    If InStr(1, s, ";hp") Then s = ReplaceFast(s, ";hp", CStr(.lHP))
    If InStr(1, s, ";mhp") Then s = ReplaceFast(s, ";mhp", CStr(.lMaxHP))
    If InStr(1, s, ";ma") Then
        s = ReplaceFast(s, ";ma", CStr(.lMana))
    End If
    If InStr(1, s, ";mma") Then
        s = ReplaceFast(s, ";mma", CStr(.lMaxMana))
    End If
    If InStr(1, s, ";texp") Then s = ReplaceFast(s, ";texp", CStr(.dTotalEXP))
    If InStr(1, s, ";nexp") Then s = ReplaceFast(s, ";nexp", CStr(.dEXPNeeded))
    If InStr(1, s, ";cexp") Then s = ReplaceFast(s, ";cexp", CStr(.dEXP))
    If InStr(1, s, ";gold") Then s = ReplaceFast(s, ";gold", CStr(.dGold))
    
    If InStr(1, s, ";%exp") Then s = ReplaceFast(s, ";%exp", CStr(((FormatNumber(.dEXP / .dEXPNeeded, 2)) * 100) & "%"))
    If InStr(1, s, ";%stamina") Then
        d = .dStamina
        If d > 100 Then d = 100
        s = ReplaceFast(s, ";%stamina", CStr(d) & "%")
    End If
    If InStr(1, s, ";%hunger") Then
        d = .dHunger
        If d > 100 Then d = 100
        s = ReplaceFast(s, ";%hunger", CStr(d) & "%")
    End If
    If InStr(1, s, ";lives") Then s = ReplaceFast(s, ";lives", CStr(.iLives))
    If InStr(1, s, ";famhp") Then s = ReplaceFast(s, ";famhp", CStr(.lFamCHP))
    If InStr(1, s, ";fammaxhp") Then s = ReplaceFast(s, ";fammaxhp", CStr(.lFamMHP))
    If InStr(1, s, ";date") Then s = ReplaceFast(s, ";date", CStr(modTime.MonthOfYear & "/" & CStr(modTime.DayOfMonth) & "/" & modTime.CurYear))
    If InStr(1, s, ";time") Then
        Set clsNtW = New clsNumsToWords
        a = Mid$(TimeOfDay, 1, 2)
        If CLng(Mid$(TimeOfDay, 4, 2)) > 29 Then a = a & ":30"
        If Val(Left$(a, 2)) > 12 Then
            m = CStr(Val(Left$(a, 2)) - 12)
            If Len(m) < 2 Then m = "0" & m
            Mid$(a, 1, 2) = m
        End If
        m = LCaseFast(clsNtW.ConvertNumberToText(Mid$(a, 4, 2)))
        n = LCaseFast(clsNtW.ConvertNumberToText(Mid$(a, 1, 2)))
        a = n & " " & m & GREEN & modTime.GetDayNight
        If clsSC.FastStringComp(Left$(a, 1), "0") Then Mid$(a, 1, 1) = " "
        s = ReplaceFast(s, ";time", a)
        Set clsNtW = Nothing
    End If
    ToSend$ = ToSend & BRIGHTWHITE & "[" & s & "]"
    If .iResting = 1 Then ToSend = ToSend & "(Resting) "
    If .iMeditating = 1 Then ToSend = ToSend & "(Meditating} "
    If .iSneaking = 1 Then ToSend = ToSend & "(Sneaking)"
    If Right$(ToSend, 1) = " " Then ToSend = Mid$(ToSend, 1, Len(ToSend) - 1)
    ToSend = ToSend & ": "
    GetStatLine = ReplaceFakeANSI(dbIndex, ToSend$ & WHITE)
End With
End Function

Public Function GetPlayersSpellIds(Index As Long) As String
GetPlayersSpellIds = dbPlayers(GetPlayerIndexNumber(Index)).sSpells 'get the players spells
End Function

Public Function GetSpellHeal(SpellMin As Long, SpellMax As Long, SpellModify As Long, MaxLevel As Long) As Long
Dim Healing&
SpellMax = SpellMax + (SpellModify * MaxLevel)
Healing& = RndNumber(CDbl(SpellMin), CDbl(SpellMax))
GetSpellHeal = Healing&
End Function

Public Function GetSpellMaxDamage(dbIndex As Long, dbSpellID) As Long
Dim lMax As Long
With dbSpells(dbSpellID)
    If dbPlayers(dbIndex).iLevel > .iLevelMax Then lMax = .iLevelMax Else lMax = dbPlayers(dbIndex).iLevel
    GetSpellMaxDamage = .lMaxDam + (.iLevelModify * lMax)
End With
End Function

Public Function GetSpellDam(dbSpellID As Long, lLevel As Long) As Long
Dim lMax As Long
With dbSpells(dbSpellID)
    If lLevel > .iLevelMax Then lMax = .iLevelMax Else lMax = l
    GetSpellDam = .lMaxDam + (.iLevelModify * lMax)
End With
End Function

Public Function GetPlayersSpellShorts(Index As Long) As String
GetPlayersSpellShorts = dbPlayers(GetPlayerIndexNumber(Index)).sSpellShorts 'get the players spells
End Function

Public Function sGetRoomExits(Index As Long) As String
Dim sExits As String
With dbMap(GetMapIndex(dbPlayers(GetPlayerIndexNumber(Index)).lLocation))
    If .lNorth <> 0 Then sExits = sExits & .lNorth & ","
    If .lSouth <> 0 Then sExits = sExits & .lSouth & ","
    If .lWest <> 0 Then sExits = sExits & .lWest & ","
    If .lEast <> 0 Then sExits = sExits & .lEast & ","
    If .lUp <> 0 Then sExits = sExits & .lUp & ","
    If .lDown <> 0 Then sExits = sExits & .lDown & ","
    If .lNorthWest <> 0 Then sExits = sExits & .lNorthWest & ","
    If .lNorthEast <> 0 Then sExits = sExits & .lNorthEast & ","
    If .lSouthWest <> 0 Then sExits = sExits & .lSouthWest & ","
    If .lSouthEast <> 0 Then sExits = sExits & .lSouthEast & ","
End With
If clsSC.FastStringComp(sExits, "") Then sGetRoomExits = "": Exit Function
sGetRoomExits = Left$(sExits, Len(sExits) - 1)
End Function

Public Function GetRoomExits2(dbRoomID As Long) As String
Dim sExits As String
If l2RoomID < 1 Then Exit Function
With dbMap(dbRoomID)
    If .lNorth <> 0 Then sExits = sExits & .lNorth & ","
    If .lSouth <> 0 Then sExits = sExits & .lSouth & ","
    If .lWest <> 0 Then sExits = sExits & .lWest & ","
    If .lEast <> 0 Then sExits = sExits & .lEast & ","
    If .lUp <> 0 Then sExits = sExits & .lUp & ","
    If .lDown <> 0 Then sExits = sExits & .lDown & ","
    If .lNorthWest <> 0 Then sExits = sExits & .lNorthWest & ","
    If .lNorthEast <> 0 Then sExits = sExits & .lNorthEast & ","
    If .lSouthWest <> 0 Then sExits = sExits & .lSouthWest & ","
    If .lSouthEast <> 0 Then sExits = sExits & .lSouthEast & ","
End With
If clsSC.FastStringComp(sExits, "") Then GetRoomExits2 = "": Exit Function
GetRoomExits2 = Left$(sExits, Len(sExits) - 1)
End Function

Public Function GetRoomExitFrom2Points(dblPoint1 As Long, lPoint2 As Long) As String
Dim sExits As String
On Error GoTo eh1:
With dbMap(dblPoint1)
    If .lNorth = lPoint2 Then sExits = "n"
    If .lSouth = lPoint2 Then sExits = "s"
    If .lWest = lPoint2 Then sExits = "w"
    If .lEast = lPoint2 Then sExits = "e"
    If .lUp = lPoint2 Then sExits = "u"
    If .lDown = lPoint2 Then sExits = "d"
    If .lNorthWest = lPoint2 Then sExits = "nw"
    If .lNorthEast = lPoint2 Then sExits = "ne"
    If .lSouthWest = lPoint2 Then sExits = "sw"
    If .lSouthEast = lPoint2 Then sExits = "se"
    GetRoomExitFrom2Points = sExits
End With
eh1:
End Function

Public Function GetPlayerRankFromNum(Num As Long) As String
Select Case Num
    Case 0
        GetPlayerRankFromNum = "none"
    Case 1
        GetPlayerRankFromNum = "frontrank"
    Case 2
        GetPlayerRankFromNum = "backrank"
End Select
End Function

Public Function GetHitPositionID() As Long
GetHitPositionID = RndNumber(0, 11)

End Function

Public Function GetUnformatedStringFromID(dbIndex As Long, i As Long) As String
With dbPlayers(dbIndex)
    Select Case i
        Case 0
            GetUnformatedStringFromID = .sArms
        Case 1
            GetUnformatedStringFromID = .sBack
        Case 2
            GetUnformatedStringFromID = .sBody
        Case 3
            GetUnformatedStringFromID = .sEars
        Case 4
            GetUnformatedStringFromID = .sFace
        Case 5
            GetUnformatedStringFromID = .sFeet
        Case 6
            GetUnformatedStringFromID = .sHands
        Case 7
            GetUnformatedStringFromID = .sHead
        Case 8
            GetUnformatedStringFromID = .sLegs
        Case 9
            GetUnformatedStringFromID = .sNeck
        Case 10
            GetUnformatedStringFromID = .sShield
        Case 11
            GetUnformatedStringFromID = .sWaist
    End Select
End With
End Function

Public Function GetItemsDurPercent(CurDur As Double, RealDur As Double) As String
GetItemsDurPercent = YELLOW & "This item appears to be in " & GetDurColor(CurDur, RealDur) & YELLOW & " condition." & vbCrLf
End Function

Public Function GetDurColor(CurDur As Double, RealDur As Double) As String
If RealDur <= 0 Then RealDur = 1
Select Case (100 - (FormatNumber(CurDur / RealDur, 2) * 100))
    Case 81 To 100, Is > 100
        GetDurColor = BRIGHTRED & "horrid"
    Case 61 To 80
        GetDurColor = RED & "horrible"
    Case 41 To 60
        GetDurColor = BRIGHTYELLOW & "worn"
    Case 21 To 40
        GetDurColor = BLUE & "pretty good"
    Case 1 To 20
        GetDurColor = GREEN & "almost new"
    Case 0, Is < 0
        GetDurColor = YELLOW & "new"
End Select
End Function

Public Function GetGenderDesc(dbIndex As Long) As String
With dbPlayers(dbIndex)
    Select Case .iGender
        Case -1
            GetGenderDesc = "It"
        Case 0
            GetGenderDesc = "He"
        Case 1
            GetGenderDesc = "She"
    End Select
End With
End Function

Public Function GetGenderPronoun(dbIndex As Long, Optional AddStoit As Boolean = False, Optional Capitalize As Boolean = False) As String
If Not Capitalize Then
    With dbPlayers(dbIndex)
        Select Case .iGender
            Case -1
                If AddStoit Then
                    GetGenderPronoun = "its"
                Else
                    GetGenderPronoun = "it"
                End If
            Case 0
                If AddStoit Then
                    GetGenderPronoun = "his"
                Else
                    GetGenderPronoun = "him"
                End If
            Case 1
                GetGenderPronoun = "her"
        End Select
    End With
Else
    With dbPlayers(dbIndex)
        Select Case .iGender
            Case -1
                If AddStoit Then
                    GetGenderPronoun = "Its"
                Else
                    GetGenderPronoun = "It"
                End If
            Case 0
                If AddStoit Then
                    GetGenderPronoun = "His"
                Else
                    GetGenderPronoun = "Him"
                End If
            Case 1
                GetGenderPronoun = "Her"
        End Select
    End With
End If
End Function

Public Function GetMoveDelay(dbIndex As Long) As Long
With dbPlayers(dbIndex)
    Select Case .dStamina
        Case Is <= 0
            GetMoveDelay = 950
        Case Is < 10
            GetMoveDelay = 800
        Case Is < 20
            GetMoveDelay = 700
        Case Is < 30
            GetMoveDelay = 600
        Case Is < 40
            GetMoveDelay = 400
        Case Is < 50
            GetMoveDelay = 350
        Case Is < 60
            GetMoveDelay = 300
        Case Is < 70
            GetMoveDelay = 290
        Case Is < 80
            GetMoveDelay = 180
        Case Is < 90
            GetMoveDelay = 160
        Case Is < 100
            GetMoveDelay = 135
        Case Is < 110
            GetMoveDelay = 118
        Case Is > 110
            GetMoveDelay = 100
    End Select
    If .iHorse > 0 Then
        Select Case .iHorse
            Case 1
                GetMoveDelay = GetMoveDelay + 75
            Case 2
                GetMoveDelay = GetMoveDelay + 50
            Case 3
                GetMoveDelay = GetMoveDelay + 25
            Case 5
                GetMoveDelay = GetMoveDelay - 25
            Case 6
                GetMoveDelay = GetMoveDelay - 50
            Case 7
                GetMoveDelay = GetMoveDelay - 75
            Case 8
                GetMoveDelay = GetMoveDelay - 85
            Case 9
                GetMoveDelay = GetMoveDelay - 90
            Case 10
                GetMoveDelay = GetMoveDelay - 95
            Case Is > 11
                GetMoveDelay = 0
        End Select
    End If
End With
End Function

Public Function GetRoomDescription(dbIndex As Long, iLocation As Long, Optional IncludeDesc As Boolean = True, Optional UpdateDB As Boolean = False)
Dim TempItemHere As String, ToSend$
Dim TempPeeps As String, TempGold$
Dim Mons$
Dim dbMapId As Long
With dbPlayers(dbIndex)
    dbMapId = GetMapIndex(iLocation)
    If UpdateDB Then If .lDBLocation <> dbMapId Then .lDBLocation = dbMapId
    Select Case .iVision + modgetdata.GetRoomLight(iLocation)
        Case Is < -3
            ToSend$ = WHITE & "This room is too dark, you can't see a thing!" & WHITE & vbCrLf
            GetRoomDescription = ToSend$
            Exit Function
    End Select
    TempPeeps = modgetdata.GetPlayersHere(iLocation, dbIndex)
    'TempPeeps = TempPeeps & modgetdata.GetFamiliarsHere(iLocation)
    TempGold$ = modgetdata.GetGoldHereDBINDEX(dbMapId)
    TempItemHere = modgetdata.GetRoomItemsFromNumsFromRoomID(CLng(iLocation), True, True, dbMapId)
    Mons$ = modgetdata.GetMonsHere(iLocation, True, dbIndex, dbMapId)
    If clsSC.FastStringComp(TempItemHere, "") Then
        If Not clsSC.FastStringComp(TempGold$, "") Then TempItemHere = Left$(TempGold$, Len(TempGold$) - 1)
    Else
        TempItemHere = TempGold$ & TempItemHere
    End If
    TempItemHere = ReplaceFast(TempItemHere, ";", YELLOW & ", " & GREEN)
    If IncludeDesc Then
        ToSend$ = BRIGHTLIGHTBLUE & modgetdata.GetRoomTitle( _
            iLocation, dbMapId) & vbCrLf & BRIGHTWHITE & modgetdata.GetRoomDesc( _
            iLocation, dbMapId) & vbCrLf
    Else
        ToSend$ = BRIGHTLIGHTBLUE & modgetdata.GetRoomTitle(iLocation, dbMapId) & vbCrLf
    End If
    If Not clsSC.FastStringComp(TempPeeps, "") And Not clsSC.FastStringComp(Mons$, "") Then
        TempPeeps = TempPeeps & Mons$
    ElseIf clsSC.FastStringComp(TempPeeps, "") And Not clsSC.FastStringComp(Mons$, "") Then
        TempPeeps = Mons$
    ElseIf Not clsSC.FastStringComp(TempPeeps, "") And clsSC.FastStringComp(Mons$, "") Then
        TempPeeps = Left$(TempPeeps, Len(TempPeeps) - 2)
    ElseIf clsSC.FastStringComp(TempPeeps, "") And clsSC.FastStringComp(Mons$, "") Then
        TempPeeps = ""
    End If
    If Not clsSC.FastStringComp(TempPeeps, "") Then ToSend$ = ToSend$ & MAGNETA & "Also here: " & TempPeeps
    If Not clsSC.FastStringComp(TempItemHere, "") Then
        ToSend$ = ToSend$ & vbCrLf & YELLOW & "You notice " _
            & GREEN & TempItemHere & YELLOW & " here." & vbCrLf _
            & BRIGHTMAGNETA & "Visible Exits: " & GREEN & modgetdata.GetRoomExits(iLocation, dbMapId) _
            & vbCrLf & WHITE
    Else
        ToSend$ = ToSend$ & vbCrLf _
            & BRIGHTMAGNETA & "Visible Exits: " & GREEN & modgetdata.GetRoomExits(iLocation, dbMapId) _
            & vbCrLf & WHITE
    End If
    If Not clsSC.FastStringComp(TempPeeps, "") Then TempPeeps = Left$(TempPeeps, Len(TempPeeps) - 1) & "."
    Select Case .iVision + modgetdata.GetRoomLight(iLocation, dbMapId)
        Case -3 To -1
            ToSend$ = ToSend$ & WHITE & "This room is barely visible." & WHITE & vbCrLf
        Case 0 To 2
            ToSend$ = ToSend$ & WHITE & "This room has little light in it." & WHITE & vbCrLf
    End Select
    GetRoomDescription = ToSend$
End With
End Function

Public Function GetClassPointLevel(dbIndex As Long, Optional ForgetPrefix As Boolean = False) As String
Dim dbClassID As Long
If ForgetPrefix Then
    With dbPlayers(dbIndex)
        dbClassID = GetClassID(.sClass)
        Select Case .dClassPoints
            Case Is > dbClass(dbClassID).dGuru
                GetClassPointLevel = BLUE & "Guru"
            Case Is > dbClass(dbClassID).dMasterMax
                GetClassPointLevel = GREEN & "Master"
            Case Is > dbClass(dbClassID).dIntermediateMax
                GetClassPointLevel = LIGHTBLUE & "Intermediate"
            Case Is > dbClass(dbClassID).dBeginnerMax
                GetClassPointLevel = BRIGHTWHITE & "Beginner"
            Case Else
                GetClassPointLevel = WHITE & "Apprentice"
        End Select
    End With
Else
    With dbPlayers(dbIndex)
        dbClassID = GetClassID(.sClass)
        Select Case .dClassPoints
            Case Is > dbClass(dbClassID).dGuru
                GetClassPointLevel = BLUE & " Guru"
            Case Is > dbClass(dbClassID).dMasterMax
                GetClassPointLevel = GREEN & " Master"
            Case Is > dbClass(dbClassID).dIntermediateMax
                GetClassPointLevel = BRIGHTWHITE & "n " & LIGHTBLUE & "Intermediate"
            Case Is > dbClass(dbClassID).dBeginnerMax
                GetClassPointLevel = BRIGHTWHITE & " Beginner"
            Case Else
                GetClassPointLevel = BRIGHTWHITE & "n " & WHITE & "Apprentice"
        End Select
    End With
End If
End Function

Public Function GetStatFromString(dbIndex As Long, s As String) As Double
With dbPlayers(dbIndex)
    Select Case s
        Case "lives"
            GetStatFromString = .iLives
        Case "str"
            GetStatFromString = .iStr
        Case "agil"
            GetStatFromString = .iAgil
        Case "int"
            GetStatFromString = .iInt
        Case "dex"
            GetStatFromString = .iDex
        Case "cha"
            GetStatFromString = .iCha
        Case "level"
            GetStatFromString = .iLevel
        Case "exp"
            GetStatFromString = .dEXP
        Case "expneeded"
            GetStatFromString = .dEXPNeeded
        Case "totalexp"
            GetStatFromString = .dTotalEXP
        Case "weapons"
            GetStatFromString = .iWeapons
        Case "armortype"
            GetStatFromString = .iArmorType
        Case "spelllevel"
            GetStatFromString = .iSpellLevel
        Case "spelltype"
            GetStatFromString = .iSpellType
        Case "gold"
            GetStatFromString = .dGold
        Case "hp"
            GetStatFromString = .lHP
        Case "maxhp"
            GetStatFromString = .lMaxHP
        Case "mana"
            GetStatFromString = .lMana
        Case "maxmana"
            GetStatFromString = .lMaxMana
        Case "ac"
            GetStatFromString = .iAC
        Case "acc"
            GetStatFromString = .iAcc
        Case "crits"
            GetStatFromString = .iCrits
        Case "dodge"
            GetStatFromString = .iDodge
        Case "maxdamage"
            GetStatFromString = .iMaxDamage
        Case "bank"
            GetStatFromString = .dBank
        Case "vision"
            GetStatFromString = .iVision
        Case "maxitems"
            GetStatFromString = modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items])
        Case "guildleader"
            GetStatFromString = .iGuildLeader
        Case "evil"
            GetStatFromString = .iEvil
        Case "element.fire"
            GetStatFromString = modResist.GetResistValue(dbIndex, Fire)
        Case "element.ice"
            GetStatFromString = modResist.GetResistValue(dbIndex, Ice)
        Case "element.water"
            GetStatFromString = modResist.GetResistValue(dbIndex, Water)
        Case "element.lightning"
            GetStatFromString = modResist.GetResistValue(dbIndex, Lightning)
        Case "element.earth"
            GetStatFromString = modResist.GetResistValue(dbIndex, Earth)
        Case "element.poison"
            GetStatFromString = modResist.GetResistValue(dbIndex, Poison)
        Case "element.wind"
            GetStatFromString = modResist.GetResistValue(dbIndex, Wind)
        Case "element.holy"
            GetStatFromString = modResist.GetResistValue(dbIndex, Holy)
        Case "element.unholy"
            GetStatFromString = modResist.GetResistValue(dbIndex, Unholy)
        Case "classpts"
            GetStatFromString = dbPlayers(dbIndex).dClassPoints
        Case "stamina"
            GetStatFromString = dbPlayers(dbIndex).dStamina
        Case "hunger"
            GetStatFromString = dbPlayers(dbIndex).dHunger
        Case "thieving"
            GetStatFromString = modMiscFlag.GetStatsPlusTotal(dbIndex, Thieving)
        Case "perception"
            GetStatFromString = modMiscFlag.GetStatsPlusTotal(dbIndex, Perception)
        Case "spellcasting"
            GetStatFromString = modMiscFlag.GetStatsPlusTotal(dbIndex, [Spell Casting])
        Case "animalrelations"
            GetStatFromString = modMiscFlag.GetStatsPlusTotal(dbIndex, [Animal Relations])
        Case "magicres"
            GetStatFromString = modMiscFlag.GetStatsPlusTotal(dbIndex, [Magic Resistance])
        Case "stealth"
            GetStatFromString = modMiscFlag.GetStatsPlusTotal(dbIndex, Steath)
        Case "pallete"
            GetStatFromString = modMiscFlag.GetStatsPlus(dbIndex, [Pallete Number])
        Case "isasysop"
            GetStatFromString = modMiscFlag.GetStatsPlus(dbIndex, [Is A Sysop])
            
    End Select
End With
End Function

Public Function GetItemNameWithS(dbItemID As Long) As String
With dbItems(dbItemID)
    If Right$(.sItemName, 1) = "s" Then GetItemNameWithS = .sItemName & " drop " Else GetItemNameWithS = .sItemName & " drops "
End With
End Function

Public Function GetItemsNameAddS(dbItemID As Long) As String
With dbItems(dbItemID)
    If Right$(.sItemName, 1) = "s" Then GetItemsNameAddS = .sItemName Else GetItemsNameAddS = .sItemName & "s"
End With
End Function

Public Function GetGibberish(s As String, lLevel As Long) As String
Dim i As Long
Dim n As String
Dim Arr() As String
Dim j As Long
If Len(s) < 2 Then GetGibberish = s: Exit Function
If lLevel = 1 Then
    For i = 1 To Len(s)
        GetGibberish = GetGibberish & Mid$(s, RndNumber(1, Len(s)), 1)
        If DE Then DoEvents
    Next
ElseIf lLevel = 2 Then
    For i = 1 To Len(s)
        n = Mid$(s, i, 1)
        Select Case n
            Case "a"
                n = "ahh"
            Case "e"
                n = "o"
            Case "i"
                n = "eye"
            Case "o"
                n = "mo"
            Case "u"
                n = "eww"
        End Select
        GetGibberish = GetGibberish & n
        If DE Then DoEvents
    Next
ElseIf lLevel = 3 Then
    For i = 1 To Len(s)
        n = Mid$(s, i, 1)
        n = n & String$(RndNumber(1, 4), n)
        GetGibberish = GetGibberish & n
        If DE Then DoEvents
    Next
ElseIf lLevel = 4 Then
    If InStr(1, s, " ") <> 0 Then
        SplitFast s, Arr, " "
        For i = LBound(Arr) To UBound(Arr)
            For j = 1 To Len(Arr(i))
                n = Mid$(Arr(i), j, 1)
                Select Case n
                    Case "a"
                        n = "uhuhuhuh"
                    Case "e"
                        n = "ehhhehhhhhhh"
                    Case "i"
                        n = "i i i i i i i i i"
                    Case "o"
                        n = "ohohohoh"
                    Case "u"
                        n = "uuuuuuu"
                End Select
                GetGibberish = GetGibberish & n
                If DE Then DoEvents
            Next
            GetGibberish = GetGibberish & " "
            If DE Then DoEvents
        Next
    Else
        For j = 1 To Len(s)
            n = Mid$(s, j, 1)
            Select Case n
                Case "a"
                    n = "uhuhuhuh"
                Case "e"
                    n = "ehhhehhhhhhh"
                Case "i"
                    n = "i i i i i i i i i"
                Case "o"
                    n = "ohohohoh"
                Case "u"
                    n = "uuuuuuu"
            End Select
            GetGibberish = GetGibberish & n
            If DE Then DoEvents
        Next
    End If
End If
End Function

Function GetPlayersPerceptionBase(dbIndex As Long, Optional CapIt As Boolean = True) As Long
Dim i As Long
With dbPlayers(dbIndex)
    i = (((((.iLevel * 2) + .iInt) / 1.8) + (.iDex / 2)) / 1.01) + .iAgil + (modMiscFlag.GetStatsPlus(dbIndex, [Spell Casting Base]) / 15)
End With
If CapIt Then
    If i > 97 Then i = 97
End If
GetPlayersPerceptionBase = i
End Function

Function GetPlayersMR(dbIndexA As Long, dbIndexV As Long) As Long
Dim i As Long
With dbPlayers(dbIndexA)
    i = modMiscFlag.GetStatsPlusTotal(dbIndexV, [Magic Resistance])
    i = i / modMiscFlag.GetStatsPlusTotal(dbIndexA, [Spell Casting])
End With
GetPlayersMR = i
End Function
