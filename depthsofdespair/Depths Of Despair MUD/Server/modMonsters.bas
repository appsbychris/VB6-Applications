Attribute VB_Name = "modMonsters"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modMonsters
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Sub AddMonsterRgn(MonsterName As String)
'Sub for when a regen monster dies, make its regen time to full
With dbMonsters(GetMonsterID(MonsterName))
    .lRegenTimeLeft = .lRegenTime
End With
End Sub

Sub ArenaMon()
'Sub to make sure the specific Arena rooms or Special regen rooms are full
Dim i As Long
For i = LBound(dbArenas) To UBound(dbArenas)
    With dbArenas(i)
        Select Case .iType
            Case 3:
                GenAMonster CLng(.lRoomID), True, .iMobGroup, , .ldbMapID
            Case 1, 2, 4, 5:
                GenAMonster CLng(.lRoomID), True, , CLng(.lSpecialMon), .ldbMapID
        End Select
    End With
    If DE Then DoEvents
Next
End Sub

Sub FillType()
'Sub to fill in the Type array of all the monsters that are in the rooms
Dim TempMons() As String, TempMonsLoc() As String 'temp arrays
Dim TempDbIds() As String
Dim LocTemp As String, TempString As String 'temp strings
Dim dbTempS As String
Dim i As Long, j As Long
For i = LBound(dbMap) To UBound(dbMap)
    With dbMap(i)
        If Not modSC.FastStringComp(.sMonsters, "0") Then
            TempString = TempString & .sMonsters
            For j = 1 To DCount(.sMonsters, ";")
                LocTemp = LocTemp & .lRoomID & ";"
                dbTempS = dbTempS & i & ";"
                If DE Then DoEvents
            Next
        End If
    End With
    If DE Then DoEvents
Next
If modSC.FastStringComp(TempString, "") Then Exit Sub
TempString = Left$(TempString, Len(TempString) - 1) 'trim off the last ";"
LocTemp = Left$(LocTemp, Len(LocTemp) - 1) 'trim off the last ";"
TempString = ReplaceFast(TempString, ":", "")
SplitFast TempString, TempMons, ";"
SplitFast LocTemp, TempMonsLoc, ";"
SplitFast dbTempS, TempDbIds, ";"
AmountMons = 0 'set to 0
For j = LBound(TempMons) To UBound(TempMons)
    For i = LBound(dbMonsters) To UBound(dbMonsters)
        'ReDim Preserve aMons(AmountMons) As Monster
        With dbMonsters(i)
            If .lID = CLng(TempMons(j)) Then
                aMons(AmountMons).mName = .sMonsterName
                aMons(AmountMons).mHP = .dHP
                aMons(AmountMons).mLoc = TempMonsLoc(j)
                dbMap(TempDbIds(AmountMons)).sAMonIds = dbMap(TempDbIds(AmountMons)).sAMonIds & AmountMons & ";"
                aMons(AmountMons).mdbMapID = TempDbIds(AmountMons)
                aMons(AmountMons).mMessage = .sMessage
                aMons(AmountMons).mAc = .iAC
                aMons(AmountMons).mEXP = .dEXP
                aMons(AmountMons).mMin = Val(Mid$(.sAttack, 1, InStr(1, .sAttack, ":")))  'its min damage
                aMons(AmountMons).mMax = Val(Mid$(.sAttack, InStr(1, .sAttack, ":") + 1, Len(.sAttack) - (InStr(1, .sAttack, ":") - 1))) 'its max damage
                aMons(AmountMons).mEnergy = .lEnergy
                aMons(AmountMons).mPEnergy = .lPEnergy
                If .lWeapon <> 0 Then aMons(AmountMons).mWeapon = dbItems(GetItemID(, .lWeapon))
                SetUpAmonSpells AmountMons, .sSpells
                aMons(AmountMons).mLevel = .lLevel
                aMons(AmountMons).mMoney = .dMoney
                aMons(AmountMons).mDeathText = .sDeathText
                aMons(AmountMons).mHostile = IIf(.iHostile = 1, True, False)
                aMons(AmountMons).mAttackable = IIf(.iAttackable = 0, True, False)
                aMons(AmountMons).mRoams = .iRoams
                aMons(AmountMons).mDontAttackIfItem = .iDontAttackIfItem
                aMons(AmountMons).mMaxHP = .dHP
                aMons(AmountMons).mAtDayMonster = .iAtDayMonster
                aMons(AmountMons).mAtNightMonster = .iAtNightMonster
                aMons(AmountMons).mScript = .sScript
                aMons(AmountMons).miID = .lID
                aMons(AmountMons).mEvil = .iEvil
                aMons(AmountMons).mDesc = .sDesc
                aMons(AmountMons).mdbMonID = i
                AmountMons = AmountMons + 1
                Exit For
            End If
        End With
        If DE Then DoEvents
    Next i
    If DE Then DoEvents
Next j
End Sub

Public Function GetTitle() As String
Select Case RndNumber(0, 7)
    Case 0
        GetTitle = ""
    Case 1
        GetTitle = "small "
    Case 2
        GetTitle = "large "
    Case 3
        GetTitle = "ugly "
    Case 4
        GetTitle = "big "
    Case 5
        GetTitle = "pretty "
    Case 6
        GetTitle = "tiny "
    Case 7
        GetTitle = ""
End Select
End Function

Public Sub SetUpAmonSpells(amonIndex As Long, sSpells As String)
Dim Arr() As String
Dim lSpellCount As Long
With aMons(amonIndex)
    If sSpells <> "0" Then
        SplitFast sSpells, Arr, ";"
        For i = LBound(Arr) To UBound(Arr)
            If lSpellCount > 4 Then Exit For
            If Not modSC.FastStringComp(Arr(i), "") Then
                With .mSpells(lSpellCount)
                    .lCurrentCast = 0
                    .lSpellID = modItemManip.GetItemIDFromUnFormattedString(Arr(i))
                    .lEnergy = modItemManip.GetItemDurFromUnFormattedString(Arr(i))
                    .lMaxCast = modItemManip.GetItemUsesFromUnFormattedString(Arr(i))
                    .ldbSpellID = GetSpellID(, .lSpellID)
                    .lCastPerRound = Val(modItemManip.GetItemEnchantsFromUnFormattedString(Arr(i)))
                End With
                lSpellCount = lSpellCount + 1
            End If
            If DE Then DoEvents
        Next
    End If
End With
End Sub

Sub GenAMonster(Location As Long, Optional IsArena As Boolean = False, _
    Optional CertainMob As Long = 0, Optional CertainMonster As Long = 0, Optional dbLoc As Long = 0)
'Sub to generate a monster in a room
Dim RndNum As Long, MakeIt As Boolean, SkipCount As Boolean
Dim MG$ 'values
Dim i As Long
Dim j As Long
Dim dbM As Long
Dim lAmonID As Long
On Error GoTo GenAMonster_Error
If AmountMons < MaxMonsters Then 'make sure there aren't too many monsters
    If RndNumber(1, 25) > 12 Or IsArena = True Then 'randmon chance, or its an arnea monster
        If CertainMob <> 0 Then 'if they don't want a certain mob group...
            For i = LBound(dbMonsters) To UBound(dbMonsters)
                With dbMonsters(i)
                    If .lMobGroup = CertainMob And RndNumber(1, 25) > 15 Then
                        RndNum = i
                        MakeIt = True
                        Exit For
                    End If
                End With
                If DE Then DoEvents
            Next
        End If
        If CertainMonster <> 0 Then 'check if they want a random monster
            RndNum = GetMonsterID(, CertainMonster)
            MakeIt = True 'flag
        End If
        If MakeIt = False Then 'if the flag didn't change
            RndNum = UBound(dbMonsters)
            RndNum = RndNumber(1, CDbl(RndNum))
        End If
        MG$ = CStr(dbMonsters(RndNum).lMobGroup)
        If dbLoc = 0 Then
            dbM = GetMapIndex(Location)
        Else
            dbM = dbLoc
        End If
        With dbMap(dbM)
            If RndNumber(0, 100) > 97 Then DropOutDoorFood dbM
            If CLng(MG$) = .iMobGroup Or MakeIt = True Then
                If .sMonsters = "0" Then SkipCount = True
                If DCount(.sMonsters, ";") < .iMaxRegen Or SkipCount = True Then
                    If dbMonsters(RndNum).lRegenTimeLeft > 0 Then Exit Sub
                    AmountMons = AmountMons + 1
                    For j = LBound(aMons) To UBound(aMons)
                        With aMons(j)
                            If .mLoc = 0 Or .mLoc = -1 Then
                                lAmonID = j
                                Exit For
                            End If
                        End With
                        If DE Then DoEvents
                    Next
                    'ReDim Preserve aMons(AmountMons) As Monster
                    .sAMonIds = .sAMonIds & lAmonID & ";"
                    aMons(lAmonID).mName = dbMonsters(RndNum).sMonsterName
                    aMons(lAmonID).mHP = dbMonsters(RndNum).dHP
                    aMons(lAmonID).mLoc = Location
                    aMons(lAmonID).mdbMapID = dbM
                    aMons(lAmonID).mMessage = dbMonsters(RndNum).sMessage
                    aMons(lAmonID).mAc = dbMonsters(RndNum).iAC
                    aMons(lAmonID).mEXP = dbMonsters(RndNum).dEXP
                    aMons(lAmonID).mMin = Val(Mid$(dbMonsters(RndNum).sAttack, 1, InStr(1, dbMonsters(RndNum).sAttack, ":")))  'its min damage
                    aMons(lAmonID).mMax = Val(Mid$(dbMonsters(RndNum).sAttack, InStr(1, dbMonsters(RndNum).sAttack, ":") + 1, Len(dbMonsters(RndNum).sAttack) - (InStr(1, dbMonsters(RndNum).sAttack, ":") - 1)))  'its max damage
                    If dbMonsters(RndNum).lWeapon <> 0 Then aMons(lAmonID).mWeapon = dbItems(GetItemID(, dbMonsters(RndNum).lWeapon))
                    aMons(lAmonID).mEnergy = dbMonsters(RndNum).lEnergy
                    aMons(lAmonID).mPEnergy = dbMonsters(RndNum).lPEnergy
                    SetUpAmonSpells lAmonID, dbMonsters(RndNum).sSpells
                    aMons(lAmonID).mLevel = dbMonsters(RndNum).lLevel
                    aMons(lAmonID).mMoney = dbMonsters(RndNum).dMoney
                    aMons(lAmonID).mDeathText = dbMonsters(RndNum).sDeathText
                    aMons(lAmonID).mDesc = dbMonsters(RndNum).sDesc
                    'aMons(AmountMons).mTitle = GetTitle
                    aMons(lAmonID).mHostile = IIf(dbMonsters(RndNum).iHostile = 1, True, False)
                    aMons(lAmonID).mAttackable = IIf(dbMonsters(RndNum).iAttackable = 0, True, False)
                    aMons(lAmonID).mRoams = dbMonsters(RndNum).iRoams
                    aMons(lAmonID).mDontAttackIfItem = dbMonsters(RndNum).iDontAttackIfItem
                    aMons(lAmonID).mMaxHP = dbMonsters(RndNum).dHP
                    aMons(lAmonID).mAtDayMonster = dbMonsters(RndNum).iAtDayMonster
                    aMons(lAmonID).mAtNightMonster = dbMonsters(RndNum).iAtNightMonster
                    aMons(lAmonID).mScript = dbMonsters(RndNum).sScript
                    aMons(lAmonID).miID = dbMonsters(RndNum).lID
                    aMons(lAmonID).mEvil = dbMonsters(RndNum).iEvil
                    aMons(lAmonID).mdbMonID = RndNum
                    If modSC.FastStringComp(.sMonsters, "0") Then .sMonsters = ""
                    .sMonsters = .sMonsters & ":" & dbMonsters(RndNum).lID & ";"
                    'send the message to everyone in the room
                    SendToAllInRoom 0, MAGNETA & "A " & aMons(lAmonID).mName & " wanders into the room." & vbCrLf & WHITE, .lRoomID
                End If
            End If
        End With
    End If
End If
On Error GoTo 0
Exit Sub
GenAMonster_Error:
End Sub

Public Sub AdjustMonList(amonIndex As Long)
Dim i As Long
Dim j As Long
Dim bD As Long
Dim lP As Long
With aMons(amonIndex)
    For i = 0 To 9
        If .mList(i) = 0 Then
            If i - 1 = lP Then bD = bD + 1
            lP = i
            For j = i To 8
                .mList(j) = .mList(j + 1)
                If DE Then DoEvents
            Next
        End If
        If bD >= 3 Then Exit For
        If DE Then DoEvents
    Next
End With
End Sub

Public Sub InsertInMonList(amonIndex As Long, PlayerID As Long, Optional IndexNum = -1)
Dim i As Long
If amonIndex > UBound(aMons) Then Exit Sub
With aMons(amonIndex)
    For i = 0 To 9
        If .mList(i) = PlayerID Then .mList(i) = 0
    Next
    If IndexNum <> -1 Then
        For i = IndexNum To 8
            .mList(i + 1) = .mList(i)
            If i = IndexNum Then .mList(i) = PlayerID
            If DE Then DoEvents
        Next
    Else
        For i = 0 To 8
            If .mList(i) = 0 Then .mList(i) = PlayerID: Exit For
            If i = 8 Then .mList(9) = PlayerID
            If DE Then DoEvents
        Next
    End If
    AdjustMonList amonIndex
End With
End Sub

Public Sub RemoveFromMonList(amonIndex As Long, PlayerID As Long)
Dim i As Long
With aMons(amonIndex)
    For i = 0 To 9
        If .mList(i) = PlayerID Then .mList(i) = 0
    Next
    AdjustMonList amonIndex
End With
End Sub

Public Sub PrintMonList(amonIndex As Long, PlayerID As Long)
Dim i As Long
With aMons(amonIndex)
        Debug.Print "--------------------------------------"
        Debug.Print .mName & "'s list (" & amonIndex & ")"
    For i = 0 To 9
        Debug.Print "(" & i & ") - " & .mList(i)
    Next
        Debug.Print "--------------------------------------"
End With
End Sub

