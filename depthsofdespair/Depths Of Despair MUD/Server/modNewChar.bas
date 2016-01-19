Attribute VB_Name = "modNewChar"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modNewChar
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'


'Sub ChooseClass(xx$, Index As Long)
'Dim Classes$, tArr() As String
'Dim intNum&, ToSend$
'For i = LBound(dbClass) To UBound(dbClass)
'    Classes$ = Classes$ & dbClass(i).sName & ";"
'    If DE Then DoEvents
'Next
'Classes$ = Left$(Classes$, Len(Classes$) - 1)
'SplitFast Classes$, tArr, ";"
'If Val(xx$) - 1 > UBound(tArr) Or Val(xx$) = 0 Or Not IsNumeric(xx$) Then
'    intNum& = 1
'    For i = LBound(dbClass) To UBound(dbClass)
'        ToSend$ = ToSend$ & intNum& & ": " & dbClass(i).sName & vbCrLf
'        intNum& = intNum& + 1
'        If DE Then DoEvents
'    Next
'    WrapAndSend Index, MAGNETA & "To which profression do you wish to follow? " & vbCrLf & ToSend$ & _
'        vbCrLf & "Choose a number that corresponds with your class (1-" & intNum& - 1 & "): " & WHITE
'    X(Index) = ""
'    Exit Sub
'End If
'End Sub

Sub ChooseRace(xx$, Index As Long)
Dim Races$, tArr() As String
Dim intNum&, ToSend$
Dim dbClassID As Long
For i = LBound(dbRaces) To UBound(dbRaces)
    Races$ = Races$ & dbRaces(i).sName & ";"
    If DE Then DoEvents
Next
Races$ = Left$(Races$, Len(Races$) - 1)
SplitFast Races$, tArr, ";"
If Val(xx$) - 1 > UBound(tArr) Or Val(xx$) = 0 Or Not IsNumeric(xx$) Then
    intNum& = 1
    For i = LBound(dbRaces) To UBound(dbRaces)
        ToSend$ = ToSend$ & intNum& & ": " & dbRaces(i).sName & vbCrLf
        intNum& = intNum& + 1
        If DE Then DoEvents
    Next
    WrapAndSend Index, MAGNETA & "What race will you live by?" & vbCrLf & ToSend$ & _
        vbCrLf & "Choose a number that corresponds with your race (1-" & intNum& - 1 & "): " & WHITE
    X(Index) = ""
    Exit Sub
End If
With dbPlayers(GetPlayerIndexNumber(Index))
    .sRace = tArr(Val(xx$) - 1)
    dbClassID = GetClassID("Apprentice")
    With dbPlayers(GetPlayerIndexNumber(Index))
        .sClass = "Apprentice"
        .iAcc = dbClass(dbClassID).iAcc
        .iCrits = dbClass(dbClassID).iCrits
        .iArmorType = tiArmorType
        .iWeapons = tiWeapons
        .iAC = dbClass(dbClassID).lACBonus
        .iMaxDamage = dbClass(dbClassID).lDamBonus
        .iDodge = dbClass(dbClassID).lDodgeBonus
        .iVision = dbClass(dbClassID).lVisionBonus
        WrapAndSend Index, MAGNETA & "Be patient while your character is brought to life." & WHITE & vbCrLf
        pPoint(Index) = 0
        X(Index) = ""
        MakeCharacter Index
    End With
    Exit Sub
End With
End Sub

Sub MakeCharacter(Index As Long)
Dim dbClassID As Long, dbIndex As Long, dbRaceID As Long
Dim Arr() As String
dbIndex = GetPlayerIndexNumber(Index)
dbRaceID = GetRaceID(dbPlayers(dbIndex).sRace)
dbClassID = GetClassID("Apprentice")
With dbPlayers(GetPlayerIndexNumber(Index))
    SplitFast dbRaces(dbRaceID).sStats, Arr, ":"
    .iStr = Arr(0)
    .iAgil = Arr(1)
    .iInt = Arr(2)
    .iCha = Arr(3)
    .iDex = Arr(4)
    .dEXPNeeded = dbRaces(dbRaceID).dEXP
    Erase Arr
    SplitFast dbRaces(dbRaceID).sHP, Arr, ":"
    .lMaxHP = RndNumber(CDbl(Arr(0)), CDbl(Arr(1)))
    .lHP = .lMaxHP
    IncreaseTrain dbIndex, 5, .lMaxHP
    .iLevel = 1
    modMiscFlag.RedoStatsPlus dbIndex
    modMiscFlag.SetStatsPlus dbIndex, [Max Items Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Max Items Bonus]) + dbClass(dbClassID).lMaxItemsBonus
    .iVision = .iVision + dbRaces(dbRaceID).iVision
    .lAge = CLng(RndNumber(CDbl(dbRaces(dbRaceID).lStartAgeMin), CDbl(dbRaces(dbRaceID).lStartAgeMax)))
    .sBirthDay = CStr((MonthOfYear + 1)) & "/" & CStr(udtMonths(MonthOfYear).CurDay) & "/" & CStr(CurYear - .lAge)
    If .dEXP <> 0 Then .dEXP = RoundFast(.dEXP * 0.1, 0)
    .iLives = 9
    IncreaseTrain dbIndex, 6, .lMaxMana
End With
pLogOn(Index) = False
pLogOnPW(Index) = False
WrapAndSend Index, MAGNETA & "Your character has been brought to life." & vbCrLf & "Welcome to the game" & vbCrLf & WHITE
X(Index) = ""
End Sub

Sub RollCharacter(X$, pw$, Index&)
Dim lID As Long

SaveMemoryToDatabase 0
SaveMemoryToDatabase 1
SaveMemoryToDatabase 2
SaveMemoryToDatabase 3
SaveMemoryToDatabase 4

Set MRS = db.OpenRecordset("SELECT * FROM Players")
lID = 1
With MRS
    .MoveFirst
    Do
        If lID = CLng(!PlayerID) Then
            lID = lID + 1
            .MoveFirst
        ElseIf Not .EOF Then
            .MoveNext
        End If
        If DE Then DoEvents
    Loop Until .EOF
    .AddNew
    !PlayerID = lID
    !PlayerName = X$
    !PlayerPW = pw$
    !OverrideDesc = "0"
    !Inv = "0"
    !Head = "0"
    !TrainStats = "0/0/0/0/0/0/0"
    !Body = "0"
    !Arms = "0"
    !Hands = "0"
    !Legs = "0"
    !Misc = "0000000000000000000000000000000"
    !Feet = "0"
    !Waist = "0"
    !Face = "0"
    !Rings = "0;0;0;0;0;0;"
    !Back = "0"
    !Ears = "0"
    !Neck = "0"
    !Shield = "0"
    !Index = Index&
    !SpellType = "0"
    !Spells = "0"
    !ClassChanges = "0"
    !Bank = "0"
    !Horse = "0"
    !QUEST1 = "0"
    !Appearance = "0"
    !Statline = "HP=;hp/;mhp,MA=;ma/;mma"
    !QUEST2 = "0"
    !QUEST3 = "0"
    !Echo = "0"
    !TotalEXP = "0"
    !QUEST4 = "0"
    !SpellLevel = "0"
    !Dodge = "0"
    !MaxDamage = "0"
    !SpellShorts = "0"
    !Stamina = "150"
    !Hunger = "100"
    !MaxMana = "0"
    !IsReadyToTrain = "0"
    !MANA = "0"
    !EXP = "0"
    !EXPneeded = "0"
    !Race = "None"
    !Class = "None"
    !Str = "0"
    !AGIL = "0"
    !Int = "0"
    !CHA = "0"
    !DEX = "0"
    !BackUpLoc = "1"
    !Location = "1"
    !HP = "1"
    !MaxHP = "1"
    !Crits = "0"
    !Acc = "0"
    !Weapon = "0"
    !AC = "0"
    !Level = "1"
    !Weapons = "0"
    !Gold = "0"
    !Stun = "0"
    !ArmorType = "0"
    !SeenAs = X$
    !Vision = "0"
    !Lives = "9"
    !BlessSpells = "0"
    !Guild = "0"
    !GuildLeader = "0"
    !Evil = "0"
    !StatsPlus = "0/0/0/0/0/0/0/0/0/0/0/0/0/0/0/0"
    !Flag1 = "0"
    !Flag2 = "0"
    !Flag3 = "0"
    !Flag4 = "0"
    !Paper = "0"
    !Index = Index
    !Birthday = "0"
    !Age = "0"
    !ClassPoints = "0"
    !Gender = "-1"
    !ClassBonusLevel = "0"
    !Resist = "0/0/0/0/0/0/0/0/0"
    !KillDurItems = "0"
    !FamFlags = "0/0/0/0/0/0/0/0/0/0/0/0"
    .Update
End With
ReloadPlayersOnly
End Sub

Sub ChooseGender(Index As Long)
With dbPlayers(GetPlayerIndexNumber(Index))
    Select Case Left$(LCaseFast(X(Index)), 1)
        Case "m"
            .iGender = 0
        Case "f"
            .iGender = 1
        Case "i"
            .iGender = -1
    End Select
End With
End Sub
