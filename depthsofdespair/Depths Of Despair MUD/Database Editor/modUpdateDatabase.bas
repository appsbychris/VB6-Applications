Attribute VB_Name = "modUpdateDatabase"
Public sIndex As String
Public Enum SaveDB
    Players = 0
    Map = 1
    A_Monster = 2
    Item = 3
    Shop = 4
    Class = 5
    Race = 6
    Familiars = 7
    Spells = 8
    Emotions = 9
    Events = 10
End Enum

Public Function GetPlayerIndexNumber(Optional lID As Long = -1, Optional sPlayerName As String = "") As Long
Dim i As Long
If Index <> -1 Then
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        If dbPlayers(i).lPlayerID = lID Then
            GetPlayerIndexNumber = i
            Exit For
        End If
        If DE Then DoEvents
    Next
ElseIf sPlayerName <> "" Then
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        If LCase$(dbPlayers(i).sPlayerName) = LCase$(sPlayerName) Then
            GetPlayerIndexNumber = i
            Exit For
        End If
        If DE Then DoEvents
    Next
End If
End Function

Public Function GetMonsterID(Optional sMonsterName As String = "", Optional MonID As Long = -1) As Long
    Dim i As Long
    If sMonsterName <> "" Then
        For i = LBound(dbMonsters) To UBound(dbMonsters)
            If LCase$(dbMonsters(i).sMonsterName) = LCase$(sMonsterName) Then
                GetMonsterID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    ElseIf MonID <> -1 Then
        For i = LBound(dbMonsters) To UBound(dbMonsters)
            If dbMonsters(i).lID = CLng(MonID) Then
                GetMonsterID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
End Function

Public Function GetEmotionID(Optional sSyntax As String = "", Optional EmoteID As Long = -1) As Long
    Dim i As Long
    If sSyntax <> "" Then
        For i = LBound(dbEmotions) To UBound(dbEmotions)
            If LCase$(dbEmotions(i).sSyntax) = LCase$(sSyntax) Then
                GetEmotionID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    ElseIf EmoteID <> -1 Then
        For i = LBound(dbEmotions) To UBound(dbEmotions)
            If dbEmotions(i).iID = CLng(EmoteID) Then
                GetEmotionID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
End Function

Public Function GetSpellID(Optional sSpellName As String = "", Optional SpellID As Long = -1) As Long
    Dim i As Long
    If sSpellName <> "" Then
        For i = LBound(dbSpells) To UBound(dbSpells)
            If LCase$(dbSpells(i).sSpellName) = LCase$(sSpellName) Then
                GetSpellID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    ElseIf SpellID <> -1 Then
        For i = LBound(dbSpells) To UBound(dbSpells)
            If dbSpells(i).lID = CLng(SpellID) Then
                GetSpellID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
End Function

Public Function GetLetterID(Optional sTitle As String = "", Optional lID As Long = -1) As Long
Dim i As Long
If lID = -1 Then
    For i = LBound(dbLetters) To UBound(dbLetters)
        With dbLetters(i)
            If LCase$(.sTitle) = sTitle Then
                GetLetterID = i
                Exit Function
            End If
        End With
        If DE Then DoEvents
    Next
Else
    For i = LBound(dbLetters) To UBound(dbLetters)
        With dbLetters(i)
            If .lID = lID Then
                GetLetterID = i
                Exit Function
            End If
        End With
        If DE Then DoEvents
    Next
End If
End Function

Public Function GetShopID(ShopID As Long) As Long
Dim i As Long
For i = LBound(dbShops) To UBound(dbShops)
    With dbShops(i)
        If .iID = ShopID Then GetShopID = i: Exit Function
    End With
Next
End Function

Public Function GetClassID(Optional sClassName As String = "", Optional ClassID As Long = -1) As Long
    Dim i As Long
    If sClassName <> "" Then
        For i = LBound(dbClass) To UBound(dbClass)
            If LCase$(dbClass(i).sName) = LCase$(sClassName) Then
                GetClassID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    ElseIf ClassID <> -1 Then
        For i = LBound(dbClass) To UBound(dbClass)
            If dbClass(i).iID = ClassID Then
                GetClassID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
End Function

Public Function GetRaceID(Optional sRaceName As String = "", Optional RaceID As Long = -1) As Long
    Dim i As Long
    If sRaceName <> "" Then
        For i = LBound(dbRaces) To UBound(dbRaces)
            If LCase$(dbRaces(i).sName) = LCase$(sRaceName) Then
                GetRaceID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    ElseIf RaceID <> -1 Then
        For i = LBound(dbRaces) To UBound(dbRaces)
            If dbRaces(i).iID = RaceID Then
                GetRaceID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
End Function

Public Function GetFamID(Optional iID As Long = -1, Optional sFamName As String = "") As Long
    Dim i As Long
    If iID <> -1 Then
        For i = LBound(dbFamiliars) To UBound(dbFamiliars)
            If dbFamiliars(i).iID = iID Then
                GetFamID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    ElseIf sFamName <> "" Then
        For i = LBound(dbFamiliars) To UBound(dbFamiliars)
            If LCase$(dbFamiliars(i).sFamName) = LCase$(sFamName) Then
                GetFamID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
End Function

Public Function GetItemID(Optional sItemName As String = "", Optional lItemID As Long = 0) As Long
    Dim i As Long
    If sItemName <> "" Then
        For i = LBound(dbItems) To UBound(dbItems)
            If LCase$(dbItems(i).sItemName) = LCase$(sItemName) Then
                GetItemID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    ElseIf lItemID <> 0 Then
        For i = LBound(dbItems) To UBound(dbItems)
            If lItemID = dbItems(i).iID Then
                GetItemID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
End Function

Public Function GetMapIndex(RoomID As Long) As Long
    Dim i As Long
    For i = LBound(dbMap) To UBound(dbMap)
        If RoomID = dbMap(i).lRoomID Then
            GetMapIndex = i
            Exit For
        End If
        If DE Then DoEvents
    Next
End Function

Public Function GetShopIndex(ShopID As Long) As Long
    Dim i As Long
    For i = LBound(dbShops) To UBound(dbShops)
        If ShopID = dbShops(i).iID Then
            GetShopIndex = i
            Exit For
        End If
        If DE Then DoEvents
    Next
End Function
'================================================================================

Public Sub SaveMemoryToDatabase(iStep As SaveDB)
'On Error GoTo SaveMemoryToDatabase_Error

Dim i As Long
Dim b As Boolean
Dim j As Long
Dim s As String
Select Case iStep
    Case 0
        Set MRS = DB.OpenRecordset("SELECT * FROM Players")
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            With MRS
                .MoveFirst
                Do
                    If dbPlayers(i).lPlayerID = CLng(!PlayerId) Then
                        .Edit
                        !Bank = dbPlayers(i).dBank
                        !Exp = dbPlayers(i).dEXP
                        !EXPNeeded = dbPlayers(i).dEXPNeeded
                        !Gold = dbPlayers(i).dGold
                        !AC = dbPlayers(i).iAC
                        !Acc = dbPlayers(i).iAcc
                        !AGIL = dbPlayers(i).iAgil
                        !CHA = dbPlayers(i).iCha
                        !Crits = dbPlayers(i).iCrits
                        !DEX = dbPlayers(i).iDex
                        !Dodge = dbPlayers(i).iDodge
                        !Echo = dbPlayers(i).iEcho
'                        !FamID = dbPlayers(i).iFamID
                        !Horse = dbPlayers(i).iHorse
                        !Int = dbPlayers(i).iInt
                        !MaxDamage = dbPlayers(i).iMaxDamage
                        !Level = dbPlayers(i).iLevel
                        !Str = dbPlayers(i).iStr
                        !Stun = dbPlayers(i).iStun
                        !BackUpLoc = dbPlayers(i).lBackUpLoc
                        !HP = dbPlayers(i).lHP
                        !Vision = dbPlayers(i).iVision
                        !Location = dbPlayers(i).lLocation
                        !Mana = dbPlayers(i).lMana
                        !MaxHP = dbPlayers(i).lMaxHP
                        !MaxMana = dbPlayers(i).lMaxMana
                        !Arms = dbPlayers(i).sArms
                        !Body = dbPlayers(i).sBody
'                        !FamName = dbPlayers(i).sFamName
                        !Feet = dbPlayers(i).sFeet
                        !Hands = dbPlayers(i).sHands
                        !Head = dbPlayers(i).sHead
                        !Inv = dbPlayers(i).sInventory
                        !Legs = dbPlayers(i).sLegs
                        !QUEST1 = dbPlayers(i).sQuest1
                        !QUEST2 = dbPlayers(i).sQuest2
                        !QUEST3 = dbPlayers(i).sQuest3
                        !QUEST4 = dbPlayers(i).sQuest4
                        !Index = dbPlayers(i).iIndex
                        !TrainStats = dbPlayers(i).sTrainStats
                        !Waist = dbPlayers(i).sWaist
                        !Weapon = dbPlayers(i).sWeapon
                        !Spells = dbPlayers(i).sSpells
                        !SpellShorts = dbPlayers(i).sSpellShorts
                        !Race = dbPlayers(i).sRace
                        !Class = dbPlayers(i).sClass
                        !Weapons = dbPlayers(i).iWeapons
                        !ArmorType = dbPlayers(i).iArmorType
                        !SpellLevel = dbPlayers(i).iSpellLevel
                        !SpellType = dbPlayers(i).iSpellType
                        !BlessSpells = dbPlayers(i).sBlessSpells
                        '!MaxItems = dbPlayers(i).iMaxItems
                        !Guild = dbPlayers(i).sGuild
                        !GuildLeader = dbPlayers(i).iGuildLeader
                        !Evil = dbPlayers(i).iEvil
                        !Face = dbPlayers(i).sFace
                        !Ears = dbPlayers(i).sEars
                        !Neck = dbPlayers(i).sNeck
                        !Back = dbPlayers(i).sBack
                        !Shield = dbPlayers(i).sShield
                        !IsReadyToTrain = dbPlayers(i).iIsReadyToTrain
                        '!sC = dbPlayers(i).iSC
                        !Flag1 = dbPlayers(i).iFlag1
                        !Flag2 = dbPlayers(i).iFlag2
                        !Flag3 = dbPlayers(i).iFlag3
                        !Flag4 = dbPlayers(i).iFlag4
                        !Lives = dbPlayers(i).iLives
                        '!Letters = dbPlayers(i).sLetters
                        !Paper = dbPlayers(i).lPaper
                        !Age = dbPlayers(i).lAge
                        !ClassPoints = dbPlayers(i).dClassPoints
                        !ClassBonusLevel = dbPlayers(i).iClassBonusLevel
                        !ClassChanges = dbPlayers(i).lClassChanges
                        !TotalEXP = dbPlayers(i).dTotalEXP
                        !Stamina = dbPlayers(i).dStamina
                        !Hunger = dbPlayers(i).dHunger
'                        !FamEXP = dbPlayers(i).dFamEXP
'                        !FamCurrentHP = dbPlayers(i).lFamCurrentHP
'                        !FamMaxHP = dbPlayers(i).lFamMaxHP
                        !KillDurItems = dbPlayers(i).sKillDurItems
                        !Gender = dbPlayers(i).iGender
                        !Statline = dbPlayers(i).sStatline
                        '!IsSysop = dbPlayers(i).iIsSysop
                        
                        '!Pal = dbPlayers(i).iPal
                        If !Birthday = "0" Then !Birthday = dbPlayers(i).sBirthDay
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
            End With
            If DE Then DoEvents
        Next i
        Set MRS = Nothing
    Case 1
        Set MRSMAP = DB.OpenRecordset("SELECT * FROM Map")
        For i = LBound(dbMap) To UBound(dbMap)
            With MRSMAP
                .MoveFirst
                b = False
                Do
                    If dbMap(i).lRoomID = CLng(!RoomID) Then
                        b = True
                        .Edit
                        '!Gold = dbMap(i).dGold
                        !Items = dbMap(i).sItems
                        !Monsters = dbMap(i).sMonsters
                        !Hidden = dbMap(i).sHidden
                        !North = dbMap(i).lNorth
                        !South = dbMap(i).lSouth
                        !East = dbMap(i).lEast
                        !West = dbMap(i).lWest
                        !NorthEast = dbMap(i).lNorthEast
                        !NorthWest = dbMap(i).lNorthWest
                        !SouthEast = dbMap(i).lSouthEast
                        !SouthWest = dbMap(i).lSouthWest
                        !Up = dbMap(i).lUp
                        !Down = dbMap(i).lDown
                        s = ":" & dbMap(i).lKN & ";"
                        s = s & ":" & dbMap(i).lKS & ";"
                        s = s & ":" & dbMap(i).lKE & ";"
                        s = s & ":" & dbMap(i).lKW & ";"
                        s = s & ":" & dbMap(i).lKNE & ";"
                        s = s & ":" & dbMap(i).lKNW & ";"
                        s = s & ":" & dbMap(i).lKSE & ";"
                        s = s & ":" & dbMap(i).lKSW & ";"
                        s = s & ":" & dbMap(i).lKU & ";"
                        s = s & ":" & dbMap(i).lKD & ";"
                        !Key = s
                        s = ":" & dbMap(i).lPN & ";"
                        s = s & ":" & dbMap(i).lPS & ";"
                        s = s & ":" & dbMap(i).lPE & ";"
                        s = s & ":" & dbMap(i).lPW & ";"
                        s = s & ":" & dbMap(i).lPNE & ";"
                        s = s & ":" & dbMap(i).lPNW & ";"
                        s = s & ":" & dbMap(i).lPSE & ";"
                        s = s & ":" & dbMap(i).lPSW & ";"
                        s = s & ":" & dbMap(i).lPU & ";"
                        s = s & ":" & dbMap(i).lPD & ";"
                        !Pick = s
                        s = ":" & dbMap(i).lBN & ";"
                        s = s & ":" & dbMap(i).lBS & ";"
                        s = s & ":" & dbMap(i).lBE & ";"
                        s = s & ":" & dbMap(i).lBW & ";"
                        s = s & ":" & dbMap(i).lBNE & ";"
                        s = s & ":" & dbMap(i).lBNW & ";"
                        s = s & ":" & dbMap(i).lBSE & ";"
                        s = s & ":" & dbMap(i).lBSW & ";"
                        s = s & ":" & dbMap(i).lBU & ";"
                        s = s & ":" & dbMap(i).lBD & ";"
                        !Bash = s
                        s = ":" & dbMap(i).lDN & ";"
                        s = s & ":" & dbMap(i).lDS & ";"
                        s = s & ":" & dbMap(i).lDE & ";"
                        s = s & ":" & dbMap(i).lDW & ";"
                        s = s & ":" & dbMap(i).lDNE & ";"
                        s = s & ":" & dbMap(i).lDNW & ";"
                        s = s & ":" & dbMap(i).lDSE & ";"
                        s = s & ":" & dbMap(i).lDSW & ";"
                        s = s & ":" & dbMap(i).lDU & ";"
                        s = s & ":" & dbMap(i).lDD & ";"
                        !Door = s
                        !Items = dbMap(i).sItems
                        !RoomTitle = dbMap(i).sRoomTitle
                        !RoomDesc = dbMap(i).sRoomDesc
                        !Monsters = dbMap(i).sMonsters
                        '!MaxRegen = dbMap(i).iMaxRegen
                        '!Type = dbMap(i).iType
                        '!ShopItems = dbMap(i).sShopItems
                        '!MobGroup = dbMap(i).iMobGroup
                        '!Gold = dbMap(i).dGold
                        !SpecialMon = dbMap(i).lSpecialMon
                        !SpecialItem = dbMap(i).lSpecialItem
                        '!Light = dbMap(i).lLight
                        !Hidden = dbMap(i).sHidden
                        !Scripting = dbMap(i).sScript
                        '!SafeRoom = dbMap(i).iSafeRoom
                        '!DeathRoom = dbMap(i).lDeathRoom
                        '!InDoor = dbMap(i).iInDoor
                        '!TrainClass = dbMap(i).iTrainClass
                        modMapFlags.UpdateMapFlags i
                        !flags = dbMap(i).sMapFlags
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
                If b = False Then
                    .AddNew
                    !RoomID = dbMap(i).lRoomID
                    '!Gold = dbMap(i).dGold
                    !Items = dbMap(i).sItems
                    !Monsters = dbMap(i).sMonsters
                    !Hidden = dbMap(i).sHidden
                    !North = dbMap(i).lNorth
                    !South = dbMap(i).lSouth
                    !East = dbMap(i).lEast
                    !West = dbMap(i).lWest
                    !NorthEast = dbMap(i).lNorthEast
                    !NorthWest = dbMap(i).lNorthWest
                    !SouthEast = dbMap(i).lSouthEast
                    !SouthWest = dbMap(i).lSouthWest
                    !Up = dbMap(i).lUp
                    !Down = dbMap(i).lDown
                    s = ":" & dbMap(i).lKN & ";"
                    s = s & ":" & dbMap(i).lKS & ";"
                    s = s & ":" & dbMap(i).lKE & ";"
                    s = s & ":" & dbMap(i).lKW & ";"
                    s = s & ":" & dbMap(i).lKNE & ";"
                    s = s & ":" & dbMap(i).lKNW & ";"
                    s = s & ":" & dbMap(i).lKSE & ";"
                    s = s & ":" & dbMap(i).lKSW & ";"
                    s = s & ":" & dbMap(i).lKU & ";"
                    s = s & ":" & dbMap(i).lKD & ";"
                    !Key = s
                    s = ":" & dbMap(i).lPN & ";"
                    s = s & ":" & dbMap(i).lPS & ";"
                    s = s & ":" & dbMap(i).lPE & ";"
                    s = s & ":" & dbMap(i).lPW & ";"
                    s = s & ":" & dbMap(i).lPNE & ";"
                    s = s & ":" & dbMap(i).lPNW & ";"
                    s = s & ":" & dbMap(i).lPSE & ";"
                    s = s & ":" & dbMap(i).lPSW & ";"
                    s = s & ":" & dbMap(i).lPU & ";"
                    s = s & ":" & dbMap(i).lPD & ";"
                    !Pick = s
                    s = ":" & dbMap(i).lBN & ";"
                    s = s & ":" & dbMap(i).lBS & ";"
                    s = s & ":" & dbMap(i).lBE & ";"
                    s = s & ":" & dbMap(i).lBW & ";"
                    s = s & ":" & dbMap(i).lBNE & ";"
                    s = s & ":" & dbMap(i).lBNW & ";"
                    s = s & ":" & dbMap(i).lBSE & ";"
                    s = s & ":" & dbMap(i).lBSW & ";"
                    s = s & ":" & dbMap(i).lBU & ";"
                    s = s & ":" & dbMap(i).lBD & ";"
                    !Bash = s
                    s = ":" & dbMap(i).lDN & ";"
                    s = s & ":" & dbMap(i).lDS & ";"
                    s = s & ":" & dbMap(i).lDE & ";"
                    s = s & ":" & dbMap(i).lDW & ";"
                    s = s & ":" & dbMap(i).lDNE & ";"
                    s = s & ":" & dbMap(i).lDNW & ";"
                    s = s & ":" & dbMap(i).lDSE & ";"
                    s = s & ":" & dbMap(i).lDSW & ";"
                    s = s & ":" & dbMap(i).lDU & ";"
                    s = s & ":" & dbMap(i).lDD & ";"
                    !Door = s
                    !Items = dbMap(i).sItems
                    !RoomTitle = dbMap(i).sRoomTitle
                    !RoomDesc = dbMap(i).sRoomDesc
                    !Monsters = dbMap(i).sMonsters
                    '!MaxRegen = dbMap(i).iMaxRegen
                    '!Type = dbMap(i).iType
                    '!ShopItems = dbMap(i).sShopItems
                    '!MobGroup = dbMap(i).iMobGroup
                    '!Gold = dbMap(i).dGold
                    !SpecialMon = dbMap(i).lSpecialMon
                    !SpecialItem = dbMap(i).lSpecialItem
                    '!Light = dbMap(i).lLight
                    !Hidden = dbMap(i).sHidden
                    !Scripting = dbMap(i).sScript
                    '!SafeRoom = dbMap(i).iSafeRoom
                    '!DeathRoom = dbMap(i).lDeathRoom
                    '!InDoor = dbMap(i).iInDoor
                    '!TrainClass = dbMap(i).iTrainClass
                    dbMap(i).sMapFlags = "0/0/2/1/0/-49/0/232/0/0/0;0;0;0;0;0;0;0;0;0;/0;"
                    modMapFlags.UpdateMapFlags i
                    !flags = dbMap(i).sMapFlags
                    .Update
                End If
             End With
            If DE Then DoEvents
        Next i
        Set MRSMAP = Nothing
    Case 2
        Set MRSMONSTER = DB.OpenRecordset("SELECT * FROM Monsters")
        For i = LBound(dbMonsters) To UBound(dbMonsters)
            With MRSMONSTER
                .MoveFirst
                b = False
                Do
                    If dbMonsters(i).lID = CLng(!ID) Then
                        b = True
                        .Edit
                        !Exp = dbMonsters(i).dEXP
                        !HP = dbMonsters(i).dHP
                        !Money = dbMonsters(i).dMoney
                        !AC = dbMonsters(i).iAC
                        !AtDayMonster = dbMonsters(i).iAtDayMonster
                        !AtNightMonster = dbMonsters(i).iAtNightMonster
                        !Attackable = dbMonsters(i).iAttackable
                        !DontAttackIfItem = dbMonsters(i).iDontAttackIfItem
                        !DropCorpse = dbMonsters(i).iDropCorpse
                        !Alignment = dbMonsters(i).iEvil
                        !Hostile = dbMonsters(i).iHostile
                        !Roams = dbMonsters(i).iRoams
                        !TameToFam = dbMonsters(i).iTameToFam
                        !Type = dbMonsters(i).iType
                        !DropItem = dbMonsters(i).sDropItem
                        !TotalEnergy = dbMonsters(i).lEnergy
                        !Level = dbMonsters(i).lLevel
                        !MobGroup = dbMonsters(i).lMobGroup
                        !PAttackEnergy = dbMonsters(i).lPEnergy
                        !RegenTime = dbMonsters(i).lRegenTime
                        !RegenTimeLeft = dbMonsters(i).lRegenTimeLeft
                        !Attack = dbMonsters(i).sAttack
                        !DeathText = dbMonsters(i).sDeathText
                        !Desc = dbMonsters(i).sDesc
                        !Message = dbMonsters(i).sMessage
                        !MonsterName = dbMonsters(i).sMonsterName
                        !OnDeathScript = dbMonsters(i).sScript
                        !CastSpells = dbMonsters(i).sSpells
                        !Weapon = dbMonsters(i).lWeapon
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
                If b = False Then
                    .MoveLast
                    j = CLng(!ID)
                    j = j + 1
                    Do
                        If CLng(!ID) = j Then
                            .MoveFirst
                            j = j + 1
                        ElseIf Not .EOF Then
                            .MoveNext
                        End If
                    Loop Until .EOF
                    .AddNew
                    !ID = j
                    !Exp = dbMonsters(i).dEXP
                    !HP = dbMonsters(i).dHP
                    !Money = dbMonsters(i).dMoney
                    !AC = dbMonsters(i).iAC
                    !AtDayMonster = dbMonsters(i).iAtDayMonster
                    !AtNightMonster = dbMonsters(i).iAtNightMonster
                    !Attackable = dbMonsters(i).iAttackable
                    !DontAttackIfItem = dbMonsters(i).iDontAttackIfItem
                    !DropCorpse = dbMonsters(i).iDropCorpse
                    !Alignment = dbMonsters(i).iEvil
                    !Hostile = dbMonsters(i).iHostile
                    !Roams = dbMonsters(i).iRoams
                    !TameToFam = dbMonsters(i).iTameToFam
                    !Type = dbMonsters(i).iType
                    !DropItem = dbMonsters(i).sDropItem
                    !TotalEnergy = dbMonsters(i).lEnergy
                    !Level = dbMonsters(i).lLevel
                    !RegenTimeLeft = "0"
                    !MobGroup = dbMonsters(i).lMobGroup
                    !PAttackEnergy = dbMonsters(i).lPEnergy
                    !RegenTime = dbMonsters(i).lRegenTime
                    !RegenTimeLeft = dbMonsters(i).lRegenTimeLeft
                    !Attack = dbMonsters(i).sAttack
                    !DeathText = dbMonsters(i).sDeathText
                    !Desc = dbMonsters(i).sDesc
                    !Message = dbMonsters(i).sMessage
                    !MonsterName = dbMonsters(i).sMonsterName
                    !OnDeathScript = dbMonsters(i).sScript
                    !CastSpells = dbMonsters(i).sSpells
                    !Weapon = dbMonsters(i).lWeapon
                    .Update
                End If
            End With
            If DE Then DoEvents
        Next i
        Set MRSMONSTER = Nothing
    Case 3
        Set MRSITEM = DB.OpenRecordset("SELECT * FROM Items")
        For i = LBound(dbItems) To UBound(dbItems)
            With MRSITEM
                .MoveFirst
                b = False
                Do
                    If dbItems(i).iID = CLng(!ID) Then
                        b = True
                        .Edit
                        !ItemName = dbItems(i).sItemName
                        !Damage = dbItems(i).sDamage
                        !Worn = dbItems(i).sWorn
                        !AC = dbItems(i).iAC
                        !Swings = dbItems(i).sSwings
                        !Message2 = dbItems(i).sMessage2
                        !MessageV = dbItems(i).sMessageV
                        !Speed = dbItems(i).iSpeed
                        !Type = dbItems(i).iType
                        !Desc = dbItems(i).sDesc
                        !Cost = dbItems(i).dCost
                        !Level = dbItems(i).lLevel
                        !ArmorType = dbItems(i).iArmorType
                        !Limit = dbItems(i).iLimit
                        !ClassRestriction = dbItems(i).sClassRestriction
                        !RaceRestriction = dbItems(i).sRaceRestriction
                        !Moveable = dbItems(i).iMoveable
                        !Magical = dbItems(i).iMagical
                        !Ledgendary = dbItems(i).iIsLedgenary
                        !Script = dbItems(i).sScript
                        !Durability = dbItems(i).lDurability
                        !Uses = dbItems(i).iUses
                        !ClassPoints = dbItems(i).dClassPoints
                        !OnEquipKillDur = dbItems(i).iOnEquipKillDur
                        !Projectile = dbItems(i).sProjectile
                        !OnLastUseDoFlags2 = dbItems(i).lOnLastUseDoFlags2
                        If i = 14 Then
                            i = i
                        End If
                        !flags = dbItems(i).sFlags
                        !Flags2 = dbItems(i).sFlags2
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
                If b = False Then
                    .MoveLast
                    j = CLng(!ID)
                    j = j + 1
                    Do
                        If CLng(!ID) = j Then
                            .MoveFirst
                            j = j + 1
                        ElseIf Not .EOF Then
                            .MoveNext
                        End If
                    Loop Until .EOF
                    .AddNew
                    !ID = j
                    !ItemName = dbItems(i).sItemName
                    !Damage = dbItems(i).sDamage
                    !Worn = dbItems(i).sWorn
                    !AC = dbItems(i).iAC
                    !Swings = dbItems(i).sSwings
                    !Message2 = dbItems(i).sMessage2
                    !MessageV = dbItems(i).sMessageV
                    !Speed = dbItems(i).iSpeed
                    !Type = dbItems(i).iType
                    !Desc = dbItems(i).sDesc
                    !Cost = dbItems(i).dCost
                    !Level = dbItems(i).lLevel
                    !ArmorType = dbItems(i).iArmorType
                    !Limit = dbItems(i).iLimit
                    !ClassRestriction = dbItems(i).sClassRestriction
                    !RaceRestriction = dbItems(i).sRaceRestriction
                    !Moveable = dbItems(i).iMoveable
                    !Magical = dbItems(i).iMagical
                    !Ledgendary = dbItems(i).iIsLedgenary
                    !Script = dbItems(i).sScript
                    !InGame = "0"
                    !Durability = dbItems(i).lDurability
                    !Uses = dbItems(i).iUses
                    !ClassPoints = dbItems(i).dClassPoints
                    !OnEquipKillDur = dbItems(i).iOnEquipKillDur
                    !Projectile = dbItems(i).sProjectile
                    !OnLastUseDoFlags2 = dbItems(i).lOnLastUseDoFlags2
                    !flags = dbItems(i).sFlags
                    !Flags2 = dbItems(i).sFlags2
                    .Update
                End If
            End With
            If DE Then DoEvents
        Next i
        Set MRSITEM = Nothing
    Case 4:
        Set MRSSHOPS = DB.OpenRecordset("SELECT * FROM SHOPS")
        For i = LBound(dbShops) To UBound(dbShops)
            With MRSSHOPS
                .MoveFirst
                b = False
                Do
                    If !ID = dbShops(i).iID Then
                        b = True
                        .Edit
                        !Markup = dbShops(i).iMarkUp
                        !ShopName = dbShops(i).sShopName
                        For j = 0 To 14
                            .Fields("Item" & CStr(j + 1)) = dbShops(i).iItems(j)
                            .Fields("Q" & CStr(j + 1)) = dbShops(i).iQ(j)
                        Next
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
                If b = False Then
                    .MoveLast
                    j = CLng(!ID)
                    j = j + 1
                    Do
                        If CLng(!ID) = j Then
                            .MoveFirst
                            j = j + 1
                        ElseIf Not .EOF Then
                            .MoveNext
                        End If
                    Loop Until .EOF
                    .AddNew
                    !ID = j
                    !Markup = dbShops(i).iMarkUp
                    !ShopName = dbShops(i).sShopName
                    For j = 0 To 14
                         .Fields("Item" & CStr(j + 1)) = dbShops(i).iItems(j)
                        .Fields("Q" & CStr(j + 1)) = dbShops(i).iQ(j)
                    Next
                    .Update
                End If
            End With
            If DE Then DoEvents
        Next
    Case 5:
        For i = LBound(dbClass) To UBound(dbClass)
            With MRSCLASS
                .MoveFirst
                b = False
                Do
                    If CLng(!ID) = dbClass(i).iID Then
                        b = True
                        .Edit
                        !Exp = dbClass(i).dEXP
                        !ArmorType = dbClass(i).iArmorType
                        !ID = dbClass(i).iID
                        !MaxMana = dbClass(i).iMaxMana
                        !MinMana = dbClass(i).iMinMana
                        !SpellLevel = dbClass(i).iSpellLevel
                        !SpellType = dbClass(i).iSpellType
                        !UseMagical = dbClass(i).iUseMagical
                        !Weapon = dbClass(i).iWeapon
                        !Name = dbClass(i).sName
                        !BeginnerMax = dbClass(i).dBeginnerMax
                        !Guru = dbClass(i).dGuru
                        !IntermediateMax = dbClass(i).dIntermediateMax
                        !MasterMax = dbClass(i).dMasterMax
                        !BBonus = dbClass(i).sBBonus
                        !GBonus = dbClass(i).sGBonus
                        !IBonus = dbClass(i).sIBonus
                        !MBonus = dbClass(i).sMBonus
                        !BaseBonus = dbClass(i).sBaseBonus
                        'If dbClass(i).sFlags = "" Then dbClass(i).sFlags = "0"
                        '!flags = dbClass(i).sFlags
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                Loop Until .EOF
                If b = False Then
                    .MoveLast
                    j = CLng(!ID)
                    j = j + 1
                    Do
                        If CLng(!ID) = j Then
                            .MoveFirst
                            j = j + 1
                        ElseIf Not .EOF Then
                            .MoveNext
                        End If
                    Loop Until .EOF
                    .AddNew
                    !ID = j
                    !Exp = dbClass(i).dEXP
                    !ArmorType = dbClass(i).iArmorType
                    !ID = dbClass(i).iID
                    !MaxMana = dbClass(i).iMaxMana
                    !MinMana = dbClass(i).iMinMana
                    !SpellLevel = dbClass(i).iSpellLevel
                    !SpellType = dbClass(i).iSpellType
                    !UseMagical = dbClass(i).iUseMagical
                    !Weapon = dbClass(i).iWeapon
                    !Name = dbClass(i).sName
                    !BeginnerMax = dbClass(i).dBeginnerMax
                    !Guru = dbClass(i).dGuru
                    !IntermediateMax = dbClass(i).dIntermediateMax
                    !MasterMax = dbClass(i).dMasterMax
                    !BBonus = dbClass(i).sBBonus
                    !GBonus = dbClass(i).sGBonus
                    !IBonus = dbClass(i).sIBonus
                    !MBonus = dbClass(i).sMBonus
                    !BaseBonus = dbClass(i).sBaseBonus
                    '!flags = dbClass(i).sFlags
                    .Update
                End If
            End With
        Next
    Case 6:
        For i = LBound(dbRaces) To UBound(dbRaces)
            With MRSRACE
                .MoveFirst
                b = False
                Do
                    If CLng(!ID) = dbRaces(i).iID Then
                        b = True
                        .Edit
                        !Name = dbRaces(i).sName
                        !Stats = dbRaces(i).sStats
                        !Exp = dbRaces(i).dEXP
                        !Vision = dbRaces(i).iVision
                        !MaxAge = dbRaces(i).lMaxAge
                        !StartAgeMin = dbRaces(i).lStartAgeMin
                        !StartAgeMax = dbRaces(i).lStartAgeMax
                        !HP = dbRaces(i).sHP
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                Loop Until .EOF
                If b = False Then
                    .MoveLast
                    j = CLng(!ID)
                    j = j + 1
                    Do
                        If CLng(!ID) = j Then
                            .MoveFirst
                            j = j + 1
                        ElseIf Not .EOF Then
                            .MoveNext
                        End If
                    Loop Until .EOF
                    .AddNew
                    !ID = j
                    !Name = dbRaces(i).sName
                    !Stats = dbRaces(i).sStats
                    !Exp = dbRaces(i).dEXP
                    !Vision = dbRaces(i).iVision
                    !MaxAge = dbRaces(i).lMaxAge
                    !StartAgeMin = dbRaces(i).lStartAgeMin
                    !StartAgeMax = dbRaces(i).lStartAgeMax
                    !HP = dbRaces(i).sHP
                    .Update
                End If
            End With
        Next
    Case 7:
        For i = LBound(dbFamiliars) To UBound(dbFamiliars)
            With MRSFAMILIARS
                .MoveFirst
                b = False
                Do
                    If CLng(!ID) = dbFamiliars(i).iID Then
                        b = True
                        .Edit
                        !Description = dbFamiliars(i).sDescription
                        !famName = dbFamiliars(i).sFamName
                        !StartHPMin = dbFamiliars(i).lStartHPMin
                        !StartHPMax = dbFamiliars(i).lStartHPMax
                        !EXPPerLevel = dbFamiliars(i).dEXPPerLevel
                        !MinDam = dbFamiliars(i).lMinDam
                        !MaxDam = dbFamiliars(i).lMaxDam
                        !LevelMod = dbFamiliars(i).lLevelMod
                        !AttackMessage = dbFamiliars(i).sAttackMessage
                        !Message2 = dbFamiliars(i).sMessage2
                        !MissMessage = dbFamiliars(i).sMissMessage
                        !MissMessage2 = dbFamiliars(i).sMissMessage2
                        !flags = dbFamiliars(i).sFlags
                        !LevelMax = dbFamiliars(i).lLevelMax
                        !Swings = dbFamiliars(i).lSwings
                        !Ridable = dbFamiliars(i).lRidable
                        !Speed = dbFamiliars(i).lSpeed
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                Loop Until .EOF
                If b = False Then
                    .MoveLast
                    j = CLng(!ID)
                    j = j + 1
                    Do
                        If CLng(!ID) = j Then
                            .MoveFirst
                            j = j + 1
                        ElseIf Not .EOF Then
                            .MoveNext
                        End If
                    Loop Until .EOF
                    .AddNew
                    !ID = j
                    !Description = dbFamiliars(i).sDescription
                    !famName = dbFamiliars(i).sFamName
                    !StartHPMin = dbFamiliars(i).lStartHPMin
                    !StartHPMax = dbFamiliars(i).lStartHPMax
                    !EXPPerLevel = dbFamiliars(i).dEXPPerLevel
                    !MinDam = dbFamiliars(i).lMinDam
                    !MaxDam = dbFamiliars(i).lMaxDam
                    !LevelMod = dbFamiliars(i).lLevelMod
                    !AttackMessage = dbFamiliars(i).sAttackMessage
                    !Message2 = dbFamiliars(i).sMessage2
                    !MissMessage = dbFamiliars(i).sMissMessage
                    !MissMessage2 = dbFamiliars(i).sMissMessage2
                    !flags = dbFamiliars(i).sFlags
                    !LevelMax = dbFamiliars(i).lLevelMax
                    !Swings = dbFamiliars(i).lSwings
                    !Ridable = dbFamiliars(i).lRidable
                    !Speed = dbFamiliars(i).lSpeed
                    .Update
                End If
            End With
        Next
    Case 8:
        For i = LBound(dbSpells) To UBound(dbSpells)
            With MRSSPELLS
                .MoveFirst
                b = False
                Do
                    If CLng(!ID) = dbSpells(i).lID Then
                        b = True
                        .Edit
                        !Cast = dbSpells(i).iCast
                        !Difficulty = dbSpells(i).iDifficulty
                        !Level = dbSpells(i).iLevel
                        !LevelMax = dbSpells(i).iLevelMax
                        !LevelModify = dbSpells(i).iLevelModify
                        !Type = dbSpells(i).iType
                        !Use = dbSpells(i).iUse
                        !Element = dbSpells(i).lElement
                        !Mana = dbSpells(i).lMana
                        !MaxDam = dbSpells(i).lMaxDam
                        !MinDam = dbSpells(i).lMinDam
                        !TimeOut = dbSpells(i).lTimeOut
                        !EndCastFlags = dbSpells(i).sEndCastFlags
                        !flags = dbSpells(i).sFlags
                        !Message = dbSpells(i).sMessage
                        !Message2 = dbSpells(i).sMessage2
                        !MessageV = dbSpells(i).sMessageV
                        !RunOutMessage = dbSpells(i).sRunOutMessage
                        !Short = dbSpells(i).sShort
                        !SpellName = dbSpells(i).sSpellName
                        !StatMessage = dbSpells(i).sStatMessage
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                Loop Until .EOF
                If b = False Then
                    .MoveLast
                    j = CLng(!ID)
                    j = j + 1
                    Do
                        If CLng(!ID) = j Then
                            .MoveFirst
                            j = j + 1
                        ElseIf Not .EOF Then
                            .MoveNext
                        End If
                    Loop Until .EOF
                    .AddNew
                    !ID = j
                    !Cast = dbSpells(i).iCast
                    !Difficulty = dbSpells(i).iDifficulty
                    !Level = dbSpells(i).iLevel
                    !LevelMax = dbSpells(i).iLevelMax
                    !LevelModify = dbSpells(i).iLevelModify
                    !Type = dbSpells(i).iType
                    !Use = dbSpells(i).iUse
                    !Element = dbSpells(i).lElement
                    !Mana = dbSpells(i).lMana
                    !MaxDam = dbSpells(i).lMaxDam
                    !MinDam = dbSpells(i).lMinDam
                    !TimeOut = dbSpells(i).lTimeOut
                    !EndCastFlags = dbSpells(i).sEndCastFlags
                    !flags = dbSpells(i).sFlags
                    !Message = dbSpells(i).sMessage
                    !Message2 = dbSpells(i).sMessage2
                    !MessageV = dbSpells(i).sMessageV
                    !RunOutMessage = dbSpells(i).sRunOutMessage
                    !Short = dbSpells(i).sShort
                    !SpellName = dbSpells(i).sSpellName
                    !StatMessage = dbSpells(i).sStatMessage
                    .Update
                End If
            End With
        Next
    Case 9:
        For i = LBound(dbEmotions) To UBound(dbEmotions)
            With MRSEMOTIONS
                .MoveFirst
                b = False
                Do
                    If CLng(!ID) = dbEmotions(i).iID Then
                        b = True
                        .Edit
                        !Syntax = dbEmotions(i).sSyntax
                        !PhraseYou = dbEmotions(i).sPhraseYou
                        !PhraseOthers = dbEmotions(i).sPhraseOthers
                        !PhraseOthers2 = dbEmotions(i).sPhraseOthers2
                        !PhraseToYou = dbEmotions(i).sPhraseToYou
                        !PhraseYouToOther = dbEmotions(i).sPhraseYouToOther
                        .Update
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                Loop Until .EOF
                If b = False Then
                    .MoveLast
                    j = CLng(!ID)
                    j = j + 1
                    Do
                        If CLng(!ID) = j Then
                            .MoveFirst
                            j = j + 1
                        ElseIf Not .EOF Then
                            .MoveNext
                        End If
                    Loop Until .EOF
                    .AddNew
                    !ID = j
                    !Syntax = dbEmotions(i).sSyntax
                    !PhraseYou = dbEmotions(i).sPhraseYou
                    !PhraseOthers = dbEmotions(i).sPhraseOthers
                    !PhraseOthers2 = dbEmotions(i).sPhraseOthers2
                    !PhraseToYou = dbEmotions(i).sPhraseToYou
                    !PhraseYouToOther = dbEmotions(i).sPhraseYouToOther
                    .Update
                End If
            End With
        Next
End Select
Load frmSplash
frmSplash.Top = Screen.Height - frmSplash.Height
frmSplash.Left = Screen.Width - frmSplash.Width
UpdateMRSSets
LoadDatabaseIntoMemory 'False
Unload frmSplash
On Error GoTo 0
Exit Sub
SaveMemoryToDatabase_Error:
'UpdateList "An error occured while saving information to the database."
'UpdateList "          " & Err.Number & " " & Err.Description
'UpdateList "          Occured on Staggered Step " & iStep
End Sub

Public Function LoadDatabaseIntoMemory()
Dim i As Long
Dim s As String
Dim t As Long
Dim Arr() As String
Dim a As Long
Dim j As Long
'On Error GoTo LoadDatabaseIntoMemory_Error
bUpdate = True
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Classes) [18%] ..."
i = 1
With MRSCLASS
    .MoveLast
    ReDim dbClass(1 To .RecordCount) As UDTClasses
    .MoveFirst
    Do
        'ReDim Preserve dbClass(1 To i) As UDTClasses
        dbClass(i).dEXP = !Exp
        dbClass(i).iArmorType = !ArmorType
        dbClass(i).iID = !ID
        dbClass(i).iMaxMana = !MaxMana
        dbClass(i).iMinMana = !MinMana
        dbClass(i).iSpellLevel = !SpellLevel
        dbClass(i).iSpellType = !SpellType
        dbClass(i).iUseMagical = !UseMagical
        dbClass(i).iWeapon = !Weapon
        'dbClass(i).sHP = !HP
        dbClass(i).sName = !Name
        dbClass(i).dBeginnerMax = !BeginnerMax
        dbClass(i).dGuru = !Guru
        dbClass(i).dIntermediateMax = !IntermediateMax
        dbClass(i).dMasterMax = !MasterMax
        dbClass(i).sBBonus = !BBonus
        dbClass(i).sGBonus = !GBonus
        dbClass(i).sIBonus = !IBonus
        dbClass(i).sMBonus = !MBonus
        dbClass(i).sBaseBonus = !BaseBonus
        s = !BaseBonus
'        cri         Adds/subtracts Critical hit chance.
'acc         adds/subtracts accuracy bonus
'h/l         Bonus to HP a level
'm/l                 Bonus to MA a level
'dam         Bonus to max damage
'sne         Can sneak (value of 0 or 1)
'cbs                     Can Backstab (value of 0 or 1)
'pts                 Train point bonus per level
'dog         Dodge bonus
'ACl         Armor Class Bonus
'Vis             Vision Bonus
'mit         Max Items bonus
'cdw         can duel wield
        If s <> "0" Then
            SplitFast LCaseFast(s), Arr, ";"
            For a = LBound(Arr) To UBound(Arr)
                If Arr(a) <> "" Then
                    t = CLng(Val(Mid$(s, 4)))
                    Select Case Left$(Arr(a), 3)
                        Case "cri"
                            dbClass(i).iCrits = t
                        Case "acc"
                            dbClass(i).iAcc = t
                        Case "h/l"
                            dbClass(i).lHPBonus = t
                        Case "m/l"
                            dbClass(i).lMABonus = t
                        Case "dam"
                            dbClass(i).lDamBonus = t
                        Case "sne"
                            dbClass(i).lCanSneak = t
                        Case "cbs"
                            dbClass(i).lCanBS = t
                        Case "pts"
                            dbClass(i).lCPBonus = t
                        Case "dog"
                            dbClass(i).lDodgeBonus = t
                        Case "acl"
                            dbClass(i).lACBonus = t
                        Case "vis"
                            dbClass(i).lVisionBonus = t
                        Case "mit"
                            dbClass(i).lMaxItemsBonus = t
                        Case "the"
                            dbClass(i).lCanSteal = t
                        Case "cdw"
                            dbClass(i).lCanDualWield = t
                    End Select
                End If
            Next
        End If
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Classes Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Emotions) [21%] ..."
i = 1
With MRSEMOTIONS
    .MoveLast
    ReDim dbEmotions(1 To .RecordCount) As UDTEmotions
    .MoveFirst
    Do
  
        dbEmotions(i).iID = !ID
        dbEmotions(i).sPhraseOthers = !PhraseOthers
        dbEmotions(i).sPhraseOthers2 = !PhraseOthers2
        dbEmotions(i).sPhraseToYou = !PhraseToYou
        dbEmotions(i).sPhraseYou = !PhraseYou
        dbEmotions(i).sPhraseYouToOther = !PhraseYouToOther
        dbEmotions(i).sSyntax = !Syntax
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Emotions Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Events) [24%] ..."
i = 1
With MRSEVENTS
    .MoveLast
    ReDim dbEvents(1 To .RecordCount) As UDTEvents
    .MoveFirst
    Do
        dbEvents(i).lEventID = !EventID
        dbEvents(i).lIsComplete = !IsComplete
        dbEvents(i).lPlayerID = !PlayerId
        dbEvents(i).sEndTime = !EndTime
        dbEvents(i).sStartTime = !StartTime
        dbEvents(i).sExpire = !Expire
        dbEvents(i).sCustomID = !CustomID
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Events Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Familiars) [27%] ..."
i = 1
With MRSFAMILIARS
    .MoveLast
    ReDim dbFamiliars(1 To .RecordCount) As UDTFamiliars
    .MoveFirst
    Do
        dbFamiliars(i).iID = !ID
        dbFamiliars(i).sDescription = !Description
        dbFamiliars(i).sFlags = !flags
        dbFamiliars(i).sMessage2 = !Message2
        dbFamiliars(i).lLevelMax = !LevelMax
        dbFamiliars(i).sFamName = !famName
        dbFamiliars(i).lLevelMod = !LevelMod
        dbFamiliars(i).dEXPPerLevel = !EXPPerLevel
        dbFamiliars(i).lMaxDam = !MaxDam
        dbFamiliars(i).lMinDam = !MinDam
        dbFamiliars(i).lStartHPMax = !StartHPMax
        dbFamiliars(i).lStartHPMin = !StartHPMin
        dbFamiliars(i).sAttackMessage = !AttackMessage
        dbFamiliars(i).lSwings = !Swings
        dbFamiliars(i).sMissMessage = !MissMessage
        dbFamiliars(i).sMissMessage2 = !MissMessage2
        dbFamiliars(i).lRidable = !Ridable
        dbFamiliars(i).lSpeed = !Speed
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Familiars Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Items) [30%] ..."
i = 1
With MRSITEM
    .MoveLast
    ReDim dbItems(1 To .RecordCount) As UDTItems
    .MoveFirst
    Do
  
        dbItems(i).dCost = !Cost
        dbItems(i).iAC = !AC
        dbItems(i).iArmorType = !ArmorType
        dbItems(i).iID = !ID
        dbItems(i).iInGame = !InGame
        dbItems(i).iLimit = !Limit
        dbItems(i).iMagical = !Magical
        dbItems(i).iMoveable = !Moveable
        dbItems(i).iSpeed = !Speed
        dbItems(i).iType = !Type
        dbItems(i).lLevel = !Level
        dbItems(i).sClassRestriction = !ClassRestriction
        dbItems(i).sDamage = !Damage
        dbItems(i).sDesc = !Desc
        dbItems(i).sItemName = !ItemName
        dbItems(i).sRaceRestriction = !RaceRestriction
        dbItems(i).sSwings = !Swings
        dbItems(i).sWorn = !Worn
        dbItems(i).iIsLedgenary = !Ledgendary
        dbItems(i).sScript = !Script
        dbItems(i).lDurability = !Durability
        dbItems(i).iUses = !Uses
        dbItems(i).dClassPoints = !ClassPoints
        dbItems(i).iOnEquipKillDur = !OnEquipKillDur
        dbItems(i).sMessage2 = !Message2
        dbItems(i).sMessageV = !MessageV
        dbItems(i).sFlags = !flags
        dbItems(i).sFlags2 = !Flags2
        dbItems(i).lOnLastUseDoFlags2 = !OnLastUseDoFlags2
        dbItems(i).sProjectile = !Projectile
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Items Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Map) [33%] ..."
i = 1
With MRSMAP
    .MoveLast
    ReDim dbMap(1 To .RecordCount) As UDTMap
    .MoveFirst
    Do
        'dbMap(i).dGold = !Gold '
        'dbMap(i).iMaxRegen = !MaxRegen '
        'dbMap(i).iMobGroup = !MobGroup '
        'dbMap(i).iType = !Type '
        dbMap(i).lDown = !Down
        dbMap(i).lEast = !East
        dbMap(i).lNorth = !North
        dbMap(i).lNorthEast = !NorthEast
        dbMap(i).lNorthWest = !NorthWest
        dbMap(i).lRoomID = !RoomID
        dbMap(i).lSouth = !South
        dbMap(i).lSouthEast = !SouthEast
        dbMap(i).lSouthWest = !SouthWest
        dbMap(i).lSpecialItem = !SpecialItem
        dbMap(i).lSpecialMon = !SpecialMon
        dbMap(i).lUp = !Up
        dbMap(i).lWest = !West
        dbMap(i).sItems = !Items
        dbMap(i).sMonsters = !Monsters
        dbMap(i).sRoomDesc = !RoomDesc
        dbMap(i).sRoomTitle = !RoomTitle
        'dbMap(i).sShopItems = !ShopItems '
        'dbMap(i).lLight = !Light '
        
        s = !Door
        s = ReplaceFast(s, ":", "")
        SplitFast s, Arr, ";"
        dbMap(i).lDD = Arr(9)
        dbMap(i).lDE = Arr(2)
        dbMap(i).lDN = Arr(0)
        dbMap(i).lDNE = Arr(4)
        dbMap(i).lDNW = Arr(5)
        dbMap(i).lDS = Arr(1)
        dbMap(i).lDSE = Arr(6)
        dbMap(i).lDSW = Arr(7)
        dbMap(i).lDU = Arr(8)
        dbMap(i).lDW = Arr(3)
        
        s = !Key
        s = ReplaceFast(s, ":", "")
        SplitFast s, Arr, ";"
        dbMap(i).lKD = Arr(9)
        dbMap(i).lKE = Arr(2)
        dbMap(i).lKN = Arr(0)
        dbMap(i).lKNE = Arr(4)
        dbMap(i).lKNW = Arr(5)
        dbMap(i).lKS = Arr(1)
        dbMap(i).lKSE = Arr(6)
        dbMap(i).lKSW = Arr(7)
        dbMap(i).lKU = Arr(8)
        dbMap(i).lKW = Arr(3)
        
        s = !Bash
        s = ReplaceFast(s, ":", "")
        SplitFast s, Arr, ";"
        dbMap(i).lBD = Arr(9)
        dbMap(i).lBE = Arr(2)
        dbMap(i).lBN = Arr(0)
        dbMap(i).lBNE = Arr(4)
        dbMap(i).lBNW = Arr(5)
        dbMap(i).lBS = Arr(1)
        dbMap(i).lBSE = Arr(6)
        dbMap(i).lBSW = Arr(7)
        dbMap(i).lBU = Arr(8)
        dbMap(i).lBW = Arr(3)
  
        s = !Pick
        s = ReplaceFast(s, ":", "")
        SplitFast s, Arr, ";"
        dbMap(i).lPD = Arr(9)
        dbMap(i).lPE = Arr(2)
        dbMap(i).lPN = Arr(0)
        dbMap(i).lPNE = Arr(4)
        dbMap(i).lPNW = Arr(5)
        dbMap(i).lPS = Arr(1)
        dbMap(i).lPSE = Arr(6)
        dbMap(i).lPSW = Arr(7)
        dbMap(i).lPU = Arr(8)
        dbMap(i).lPW = Arr(3)
        
        dbMap(i).sHidden = !Hidden
        dbMap(i).sScript = !Scripting
        'If InStr(LCaseFast(dbMap(i).sScript), "mybase.timer(") <> 0 Then
         '   ReDim Preserve dbMBTimer(UBound(dbMBTimer) + 1)
         '   sScripting 0, dbMap(i).lRoomID, , , , True, dbMBTimer(UBound(dbMBTimer)).lInterval, , dbMBTimer(UBound(dbMBTimer)).sScript
         '   With dbMBTimer(UBound(dbMBTimer))
       '         .lRoomID = dbMap(i).lRoomID
                
'                Debug.Print "------------------------"
'                Debug.Print "INDEX NUMBER : " & UBound(dbMBTimer)
'                Debug.Print "INTERVAL     : " & .lInterval
'                Debug.Print "SCRIPT       : " & .sScript
'                Debug.Print "ROOM ID      : " & .lRoomID
'                Debug.Print "------------------------"
    '        End With
    '    End If
        'If InStr(LCaseFast(dbMap(i).sScript), "begin.usescript ") <> 0 Then
      '      j = 0
       '     s = ""
       '     sScripting 0, dbMap(i).lRoomID, , , , True, j, , s
     '       If j <> 0 Then
      '          ReDim Preserve dbMBTimer(UBound(dbMBTimer) + 1)
      '          With dbMBTimer(UBound(dbMBTimer))
      '              .lRoomID = dbMap(i).lRoomID
      '              .lInterval = j
      '              .sScript = s
    '                Debug.Print "------------------------"
    '                Debug.Print "INDEX NUMBER : " & UBound(dbMBTimer)
    '                Debug.Print "INTERVAL     : " & .lInterval
    '                Debug.Print "SCRIPT       : " & .sScript
    '                Debug.Print "ROOM ID      : " & .lRoomID
    '                Debug.Print "------------------------"
         '       End With
         '   End If
    '    End If
        'dbMap(i).iSafeRoom = !Flags '
        'dbMap(i).lDeathRoom = !DeathRoom '
        'dbMap(i).iInDoor = !InDoor '
        'dbMap(i).iTrainClass = !TrainClass '
        dbMap(i).sMapFlags = !flags
        'dbMap(i).sMapFlags = dbMap(i).sMapFlags & "/0;"
        'modMapFlags.UpdateMapFlags i
        modMapFlags.LoadMapFlags i
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Map Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Monsters) [36%] ..."
i = 1
With MRSMONSTER
    .MoveLast
    ReDim dbMonsters(1 To .RecordCount) As UDTMonsters
    .MoveFirst
    Do
    
        dbMonsters(i).dEXP = !Exp
        dbMonsters(i).dHP = !HP
        dbMonsters(i).dMoney = !Money
        dbMonsters(i).iAC = !AC
        dbMonsters(i).iAttackable = !Attackable
        dbMonsters(i).iHostile = !Hostile
        dbMonsters(i).iType = !Type
        dbMonsters(i).sDropItem = !DropItem
        dbMonsters(i).lID = !ID
        dbMonsters(i).lMobGroup = !MobGroup
        dbMonsters(i).lRegenTime = !RegenTime
        dbMonsters(i).lRegenTimeLeft = !RegenTimeLeft
        dbMonsters(i).sAttack = !Attack
        dbMonsters(i).sDeathText = !DeathText
        dbMonsters(i).sDesc = !Desc
        dbMonsters(i).sMessage = !Message
        dbMonsters(i).sMonsterName = !MonsterName
        dbMonsters(i).iRoams = !Roams
        dbMonsters(i).iEvil = !Alignment
        dbMonsters(i).iDontAttackIfItem = !DontAttackIfItem
        dbMonsters(i).iAtDayMonster = !AtDayMonster
        dbMonsters(i).iAtNightMonster = !AtNightMonster
        dbMonsters(i).iDropCorpse = !DropCorpse
        dbMonsters(i).iTameToFam = !TameToFam
        dbMonsters(i).sScript = !OnDeathScript
        dbMonsters(i).lLevel = !Level
        dbMonsters(i).lEnergy = !TotalEnergy
        dbMonsters(i).lWeapon = !Weapon
        dbMonsters(i).lPEnergy = !PAttackEnergy
        dbMonsters(i).sSpells = !CastSpells
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Monsters Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Players) [39%] ..."
i = 1
With MRS
    .MoveLast
    ReDim dbPlayers(1 To .RecordCount) As UDTPlayers
    .MoveFirst
    Do
  
        dbPlayers(i).dBank = !Bank
        dbPlayers(i).dEXP = !Exp
        dbPlayers(i).dEXPNeeded = !EXPNeeded
        dbPlayers(i).dGold = !Gold
        dbPlayers(i).dMonsterID = 99999
        dbPlayers(i).iAC = !AC
        dbPlayers(i).iAcc = !Acc
        dbPlayers(i).iAgil = !AGIL
        dbPlayers(i).iArmorType = !ArmorType
        dbPlayers(i).iCasting = 0
        dbPlayers(i).iCha = !CHA
        dbPlayers(i).iCrits = !Crits
        dbPlayers(i).iDex = !DEX
        dbPlayers(i).iDodge = !Dodge
        dbPlayers(i).iVision = !Vision
        dbPlayers(i).iEcho = !Echo
        'dbPlayers(i).lFamID = !FamID
        dbPlayers(i).iHorse = !Horse
        dbPlayers(i).iInt = !Int
        dbPlayers(i).iInvitedBy = 0
        dbPlayers(i).sSeenAs = !PlayerName
        dbPlayers(i).sPlayerName = !SeenAs
        dbPlayers(i).iLeadingParty = 0
        dbPlayers(i).iLevel = !Level
        dbPlayers(i).iMaxDamage = !MaxDamage
        dbPlayers(i).iPartyLeader = 0
        dbPlayers(i).iResting = 0
        dbPlayers(i).iMeditating = 0
        dbPlayers(i).iSpellLevel = !SpellLevel
        dbPlayers(i).iSpellType = !SpellType
        dbPlayers(i).sSpells = !Spells
        dbPlayers(i).sSpellShorts = !SpellShorts
        dbPlayers(i).sTrainStats = !TrainStats
        dbPlayers(i).iStr = !Str
        dbPlayers(i).iStun = !Stun
        dbPlayers(i).sMiscFlag = !Misc
        dbPlayers(i).iWeapons = !Weapons
        dbPlayers(i).lBackUpLoc = !BackUpLoc
        dbPlayers(i).lHP = !HP
        dbPlayers(i).lLocation = !Location
        dbPlayers(i).lDBLocation = GetMapIndex(dbPlayers(i).lLocation)
        dbPlayers(i).lMana = !Mana
        dbPlayers(i).lMaxHP = !MaxHP
        dbPlayers(i).sStatline = !Statline
        dbPlayers(i).lMaxMana = !MaxMana
        dbPlayers(i).lPlayerID = !PlayerId
        dbPlayers(i).sArms = !Arms
        dbPlayers(i).sBody = !Body
        dbPlayers(i).sClass = !Class
        dbPlayers(i).sAppearance = !Appearance
        'dbPlayers(i).sFamName = !famName
        dbPlayers(i).sFeet = !Feet
        dbPlayers(i).sHands = !Hands
        dbPlayers(i).sHead = !Head
        dbPlayers(i).sInventory = !Inv
        dbPlayers(i).sLegs = !Legs
        dbPlayers(i).sParty = 0
        dbPlayers(i).sElements = !Resist
        dbPlayers(i).sPlayerPW = !PlayerPW
        dbPlayers(i).sOverrideDesc = !OverrideDesc
        dbPlayers(i).sQuest1 = !QUEST1
        dbPlayers(i).sQuest2 = !QUEST2
        dbPlayers(i).sQuest3 = !QUEST3
        dbPlayers(i).sQuest4 = !QUEST4
        dbPlayers(i).sRace = !Race
        dbPlayers(i).sWaist = !Waist
        dbPlayers(i).sWeapon = !Weapon
        dbPlayers(i).sBlessSpells = !BlessSpells
        dbPlayers(i).iLives = !Lives
        dbPlayers(i).iGuildLeader = !GuildLeader
        dbPlayers(i).sGuild = !Guild
        dbPlayers(i).sInvitedToGuild = "0"
        dbPlayers(i).iEvil = !Evil
        dbPlayers(i).sFace = !Face
        dbPlayers(i).sEars = !Ears
        dbPlayers(i).sNeck = !Neck
        dbPlayers(i).sBack = !Back
        dbPlayers(i).sShield = !Shield
        dbPlayers(i).iIsReadyToTrain = !IsReadyToTrain
        dbPlayers(i).sStatsPlus = !StatsPlus
        dbPlayers(i).iFlag1 = !Flag1
        dbPlayers(i).iFlag2 = !Flag2
        dbPlayers(i).iFlag3 = !Flag3
        dbPlayers(i).iFlag4 = !Flag4
'        dbPlayers(i).sLetters = !Letters
        dbPlayers(i).lPaper = !Paper
        dbPlayers(i).sBirthDay = !Birthday
        dbPlayers(i).lAge = !Age
        dbPlayers(i).dClassPoints = !ClassPoints
        dbPlayers(i).iClassBonusLevel = !ClassBonusLevel
        dbPlayers(i).iGender = !Gender
        dbPlayers(i).lClassChanges = !ClassChanges
        dbPlayers(i).dTotalEXP = !TotalEXP
        dbPlayers(i).dStamina = !Stamina
        dbPlayers(i).dHunger = !Hunger
        'dbPlayers(i).dFamEXP = !FamEXP
        'dbPlayers(i).lFamCurrentHP = !FamCurrentHP
        'dbPlayers(i).lFamMaxHP = !FamMaxHP
        dbPlayers(i).sKillDurItems = !KillDurItems
        dbPlayers(i).sFamFlags = !FamFlags
        modFamiliars.LoadFamFlags i
        SplitFast !Rings, Arr, ";"
        For j = 0 To 5
            dbPlayers(i).sRings(j) = Arr(j)
        Next
        If dbPlayers(i).sShield <> "0" Then
            With dbItems(GetItemID(, modItemManip.GetItemIDFromUnFormattedString(dbPlayers(i).sShield)))
                If .sWorn = "weapon" Then
                    dbPlayers(i).iDualWield = 1
                Else
                    dbPlayers(i).iDualWield = 0
                End If
            End With
        End If
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Players Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Races) [42%] ..."
i = 1
With MRSRACE
    .MoveLast
    ReDim dbRaces(1 To .RecordCount) As UDTRaces
    .MoveFirst
    Do
  
        dbRaces(i).dEXP = !Exp
        dbRaces(i).iID = !ID
        dbRaces(i).sName = !Name
        dbRaces(i).sStats = !Stats
        dbRaces(i).iVision = !Vision
        dbRaces(i).lMaxAge = !MaxAge
        dbRaces(i).lStartAgeMax = !StartAgeMax
        dbRaces(i).lStartAgeMin = !StartAgeMin
        dbRaces(i).sHP = !HP
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Races Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Spells) [45%] ..."
i = 1
With MRSSPELLS
    .MoveLast
    ReDim dbSpells(1 To .RecordCount) As UDTSpells
    .MoveFirst
    Do
  
        dbSpells(i).iUse = !Use
        dbSpells(i).iCast = !Cast
        dbSpells(i).iLevel = !Level
        dbSpells(i).iLevelMax = !LevelMax
        dbSpells(i).iLevelModify = !LevelModify
        dbSpells(i).iType = !Type
        dbSpells(i).lID = !ID
        dbSpells(i).lMana = !Mana
        dbSpells(i).lMaxDam = !MaxDam
        dbSpells(i).lMinDam = !MinDam
        dbSpells(i).sMessage = !Message
        dbSpells(i).sShort = !Short
        dbSpells(i).sSpellName = !SpellName
        dbSpells(i).lTimeOut = !TimeOut
        dbSpells(i).sRunOutMessage = !RunOutMessage
        dbSpells(i).sStatMessage = !StatMessage
        dbSpells(i).iDifficulty = !Difficulty
        dbSpells(i).lElement = !Element
        dbSpells(i).sFlags = !flags
        dbSpells(i).sMessage2 = !Message2
        dbSpells(i).sMessageV = !MessageV
        dbSpells(i).sEndCastFlags = !EndCastFlags
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Spells Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Shops) [48%] ..."
i = 1
With MRSSHOPS
    .MoveLast
    ReDim dbShops(1 To .RecordCount) As UDTShops
    .MoveFirst
    Do
        dbShops(i).iID = !ID
        dbShops(i).iMarkUp = !Markup
        dbShops(i).sShopName = !ShopName
        For j = 0 To 14
            dbShops(i).iItems(j) = .Fields("Item" & CStr(j + 1))
            dbShops(i).iQ(j) = .Fields("Q" & CStr(j + 1))
        Next
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
'If Not UpdateLog Then UpdateList "Shops Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Storing Arenas And Boss rooms) [51%] ..."
i = 1
j = 0
'For i = 1 To UBound(dbMap)
'    With dbMap(i)
 '       Select Case .iType
 '           Case 1 To 5
 '               j = j + 1
  '      End Select
  '  End With
'Next
'i = 1
'ReDim dbArenas(1 To j)
'j = 1
'For i = 1 To UBound(dbMap)
'    With dbMap(i)
 '       Select Case .iType
 '           Case 1 To 5
  '              dbArenas(j) = dbMap(i)
  '              dbArenas(j).ldbMapID = i
  '              j = j + 1
  '      End Select
 ''   End With
'Next
'If Not UpdateLog Then UpdateList "Arenas And Boss Rooms Stored... }b(}n}i" & Time & "}n}b)"
'If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Storing Door Rooms) [54%] ..."
'i = 1
'j = 0
'For i = 1 To UBound(dbMap)
'    With dbMap(i)
'        If .lDD <> 0 Then j = j + 1: GoTo nNextRecord
'        If .lDU <> 0 Then j = j + 1: GoTo nNextRecord
 ''       If .lDN <> 0 Then j = j + 1: GoTo nNextRecord
 '       If .lDS <> 0 Then j = j + 1: GoTo nNextRecord
 '       If .lDE <> 0 Then j = j + 1: GoTo nNextRecord
 '       If .lDW <> 0 Then j = j + 1: GoTo nNextRecord
 '       If .lDNW <> 0 Then j = j + 1: GoTo nNextRecord
  '      If .lDSW <> 0 Then j = j + 1: GoTo nNextRecord
 '       If .lDNE <> 0 Then j = j + 1: GoTo nNextRecord
 '       If .lDSE <> 0 Then j = j + 1: GoTo nNextRecord
'nNextRecord:
'    End With
'Next
'i = 1
'ReDim dbDoor(1 To j)
'j = 1
'For i = 1 To UBound(dbMap)
 '   With dbMap(i)
  '      If .lDD <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
 '       If .lDU <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
  '      If .lDN <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
  '      If .lDS <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
  '      If .lDE <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
  '      If .lDW <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
   '     If .lDNW <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
    '    If .lDSW <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
   '     If .lDNE <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
   ''     If .lDSE <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
'nNextRecord2:
 '   End With
'Next
'If Not UpdateLog Then UpdateList "Rooms With Doors Stored... }b(}n}i" & Time & "}n}b)"
bUpdate = False
i = 1
'If Not UpdateLog Then UpdateList "Database succesfully loaded to memory... }b(}n}i" & Time & "}n}b)"
On Error GoTo 0
Exit Function
LoadDatabaseIntoMemory_Error:
'MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: LoadDatabaseIntoMemory in Module, modMain"
bUpdate = False
'UpdateList "}b}uERROR ON LOAD }n}b" & Err.Number & "}n}i " & Err.Description, True
End Function
Sub UpdateMRSSets()
'update all the default recordsets
Set MRS = DB.OpenRecordset("SELECT * FROM Players")
Set MRSITEM = DB.OpenRecordset("SELECT * FROM Items")
Set MRSMAP = DB.OpenRecordset("SELECT * FROM Map")
Set MRSMONSTER = DB.OpenRecordset("SELECT * FROM Monsters")
Set MRSCLASS = DB.OpenRecordset("SELECT * FROM Class")
Set MRSRACE = DB.OpenRecordset("SELECT * FROM Races")
Set MRSEMOTIONS = DB.OpenRecordset("SELECT * FROM Emotions")
Set MRSSPELLS = DB.OpenRecordset("SELECT * FROM Spells")
Set MRSFAMILIARS = DB.OpenRecordset("SELECT * FROM Familiars")
Set MRSSHOPS = DB.OpenRecordset("SELECT * FROM Shops")
Set MRSEVENTS = DB.OpenRecordset("SELECT * FROM Events")
End Sub

Sub AddBonusStatsInverse(dbIndex As Long, sBonuses As String)
Dim tArr() As String
tArr = Split(sBonuses, ":")
For i = LBound(tArr) To UBound(tArr)
    With dbPlayers(dbIndex)
        If tArr(i) <> "" Then
            Select Case Left$(tArr(i), 3)
                Case "mhp" 'Max Hitpoints
                    .lMaxHP = .lMaxHP - CLng(Replace$(tArr(i), "mhp", ""))
                Case "str" 'Strength
                    .iStr = .iStr - CLng(Replace$(tArr(i), "str", ""))
                Case "agi" 'Agility
                    .iAgil = .iAgil - CLng(Replace$(tArr(i), "agi", ""))
                Case "int" 'Intellect
                    .iInt = .iInt - CLng(Replace$(tArr(i), "int", ""))
                Case "cha" 'Charm
                    .iCha = .iCha - CLng(Replace$(tArr(i), "cha", ""))
                Case "dex" 'Dexterity
                    .iDex = .iDex - CLng(Replace$(tArr(i), "dex", ""))
                Case "pac" 'Armor Class
                    .iAC = .iAC - CLng(Replace$(tArr(i), "pac", ""))
                Case "acc" 'Accurracy
                    .iAcc = .iAcc - CLng(Replace$(tArr(i), "acc", ""))
                Case "cri" 'Crits
                    .iCrits = .iCrits - CLng(Replace$(tArr(i), "cri", ""))
                Case "mma" 'Max Mana
                    .lMaxMana = .lMaxMana - CLng(Replace$(tArr(i), "mma", ""))
                Case "dam" 'damage bonus
                    .iMaxDamage = .iMaxDamage - CLng(Replace$(tArr(i), "dam", ""))
                Case "dod" 'dodge
                    .iDodge = .iDodge - CLng(Replace$(tArr(i), "dod", ""))
            End Select
        End If
    End With
    If DE Then DoEvents
Next
End Sub

Sub AddBonusStats(dbIndex As Long, sBonuses As String)
Dim tArr() As String
Arr = Split(sBonuses, ":")
For i = LBound(tArr) To UBound(tArr)
    With dbPlayers(dbIndex)
        If tArr(i) <> "" Then
            Select Case Left$(tArr(i), 3)
                Case "mhp" 'Max Hitpoints
                    .lMaxHP = .lMaxHP + CLng(Replace$(tArr(i), "mhp", ""))
                Case "str" 'Strength
                    .iStr = .iStr + CLng(Replace$(tArr(i), "str", ""))
                Case "agi" 'Agility
                    .iAgil = .iAgil + CLng(Replace$(tArr(i), "agi", ""))
                Case "int" 'Intellect
                    .iInt = .iInt + CLng(Replace$(tArr(i), "int", ""))
                Case "cha" 'Charm
                    .iCha = .iCha + CLng(Replace$(tArr(i), "cha", ""))
                Case "dex" 'Dexterity
                    .iDex = .iDex + CLng(Replace$(tArr(i), "dex", ""))
                Case "pac" 'Armor Class
                    .iAC = .iAC + CLng(Replace$(tArr(i), "pac", ""))
                Case "acc" 'Accurracy
                    .iAcc = .iAcc + CLng(Replace$(tArr(i), "acc", ""))
                Case "cri" 'Crits
                    .iCrits = .iCrits + CLng(Replace$(tArr(i), "cri", ""))
                Case "mma" 'Max Mana
                    .lMaxMana = .lMaxMana + CLng(Replace$(tArr(i), "mma", ""))
                Case "dam" 'damage bonus
                    .iMaxDamage = .iMaxDamage + CLng(Replace$(tArr(i), "dam", ""))
                Case "dod" 'dodge
                    .iDodge = .iDodge + CLng(Replace$(tArr(i), "dod", ""))
            End Select
        End If
    End With
    If DE Then DoEvents
Next
End Sub

Public Sub ReverseEffects(lcID As Long, WhichDB As SaveDB)
Dim i As Long
Dim l As Long
Dim j As Long
Dim Arr() As String

Select Case WhichDB
    Case Class
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            l = 0
            If dbPlayers(i).sClass = dbClass(lcID).sName Then
                'do it
                With dbPlayers(i)
                    If .dClassPoints > dbClass(lcID).dBeginnerMax Then AddBonusStatsInverse i, dbClass(lcID).sBBonus
                    If .dClassPoints > dbClass(lcID).dIntermediateMax Then AddBonusStatsInverse i, dbClass(lcID).sIBonus
                    If .dClassPoints > dbClass(lcID).dMasterMax Then AddBonusStatsInverse i, dbClass(lcID).sMBonus
                    If .dClassPoints > dbClass(lcID).dGuru Then AddBonusStatsInverse i, dbClass(lcID).sGBonus
                    .iClassBonusLevel = 0
                    AddBonusStatsInverse i, dbClass(lcID).sBaseBonus
                    .iAcc = .iAcc - dbClass(lcID).iAcc
                    .iArmorType = dbClass(lcID).iArmorType
                    .iWeapons = dbClass(lcID).iWeapon
                    .iDodge = .iDodge - dbClass(lcID).lDodgeBonus
                    .lMaxHP = .lMaxHP - dbClass(lcID).lHPBonus
                    .lMaxMana = .lMaxMana - dbClass(lcID).lMABonus
                    .iCrits = .iCrits - dbClass(lcID).iCrits
                    .iMaxDamage = .iMaxDamage - dbClass(lcID).lDamBonus
                    .iSpellLevel = dbClass(lcID).iSpellLevel
                    .iSpellType = dbClass(lcID).iSpellType
                    .iAC = .iAC - dbClass(lcID).lACBonus
                    modMiscFlag.SetMiscFlag i, [Can Sneak], dbClass(lcID).lCanSneak
                    modMiscFlag.SetStatsPlus i, [Max Items Bonus], modMiscFlag.GetStatsPlus(i, [Max Items Bonus]) - dbClass(lcID).lMaxItemsBonus
                    .iVision = .iVision - dbClass(lcID).lVisionBonus
                    .dEXPNeeded = dbRaces(GetRaceID(.sRace)).dEXP
                    Arr = Split(.sTrainStats, "/")
                    .iStr = .iStr - Val(Arr(0))
                    .iInt = .iInt - Val(Arr(1))
                    .iDex = .iDex - Val(Arr(2))
                    .iAgil = .iAgil - Val(Arr(3))
                    .iCha = .iCha - Val(Arr(4))
                    j = 0
                    For l = 0 To 4
                        j = j + Val(Arr(l))
                    Next
                    .lMaxHP = Val(Arr(5)) \ .iLevel
                    .lMaxMana = Val(Arr(6)) \ .iLevel
                    .iLevel = 1
                    If dbClass(lcID).lCPBonus > 0 Then j = j \ dbClass(lcID).lCPBonus
                    .sTrainStats = j
                    .sClass = lcID
                End With
            End If
        Next
    Case Race
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            If dbPlayers(i).sRace = dbRaces(lcID).sName Then
                'do it
                With dbPlayers(i)
                    .iVision = .iVision - dbRaces(lcID).iVision
                    If .lAge > dbRaces(lcID).lMaxAge Then .lAge = dbRaces(lcID).lStartAgeMin
                    Arr = Split(dbRaces(lcID).sStats, ":")
                    .iStr = .iStr - Val(Arr(0))
                    .iInt = .iInt - Val(Arr(2))
                    .iDex = .iDex - Val(Arr(4))
                    .iAgil = .iAgil - Val(Arr(1))
                    .iCha = .iCha - Val(Arr(3))
                    .dEXPNeeded = dbClass(GetClassID(.sClass)).dEXP
                    Erase Arr
                    Arr = Split(.sTrainStats, "/")
                    .iStr = .iStr - Val(Arr(0))
                    .iInt = .iInt - Val(Arr(1))
                    .iDex = .iDex - Val(Arr(2))
                    .iAgil = .iAgil - Val(Arr(3))
                    .iCha = .iCha - Val(Arr(4))
                    j = 0
                    For l = 0 To 4
                        j = j + Val(Arr(l))
                    Next
                    .lMaxHP = Val(Arr(5)) \ .iLevel
                    .lMaxMana = Val(Arr(6)) \ .iLevel
                    .iLevel = 1
                    .sTrainStats = j
                    .sRace = lcID
                End With
            End If
        Next
    Case Familiars
        'For i = LBound(dbPlayers) To UBound(dbPlayers)
        '    With dbPlayers(i)
         '       If .iFamID = dbFamiliars(lcID).iID Then
           '         DoItemFlags i, 0, 0, , True, , , True, dbFamiliars(lcID).sFlags
           '         .sFamName = "0"
           '     End If
         '   End With
        'Next
    Case Item
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            With dbPlayers(i)
                Select Case dbItems(lcID).sWorn
                    Case "arms"
                        If GetIDByStr(.sArms) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "back"
                        If GetIDByStr(.sBack) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "body"
                        If GetIDByStr(.sBody) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "ears"
                        If GetIDByStr(.sEars) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "face"
                        If GetIDByStr(.sFace) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "hands"
                        If GetIDByStr(.sHands) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "head"
                        If GetIDByStr(.sHead) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "legs"
                        If GetIDByStr(.sLegs) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "neck"
                        If GetIDByStr(.sNeck) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "shield"
                        If GetIDByStr(.sShield) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "waist"
                        If GetIDByStr(.sWaist) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "weapon"
                        If GetIDByStr(.sWeapon) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0, , True
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                If .sKillDurItems = "" Then .sKillDurItems = "0"
                            End If
                        End If
                    Case "ring"
                        For j = 0 To 5
                            If .sRings(i) <> "0" Then
                                If GetIDByStr(.sRings(i)) = dbItems(lcID).iID Then
                                    DoItemFlags i, lcID, 0, , True
                                    If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                        .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                        If .sKillDurItems = "" Then .sKillDurItems = "0"
                                    End If
                                End If
                            End If
                        Next
                End Select
            End With
        Next
    Case Spells
        
End Select
End Sub

Public Function GetIDByStr(ByVal s As String) As Long
Dim m As Long
Dim n As Long

   On Error GoTo GetItemIDFromUnFormattedString_Error

m = InStr(1, s, ":")
n = InStr(m, s, "/")
GetIDByStr = CLng(Mid$(s, m + 1, n - m - 1))

   On Error GoTo 0
   Exit Function

GetItemIDFromUnFormattedString_Error:

    GetItemIDbyStr = -1
End Function

Public Sub DoEffects(lcID As Long, WhichDB As SaveDB)
Dim i As Long
Dim b As Boolean
Dim l As Long
Dim n As Long
Dim j As Long

Select Case WhichDB
    Case Class
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            If dbPlayers(i).sClass = CStr(lcID) Then
                'do it
                With dbPlayers(i)
                    AddBonusStats i, dbClass(lcID).sBaseBonus
                    .iAcc = .iAcc + dbClass(lcID).iAcc
                    .iArmorType = dbClass(lcID).iArmorType
                    .iWeapons = dbClass(lcID).iWeapon
                    .iDodge = .iDodge + dbClass(lcID).lDodgeBonus
                    .lMaxHP = .lMaxHP + dbClass(lcID).lHPBonus
                    .lMaxMana = .lMaxMana + dbClass(lcID).lMABonus
                    .iCrits = .iCrits + dbClass(lcID).iCrits
                    .iMaxDamage = .iMaxDamage + dbClass(lcID).lDamBonus
                    .iSpellLevel = dbClass(lcID).iSpellLevel
                    .iSpellType = dbClass(lcID).iSpellType
                    .iAC = .iAC + dbClass(lcID).lACBonus
                    modMiscFlag.SetMiscFlag i, [Can Sneak], dbClass(lcID).lCanSneak
                    modMiscFlag.SetStatsPlus i, [Max Items Bonus], modMiscFlag.GetStatsPlus(i, [Max Items Bonus]) + dbClass(lcID).lMaxItemsBonus
                    .iVision = .iVision + dbClass(lcID).lVisionBonus
                    .dEXPNeeded = dbRaces(GetRaceID(.sRace)).dEXP + dbClass(lcID).dEXP
                    l = .dTotalEXP
                    b = False
                    n = 0
                    Do Until b
                        l = l - .dEXPNeeded - n
                        If l >= 0 Then .iLevel = .iLevel + 1: n = n + 1
                        If l <= 0 Then b = True
                    Loop
                    .dEXPNeeded = .dEXPNeeded + n
                    If l < 0 Then .dEXP = l + .dEXPNeeded
                    .sTrainStats = .iLevel - 1
                    If dbClass(lcID).lCPBonus > 0 Then .sTrainStats = Val(.sTrainStats) * dbClass(lcID).lCPBonus
                    .iIsReadyToTrain = .iIsReadyToTrain + Val(.sTrainStats)
                    .lMaxHP = .lMaxHP * .iLevel
                    .lMaxMana = .lMaxMana * .iLevel
                    .sTrainStats = "0/0/0/0/0/" & .lMaxHP & "/" & .lMaxMana
                    If .dClassPoints > dbClass(lcID).dBeginnerMax Then AddBonusStats i, dbClass(lcID).sBBonus: .iClassBonusLevel = 1
                    If .dClassPoints > dbClass(lcID).dIntermediateMax Then AddBonusStats i, dbClass(lcID).sIBonus: .iClassBonusLevel = 2
                    If .dClassPoints > dbClass(lcID).dMasterMax Then AddBonusStats i, dbClass(lcID).sMBonus: .iClassBonusLevel = 3
                    If .dClassPoints > dbClass(lcID).dGuru Then AddBonusStats i, dbClass(lcID).sGBonus: .iClassBonusLevel = 4
                    .sClass = dbClass(lcID).sName
                End With
            End If
        Next
    Case Race
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            If dbPlayers(i).sRace = CStr(lcID) Then
                'do it
                With dbPlayers(i)
                    .iVision = .iVision + dbRaces(lcID).iVision
                    If .lAge > dbRaces(lcID).lMaxAge Then .lAge = dbRaces(lcID).lStartAgeMin
                    Arr = Split(dbRaces(lcID).sStats, ":")
                    .iStr = .iStr + Val(Arr(0))
                    .iInt = .iInt + Val(Arr(2))
                    .iDex = .iDex + Val(Arr(4))
                    .iAgil = .iAgil + Val(Arr(1))
                    .iCha = .iCha + Val(Arr(3))
                    .dEXPNeeded = dbRaces(lcID).dEXP + dbClass(GetClassID(.sClass)).dEXP
                    l = .dTotalEXP
                    b = False
                    n = 0
                    Do Until b
                        l = l - .dEXPNeeded - n
                        If l >= 0 Then .iLevel = .iLevel + 1: n = n + 1
                        If l <= 0 Then b = True
                    Loop
                    .dEXPNeeded = .dEXPNeeded + n
                    If l < 0 Then .dEXP = l + .dEXPNeeded
                    .sTrainStats = .iLevel - 1
                    .iIsReadyToTrain = .iIsReadyToTrain + Val(.sTrainStats)
                    .lMaxHP = .lMaxHP * .iLevel
                    .lMaxMana = .lMaxMana * .iLevel
                    .sTrainStats = "0/0/0/0/0/" & .lMaxHP & "/" & .lMaxMana
                    .sRace = dbRaces(lcID).sName
                End With
            End If
        Next
    Case Familiars
        'For i = LBound(dbPlayers) To UBound(dbPlayers)
        '    With dbPlayers(i)
        '        If .iFamID = dbFamiliars(lcID).iID Then
        '            DoItemFlags i, 0, 0, , , , , True, dbFamiliars(lcID).sFlags
        '            .sFamName = dbFamiliars(lcID).sFamName
       '         End If
        '    End With
       ' Next
    Case Item
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            With dbPlayers(i)
                Select Case dbItems(lcID).sWorn
                    Case "arms"
                        If GetIDByStr(.sArms) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "back"
                        If GetIDByStr(.sBack) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "body"
                        If GetIDByStr(.sBody) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "ears"
                        If GetIDByStr(.sEars) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "face"
                        If GetIDByStr(.sFace) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "hands"
                        If GetIDByStr(.sHands) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "head"
                        If GetIDByStr(.sHead) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "legs"
                        If GetIDByStr(.sLegs) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "neck"
                        If GetIDByStr(.sNeck) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "shield"
                        If GetIDByStr(.sShield) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "waist"
                        If GetIDByStr(.sWaist) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "weapon"
                        If GetIDByStr(.sWeapon) = dbItems(lcID).iID Then
                            DoItemFlags i, lcID, 0
                            If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                If .sKillDurItems = "0" Then .sKillDurItems = ""
                                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                            End If
                        End If
                    Case "ring"
                        For j = 0 To 5
                            If .sRings(i) <> "0" Then
                                If GetIDByStr(.sRings(i)) = dbItems(lcID).iID Then
                                    DoItemFlags i, lcID, 0
                                    If dbItems(lcID).iOnEquipKillDur <> 0 Then
                                        If .sKillDurItems = "0" Then .sKillDurItems = ""
                                        .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(lcID).sWorn & "/" & dbItems(lcID).iOnEquipKillDur & ";", "", 1, 1)
                                    End If
                                End If
                            End If
                        Next
                End Select
            End With
        Next
End Select
End Sub
