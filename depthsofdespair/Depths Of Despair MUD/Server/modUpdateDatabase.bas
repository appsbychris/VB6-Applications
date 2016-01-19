Attribute VB_Name = "modUpdateDatabase"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modUpdateDatabase
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public sIndex As String

Public Function GetPlayerIndexNumber(Optional Index As Long = -1, Optional sPlayerName As String = "", Optional PlayerID As Long = -1) As Long
Dim i As Long
If Index <> -1 Then
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        If dbPlayers(i).iIndex = Index Then
            GetPlayerIndexNumber = i
            Exit For
        End If
        If DE Then DoEvents
    Next
ElseIf Not modSC.FastStringComp(sPlayerName, "") Then
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        If modSC.FastStringComp(LCaseFast(dbPlayers(i).sPlayerName), LCaseFast(sPlayerName)) Then
            GetPlayerIndexNumber = i
            Exit For
        End If
        If DE Then DoEvents
    Next
ElseIf PlayerID <> -1 Then
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        If dbPlayers(i).lPlayerID = PlayerID Then
            GetPlayerIndexNumber = i
            Exit For
        End If
        If DE Then DoEvents
    Next
End If
End Function

Public Function GetEventID(Optional CustomID As String = "-1", Optional PlayerID As Long = -1, Optional EventID As Long = -1) As Long
Dim i As Long
If Not modSC.FastStringComp(CustomID, "-1") Then
    For i = LBound(dbEvents) To UBound(dbEvents)
        If modSC.FastStringComp(dbEvents(i).sCustomID, CustomID) And dbEvents(i).lPlayerID = PlayerID Then
            GetEventID = i
            Exit For
        End If
        If DE Then DoEvents
    Next
ElseIf EventID <> -1 Then
    For i = LBound(dbEvents) To UBound(dbEvents)
        If dbEvents(i).lEventID = EventID And dbEvents(i).lPlayerID = PlayerID Then
            GetEventID = i
            Exit For
        End If
        If DE Then DoEvents
    Next
End If
End Function

Public Function GetMonsterID(Optional sMonsterName As String = "", Optional MonID As Long = -1) As Long
    Dim i As Long
    If Not modSC.FastStringComp(sMonsterName, "") Then
        For i = LBound(dbMonsters) To UBound(dbMonsters)
            If modSC.FastStringComp(LCaseFast(dbMonsters(i).sMonsterName), LCaseFast(sMonsterName)) Then
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

Public Function GetSpellID(Optional sSpellName As String = "", Optional SpellID As Long = -1) As Long
    Dim i As Long
    If Not modSC.FastStringComp(sSpellName, "") Then
        For i = LBound(dbSpells) To UBound(dbSpells)
            If modSC.FastStringComp(LCaseFast(dbSpells(i).sSpellName), LCaseFast(sSpellName)) Then
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
            If LCaseFast(.sTitle) = sTitle Then
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

Public Function GetClassID(Optional sClassName As String = "", Optional ClassID As Long = -1) As Long
    Dim i As Long
    If Not modSC.FastStringComp(sClassName, "") Then
        For i = LBound(dbClass) To UBound(dbClass)
            If modSC.FastStringComp(LCaseFast(dbClass(i).sName), LCaseFast(sClassName)) Then
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
    If Not modSC.FastStringComp(sRaceName, "") Then
        For i = LBound(dbRaces) To UBound(dbRaces)
            If modSC.FastStringComp(LCaseFast(dbRaces(i).sName), LCaseFast(sRaceName)) Then
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
    ElseIf Not modSC.FastStringComp(sFamName, "") Then
        For i = LBound(dbFamiliars) To UBound(dbFamiliars)
            If modSC.FastStringComp(LCaseFast(dbFamiliars(i).sFamName), LCaseFast(sFamName)) Then
                GetFamID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
End Function

Public Function GetItemID(Optional sItemName As String = "", Optional lItemID As Long = 0) As Long
    Dim i As Long
    Dim b As Boolean
    Dim s As String
    If Not modSC.FastStringComp(sItemName, "") Then
        s = LCaseFast(sItemName)
        For i = LBound(dbItems) To UBound(dbItems)
            If modSC.FastStringComp(LCaseFast(dbItems(i).sItemName), s) Then
                GetItemID = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    ElseIf lItemID > 0 Then
        On Error GoTo eh1:
        b = False
        If dbItems(lItemID).iID = lItemID Then
            GetItemID = lItemID
            b = True
        End If
        If lItemID > 1 And Not b Then
            If dbItems(lItemID - 1).iID = lItemID Then
                GetItemID = lItemID
                b = True
            End If
        End If
        If lItemID < UBound(dbItems) And Not b Then
            If dbItems(lItemID + 1).iID = lItemID Then
                GetItemID = lItemID
                b = True
            End If
        End If
        If Not b Then
eh1:
            On Error GoTo eh2:
            For i = LBound(dbItems) To UBound(dbItems)
                If lItemID = dbItems(i).iID Then
                    GetItemID = i
                    Exit For
                End If
                If DE Then DoEvents
            Next
        End If
    End If
eh2:
End Function

Public Function GetMapIndex(RoomID As Long) As Long
    Dim i As Long
    Dim b As Boolean
    b = False
    On Error GoTo eh1
    If dbMap(RoomID).lRoomID = RoomID Then
        GetMapIndex = RoomID
        b = True
    End If
    If RoomID > 1 Then
        If dbMap(RoomID - 1).lRoomID = RoomID Then
            GetMapIndex = RoomID - 1
            b = True
        End If
    End If
    If RoomID < UBound(dbMap) Then
        If dbMap(RoomID + 1).lRoomID = RoomID Then
            GetMapIndex = RoomID + 1
            b = True
        End If
    End If
    If Not b Then
eh1:
        On Error GoTo eh2:
        For i = LBound(dbMap) To UBound(dbMap)
            If RoomID = dbMap(i).lRoomID Then
                GetMapIndex = i
                Exit For
            End If
            If DE Then DoEvents
        Next
    End If
eh2:
End Function

'Public Function GetDoorIndex(RoomID As Long) As Long
'    Dim i As Long
'    For i = LBound(dbDoor) To UBound(dbDoor)
'        If RoomID = dbDoor(i).lRoomID Then
'            GetMapIndex = i
'            Exit For
'        End If
'        If DE Then DoEvents
'    Next
'End Function

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

Public Sub SaveMemoryToDatabase(iStep As Long)
'On Error GoTo SaveMemoryToDatabase_Error
Dim bFound As Boolean
Dim lIDn As Long
Dim i As Long
Dim s As String
Dim j As Long
Select Case iStep
    Case 0
        Set MRS = db.OpenRecordset("SELECT * FROM Players")
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            With MRS
                .MoveFirst
                Do
                    If dbPlayers(i).lPlayerID = CLng(!PlayerID) Then
                        .Edit
                        modFamiliars.UpdateFamFlags i
                        !FamFlags = dbPlayers(i).sFamFlags
                        !PlayerName = dbPlayers(i).sSeenAs
                        !SeenAs = dbPlayers(i).sPlayerName
                        !Bank = dbPlayers(i).dBank
                        !EXP = dbPlayers(i).dEXP
                        !EXPneeded = dbPlayers(i).dEXPNeeded
                        !Gold = dbPlayers(i).dGold
                        !AC = dbPlayers(i).iAC
                        !Acc = dbPlayers(i).iAcc
                        !AGIL = dbPlayers(i).iAgil
                        !TrainStats = dbPlayers(i).sTrainStats
                        !CHA = dbPlayers(i).iCha
                        !Crits = dbPlayers(i).iCrits
                        !Misc = dbPlayers(i).sMiscFlag
                        !DEX = dbPlayers(i).iDex
                        !Dodge = dbPlayers(i).iDodge
                        !Echo = dbPlayers(i).iEcho
                        '!FamID = dbPlayers(i).lFamID
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
                        !MANA = dbPlayers(i).lMana
                        !MaxHP = dbPlayers(i).lMaxHP
                        !MaxMana = dbPlayers(i).lMaxMana
                        !Arms = dbPlayers(i).sArms
                        !Body = dbPlayers(i).sBody
                        '!famName = dbPlayers(i).sFamName
                        !Feet = dbPlayers(i).sFeet
                        !Hands = dbPlayers(i).sHands
                        !Head = dbPlayers(i).sHead
                        !OverrideDesc = dbPlayers(i).sOverrideDesc
                        !Inv = dbPlayers(i).sInventory
                        !Legs = dbPlayers(i).sLegs
                        !QUEST1 = dbPlayers(i).sQuest1
                        !QUEST2 = dbPlayers(i).sQuest2
                        !QUEST3 = dbPlayers(i).sQuest3
                        !QUEST4 = dbPlayers(i).sQuest4
                        !Index = dbPlayers(i).iIndex
                        !Appearance = dbPlayers(i).sAppearance
                        !Waist = dbPlayers(i).sWaist
                        !Weapon = dbPlayers(i).sWeapon
                        !Spells = dbPlayers(i).sSpells
                        !SpellShorts = dbPlayers(i).sSpellShorts
                        !Race = dbPlayers(i).sRace
                        !Class = dbPlayers(i).sClass
                        !Resist = dbPlayers(i).sElements
                        !Weapons = dbPlayers(i).iWeapons
                        !ArmorType = dbPlayers(i).iArmorType
                        !SpellLevel = dbPlayers(i).iSpellLevel
                        !SpellType = dbPlayers(i).iSpellType
                        !BlessSpells = dbPlayers(i).sBlessSpells
                        !Guild = dbPlayers(i).sGuild
                        !GuildLeader = dbPlayers(i).iGuildLeader
                        !Evil = dbPlayers(i).iEvil
                        !Face = dbPlayers(i).sFace
                        !Ears = dbPlayers(i).sEars
                        !Neck = dbPlayers(i).sNeck
                        !Back = dbPlayers(i).sBack
                        !Shield = dbPlayers(i).sShield
                        !IsReadyToTrain = dbPlayers(i).iIsReadyToTrain
                        !StatsPlus = dbPlayers(i).sStatsPlus
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
                        '!FamEXP = dbPlayers(i).dFamEXP
                        '!FamCurrentHP = dbPlayers(i).lFamCurrentHP
                        '!FamMaxHP = dbPlayers(i).lFamMaxHP
                        !KillDurItems = dbPlayers(i).sKillDurItems
                        !Gender = dbPlayers(i).iGender
                        !Statline = dbPlayers(i).sStatline
                        If !Birthday = "0" Then !Birthday = dbPlayers(i).sBirthDay
                        s = ""
                        For j = 0 To 5
                            s = s & dbPlayers(i).sRings(j) & ";"
                            If DE Then DoEvents
                        Next
                        !Rings = s
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
        Set MRSMAP = db.OpenRecordset("SELECT * FROM Map")
        For i = LBound(dbMap) To UBound(dbMap)
            With MRSMAP
                .MoveFirst
                Do
                    If dbMap(i).lRoomID = CLng(!RoomID) Then
                        .Edit
                        modMapFlags.UpdateMapFlags i
                        !Flags = dbMap(i).sMapFlags
                        If dbMap(i).sItems = "" Then dbMap(i).sItems = "0"
                        !Items = dbMap(i).sItems
                        !Monsters = dbMap(i).sMonsters
                        !Hidden = dbMap(i).sHidden
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
        Set MRSMAP = Nothing
    Case 2
        Set MRSMONSTER = db.OpenRecordset("SELECT * FROM Monsters")
        For i = LBound(dbMonsters) To UBound(dbMonsters)
            With MRSMONSTER
                .MoveFirst
                Do
                    If dbMonsters(i).lID = CLng(!id) Then
                        If dbMonsters(i).lRegenTimeLeft <> Val(!RegenTimeLeft) Then
                            .Edit
                            !RegenTimeLeft = dbMonsters(i).lRegenTimeLeft
                            .Update
                        End If
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
            End With
            If DE Then DoEvents
        Next i
        Set MRSMONSTER = Nothing
    Case 3
        Set MRSITEM = db.OpenRecordset("SELECT * FROM Items")
        For i = LBound(dbItems) To UBound(dbItems)
            With MRSITEM
                .MoveFirst
                Do
                    If dbItems(i).iID = CLng(!id) Then
                        If CLng(!InGame) <> dbItems(i).iInGame Then
                            .Edit
                            !InGame = dbItems(i).iInGame
                            .Update
                        End If
                        Exit Do
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
            End With
            If DE Then DoEvents
        Next i
        Set MRSITEM = Nothing
        Set MRSEVENTS = db.OpenRecordset("SELECT * FROM Events")
        For i = LBound(dbEvents) To UBound(dbEvents)
            With MRSEVENTS
                .MoveFirst
                Do
                    If dbEvents(i).lEventID = CLng(!EventID) Then
                        If dbEvents(i).lPlayerID <> CLng(!PlayerID) Then
                            .Edit
                            !IsComplete = dbEvents(i).lIsComplete
                            !PlayerID = dbEvents(i).lPlayerID
                            !EndTime = dbEvents(i).sEndTime
                            !StartTime = dbEvents(i).sStartTime
                            !Expire = dbEvents(i).sExpire
                            !CustomID = dbEvents(i).sCustomID
                            .Update
                        End If
                        .MoveNext
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
'                If Not bFound Then
'                    .AddNew
'                    lID = 1
'                    Do
'                        If lID = CLng(!EventID) Then
'                            lID = lID + 1
'                            .MoveFirst
'                        ElseIf Not .EOF Then
'                            .MoveNext
'                        End If
'                        If DE Then DoEvents
'                    Loop Until .EOF
'                    .AddNew
'                    !EventID = lID
'                    dbEvents(i).lEventID = lID
'                    !IsComplete = dbEvents(i).lIsComplete
'                    !PlayerID = dbEvents(i).lPlayerID
'                    !EndTime = dbEvents(i).sEndTime
'                    !StartTime = dbEvents(i).sStartTime
'                    !Expire = dbEvents(i).sExpire
'                    .Update
'                End If
            End With
            If DE Then DoEvents
        Next
    Case 4:
        Set MRSSHOPS = db.OpenRecordset("SELECT * FROM SHOPS")
        For i = LBound(dbShops) To UBound(dbShops)
            With MRSSHOPS
                .MoveFirst
                Do
                    If !id = dbShops(i).iID Then
                        .Edit
                        !Q1 = dbShops(i).iQ(0)
                        !Q2 = dbShops(i).iQ(1)
                        !Q3 = dbShops(i).iQ(2)
                        !Q4 = dbShops(i).iQ(3)
                        !Q5 = dbShops(i).iQ(4)
                        !Q6 = dbShops(i).iQ(5)
                        !Q7 = dbShops(i).iQ(6)
                        !Q8 = dbShops(i).iQ(7)
                        !Q9 = dbShops(i).iQ(8)
                        !Q10 = dbShops(i).iQ(9)
                        !Q11 = dbShops(i).iQ(10)
                        !Q12 = dbShops(i).iQ(11)
                        !Q13 = dbShops(i).iQ(12)
                        !Q14 = dbShops(i).iQ(13)
                        !Q15 = dbShops(i).iQ(14)
                        .Update
                        .MoveNext
                    ElseIf Not .EOF Then
                        .MoveNext
                    End If
                    If DE Then DoEvents
                Loop Until .EOF
            End With
            If DE Then DoEvents
        Next
End Select
On Error GoTo 0
Exit Sub
SaveMemoryToDatabase_Error:
UpdateList "}b}uAn error occured while saving information to the database.", True
UpdateList "          }b" & Err.Number & " }n}i" & Err.Description, True
UpdateList "          }iOccured on Staggered Step " & iStep, True
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
        dbClass(i).dEXP = !EXP
        dbClass(i).iArmorType = !ArmorType
        dbClass(i).iID = !id
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
        s = !Flags
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
                If Not modSC.FastStringComp(Arr(a), "") Then
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
If Not UpdateLog Then UpdateList "Classes Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Emotions) [21%] ..."
i = 1
With MRSEMOTIONS
    .MoveLast
    ReDim dbEmotions(1 To .RecordCount) As UDTEmotions
    .MoveFirst
    Do
  
        dbEmotions(i).iID = !id
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
If Not UpdateLog Then UpdateList "Emotions Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Events) [24%] ..."
i = 1
With MRSEVENTS
    .MoveLast
    ReDim dbEvents(1 To .RecordCount) As UDTEvents
    .MoveFirst
    Do
        dbEvents(i).lEventID = !EventID
        dbEvents(i).lIsComplete = !IsComplete
        dbEvents(i).lPlayerID = !PlayerID
        dbEvents(i).sEndTime = !EndTime
        dbEvents(i).sStartTime = !StartTime
        dbEvents(i).sExpire = !Expire
        dbEvents(i).sCustomID = !CustomID
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
If Not UpdateLog Then UpdateList "Events Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Familiars) [27%] ..."
i = 1
With MRSFAMILIARS
    .MoveLast
    ReDim dbFamiliars(1 To .RecordCount) As UDTFamiliars
    .MoveFirst
    Do
        dbFamiliars(i).iID = !id
        dbFamiliars(i).sDescription = !Description
        dbFamiliars(i).sFlags = !Flags
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
If Not UpdateLog Then UpdateList "Familiars Loaded... }b(}n}i" & Time & "}n}b)"
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
        dbItems(i).iID = !id
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
        dbItems(i).sFlags = !Flags
        dbItems(i).sFlags2 = !Flags2
        dbItems(i).lOnLastUseDoFlags2 = !OnLastUseDoFlags2
        dbItems(i).sProjectile = !Projectile
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
If Not UpdateLog Then UpdateList "Items Loaded... }b(}n}i" & Time & "}n}b)"
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
        If InStr(LCaseFast(dbMap(i).sScript), "mybase.timer(") <> 0 Then
            ReDim Preserve dbMBTimer(UBound(dbMBTimer) + 1)
            sScripting 0, dbMap(i).lRoomID, , , , True, dbMBTimer(UBound(dbMBTimer)).lInterval, , dbMBTimer(UBound(dbMBTimer)).sScript
            With dbMBTimer(UBound(dbMBTimer))
                .lRoomID = dbMap(i).lRoomID
                
'                Debug.Print "------------------------"
'                Debug.Print "INDEX NUMBER : " & UBound(dbMBTimer)
'                Debug.Print "INTERVAL     : " & .lInterval
'                Debug.Print "SCRIPT       : " & .sScript
'                Debug.Print "ROOM ID      : " & .lRoomID
'                Debug.Print "------------------------"
            End With
        End If
        If InStr(LCaseFast(dbMap(i).sScript), "begin.usescript ") <> 0 Then
            j = 0
            s = ""
            sScripting 0, dbMap(i).lRoomID, , , , True, j, , s
            If j <> 0 Then
                ReDim Preserve dbMBTimer(UBound(dbMBTimer) + 1)
                With dbMBTimer(UBound(dbMBTimer))
                    .lRoomID = dbMap(i).lRoomID
                    .lInterval = j
                    .sScript = s
    '                Debug.Print "------------------------"
    '                Debug.Print "INDEX NUMBER : " & UBound(dbMBTimer)
    '                Debug.Print "INTERVAL     : " & .lInterval
    '                Debug.Print "SCRIPT       : " & .sScript
    '                Debug.Print "ROOM ID      : " & .lRoomID
    '                Debug.Print "------------------------"
                End With
            End If
        End If
        'dbMap(i).iSafeRoom = !Flags '
        'dbMap(i).lDeathRoom = !DeathRoom '
        'dbMap(i).iInDoor = !InDoor '
        'dbMap(i).iTrainClass = !TrainClass '
        dbMap(i).sMapFlags = !Flags
        'dbMap(i).sMapFlags = dbMap(i).sMapFlags & "/0;"
        'modMapFlags.UpdateMapFlags i
        modMapFlags.LoadMapFlags i
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
If Not UpdateLog Then UpdateList "Map Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Monsters) [36%] ..."
i = 1
With MRSMONSTER
    .MoveLast
    ReDim dbMonsters(1 To .RecordCount) As UDTMonsters
    .MoveFirst
    Do
    
        dbMonsters(i).dEXP = !EXP
        dbMonsters(i).dHP = !HP
        dbMonsters(i).dMoney = !Money
        dbMonsters(i).iAC = !AC
        dbMonsters(i).iAttackable = !Attackable
        dbMonsters(i).iHostile = !Hostile
        dbMonsters(i).iType = !Type
        dbMonsters(i).sDropItem = !DropItem
        dbMonsters(i).lID = !id
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
If Not UpdateLog Then UpdateList "Monsters Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Players) [39%] ..."
i = 1
With MRS
    .MoveLast
    ReDim dbPlayers(1 To .RecordCount) As UDTPlayers
    .MoveFirst
    Do
  
        dbPlayers(i).dBank = !Bank
        dbPlayers(i).dEXP = !EXP
        dbPlayers(i).dEXPNeeded = !EXPneeded
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
        dbPlayers(i).lMana = !MANA
        dbPlayers(i).lMaxHP = !MaxHP
        dbPlayers(i).sStatline = !Statline
        dbPlayers(i).lMaxMana = !MaxMana
        dbPlayers(i).lPlayerID = !PlayerID
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
If Not UpdateLog Then UpdateList "Players Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Races) [42%] ..."
i = 1
With MRSRACE
    .MoveLast
    ReDim dbRaces(1 To .RecordCount) As UDTRaces
    .MoveFirst
    Do
  
        dbRaces(i).dEXP = !EXP
        dbRaces(i).iID = !id
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
If Not UpdateLog Then UpdateList "Races Loaded... }b(}n}i" & Time & "}n}b)"
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
        dbSpells(i).lID = !id
        dbSpells(i).lMana = !MANA
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
        dbSpells(i).sFlags = !Flags
        dbSpells(i).sMessage2 = !Message2
        dbSpells(i).sMessageV = !MessageV
        dbSpells(i).sEndCastFlags = !EndCastFlags
        i = i + 1
        .MoveNext
    Loop Until .EOF
End With
If Not UpdateLog Then UpdateList "Spells Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Loading Shops) [48%] ..."
i = 1
With MRSSHOPS
    .MoveLast
    ReDim dbShops(1 To .RecordCount) As UDTShops
    .MoveFirst
    Do
        dbShops(i).iID = !id
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
If Not UpdateLog Then UpdateList "Shops Loaded... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Storing Arenas And Boss rooms) [51%] ..."
i = 1
j = 0
For i = 1 To UBound(dbMap)
    With dbMap(i)
        Select Case .iType
            Case 1 To 5
                j = j + 1
        End Select
    End With
Next
i = 1
ReDim dbArenas(1 To j)
j = 1
For i = 1 To UBound(dbMap)
    With dbMap(i)
        Select Case .iType
            Case 1 To 5
                dbArenas(j) = dbMap(i)
                dbArenas(j).ldbMapID = i
                j = j + 1
        End Select
    End With
Next
If Not UpdateLog Then UpdateList "Arenas And Boss Rooms Stored... }b(}n}i" & Time & "}n}b)"
If frmSplash.Visible = True Then frmSplash.lblPer = " Loading (Storing Door Rooms) [54%] ..."
i = 1
j = 0
For i = 1 To UBound(dbMap)
    With dbMap(i)
        If .lDD <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDU <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDN <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDS <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDE <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDW <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDNW <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDSW <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDNE <> 0 Then j = j + 1: GoTo nNextRecord
        If .lDSE <> 0 Then j = j + 1: GoTo nNextRecord
nNextRecord:
    End With
Next
i = 1
ReDim dbDoor(1 To j)
j = 1
For i = 1 To UBound(dbMap)
    With dbMap(i)
        If .lDD <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDU <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDN <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDS <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDE <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDW <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDNW <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDSW <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDNE <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
        If .lDSE <> 0 Then dbDoor(j) = dbMap(i): dbDoor(j).ldbDoorsMapID = i: j = j + 1: GoTo nNextRecord2
nNextRecord2:
    End With
Next
If Not UpdateLog Then UpdateList "Rooms With Doors Stored... }b(}n}i" & Time & "}n}b)"
bUpdate = False
i = 1
If Not UpdateLog Then UpdateList "Database succesfully loaded to memory... }b(}n}i" & Time & "}n}b)"
On Error GoTo 0
Exit Function
LoadDatabaseIntoMemory_Error:
'MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: LoadDatabaseIntoMemory in Module, modMain"
bUpdate = False
UpdateList "}b}uERROR ON LOAD }n}b" & Err.Number & "}n}i " & Err.Description, True
End Function

Sub ReloadPlayersOnly()
Dim i As Long
Dim j As Long
Dim Arr() As String
Set MRS = db.OpenRecordset("SELECT * FROM Players")
i = 1
With MRS
    .MoveLast
    ReDim dbPlayers(1 To .RecordCount) As UDTPlayers
    .MoveFirst
    Do
  
        dbPlayers(i).dBank = !Bank
        dbPlayers(i).dEXP = !EXP
        dbPlayers(i).dEXPNeeded = !EXPneeded
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
        dbPlayers(i).sTrainStats = !TrainStats
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
        dbPlayers(i).iStr = !Str
        dbPlayers(i).iStun = !Stun
        dbPlayers(i).sMiscFlag = !Misc
        dbPlayers(i).iWeapons = !Weapons
        dbPlayers(i).lBackUpLoc = !BackUpLoc
        dbPlayers(i).lHP = !HP
        dbPlayers(i).lLocation = !Location
        dbPlayers(i).lDBLocation = GetMapIndex(dbPlayers(i).lLocation)
        dbPlayers(i).lMana = !MANA
        dbPlayers(i).lMaxHP = !MaxHP
        dbPlayers(i).sStatline = !Statline
        dbPlayers(i).lMaxMana = !MaxMana
        dbPlayers(i).lPlayerID = !PlayerID
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
        dbPlayers(i).sFamFlags = !FamFlags
        modFamiliars.LoadFamFlags i
        dbPlayers(i).sKillDurItems = !KillDurItems
        dbPlayers(i).iIndex = !Index
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
Set MRS = Nothing
End Sub

Public Sub mRemoveItem(lngIndexToRemove As Long)
'removes a monster from the type array
'If lngIndexToRemove > UBound(aMons) - 1 Then
'    ReDim Preserve aMons(UBound(aMons) - 1)
'    Exit Sub
'End If
''On Error GoTo Err10
''ReDim Preserve aMons(UBound(aMons))
'For lngIndex = lngIndexToRemove To UBound(aMons) - 1
'    aMons(lngIndex) = aMons(lngIndex + 1)
'    If DE Then DoEvents
'Next lngIndex
'ReDim Preserve aMons(UBound(aMons) - 1)
With dbMap(aMons(lngIndexToRemove).mdbMapID)
    .sAMonIds = ReplaceFast(.sAMonIds, lngIndexToRemove & ";", "")
    .sMonsters = ReplaceFast(.sMonsters, ":" & aMons(lngIndexToRemove).miID & ";", "", 1, 1)
    If modSC.FastStringComp(.sMonsters, "") Then .sMonsters = "0"
End With
aMons(lngIndexToRemove).mLoc = -1
'Exit Sub
'Err10:
'If Err.Number = 10 Then sIndex = sIndex & CStr(lngIndexToRemove) & ";"
End Sub

Public Sub RemoveItemFromArray(lngIndexToRemove As Long)
'removes a monster from the type array
On Error GoTo eh1
Dim lngIndex As Long
If lngIndexToRemove > UBound(aMons) - 1 Then
    ReDim Preserve aMons(UBound(aMons) - 1)
    Exit Sub
End If
For lngIndex = lngIndexToRemove To UBound(aMons) - 1
    With aMons(lngIndex + 1)
        With dbMap(.mdbMapID)
            .sAMonIds = ReplaceFast$(.sAMonIds, CStr(lngIndex + 1) & ";", "")
        End With
    End With
    aMons(lngIndex) = aMons(lngIndex + 1)
    With aMons(lngIndex)
        With dbMap(.mdbMapID)
            .sAMonIds = .sAMonIds & CStr(lngIndex) & ";"
        End With
    End With
    If DE Then DoEvents
Next lngIndex
ReDim Preserve aMons(UBound(aMons) - 1)
eh1:
End Sub

Sub CheckSpecialItems()
Dim sItemsHere$
UpdateList "Checking Special Items }b(}n}i" & Time & "}n}b)"
With MRSMAP
    .MoveFirst
    Do
        If !SpecialItem <> "0" Then
            sItemsHere$ = !Items
            If modSC.FastStringComp(sItemsHere$, "0") Then sItemsHere$ = ""
            If InStr(1, sItemsHere$, ":" & !SpecialItem & "/") = 0 Then
                sItemsHere$ = sItemsHere$ & ":" & !SpecialItem & "/99/E{}F{}A{}B{0|0|0|0}/1;"
                .Edit
                !Items = sItemsHere$
                .Update
            End If
            .MoveNext
        ElseIf Not .EOF Then
            .MoveNext
        End If
        If DE Then DoEvents
    Loop Until .EOF
End With
UpdateList "Special Item Check Complete }b(}n}i" & Time & "}n}b)"
End Sub

Sub RemoveCorpses()
Dim sItemsHere$, tArr$(), i&, j&
UpdateList "Removing Monster Corpses }b(}n}i" & Time & "}n}b)"
j& = 0
With MRSMAP
    .MoveFirst
    Do
        sItemsHere$ = !Items
        If sItemsHere$ <> "0" Then
            SplitFast sItemsHere, tArr$(), ";"
            For i& = LBound(tArr$()) To UBound(tArr$())
                With MRSITEM
                    If tArr$(i&) <> "" Then
                        .MoveFirst
                        Do
                            If !id = modItemManip.GetItemIDFromUnFormattedString(tArr$(i&)) Then
                                If !Worn = "corpse" Then
                                    tArr$(i&) = ""
                                    j& = j& + 1
                                End If
                            End If
                            .MoveNext
                            If DE Then DoEvents
                        Loop Until .EOF
                    End If
                End With
                If DE Then DoEvents
            Next
            sItemsHere$ = ""
            For i& = LBound(tArr$()) To UBound(tArr$())
                If tArr$(i&) <> "" Then
                    sItemsHere$ = sItemsHere$ & tArr$(i&) & ";"
                End If
                If DE Then DoEvents
            Next
            If sItemsHere$ = "" Then sItemsHere$ = "0"
            .Edit
            !Items = sItemsHere$
            .Update
        End If
        .MoveNext
    If DE Then DoEvents
    Loop Until .EOF
End With
UpdateList CStr(j&) & " Monster Corpses Removed }b(}n}i" & Time & "}n}b)"
End Sub

Sub RemoveOutDoorFood()
Dim sItemsHere$, tArr$(), i&, j&
UpdateList "Removing Outdoor Food }b(}n}i" & Time & "}n}b)"
j& = 0
With MRSMAP
    .MoveFirst
    Do
        sItemsHere$ = !Items
        If sItemsHere$ <> "0" Then
            SplitFast sItemsHere, tArr$(), ";"
            For i& = LBound(tArr$()) To UBound(tArr$())
                With MRSITEM
                    If tArr$(i&) <> "" Then
                        .MoveFirst
                        Do
                            If !id = modItemManip.GetItemIDFromUnFormattedString(tArr$(i&)) Then
                                If !Worn = "ofood" Then
                                    tArr$(i&) = ""
                                    j& = j& + 1
                                End If
                            End If
                            .MoveNext
                            If DE Then DoEvents
                        Loop Until .EOF
                    End If
                End With
                If DE Then DoEvents
            Next
            sItemsHere$ = ""
            For i& = LBound(tArr$()) To UBound(tArr$())
                If tArr$(i&) <> "" Then
                    sItemsHere$ = sItemsHere$ & tArr$(i&) & ";"
                End If
                If DE Then DoEvents
            Next
            If sItemsHere$ = "" Then sItemsHere$ = "0"
            .Edit
            !Items = sItemsHere$
            .Update
        End If
        .MoveNext
    If DE Then DoEvents
    Loop Until .EOF
End With
UpdateList CStr(j&) & " Pieces Of Outdoor Food Removed }b(}n}i" & Time & "}n}b)"
End Sub

Sub RemoveNormalFood()
Dim sItemsHere$, tArr$(), i&, j&
UpdateList "Removing Normal Food }b(}n}i" & Time & "}n}b)"
j& = 0
With MRSMAP
    .MoveFirst
    Do
        sItemsHere$ = !Items
        If sItemsHere$ <> "0" Then
            SplitFast sItemsHere, tArr$(), ";"
            For i& = LBound(tArr$()) To UBound(tArr$())
                With MRSITEM
                    If tArr$(i&) <> "" Then
                        .MoveFirst
                        Do
                            If !id = modItemManip.GetItemIDFromUnFormattedString(tArr$(i&)) Then
                                If !Worn = "food" Then
                                    tArr$(i&) = ""
                                    j& = j& + 1
                                End If
                            End If
                            .MoveNext
                            If DE Then DoEvents
                        Loop Until .EOF
                    End If
                End With
                If DE Then DoEvents
            Next
            sItemsHere$ = ""
            For i& = LBound(tArr$()) To UBound(tArr$())
                If tArr$(i&) <> "" Then
                    sItemsHere$ = sItemsHere$ & tArr$(i&) & ";"
                End If
                If DE Then DoEvents
            Next
            If sItemsHere$ = "" Then sItemsHere$ = "0"
            .Edit
            !Items = sItemsHere$
            .Update
        End If
        .MoveNext
    If DE Then DoEvents
    Loop Until .EOF
End With
UpdateList CStr(j&) & " Pieces Of Normal Food Removed }b(}n}i" & Time & "}n}b)"
End Sub


