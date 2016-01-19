Attribute VB_Name = "modUseItems"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modUseItems
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function UseItem(Index As Long, Optional dbIndex As Long = 0) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 4)), "use ") Then
    UseItem = True
    Dim dbItemID As Long
    Dim s As String
    Dim sUF As String
    Dim b As Boolean
    Dim Arr() As String
    Dim lRoll As Long
    
    s = X(Index)
    X(Index) = ""
    s = Mid$(s, InStr(1, s, " ") + 1, Len(s) - InStr(1, s, " "))
    s = SmartFind(Index, s, Inventory_Item, True, sUF)
    If InStr(1, s, Chr$(0)) > 0 Then s = Mid$(s, InStr(1, s, Chr$(0)) + 1)
    
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    dbItemID = GetItemID(s)
    
    If dbItemID = 0 Or sUF = "" Then
        WrapAndSend Index, RED & "You can't seem to find that in your inventory." & WHITE & vbCrLf
        Exit Function
    End If
    
    With dbItems(dbItemID)
        If .lLevel <= dbPlayers(dbIndex).iLevel And _
           .dClassPoints <= dbPlayers(dbIndex).dClassPoints And _
           (modSC.FastStringComp(.sWorn, "item") Or modSC.FastStringComp(.sWorn, "scroll")) And _
           ClassCanuseMagical(dbPlayers(dbIndex).sClass, .iMagical) And _
           ClassCanWear(.iID, dbPlayers(dbIndex).sClass, dbItemID) And _
           RaceCanWear(.iID, dbPlayers(dbIndex).sRace, dbItemID) _
           Then
                Select Case .sWorn
                    Case "item"
                        If Not modSC.FastStringComp(.sDamage, "0:0") Then
                            SplitFast .sDamage, Arr, ":"
                            lRoll = RndNumber(CDbl(Arr(0)), CDbl(Arr(1)))
                        End If
                        modUseItems.DoFlags dbIndex, .sFlags, , lRoll
                    Case "scroll"
                        modUseItems.DoFlags dbIndex, .sFlags, , , , True, b
                        If Not b Then
                            WrapAndSend Index, RED & "Something is stoping you from using this item." & WHITE & vbCrLf
                            Exit Function
                        End If
                End Select
                If .sScript <> "0" Then sScripting Index, UseThisScript:=.sScript
                modItemManip.SubtractOneFromItemUseINV dbIndex, _
                                                       dbItemID, _
                                                       modItemManip.GetItemUsesFromUnFormattedString(sUF), _
                                                       modItemManip.GetItemDurFromUnFormattedString(sUF)
                s = .sSwings
                s = ReplaceFast(s, "<%n>", .sItemName)
                s = ReplaceFast(s, "<%d>", CStr(lRoll))
                WrapAndSend Index, LIGHTBLUE & s & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " uses " & .sItemName & "." & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
                
                
                
        Else
            WrapAndSend Index, RED & "Something is stoping you from using this item." & WHITE & vbCrLf
            Exit Function
        End If
    End With
End If
End Function

Public Sub DoFlags(dbIndex As Long, sFlags As String, Optional Delmin As String = ";", Optional lRoll As Long = 0, Optional Inverse As Boolean = False, Optional AllowSpell As Boolean = False, Optional ByRef bLearnedSpell As Boolean, Optional UseMessages As Boolean = False, Optional ByRef Message As String, Optional ByRef Message2 As String, Optional ByRef Send2InThisRoom As Long, Optional ByRef Message3 As String, Optional ByRef Send3InThisRoom As Long)
Dim i As Long
Dim j As Long
Dim Arr() As String
Dim Arr2() As String
Dim Arr3() As String
Dim dVal As Double
Dim sVal As String
Dim dbID As Long
Dim s As String
If sFlags <> "" Then
    SplitFast sFlags, Arr, Delmin
    If UseMessages Then Send2InThisRoom = dbPlayers(dbIndex).lLocation
    For i = LBound(Arr) To UBound(Arr)
        If Not modSC.FastStringComp(Arr(i), "") Then
            sVal = Mid$(Arr(i), 4)
            If IsNumeric(sVal) Then
                dVal = CDbl(Val(sVal))
                If sVal = "-0" Then dVal = lRoll
                If Inverse Then dVal = -dVal
            End If
            Select Case Left$(Arr(i), 3)
                Case "lig"
                    With dbPlayers(dbIndex)
                        .iVision = .iVision + dVal
                    End With
                Case "acl"
                    With dbPlayers(dbIndex)
                        .iAC = .iAC + dVal
                    End With
                Case "cri"
                    With dbPlayers(dbIndex)
                        .iCrits = .iCrits + dVal
                    End With
                Case "acc"
                    With dbPlayers(dbIndex)
                        .iAcc = .iAcc + dVal
                     End With
                Case "dam"
                    With dbPlayers(dbIndex)
                        .iMaxDamage = .iMaxDamage + dVal
                    End With
                Case "str"
                    With dbPlayers(dbIndex)
                        .iStr = .iStr + dVal
                    End With
                Case "agi"
                    With dbPlayers(dbIndex)
                        .iAgil = .iAgil + dVal
                    End With
                Case "cha"
                    With dbPlayers(dbIndex)
                        .iCha = .iCha + dVal
                    End With
                Case "dex"
                    With dbPlayers(dbIndex)
                        .iDex = .iDex + dVal
                    End With
                Case "int"
                    With dbPlayers(dbIndex)
                        .iInt = .iInt + dVal
                    End With
                Case "chp"
                    With dbPlayers(dbIndex)
                        .lHP = .lHP + dVal
                        If .lHP > .lMaxHP Then .lHP = .lMaxHP
                    End With
                Case "mhp"
                    With dbPlayers(dbIndex)
                        .lMaxHP = .lMaxHP + dVal
                    End With
                Case "cma"
                    With dbPlayers(dbIndex)
                        .lMana = .lMana + dVal
                        If .lMana > .lMaxMana Then .lMana = .lMaxMana
                    End With
                Case "mma"
                    With dbPlayers(dbIndex)
                        .lMaxMana = .lMaxMana + dVal
                    End With
                Case "hun"
                    With dbPlayers(dbIndex)
                        .dHunger = .dHunger + dVal
                    End With
                Case "sta"
                    With dbPlayers(dbIndex)
                        .dStamina = .dStamina + dVal
                    End With
                Case "cac"
                    With dbPlayers(dbIndex)
                        .iAC = .iAC + dVal
                    End With
                Case "dod"
                    With dbPlayers(dbIndex)
                        .iDodge = .iDodge + dVal
                    End With
                Case "exp"
                    With dbPlayers(dbIndex)
                        .dEXP = .dEXP + dVal
                    End With
                Case "txp"
                    With dbPlayers(dbIndex)
                        .dTotalEXP = .dTotalEXP + dVal
                    End With
                Case "gol"
                    With dbPlayers(dbIndex)
                        .dGold = .dGold + dVal
                    End With
                Case "ban"
                    With dbPlayers(dbIndex)
                        .dBank = .dBank + dVal
                    End With
                Case "vis"
                    With dbPlayers(dbIndex)
                        .iVision = .iVision + dVal
                    End With
                Case "clp"
                    With dbPlayers(dbIndex)
                        .dClassPoints = .dClassPoints + dVal
                    End With
                Case "ccp"
                    With dbPlayers(dbIndex)
                        .iIsReadyToTrain = .iIsReadyToTrain + dVal
                    End With
                Case "mit"
                    modMiscFlag.SetStatsPlus dbIndex, [Max Items Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Max Items Bonus]) + CLng(dVal)
                Case "evi"
                    With dbPlayers(dbIndex)
                        .iEvil = .iEvil + dVal
                    End With
                Case "rms"
                    With dbPlayers(dbIndex)
                        SplitFast Left$(.sBlessSpells, Len(.sBlessSpells) - 1), Arr2, "Œ"
                        For j = LBound(Arr2) To UBound(Arr2)
                            Erase Arr3
                            SplitFast Arr2(j), Arr3, "~"
                            If modSC.FastStringComp(CStr(dbSpells(Val(Arr3(3))).lID), CStr(dVal)) Then
                                If Not modSC.FastStringComp(dbSpells(Val(Arr3(3))).sFlags, "0") Then
                                    DoFlags dbIndex, dbSpells(Val(Arr3(3))).sFlags, lRoll:=CLng(Val(Arr3(1))), Inverse:=True
                                End If
                                .sBlessSpells = ReplaceFast(.sBlessSpells, Arr3(0) & "~" & Arr3(1) & "~" & Arr3(2) & "~" & Arr3(3) & "Œ", "", 1, 1)
                                If modSC.FastStringComp(.sBlessSpells, "") Then .sBlessSpells = "0"
                                sSend .iIndex, LIGHTBLUE & dbSpells(Val(Arr3(3))).sRunOutMessage
                            End If
                            If DE Then DoEvents
                        Next
                    End With
                Case "pap"
                    With dbPlayers(dbIndex)
                        .lPaper = .lPaper + dVal
                        If .lPaper < 0 Then .lPaper = 0
                    End With
                Case "mat" 'MakeItem
                    dbID = 0
                    dbID = GetItemID(, CLng(dVal))
                    If dbID = 0 Then GoTo nNext
                    With dbPlayers(dbIndex)
                        If modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) + 1 < modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) Then
                            If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                            .sInventory = .sInventory & ":" & dbItems(dbID).iID & "/" & dbItems(dbID).lDurability & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(dbID).iUses & ";"
                        Else
                            With dbMap(.lDBLocation)
                                If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
                                .sItems = .sItems & ":" & dbItems(dbID).iID & "/" & dbItems(dbID).lDurability & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(dbID).iUses & ";"
                            End With
                        End If
                    End With
                Case "snd"
                    If Not Inverse Then sSend dbPlayers(dbIndex).iIndex, sVal
                Case "sro"
                    If Not Inverse Then sSendRoom dbPlayers(dbIndex).iIndex, sVal
                Case "gsp"
                    If Not Inverse And AllowSpell Then
                        dbID = 0
                        dbID = GetSpellID(, CLng(dVal))
                        If dbID = 0 Then bLearnedSpell = False: GoTo nNext
                        If dbPlayers(dbIndex).iSpellLevel >= dbSpells(dbID).iLevel And _
                           dbPlayers(dbIndex).iSpellType = dbSpells(dbID).iType And _
                           InStr(1, dbPlayers(dbIndex).sSpells, ":" & dbSpells(dbID).lID & ";") = 0 _
                           Then
                                
                                If modSC.FastStringComp(dbPlayers(dbIndex).sSpells, "0") Then dbPlayers(dbIndex).sSpells = ""
                                dbPlayers(dbIndex).sSpells = dbPlayers(dbIndex).sSpells & ":" & dbSpells(dbID).lID & ";"
                                If modSC.FastStringComp(dbPlayers(dbIndex).sSpellShorts, "0") Then dbPlayers(dbIndex).sSpellShorts = ""
                                dbPlayers(dbIndex).sSpellShorts = dbPlayers(dbIndex).sSpellShorts & dbSpells(dbID).sShort & ";"
                                bLearnedSpell = True
                        End If
                    End If
                Case "gfa"
                    'dbID = CLng(dVal)
                    dbID = GetFamID(CLng(dVal))
                    If dbID = 0 Then GoTo nNext
                    RemoveStats dbPlayers(dbIndex).iIndex
                    With dbPlayers(dbIndex)
                        .sFamName = dbFamiliars(dbID).sFamName
                        .lFamID = CLng(dVal)
                        .lFamMHP = RndNumber(CDbl(dbFamiliars(dbID).lStartHPMin), CDbl(dbFamiliars(dbID).lStartHPMax))
                        .lFamAcc = 0
                        .lFamLevel = 1
                        .dFamCEXP = 0
                        .dFamEXPN = dbFamiliars(dbID).dEXPPerLevel
                        .dFamTEXP = 0
                        .lFamCHP = .lFamMHP
                        .lFamMin = dbFamiliars(dbID).lMinDam
                        .lFamMax = dbFamiliars(dbID).lMaxDam
                    End With
                    AddStats dbPlayers(dbIndex).iIndex
                Case "sas"
                    If Inverse Then
                        With dbPlayers(dbIndex)
                            .sPlayerName = .sSeenAs
                        End With
                    Else
                        With dbPlayers(dbIndex)
                            .sPlayerName = sVal
                        End With
                    End If
                Case "des"
                    If Inverse Then
                        With dbPlayers(dbIndex)
                            .sOverrideDesc = "0"
                        End With
                    Else
                        With dbPlayers(dbIndex)
                            .sOverrideDesc = sVal
                        End With
                    End If
                Case "csp"
                    dbPlayers(dbIndex).lHasCasted = 0
                    modSpells.DoNonCombatSpell dbPlayers(dbIndex).iIndex, dbIndex, Abs(dVal)
                Case "thi"
                    modMiscFlag.SetStatsPlus dbIndex, [Thieving Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Thieving Bonus]) + CLng(dVal)
                Case "stu"
                    If Not Inverse Then
                        With dbPlayers(dbIndex)
                            .iStun = .iStun + dVal
                            If UseMessages Then
                                Message = Message & BRIGHTYELLOW & "You are stunned!" & WHITE & vbCrLf
                                Message2 = Message2 & YELLOW & .sPlayerName & " is stunned!" & WHITE & vbCrLf
                            End If
                        End With
                    End If
                Case "tel"
                    If Not Inverse Then
                        Select Case dVal
                            Case "-1"
                                s = modGetData.sGetRoomExits(dbPlayers(dbIndex).iIndex, dbIndex)
                                If Not modSC.FastStringComp(s, "") Then
                                    Erase Arr2
                                    SplitFast s, Arr2, ","
                                Else
                                    If UseMessages Then Message = Message & BRIGHTBLUE & "You fail to teleport!" & WHITE & vbCrLf
                                    GoTo nNext
                                End If
                                dbID = CLng(RndNumber(LBound(Arr2), UBound(Arr2)))
                                dbID = CLng(Val(Arr2(dbID)))
                            Case "-2"
                                dbID = dbMap(RndNumber(LBound(dbMap), UBound(dbMap))).lRoomID
                            Case Else
                                If dVal <= 0 Then
                                    If UseMessages Then Message = Message & BRIGHTBLUE & "You fail to teleport!" & WHITE & vbCrLf
                                    GoTo nNext
                                End If
                                dbID = CLng(dVal)
                        End Select
                        dbPlayers(dbIndex).lLocation = dbID
                        dbPlayers(dbIndex).lDBLocation = GetMapIndex(dbID)
                        If UseMessages Then
                            Message3 = Message3 & BLUE & dbPlayers(dbIndex).sPlayerName & " appears in the room!" & WHITE & vbCrLf
                            Send3InThisRoom = lFlg
                        End If
                    End If
            End Select
            If Left$(Arr(i), 3) Like "el#" Then
                modResist.UpdateResistValue dbIndex, CLng(Val(Mid$(Arr(i), 3, 1))), CLng(dVal)
            End If
            If Left$(Arr(i), 3) Like "m##" Then
                If Inverse Then
                    Select Case dVal
                        Case -1
                            dVal = 0
                        Case 0
                            dVal = 1
                        Case Else
                            dVal = 0
                    End Select
                End If
                modMiscFlag.SetMiscFlag dbIndex, CLng(Val(Mid$(Arr(i), 2, 2))), CLng(dVal)
            End If
            If Left(Arr(i), 3) Like "s##" Then
                Select Case Val(Mid$(Arr(i), 2, 2))
                    Case 1, 3, 5, 9, 11, 13
                        modMiscFlag.SetStatsPlus dbIndex, CLng(Val(Mid$(Arr(i), 2, 2))), modMiscFlag.GetStatsPlus(dbIndex, CLng(Val(Mid$(Arr(i), 2, 2)))) + CLng(dVal)
                End Select
            End If
        End If
nNext:
        If DE Then DoEvents
    Next
End If
End Sub

'Public Sub DoItemFlags(dbIndex As Long, dbItemID As Long, lRoll As Long, Optional ByRef WasUsed As Long, Optional Inverse As Boolean = False, Optional Flags2 As Boolean = False, Optional AllowSpell As Boolean = True, Optional ThisISNotAnItem As Boolean = False, Optional FlagsAsString As String = "")
''tel# 'teleport, -1,-2, room num
''stu# 'stun
''lig# 'light
''cri# 'crits
''Acc# 'acc
''dam# 'damage
''Str# 'strengh
''agi# 'agility
''cha# 'charm
''dex# 'dexterity
''int# 'intelect
''chp# 'current hp
''mHP# 'max hp
''cma# 'current mana
''mma# 'max mana
''hun# 'hunger
''sta# 'stamina
''cac# 'current AC
''dod# 'dodge
''vis# 'vision
''mit# 'max items
''evi# 'evil points
''pap# 'paper
''mat# 'make item
''SndABC 'Send a message
''sRoABC 'Send a message to the room
''gsp# 'Give spell
''gfa# 'give familiar
''el0
''el1
''el2
''el3
''el4
''el5
''el6
''el7
''el8
''m01-m19
'Dim i As Long
'Dim sVal As String
'Dim dVal As Double
'Dim aFlgs() As String
'Dim iSpellID As Long
'Dim FamID As Long
'Dim dbFamId As Long
'If Not Flags2 And Not ThisISNotAnItem And dbItemID <> 0 Then
'    SplitFast dbItems(dbItemID).sFlags, aFlgs, ";"
'    If InStr(1, dbItems(dbItemID).sFlags, "gsp") <> 0 And Not Inverse And AllowSpell And modSC.FastStringComp(dbItems(dbItemID).sWorn, "scroll") Then
'        For i = LBound(aFlgs) To UBound(aFlgs)
'            dVal = CDbl(Val(Mid$(aFlgs(i), 4)))
'            Select Case Left$(aFlgs(i), 3)
'                Case "gsp"
'                    iSpellID = GetSpellID(, CLng(dVal))
'                    If iSpellID = 0 Then GoTo tNext
'                    If dbPlayers(dbIndex).iSpellLevel >= dbSpells(iSpellID).iLevel Then
'                        If dbPlayers(dbIndex).iSpellType = dbSpells(iSpellID).iType Then
'                            If InStr(1, dbPlayers(dbIndex).sSpells, ":" & dbSpells(iSpellID).lID & ";") Then
'
'                                WasUsed = 1
'                            End If
'
'                        Else
'
'                            WasUsed = 1
'                        End If
'                    Else
'
'                        WasUsed = 1
'                    End If
'                Case Else
'
'            End Select
'tNext:
'            If DE Then DoEvents
'        Next
'    End If
'    If WasUsed = 1 Then Exit Sub
'ElseIf Not ThisISNotAnItem And dbItemID <> 0 Then
'    SplitFast dbItems(dbItemID).sFlags2, aFlgs, ";"
'Else
'    If InStr(1, FlagsAsString, ";") > 0 Then
'        SplitFast FlagsAsString, aFlgs, ";"
'    Else
'        SplitFast FlagsAsString, aFlgs, "|"
'    End If
'End If
'For i = LBound(aFlgs) To UBound(aFlgs)
'    If Not modSC.FastStringComp(aFlgs(i), "") Then
'        sVal = Mid$(aFlgs(i), 4)
'        If IsNumeric(sVal) Then
'            dVal = CDbl(Val(sVal))
'            If dVal = -3 Then dVal = lRoll
'            If Inverse Then dVal = -dVal
'        End If
'        Select Case Left$(aFlgs(i), 3)
'            Case "lig"
'                With dbPlayers(dbIndex)
'                    .iVision = .iVision + dVal
'
'                End With
'            Case "acl"
'                With dbPlayers(dbIndex)
'                    .iAC = .iAC + dVal
'
'                End With
'            Case "cri"
'                With dbPlayers(dbIndex)
'                    .iCrits = .iCrits + dVal
'
'                End With
'            Case "acc"
'                With dbPlayers(dbIndex)
'                    .iAcc = .iAcc + dVal
'
'                End With
'            Case "dam"
'                With dbPlayers(dbIndex)
'                    .iMaxDamage = .iMaxDamage + dVal
'
'                End With
'            Case "str"
'                With dbPlayers(dbIndex)
'                    .iStr = .iStr + dVal
'
'                End With
'            Case "agi"
'                With dbPlayers(dbIndex)
'                    .iAgil = .iAgil + dVal
'
'                End With
'            Case "cha"
'                With dbPlayers(dbIndex)
'                    .iCha = .iCha + dVal
'
'                End With
'            Case "dex"
'                With dbPlayers(dbIndex)
'                    .iDex = .iDex + dVal
'
'                End With
'            Case "int"
'                With dbPlayers(dbIndex)
'                    .iInt = .iInt + dVal
'
'                End With
'            Case "chp"
'                With dbPlayers(dbIndex)
'                    .lHP = .lHP + dVal
'                    If .lHP > .lMaxHP Then .lHP = .lMaxHP
'
'                End With
'            Case "mhp"
'                With dbPlayers(dbIndex)
'                    .lMaxHP = .lMaxHP + dVal
'
'                End With
'            Case "cma"
'                With dbPlayers(dbIndex)
'                    .lMana = .lMana + dVal
'                    If .lMana > .lMaxMana Then .lMana = .lMaxMana
'
'                End With
'            Case "mma"
'                With dbPlayers(dbIndex)
'                    .lMaxMana = .lMaxMana + dVal
'
'                End With
'            Case "hun"
'                With dbPlayers(dbIndex)
'                    .dHunger = .dHunger + dVal
'
'                End With
'            Case "sta"
'                With dbPlayers(dbIndex)
'                    .dStamina = .dStamina + dVal
'
'                End With
'            Case "cac"
'                With dbPlayers(dbIndex)
'                    .iAC = .iAC + dVal
'
'                End With
'            Case "dod"
'                With dbPlayers(dbIndex)
'                    .iDodge = .iDodge + dVal
'
'                End With
'            Case "vis"
'                With dbPlayers(dbIndex)
'                    .iVision = .iVision + dVal
'
'                End With
'            Case "mit"
'                modMiscFlag.SetStatsPlus dbIndex, [Max Items Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Max Items Bonus]) + CLng(dVal)
'            Case "evi"
'                With dbPlayers(dbIndex)
'                    .iEvil = .iEvil + dVal
'
'                End With
'            Case "pap"
'                With dbPlayers(dbIndex)
'
'                    .lPaper = .lPaper + dVal
'                    If .lPaper < 0 Then .lPaper = 0
'
'                End With
'            Case "mat"
'                dbFamId = GetItemID(, CLng(dVal))
'                If dbFamId = 0 Then GoTo nNext
'                With dbPlayers(dbIndex)
'                    If modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) + 1 < modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) Then
'                        If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
'                        .sInventory = .sInventory & ":" & dbItems(dbFamId).iID & "/" & dbItems(dbFamId).iUses & "/" & dbItems(dbFamId).lDurability & ";"
'                    Else
'                        With dbMap(GetMapIndex(.lLocation))
'                            If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
'                            .sItems = .sItems & ":" & dbItems(dbFamId).iID & "/" & dbItems(dbFamId).iUses & "/" & dbItems(dbFamId).lDurability & ";"
'                        End With
'                    End If
'                End With
'            Case "snd"
'                If Not Inverse Then sSend dbPlayers(dbIndex).iIndex, sVal
'            Case "sro"
'                If Not Inverse Then sSendRoom dbPlayers(dbIndex).iIndex, sVal
'            Case "gsp"
'                If Not Inverse And Not ThisISNotAnItem And Not Flags2 And AllowSpell And modSC.FastStringComp(dbItems(dbItemID).sWorn, "scroll") Then
'                    iSpellID = GetSpellID(, CLng(dVal))
'                    If iSpellID = 0 Then GoTo nNext
'                    If modSC.FastStringComp(dbPlayers(dbIndex).sSpells, "0") Then dbPlayers(dbIndex).sSpells = ""
'                    dbPlayers(dbIndex).sSpells = dbPlayers(dbIndex).sSpells & ":" & dbSpells(iSpellID).lID & ";"
'                    If modSC.FastStringComp(dbPlayers(dbIndex).sSpellShorts, "0") Then dbPlayers(dbIndex).sSpellShorts = ""
'                    dbPlayers(dbIndex).sSpellShorts = dbPlayers(dbIndex).sSpellShorts & dbSpells(iSpellID).sShort & ";"
'                End If
'            Case "gfa"
'                FamID = CLng(dVal)
'                dbFamId = GetFamID(FamID)
'                If dbFamId = 0 Then GoTo nNext
'                RemoveStats dbPlayers(dbIndex).iIndex
'                With dbPlayers(dbIndex)
'                    .sFamName = dbFamiliars(dbFamId).sFamName
'                    .lFamID = FamID
'                    .lFamMHP = RndNumber(CDbl(dbFamiliars(dbFamId).lStartHPMin), CDbl(dbFamiliars(dbFamId).lStartHPMax))
'                    .lFamAcc = 0
'                    .lFamLevel = 1
'                    .dFamCEXP = 0
'                    .dFamEXPN = dbFamiliars(dbFamId).dEXPPerLevel
'                    .dFamTEXP = 0
'                    .lFamCHP = .lFamMHP
'                    .lFamMin = dbFamiliars(dbFamId).lMinDam
'                    .lFamMax = dbFamiliars(dbFamId).lMaxDam
'                End With
'                AddStats dbPlayers(dbIndex).iIndex
'            Case "sas"
'                If Inverse Then
'                    With dbPlayers(dbIndex)
'                        .sPlayerName = .sSeenAs
'                    End With
'                Else
'                    With dbPlayers(dbIndex)
'                        .sPlayerName = sVal
'                    End With
'                End If
'            Case "des"
'                If Inverse Then
'                    With dbPlayers(dbIndex)
'                        .sOverrideDesc = "0"
'                    End With
'                Else
'                    With dbPlayers(dbIndex)
'                        .sOverrideDesc = sVal
'                    End With
'                End If
'            Case "thi"
'                modMiscFlag.SetStatsPlus dbIndex, [Thieving Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Thieving Bonus]) + CLng(dVal)
'        End Select
'        If Left$(aFlgs(i), 2) Like "el#" Then
'            If Inverse Then
'                Select Case dVal
'                    Case -1
'                        dVal = 0
'                    Case 0
'                        dVal = 1
'                    Case Else
'                        dVal = 0
'                End Select
'            End If
'            modResist.UpdateResistValue dbIndex, CLng(Val(Mid$(aFlgs(i), 3, 1))), CLng(dVal)
'        End If
'        If Left$(aFlgs(i), 3) Like "m##" Then
'            If Inverse Then
'                Select Case dVal
'                    Case -1
'                        dVal = 0
'                    Case 0
'                        dVal = 1
'                    Case Else
'                        dVal = 0
'                End Select
'            End If
'            modMiscFlag.SetMiscFlag dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2))), CLng(dVal)
'        End If
'        If Left(aFlgs(i), 3) Like "s##" Then
'            Select Case Val(Mid$(aFlgs(i), 2, 2))
'                Case 1, 3, 5, 9, 11, 13
'                    modMiscFlag.SetStatsPlus dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2))), modMiscFlag.GetStatsPlus(dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2)))) + CLng(dVal)
'            End Select
'        End If
'    End If
'nNext:
'    If DE Then DoEvents
'Next
'If Not ThisISNotAnItem Then
'    If Inverse And dbItemID <> 0 Then
'        dbPlayers(dbIndex).iAC = dbPlayers(dbIndex).iAC - dbItems(dbItemID).iAC
'    ElseIf dbItemID <> 0 Then
'        dbPlayers(dbIndex).iAC = dbPlayers(dbIndex).iAC + dbItems(dbItemID).iAC
'    End If
'End If
'End Sub
