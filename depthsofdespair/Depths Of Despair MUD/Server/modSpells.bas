Attribute VB_Name = "modSpells"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modSpells
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function ListSpells(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 6)), "spells") Then
    ListSpells = True
    Dim Spells$, ToSend$
    With dbPlayers(GetPlayerIndexNumber(Index))
        Spells$ = .sSpells
    End With
    If modSC.FastStringComp(Spells$, "0") Then
        WrapAndSend Index, RED & "You just remembered that you don't know any spells." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    Dim tArr() As String
    Spells$ = ReplaceFast(Spells$, ":", "")
    SplitFast Left$(Spells$, Len(Spells$) - 1), tArr, ";"
    Dim ShortLen&, NameLen&
    For i = 0 To UBound(tArr)
        With dbSpells(GetSpellID(, CLng(tArr(i))))
            ToSend$ = ToSend$ & YELLOW & .lMana & Space(10 - Len(CStr(.lMana))) & GREEN & .sShort & LIGHTBLUE & Space(20 - Len(.sShort)) & .sSpellName & vbCrLf
            If 8 - Len(CStr(.lMana)) > ShortLen& Then ShortLen& = 8 - Len(CStr(.lMana))
            If 21 - Len(.sShort) > NameLen& Then NameLen& = 21 - Len(.sShort)
        End With
        If DE Then DoEvents
    Next
    X(Index) = ""
    ToSend$ = MAGNETA & "You know the following spells:" & vbCrLf & "Cost" & Space(ShortLen& - 1) & "Short" & Space(NameLen& - 2) & "Name" & vbCrLf & ToSend$ & WHITE & vbCrLf
    WrapAndSend Index, ToSend$
End If
End Function

Sub DoNonCombatSpell(Index As Long, dbIndex As Long, dbSpellID As Long, Optional Target As Boolean = False, Optional dbTIndex As Long = 0, Optional NoPlayer As Boolean = False, Optional Caster As String = "Unknown", Optional EndCast As Boolean = False)
Dim aFlg() As String
Dim lRoll As Long
Dim Message1 As String
Dim Message2 As String
Dim Message3 As String
Dim Message4 As String
Dim Loc2 As Long
Dim Loc3 As Long
Dim bMe As Boolean
Dim Arr() As String
Dim i As Long
With dbSpells(dbSpellID)
    If Not NoPlayer Then
        If dbPlayers(dbIndex).lHasCasted = 1 Then
            WrapAndSend Index, BRIGHTBLUE & "You are still wore out from the last spell you casted." & WHITE & vbCrLf
            X(Index) = ""
            Exit Sub
        End If
        If dbPlayers(dbIndex).lMana < .lMana Then
            WrapAndSend Index, BRIGHTBLUE & "You don't have enough mana!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Sub
        End If
        BreakOffCombat dbIndex
        If dbPlayers(dbIndex).iSneaking <> 0 Then
            dbPlayers(dbIndex).iSneaking = 0
            WrapAndSend Index, BRIGHTRED & "You are no longer sneaking around." & WHITE, False
        End If
        dbPlayers(dbIndex).iResting = 0
        dbPlayers(dbIndex).iMeditating = 0
        dbPlayers(dbIndex).lHasCasted = 1
    End If
    If Target = False Then
        If (RndNumber(0, 100) < modGetData.GetSpellChanceFromdbSpell(dbIndex, dbSpellID)) Or NoPlayer Then
            If Not EndCast Then
                Message1 = LIGHTBLUE & .sMessage
                Message2 = LIGHTBLUE & .sMessage2
            End If
            If Not NoPlayer Then
                With dbPlayers(dbIndex)
                    .lMana = .lMana - dbSpells(dbSpellID).lMana
                End With
                lRoll = RndNumber(CDbl(.lMinDam), CDbl(modGetData.GetSpellMaxDamage(dbIndex, dbSpellID)))
            Else
                lRoll = RndNumber(CDbl(.lMinDam), CDbl(.lMaxDam))
            End If
            If Not EndCast Then
                If .iUse = 3 Or .iUse = 5 Then
                    If InStr(1, dbPlayers(dbIndex).sBlessSpells, .sSpellName & "~" & dbSpellID & "Œ") And .iUse <> 5 Then
                        If Not NoPlayer Then
                            WrapAndSend dbPlayers(dbIndex).iIndex, RED & "You already have this on you!" & WHITE & vbCrLf
                            X(Index) = ""
                        End If
                        Exit Sub
                    Else
                        If .iUse = 3 Then
BlessSelf:
                            If InStr(1, dbPlayers(dbIndex).sBlessSpells, .sSpellName & "~" & dbSpellID & "Œ") Then
                                WrapAndSend dbPlayers(dbIndex).iIndex, RED & "You already have this on you!" & WHITE & vbCrLf
                                X(Index) = ""
                                Exit Sub
                            End If
                            With dbPlayers(dbIndex)
                                If modSC.FastStringComp(.sBlessSpells, "0") Then .sBlessSpells = ""
                                .sBlessSpells = .sBlessSpells & dbSpells(dbSpellID).lTimeOut & "~" & lRoll & "~" & dbSpells(dbSpellID).sSpellName & "~" & dbSpellID & "Œ"
                            End With
                            bMe = True
                        ElseIf .iUse = 5 Then
                            With dbPlayers(dbIndex)
                                If .sParty = "0" Then GoTo BlessSelf
                                SplitFast ReplaceFast(.sParty, ":", ""), Arr, ";"
                                For i = 0 To UBound(Arr)
                                    If Arr(i) <> "" And Arr(i) <> "0" Then
                                        Arr(i) = CStr(GetPlayerIndexNumber(CLng(Val(Arr(i)))))
                                        With dbPlayers(Val(Arr(i)))
                                            If InStr(1, .sBlessSpells, dbSpells(dbSpellID).sSpellName & "~" & dbSpellID & "Œ") = 0 Then
                                                If modSC.FastStringComp(.sBlessSpells, "0") Then .sBlessSpells = ""
                                                .sBlessSpells = .sBlessSpells & dbSpells(dbSpellID).lTimeOut & "~" & lRoll & "~" & dbSpells(dbSpellID).sSpellName & "~" & dbSpellID & "Œ"
                                            Else
                                                Arr(i) = "0"
                                            End If
                                        End With
                                    End If
                                    If DE Then DoEvents
                                Next
                                If InStr(1, .sBlessSpells, dbSpells(dbSpellID).sSpellName & "~" & dbSpellID & "Œ") = 0 Then
                                    If modSC.FastStringComp(.sBlessSpells, "0") Then .sBlessSpells = ""
                                    .sBlessSpells = .sBlessSpells & dbSpells(dbSpellID).lTimeOut & "~" & lRoll & "~" & dbSpells(dbSpellID).sSpellName & "~" & dbSpellID & "Œ"
                                    bMe = True
                                Else
                                    bMe = False
                                End If
                            End With
                        End If
                    End If
                End If
                Message1 = ReplaceFast(Message1, "<%s>", dbSpells(dbSpellID).sSpellName)
                Message1 = ReplaceFast(Message1, "<%d>", CStr(lRoll))
                If Not NoPlayer Then
                    If .iUse <> 5 Then
                        Message1 = ReplaceFast(Message1, "<%v>", "yourself")
                        Message1 = ReplaceFast(Message1, "<%c>", "You")
                    ElseIf .iUse = 5 Then
                        Message1 = ReplaceFast(Message1, "<%v>", "the party")
                        Message1 = ReplaceFast(Message1, "<%c>", "You")
                    End If
                Else
                    Message1 = ReplaceFast(Message1, "<%v>", "you")
                    Message1 = ReplaceFast(Message1, "<%c>", Caster)
                End If
                Message2 = ReplaceFast(Message2, "<%s>", dbSpells(dbSpellID).sSpellName)
                Message2 = ReplaceFast(Message2, "<%d>", CStr(lRoll))
                If Not NoPlayer Then
                    If .iUse <> 5 Then
                        Message2 = ReplaceFast(Message2, "<%v>", modGetData.GetGenderPronoun(dbIndex) & "self")
                        Message2 = ReplaceFast(Message2, "<%c>", dbPlayers(dbIndex).sPlayerName)
                    ElseIf .iUse = 5 Then
                        Message2 = ReplaceFast(Message2, "<%v>", modGetData.GetGenderPronoun(dbIndex, True) & " party")
                        Message2 = ReplaceFast(Message2, "<%c>", dbPlayers(dbIndex).sPlayerName)
                    End If
                Else
                    Message2 = ReplaceFast(Message2, "<%v>", dbPlayers(dbIndex).sPlayerName)
                    Message2 = ReplaceFast(Message2, "<%c>", Caster)
                End If
                If Not modSC.FastStringComp(.sFlags, "0") Then
                    If .iUse <> 5 Then
                        modUseItems.DoFlags dbIndex, .sFlags, , lRoll, , , , True, Message4, Message2, Loc2, Message3, Loc3
                    ElseIf .iUse = 5 Then
                        If dbPlayers(dbIndex).sParty <> "0" Then
                            For i = 0 To UBound(Arr)
                                If Arr(i) <> "0" And Arr(i) <> "" Then
                                    modUseItems.DoFlags GetPlayerIndexNumber(Val(Arr(i))), .sFlags, , lRoll, , , , True, Message4, Message2, Loc2, Message3, Loc3
                                End If
                                If DE Then DoEvents
                            Next
                        End If
                        If bMe Then modUseItems.DoFlags dbIndex, .sFlags, , lRoll, , , , True, Message4, Message2, Loc2, Message3, Loc3
                    End If
                End If
            Else
                If Not modSC.FastStringComp(.sEndCastFlags, "0") Then
                    modUseItems.DoFlags dbIndex, .sEndCastFlags, , lRoll, , , , True, Message1, Message2, Loc2, Message3, Loc3
                End If
            End If
            If Message1 <> "" Then
                Message1 = Message1 & WHITE & vbCrLf
                WrapAndSend Index, Message1
            End If
            If Message2 <> "" Then
                Message2 = Message2 & WHITE & vbCrLf
                If Loc2 = 0 Then Loc2 = dbPlayers(dbIndex).lLocation
                SendToAllInRoom Index, Message2, Loc2
            End If
            If Message3 <> "" Then
                Message3 = Message3 & WHITE & vbCrLf
                If Loc3 <> 0 Then SendToAllInRoom Index, Message3, Loc3
            End If
        Else
            If Not NoPlayer Then
                With dbPlayers(dbIndex)
                    .lMana = .lMana - dbSpells(dbSpellID).lMana \ 2
                End With
                WrapAndSend Index, LIGHTBLUE & "You attempt to cast " & .sSpellName & ", but fail!" & WHITE & vbCrLf
                SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to cast " & .sSpellName & ", but fails" & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
                X(Index) = ""
            End If
            Exit Sub
        End If
    Else
        If RndNumber(0, 100) < modGetData.GetSpellChanceFromdbSpell(dbIndex, dbSpellID) Then
            If Not EndCast Then
                Message1 = LIGHTBLUE & .sMessage
                Message2 = LIGHTBLUE & .sMessage2
                Message4 = LIGHTBLUE & .sMessageV
            End If
            lRoll = RndNumber(CDbl(.lMinDam), CDbl(modGetData.GetSpellMaxDamage(dbIndex, dbSpellID)))
            If Not EndCast Then
                If .iUse = 3 Or .iUse = 5 Then
                    If InStr(1, dbPlayers(dbTIndex).sBlessSpells, .sSpellName & "~" & dbSpellID & "Œ") Then
                        WrapAndSend dbPlayers(dbIndex).iIndex, RED & dbPlayers(dbTIndex).sPlayerName & " already has this on them!" & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Sub
                    Else
                        With dbPlayers(dbTIndex)
                            If modSC.FastStringComp(.sBlessSpells, "0") Then .sBlessSpells = ""
                            .sBlessSpells = .sBlessSpells & dbSpells(dbSpellID).lTimeOut & "~" & lRoll & "~" & dbSpells(dbSpellID).sSpellName & "~" & dbSpellID & "Œ"
                        End With
                    End If
                End If
                With dbPlayers(dbIndex)
                    .lMana = .lMana - dbSpells(dbSpellID).lMana
                End With
                Message1 = ReplaceFast(Message1, "<%s>", dbSpells(dbSpellID).sSpellName)
                Message1 = ReplaceFast(Message1, "<%v>", dbPlayers(dbTIndex).sPlayerName)
                Message1 = ReplaceFast(Message1, "<%d>", CStr(lRoll))
                Message1 = ReplaceFast(Message1, "<%c>", "You")
                
                Message2 = ReplaceFast(Message2, "<%s>", dbSpells(dbSpellID).sSpellName)
                Message2 = ReplaceFast(Message2, "<%v>", dbPlayers(dbTIndex).sPlayerName)
                Message2 = ReplaceFast(Message2, "<%d>", CStr(lRoll))
                Message2 = ReplaceFast(Message2, "<%c>", dbPlayers(dbIndex).sPlayerName)
                
                Message4 = ReplaceFast(Message4, "<%s>", dbSpells(dbSpellID).sSpellName)
                Message4 = ReplaceFast(Message4, "<%v>", dbPlayers(dbTIndex).sPlayerName)
                Message4 = ReplaceFast(Message4, "<%d>", CStr(lRoll))
                Message4 = ReplaceFast(Message4, "<%c>", dbPlayers(dbIndex).sPlayerName)
                If Not modSC.FastStringComp(.sFlags, "0") Then
                    modUseItems.DoFlags dbIndex, .sFlags, , lRoll, , , , True, Message4, Message2, Loc2, Message3, Loc3
                End If
            Else
                If Not modSC.FastStringComp(.sEndCastFlags, "0") Then
                    modUseItems.DoFlags dbTIndex, .sEndCastFlags, , lRoll, , , , True, Message1, Message2, Loc2, Message3, Loc3
                End If
            End If
            If Message1 <> "" Then
                Message1 = Message1 & WHITE & vbCrLf
                WrapAndSend Index, Message1
            End If
            If Message2 <> "" Then
                Message2 = Message2 & WHITE & vbCrLf
                If Loc2 = 0 Then Loc2 = dbPlayers(dbIndex).lLocation
                SendToAllInRoom Index, Message2, Loc2, dbPlayers(dbTIndex).iIndex
            End If
            If Message3 <> "" Then
                Message3 = Message3 & WHITE & vbCrLf
                If Loc3 <> 0 Then SendToAllInRoom Index, Message3, Loc3, dbPlayers(dbTIndex).iIndex
            End If
            If Message4 <> "" Then
                Message4 = Message4 & WHITE & vbCrLf
                WrapAndSend dbPlayers(dbTIndex).iIndex, Message4
            End If
        Else
            With dbPlayers(dbIndex)
                .lMana = .lMana - dbSpells(dbSpellID).lMana \ 2
            End With
            WrapAndSend Index, LIGHTBLUE & "You attempt to cast " & .sSpellName & " on " & dbPlayers(dbTIndex).sPlayerName & ", but fail!" & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " attempts to cast " & .sSpellName & " on " & dbPlayers(dbTIndex).sPlayerName & ", but fails!" & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
            X(Index) = ""
            Exit Sub
        End If
    End If
    If .sEndCastFlags <> "0" And Not EndCast Then
        dbPlayers(dbIndex).lHasCasted = 0
        DoNonCombatSpell Index, dbIndex, dbSpellID, Target, dbTIndex, NoPlayer, Caster, True
    End If
End With
    
End Sub


'Public Sub DoFlagsInVerse(dbIndex As Long, ByVal lRoll As Long, ByRef aFlgs() As String)
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
''EXP# 'current EXP
''txp# 'total exp
''gol# 'gold
''dod# 'dodge
''ban# 'bank
''vis# 'vision
''mit# 'max items
''ccp# 'current CP
''evi# 'evil points
''pap# 'paper
''mat# 'make item
''clp# 'class points
'Dim i As Long
'Dim dVal As Double
'For i = LBound(aFlgs) To UBound(aFlgs)
'    If Not modSC.FastStringComp(aFlgs(i), "") Then
'        If Mid$(aFlgs(i), 4) = "-3" Then
'            dVal = -lRoll
'        ElseIf Mid(aFlgs(i), 4) = "--3" Then
'            dVal = lRoll
'        Else
'            dVal = -CDbl(Val(Mid$(aFlgs(i), 4)))
'        End If
'        Select Case Left$(aFlgs(i), 3)
'            Case "lig"
'                With dbPlayers(dbIndex)
'                    .iVision = .iVision + dVal
'
'                End With
'            Case "cri"
'                With dbPlayers(dbIndex)
'                    .iCrits = .iCrits + 1
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
'            Case "acl"
'                With dbPlayers(dbIndex)
'                    .iAC = .iAC + dVal
'                    Send2InThisRoom = .lLocation
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
'            Case "sas"
'                With dbPlayers(dbIndex)
'                    .sPlayerName = .sSeenAs
'                End With
'            Case "des"
'                With dbPlayers(dbIndex)
'                    .sOverrideDesc = "0"
'                End With
'            Case "csp"
'                dbPlayers(dbIndex).lHasCasted = 0
'                modSpells.DoNonCombatSpell dbPlayers(dbIndex).iIndex, dbIndex, Abs(dVal)
'            Case "thi"
'                modMiscFlag.SetStatsPlus dbIndex, [Thieving Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Thieving Bonus]) + CLng(dVal)
'        End Select
'        If Left$(aFlgs(i), 2) Like "el#" Then
'            Select Case dVal
'                Case Is < 1
'                    dVal = 0
'                Case Is >= 1
'                    dVal = 1
'            End Select
'            modResist.UpdateResistValue dbIndex, CLng(Val(Mid$(aFlgs(i), 3, 1))), CLng(dVal)
'        End If
'        If Left$(aFlgs(i), 3) Like "m##" Then
'            Select Case dVal
'                Case Is < 1
'                    dVal = 0
'                Case Is >= 1
'                    dVal = 1
'            End Select
'            modMiscFlag.SetMiscFlag dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2))), CLng(dVal)
'        End If
'        If Left(aFlgs(i), 3) Like "s##" Then
'            Select Case Val(Mid$(aFlgs(i), 2, 2))
'                Case 1, 3, 5, 9, 11, 13
'                    modMiscFlag.SetStatsPlus dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2))), modMiscFlag.GetStatsPlus(dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2)))) + CLng(dVal)
'            End Select
'        End If
'    End If
'    If DE Then DoEvents
'Next
'End Sub

'Public Sub DoFlags(dbIndex As Long, ByVal lRoll As Long, ByRef aFlgs() As String, ByRef Message As String, ByRef Message2 As String, ByRef Send2InThisRoom As Long, ByRef Message3 As String, ByRef Send3InThisRoom As Long)
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
''EXP# 'current EXP
''txp# 'total exp
''gol# 'gold
''dod# 'dodge
''ban# 'bank
''vis# 'vision
''mit# 'max items
''ccp# 'current CP
''evi# 'evil points
''pap# 'paper
''mit# 'make item
''clp# 'class points
'Dim i As Long
'Dim dVal As Double
'Dim tArr() As String
'Dim lFlg As Long
'Dim sExits As String
'Dim dbFamID As Long
'Dim j As Long
'Dim tArr2() As String
'Dim bFlgs() As String
'For i = LBound(aFlgs) To UBound(aFlgs)
'    If Not modSC.FastStringComp(aFlgs(i), "") Then
'        If Mid$(aFlgs(i), 4) = "-3" Then
'            dVal = lRoll
'        ElseIf Mid(aFlgs(i), 4) = "--3" Then
'            dVal = -lRoll
'        Else
'            dVal = CDbl(Val(Mid$(aFlgs(i), 4)))
'        End If
'        Select Case Left$(aFlgs(i), 3)
'            Case "tel"
'                Select Case dVal
'                    Case "-1"
'                        sExits = modGetData.sGetRoomExits(dbPlayers(dbIndex).iIndex)
'                        If Not modSC.FastStringComp(sExits, "") Then
'                            SplitFast sExits, tArr, ","
'                        Else
'                            Message = Message & BRIGHTBLUE & "You fail to teleport!" & WHITE & vbCrLf
'                            GoTo nNext
'                        End If
'                        lFlg = CLng(RndNumber(LBound(tArr), UBound(tArr)))
'                        lFlg = CLng(tArr(lFlg))
'                    Case "-2"
'                        lFlg = dbMap(RndNumber(LBound(dbMap), UBound(dbMap))).lRoomID
'                    Case Else
'                        If dVal > 0 Then
'                            lFlg = CLng(dVal)
'                        Else
'                            Message = Message & BRIGHTBLUE & "You fail to teleport!" & WHITE & vbCrLf
'                            GoTo nNext
'                        End If
'                End Select
'                Send2InThisRoom = dbPlayers(dbIndex).lLocation
'                dbPlayers(dbIndex).lLocation = lFlg
'                dbPlayers(dbIndex).lBackUpLoc = lFlg
'                Message3 = Message3 & BLUE & dbPlayers(dbIndex).sPlayerName & " appears in the room!" & WHITE & vbCrLf
'                Send3InThisRoom = lFlg
'            Case "stu"
'                With dbPlayers(dbIndex)
'                    .iStun = .iStun + dVal
'                    Message = Message & BRIGHTYELLOW & "You are stunned!" & WHITE & vbCrLf
'                    Message2 = Message2 & YELLOW & .sPlayerName & " is stunned!" & WHITE & vbCrLf
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "acl"
'                With dbPlayers(dbIndex)
'                    .iAC = .iAC + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "lig"
'                With dbPlayers(dbIndex)
'                    .iVision = .iVision + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "cri"
'                With dbPlayers(dbIndex)
'                    .iCrits = .iCrits + 1
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "acc"
'                With dbPlayers(dbIndex)
'                    .iAcc = .iAcc + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "dam"
'                With dbPlayers(dbIndex)
'                    .iMaxDamage = .iMaxDamage + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "str"
'                With dbPlayers(dbIndex)
'                    .iStr = .iStr + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "agi"
'                With dbPlayers(dbIndex)
'                    .iAgil = .iAgil + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "cha"
'                With dbPlayers(dbIndex)
'                    .iCha = .iCha + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "dex"
'                With dbPlayers(dbIndex)
'                    .iDex = .iDex + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "int"
'                With dbPlayers(dbIndex)
'                    .iInt = .iInt + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "chp"
'                With dbPlayers(dbIndex)
'                    .lHP = .lHP + dVal
'                    If .lHP > .lMaxHP Then .lHP = .lMaxHP
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "mhp"
'                With dbPlayers(dbIndex)
'                    .lMaxHP = .lMaxHP + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "cma"
'                With dbPlayers(dbIndex)
'                    .lMana = .lMana + dVal
'                    If .lMana > .lMaxMana Then .lMana = .lMaxMana
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "mma"
'                With dbPlayers(dbIndex)
'                    .lMaxMana = .lMaxMana + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "hun"
'                With dbPlayers(dbIndex)
'                    .dHunger = .dHunger + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "sta"
'                With dbPlayers(dbIndex)
'                    .dStamina = .dStamina + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "cac"
'                With dbPlayers(dbIndex)
'                    .iAC = .iAC + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "exp"
'                With dbPlayers(dbIndex)
'                    .dEXP = .dEXP + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "txp"
'                With dbPlayers(dbIndex)
'                    .dTotalEXP = .dTotalEXP + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "gol"
'                With dbPlayers(dbIndex)
'                    .dGold = .dGold + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "dod"
'                With dbPlayers(dbIndex)
'                    .iDodge = .iDodge + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "ban"
'                With dbPlayers(dbIndex)
'                    .dBank = .dBank + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "vis"
'                With dbPlayers(dbIndex)
'                    .iVision = .iVision + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "mit"
'                With dbPlayers(dbIndex)
'                    modMiscFlag.SetStatsPlus dbIndex, [Max Items Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Max Items Bonus]) + CLng(dVal)
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "ccp"
'                With dbPlayers(dbIndex)
'                    .iIsReadyToTrain = .iIsReadyToTrain + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "evi"
'                With dbPlayers(dbIndex)
'                    .iEvil = .iEvil + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "pap"
'                With dbPlayers(dbIndex)
'                    .lPaper = .lPaper + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "mat"
'                dbFamID = GetItemID(, CLng(dVal))
'                If dbFamID = 0 Then GoTo nNext
'                With dbPlayers(dbIndex)
'                    If modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) + 1 < modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) Then
'                        If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
'                        .sInventory = .sInventory & ":" & dbItems(dbFamID).iID & "/" & dbItems(dbFamID).iUses & "/" & dbItems(dbFamID).lDurability & ";"
'                    Else
'                        With dbMap(GetMapIndex(.lLocation))
'                            If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
'                            .sItems = .sItems & ":" & dbItems(dbFamID).iID & "/" & dbItems(dbFamID).iUses & "/" & dbItems(dbFamID).lDurability & ";"
'                        End With
'                    End If
'                End With
'            Case "clp"
'                With dbPlayers(dbIndex)
'                    .dClassPoints = .dClassPoints + dVal
'                    Send2InThisRoom = .lLocation
'                End With
'            Case "sas"
'                With dbPlayers(dbIndex)
'                    .sPlayerName = Mid$(aFlgs(i), 4)
'                End With
'            Case "des"
'                With dbPlayers(dbIndex)
'                    .sOverrideDesc = Mid$(aFlgs(i), 4)
'                End With
'            Case "rms"
'                With dbPlayers(dbIndex)
'                    SplitFast Left$(.sBlessSpells, Len(.sBlessSpells) - 1), tArr, "Œ"
'                    For j = LBound(tArr) To UBound(tArr)
'                        Erase tArr2
'                        SplitFast tArr(j), tArr2, "~"
'                        If modSC.FastStringComp(CStr(dbSpells(Val(tArr2(3))).lID), CStr(dVal)) Then
'                            If Not modSC.FastStringComp(dbSpells(CLng(tArr2(3))).sFlags, "0") Then
'                                SplitFast dbSpells(CLng(tArr2(3))).sFlags, bFlgs, ";"
'                                modSpells.DoFlagsInVerse dbIndex, CLng(tArr2(1)), bFlgs
'                            End If
'                            .sBlessSpells = ReplaceFast(.sBlessSpells, tArr2(0) & "~" & tArr2(1) & "~" & tArr2(2) & "~" & tArr2(3) & "Œ", "", 1, 1)
'                            If modSC.FastStringComp(.sBlessSpells, "") Then .sBlessSpells = "0"
'                            sSend .iIndex, LIGHTBLUE & dbSpells(CLng(tArr2(3))).sRunOutMessage
'                        End If
'                        If DE Then DoEvents
'                    Next
'                End With
'            Case "csp"
'                dbPlayers(dbIndex).lHasCasted = 0
'                modSpells.DoNonCombatSpell dbPlayers(dbIndex).iIndex, dbIndex, CLng(dVal)
'            Case "thi"
'                modMiscFlag.SetStatsPlus dbIndex, [Thieving Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Thieving Bonus]) + CLng(dVal)
'        End Select
'        If Left$(aFlgs(i), 2) Like "el#" Then
''            Select Case dVal
''                Case Is < 1
''                    dVal = 0
''                Case Is >= 1
''                    dVal = 1
''            End Select
'            modResist.UpdateResistValue dbIndex, CLng(Val(Mid$(aFlgs(i), 3, 1))), CLng(dVal)
'        End If
'        If Left$(aFlgs(i), 3) Like "m##" Then
'            Select Case dVal
'                Case Is < 1
'                    dVal = 0
'                Case Is >= 1
'                    dVal = 1
'            End Select
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
'End Sub



'Public Function ItemSpell(iID As Long, Index As Long) As Boolean
'Dim dbpID As Long
'With dbSpells(GetSpellID(, iID))
'    dbpID = GetPlayerIndexNumber(Index)
'    If .lTimeOut <> 0 Then
'        If modSC.FastStringComp(dbPlayers(dbpID).sBlessSpells, "0") Then dbPlayers(dbpID).sBlessSpells = ""
'        If InStr(1, dbPlayers(dbpID).sBlessSpells, .lTimeOut & "~" & .sSpellName & "~" & GetSpellID(, iID) & "Œ") Then
'            WrapAndSend Index, RED & "You already have this on you!" & WHITE & vbCrLf
'            ItemSpell = False
'            Exit Function
'        End If
'        ItemSpell = True
'        dbPlayers(dbpID).sBlessSpells = dbPlayers(dbpID).sBlessSpells & .lTimeOut & "~" & .sSpellName & "~" & GetSpellID(, iID) & "Œ"
''        AssignBlessBonus GetSpellID(, iID), dbpID
'        WrapAndSend Index, LIGHTBLUE & .sMessage & WHITE & vbCrLf
'        'If SendM Then SendToAllInRoom dbPlayers(dbpID).iIndex, LIGHTBLUE & dbPlayers(dbpID).sPlayerName & " cast " & .sSpellName & "." & WHITE & vbCrLf, CStr(dbPlayers(dbpID).lLocation)
'    Else
'        X(Index) = ""
'    End If
'End With
'End Function

'Sub DoDamageToPlayerWithSpell(Index As Long, SpellID As Long)
'Dim Min As Long, Max As Long
'Dim Damage&, Level&
'Dim pID&
'pID& = GetPlayerIndexNumber(Index)
'With dbPlayers(pID&)
'    Level& = .iLevel
'End With
'With dbSpells(GetSpellID(, SpellID&))
'    Min = .lMinDam
'    If Level& > .iLevelMax Then Level& = .iLevelMax
'    Damage& = CLng(.lMaxDam) + (Level& * .iLevelModify)
'    Max = Damage&
'End With
'With dbPlayers(pID&)
'    .lHP = .lHP - RndNumber(CDbl(Min), CDbl(Max))
'    CheckDeath Index
'End With
'End Sub
