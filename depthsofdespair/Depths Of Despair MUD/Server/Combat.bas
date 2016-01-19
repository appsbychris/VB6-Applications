Attribute VB_Name = "modCombat"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modCombat
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private Type CombatDef
    dbAIndex As Long
    lPlayerID As Long
    amonIndex As Long
    bIs_Mon_Attacking As Boolean
    bIs_Room_Spell As Boolean
    bIs_Player_Attacking As Boolean
    bVic_Is_Attacking As Boolean
    bIs_PvP As Boolean
    dbVIndex As Long
    lVPlayerID As Long
End Type
Public Enum enFirst
    Player = 0
    Monster = 1
    PvPA = 2
    PvPV = 3
End Enum

Public Enum enChoice
    PlayerAndMon = 0
    PlayerVsPlayer = 1
End Enum

Private iWhichDir As Long

Public Function WhoGoesFirst(eChoice As enChoice, Optional dbAIndex As Long, Optional amonIndex As Long, Optional dbVIndex As Long) As enFirst
Dim d1 As Double
Dim d2 As Double
If eChoice = PlayerAndMon Then
    With dbPlayers(dbAIndex)
        d1 = (.iInt + .iCha) / 1.88
        d1 = d1 + (.iDex * 2) + (.iStr / 2)
        d1 = d1 / 0.44
        d1 = d1 + (.iAgil * 44) / 17.62
        d1 = d1 * (.dHunger / 22.782)
        d1 = d1 * .dStamina
        d1 = d1 / 10753.12953
    End With
    With aMons(amonIndex)
        d2 = (.mAc * 62) / 1.88
        d2 = d2 + (.mEnergy + (.mEXP / 3))
        d2 = d2 / 0.32
        d2 = d2 + (.mHP * 44) / 23.827
        d2 = d2 + (.mPEnergy) * 42
        d2 = d2 / 926.25458
    End With
    If d1 >= d2 Then WhoGoesFirst = Player Else WhoGoesFirst = Monster
Else
    With dbPlayers(dbAIndex)
        d1 = (.iAC * 72) + (.iInt / 3.12244)
        d1 = d1 + (.iStr * (.iAgil * 14)) / 1788
        d1 = d1 + (.iDex * 2.33345)
        d1 = d1 * 3.4568 + .iCha
        d1 = d1 / 1400.83749
    End With
    With dbPlayers(dbVIndex)
        d2 = (.iAC * 72) + (.iInt / 3.12244)
        d2 = d2 + (.iStr * (.iAgil * 14)) / 1788
        d2 = d2 + (.iDex * 2.33345)
        d2 = d2 * 3.4568 + .iCha
        d2 = d2 / 1400.83749
    End With
    If d1 >= d2 Then WhoGoesFirst = PvPA Else WhoGoesFirst = PvPV
End If
End Function

Sub GenCombat()
On Error GoTo GenCombat_Error

InitMonsters
InitCombat
CleanUpAMons
'InitPvP

   On Error GoTo 0
   Exit Sub

GenCombat_Error:

    
End Sub

Sub InitMonsters()
Dim i As Long
Dim j As Long
Dim m As Long
Dim n As Long
Dim s As String
Dim spHere As String
Dim tArr() As String
Dim aHere() As String
On Error GoTo eh1:
For i = LBound(dbPlayers) To UBound(dbPlayers)
    If dbPlayers(i).iIndex <> 0 Then
        s = ""
        Erase tArr
        s = modGetData.GetAllMonstersInRoom(dbPlayers(i).lLocation, dbPlayers(i).lDBLocation)
        If s <> "" Then
            SplitFast s, tArr, ";"
            For j = LBound(tArr) To UBound(tArr)
                If Not modSC.FastStringComp(tArr(j), "") Then
                    n = CLng(Val(tArr(j)))
                    If n <= UBound(aMons) Then
                        If aMons(n).mLoc = dbPlayers(i).lLocation Then
                            If aMons(n).mAttackable = True And aMons(n).mHostile = True And aMons(n).mIsAttacking = False Then
                                Erase aHere
                                spHere = modGetData.GetPlayersIDsHere(aMons(n).mLoc)
                                If DCount(spHere, ";") > 1 Then
                                    SplitFast Left$(spHere, Len(spHere) - 1), aHere, ";"
                                    For m = LBound(aHere) To UBound(aHere)
                                        modMonsters.InsertInMonList n, CLng(Val(aHere(m)))
                                        If DE Then DoEvents
                                    Next
                                    For m = 0 To 9
                                        If aMons(n).mList(m) <> 0 Then
                                            With dbPlayers(GetPlayerIndexNumber(, , aMons(n).mList(m)))
                                                If InStr(1, .sInventory, ":" & CStr(aMons(n).mDontAttackIfItem) & "/") = 0 Then
                                                    If aMons(n).mEvil > 39 And .iEvil > 39 Then
                                                        '
                                                    ElseIf aMons(n).mEvil < -40 And .iEvil < -40 Then
                                                        '
                                                    Else
                                                        If .iGhostMode = 1 Then Exit For
                                                        aMons(n).mIsAttacking = True
                                                        aMons(n).mPlayerAttacking = .iIndex
                                                        WrapAndSend .iIndex, BRIGHTRED & aMons(n).mName & " moves to attack you!" & WHITE & vbCrLf
                                                        SendToAllInRoom .iIndex, BRIGHTRED & aMons(n).mName & " moves to attack " & .sPlayerName & "." & WHITE & vbCrLf, .lLocation
                                                        Exit For
                                                    End If
                                                End If
                                            End With
                                        End If
                                        If DE Then DoEvents
                                    Next
                                Else
                                    If dbPlayers(i).iSneaking = 0 And modMiscFlag.GetMiscFlag(i, Invisible) = 0 And aMons(n).mIsAttacking = False Then
                                        If InStr(1, dbPlayers(i).sInventory, ":" & CStr(aMons(n).mDontAttackIfItem) & "/") = 0 Then
                                            If dbPlayers(i).iSneaking = 0 Then
                                                If dbPlayers(i).iGhostMode = 0 Then
                                                    spHere = ReplaceFast(spHere, ";", "")
                                                    modMonsters.InsertInMonList n, CLng(Val(spHere))
                                                    aMons(n).mIsAttacking = True
                                                    aMons(n).mPlayerAttacking = dbPlayers(i).iIndex
                                                    WrapAndSend dbPlayers(i).iIndex, BRIGHTRED & aMons(n).mName & " moves to attack you!" & WHITE & vbCrLf
                                                    SendToAllInRoom dbPlayers(i).iIndex, BRIGHTRED & aMons(n).mName & " moves to attack " & dbPlayers(i).sPlayerName & "." & WHITE & vbCrLf, dbPlayers(i).lLocation
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf aMons(n).mHostile = False And aMons(n).mIsAttacking = False Then
                                If dbPlayers(i).iSneaking = 0 And modMiscFlag.GetMiscFlag(i, Invisible) = 0 And dbPlayers(i).iGhostMode = 0 Then
                                    Select Case dbPlayers(i).iEvil
                                        Case Is >= 40
                                            If aMons(n).mEvil < -41 Then
                                                modMonsters.InsertInMonList n, dbPlayers(i).lPlayerID, 0
                                            End If
                                        Case Is <= -41
                                            If aMons(n).mEvil > 40 Then
                                                modMonsters.InsertInMonList n, dbPlayers(i).lPlayerID, 0
                                            End If
                                    End Select
                                    For m = 0 To 9
                                        If aMons(n).mList(m) <> 0 Then
                                            With dbPlayers(GetPlayerIndexNumber(, , aMons(n).mList(m)))
                                                If InStr(1, .sInventory, ":" & CStr(aMons(n).mDontAttackIfItem) & "/") = 0 Then
                                                    If aMons(n).mEvil > 39 And .iEvil > 39 Then
                                                        '
                                                    ElseIf aMons(n).mEvil < -40 And .iEvil < -40 Then
                                                        '
                                                    Else
                                                        aMons(n).mIsAttacking = True
                                                        aMons(n).mPlayerAttacking = .iIndex
                                                        WrapAndSend .iIndex, BRIGHTRED & aMons(n).mName & " moves to attack you!" & WHITE & vbCrLf
                                                        SendToAllInRoom .iIndex, BRIGHTRED & aMons(n).mName & " moves to attack " & .sPlayerName & "." & WHITE & vbCrLf, .lLocation
                                                        Exit For
                                                    End If
                                                End If
                                            End With
                                        End If
                                        If DE Then DoEvents
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
                If DE Then DoEvents
            Next
        End If
    End If
    If DE Then DoEvents
Next
eh1:
End Sub



Sub InitCombat()
Dim sPAMessage As String
Dim lPADamage As Long
Dim lMADamage As Long
Dim sMonMessage As String
Dim sPVMessage As String
Dim sMessages2 As String
Dim sMessages As String
Dim sS As String
Dim bDidA As Boolean
Dim i As Long
Dim j As Long
Dim xxx As Double
Dim aList() As CombatDef
Dim bPFlag As Boolean
ReDim aList(0)
Dim sAIds As String
Dim tmpID() As String
Dim Y As String

For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iIndex > 0 And pPoint(.iIndex) = 0 Then
            sAIds = dbMap(.lDBLocation).sAMonIds
            If .dMonsterID = 99998 Then
                Erase tmpID
                Y = ""
                Y = "P;"
                If sAIds <> "" Then
                    SplitFast sAIds, tmpID, ";"
                    For j = LBound(tmpID) To UBound(tmpID)
                        If tmpID(j) <> "" Then
                            If aMons(Val(tmpID(j))).mHostile = True Or _
                               aMons(Val(tmpID(j))).mIs_Being_Attacked = True Or _
                               aMons(Val(tmpID(j))).mIsAttacking = True Then
                            
                                    If WhoGoesFirst(PlayerAndMon, i, Val(tmpID(j))) = Monster Then
                                        Y = tmpID(j) & ";" & Y
                                    Else
                                        Y = Y & tmpID(j) & ";"
                                    End If
                            End If
                        End If
                        If DE Then DoEvents
                    Next
                End If
                sAIds = modGetData.GetPlayersDBIndexesHereNotInParty(i, .lLocation)
                If sAIds <> "" Then
                    SplitFast sAIds, tmpID, ";"
                    For j = LBound(tmpID) To UBound(tmpID)
                        If tmpID(j) <> "" Then
                            If .iIndex = dbPlayers(Val(tmpID(j))).iPlayerAttacking Then
                                If WhoGoesFirst(PlayerVsPlayer, i, , Val(tmpID(j))) = PvPV Then
                                    Y = "V" & tmpID(j) & ";" & Y
                                Else
                                    Y = Y & "V" & tmpID(j) & ";"
                                End If
                            End If
                        End If
                    Next
                End If
                Erase tmpID
                If Y <> "P;" Then
                    SplitFast Y, tmpID, ";"
                    For j = LBound(tmpID) To UBound(tmpID)
                        If tmpID(j) <> "" Then
                            ReDim Preserve aList(UBound(aList) + 1)
                            With aList(UBound(aList))
                                If modSC.FastStringComp(Left$(tmpID(j), 1), "V") Then
                                    tmpID(j) = Mid$(tmpID(j), 2)
                                    .bIs_PvP = True
                                    .bIs_Player_Attacking = False
                                    .bVic_Is_Attacking = True
                                    .dbAIndex = i
                                    .lPlayerID = dbPlayers(i).lPlayerID
                                    .dbVIndex = CLng(Val(tmpID(j)))
                                    .lVPlayerID = dbPlayers(Val(tmpID(j))).lPlayerID
                                    .bIs_Room_Spell = True
                                ElseIf modSC.FastStringComp(Left$(tmpID(j), 1), "P") Then
                                    .bIs_Player_Attacking = True
                                    .bIs_Room_Spell = True
                                    .dbAIndex = i
                                    .lPlayerID = dbPlayers(i).lPlayerID
                                Else
                                    .amonIndex = CLng(Val(tmpID(j)))
                                    .bIs_Mon_Attacking = True
                                    .bIs_Player_Attacking = False
                                    .bIs_Room_Spell = True
                                    .dbAIndex = i
                                    .lPlayerID = dbPlayers(i).lPlayerID
                                End If
                            End With
                        End If
                        If DE Then DoEvents
                    Next
                Else
                    ReDim Preserve aList(UBound(aList) + 1)
                    With aList(UBound(aList))
                        .bIs_Player_Attacking = True
                        .bIs_Room_Spell = True
                        .dbAIndex = i
                        .lPlayerID = dbPlayers(i).lPlayerID
                    End With
                End If
            Else
                If sAIds <> "" Then
                    Erase tmpID
                    SplitFast sAIds, tmpID, ";"
                    For j = LBound(tmpID) To UBound(tmpID)
                        If tmpID(j) <> "" Then
                            
                            If (aMons(Val(tmpID(j))).mPlayerAttacking = .iIndex And .iIndex <> 0) Or _
                               (.dMonsterID <> 99998 And .iIndex <> 0) Or _
                               (.dMonsterID = i) Then
                            '
                                    If aMons(Val(tmpID(j))).mHostile = True Or _
                                       aMons(Val(tmpID(j))).mIs_Being_Attacked = True Or _
                                       aMons(Val(tmpID(j))).mIsAttacking = True Then
                                    '
                                            If (CheckIsThere(.iIndex, Val(tmpID(j))) = True) Or _
                                               (.lLocation = aMons(Val(tmpID(j))).mLoc) Then
                                            '
                                                    If aMons(Val(tmpID(j))).mIsAttacking = True Or _
                                                       aMons(Val(tmpID(j))).mIs_Being_Attacked Then
                                                    '
                                                            ReDim Preserve aList(UBound(aList) + 1)
                                                            With aList(UBound(aList))
                                                                .amonIndex = CLng(Val(tmpID(j)))
                                                                .dbAIndex = i
                                                                .lPlayerID = dbPlayers(i).lPlayerID
                                                                If aMons(Val(tmpID(j))).mList(0) = .lPlayerID Then .bIs_Mon_Attacking = True
                                                                If dbPlayers(i).dMonsterID <> 99999 Then
                                                                    .bIs_Player_Attacking = True
                                                                Else
                                                                    .bIs_Player_Attacking = False
                                                                End If
                                                                .bIs_Room_Spell = False
                                                            End With
                                                    '
                                            
                                            '
                                                    End If
                                    '
                                            End If
                            '
                                    End If
                                    
                            End If
                            
                        End If
                        If DE Then DoEvents
                    Next
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next

'For i = LBound(aMons) To UBound(aMons)
'    For j = LBound(dbPlayers) To UBound(dbPlayers)
'        If (aMons(i).mPlayerAttacking = dbPlayers(j).iIndex And dbPlayers(j).iIndex <> 0) Or (dbPlayers(j).dMonsterID = 99998 And dbPlayers(j).iIndex <> 0) Or (dbPlayers(j).dMonsterID = i) Then
'            If aMons(i).mHostile = True Or aMons(i).mIs_Being_Attacked = True Or aMons(i).mIsAttacking = True Or dbPlayers(j).dMonsterID = 99998 Then
'                If (CheckIsThere(dbPlayers(j).iIndex, i) = True) Or (dbPlayers(j).dMonsterID = 99998 And dbPlayers(j).lLocation = aMons(i).mLoc) Then
'                    If aMons(i).mIsAttacking = True Or aMons(i).mIs_Being_Attacked Or (dbPlayers(j).dMonsterID = 99998) Then
'                        If dbPlayers(j).dMonsterID = 99998 And dbPlayers(j).iHasAttacked = 0 Then
'                            ReDim Preserve aList(UBound(aList) + 1)
'                            With aList(UBound(aList))
'                                .bIs_PvP = False
'                                .dbAIndex = j
'                                .lPlayerID = dbPlayers(j).lPlayerID
'                                .bIs_Mon_Attacking = False
'                                If aMons(i).mList(0) = .lPlayerID Then .bIs_Mon_Attacking = True
'                                .bIs_Room_Spell = True
'                                dbPlayers(j).iHasAttacked = 1
'                            End With
'                        ElseIf dbPlayers(j).dMonsterID <> 99998 Then
'                            ReDim Preserve aList(UBound(aList) + 1)
'                            With aList(UBound(aList))
'                                .amonIndex = i
'                                .dbAIndex = j
'                                .lPlayerID = dbPlayers(j).lPlayerID
'                                If aMons(i).mList(0) = .lPlayerID Then .bIs_Mon_Attacking = True
'                                If dbPlayers(j).dMonsterID <> 99999 Then
'                                    .bIs_Player_Attacking = True
'                                Else
'                                    .bIs_Player_Attacking = False
'                                End If
'                                .bIs_Room_Spell = False
'                            End With
'                        ElseIf dbPlayers(j).dMonsterID = 99998 And dbPlayers(j).iHasAttacked = 1 Then
'                            ReDim Preserve aList(UBound(aList) + 1)
'                            With aList(UBound(aList))
'                                .bIs_PvP = False
'                                .dbAIndex = j
'                                .lPlayerID = dbPlayers(j).lPlayerID
'                                If aMons(i).mList(0) = .lPlayerID Then .bIs_Mon_Attacking = True
'                                .bIs_Room_Spell = True
'                                .bIs_Player_Attacking = False
'                                'dbPlayers(j).iHasAttacked = 1
'                            End With
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        If DE Then DoEvents
'    Next
'    If DE Then DoEvents
'Next
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If .iPlayerAttacking <> 0 Or .dMonsterID = 99998 Then
            If .iHasAttacked = 0 Then
                If PvPStillThere(i, .iPlayerAttacking) Or (.dMonsterID = 99998) Then
                    If dbPlayers(i).dMonsterID = 99998 Then
'                        ReDim Preserve aList(UBound(aList) + 1)
'                        With aList(UBound(aList))
'                            .dbAIndex = i
'                            .lPlayerID = dbPlayers(i).lPlayerID
'                            .bIs_Room_Spell = True
'                            .bIs_Mon_Attacking = False
'                            .bIs_PvP = False
'                            dbPlayers(i).iHasAttacked = 1
'                        End With
                    Else
                        ReDim Preserve aList(UBound(aList) + 1)
                        j = GetPlayerIndexNumber(.iPlayerAttacking)
                        If dbPlayers(j).iPlayerAttacking = .iIndex Then dbPlayers(j).iHasAttacked = 1
                        With aList(UBound(aList))
                            .bIs_PvP = True
                            .dbAIndex = i
                            .dbVIndex = j
                            .lPlayerID = dbPlayers(i).lPlayerID
                            .lVPlayerID = dbPlayers(j).lPlayerID
                            .bIs_Room_Spell = False
                            If dbPlayers(j).iPlayerAttacking = dbPlayers(i).iIndex Then .bVic_Is_Attacking = True: dbPlayers(j).iHasAttacked = 1
                        End With
                    End If
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next
If UBound(aList) < 1 Then Exit Sub
For i = 1 To UBound(aList)
    With aList(i)
        sPAMessage = ""
        sPVMessage = ""
        sMessages = ""
        sMonMessage = ""
        lPADamage = 0
        lMADamage = 0
        bPFlag = False
        If .bIs_Mon_Attacking And .bIs_Room_Spell = False Then
            'Deter who goes first
            If WhoGoesFirst(PlayerAndMon, .dbAIndex, .amonIndex) = Monster Then
                DoMonAttack .dbAIndex, .amonIndex, sMonMessage, sMessages2, lMADamage
                With dbPlayers(.dbAIndex)
                    If .lHP - lMADamage < 1 Then
                        .lHP = .lHP - lMADamage
                        'WrapAndSend dbPlayers(aList(i).dbAIndex).iIndex, sMonMessage
                        'SendToAllInRoom dbPlayers(aList(i).dbAIndex).iIndex, sMessages2, dbPlayers(aList(i).dbAIndex).lLocation
                        If CheckDeath(.iIndex, , True, sMonMessage, , sMessages2, bPFlag, , , aList(i).dbAIndex, aList(i).amonIndex) = False Then
                            WrapAndSend .iIndex, sMonMessage
                            SendToAllInRoom .iIndex, sMessages2, .lLocation
                        Else
                            WrapAndSend .iIndex, sMonMessage
                            SendToAllInRoom .iIndex, sMessages2, aMons(aList(i).amonIndex).mLoc
                        End If
                        If bPFlag Then RemoveFromParty .iIndex
                        GoTo nNext
                    End If
                End With
                If .bIs_Player_Attacking Then DoFamAttackMon .dbAIndex, .amonIndex, sPAMessage, sMessages, lPADamage
                DoPlayerAttackMon .dbAIndex, .amonIndex, sPAMessage, sMessages, lPADamage, bDidA
            Else
                DoPlayerAttackMon .dbAIndex, .amonIndex, sPAMessage, sMessages, lPADamage, bDidA
                DoMonAttack .dbAIndex, .amonIndex, sMonMessage, sMessages2, lMADamage
                If .bIs_Player_Attacking Then DoFamAttackMon .dbAIndex, .amonIndex, sPAMessage, sMessages, lPADamage
            End If
            With aMons(.amonIndex)
                .mHP = .mHP - lPADamage
                If .mHP <= 0 Then
                    For j = 1 To UBound(aList)
                        With aList(j)
                            If .amonIndex = aList(i).amonIndex Then
                                '.amonIndex = 0
                                .bIs_Mon_Attacking = False
                                .bIs_Player_Attacking = False
                            End If
                        End With
                        If DE Then DoEvents
                    Next
                    DropMonGold aList(i).amonIndex, sS
                    sPAMessage = sPAMessage & sS
                    sMessages = sMessages & sS
                    
                    sS = ""
                    SendDeathText aList(i).amonIndex, sS
                    sPAMessage = sPAMessage & sS
                    sMessages = sMessages & sS
                    
                    sS = ""
                    DropMonItem aList(i).amonIndex, sS
                    sPAMessage = sPAMessage & sS
                    sMessages = sMessages & sS
                    
                    sPAMessage = sPAMessage & WHITE & "You have slain " & .mName & "!" & vbCrLf
                    sMessages = sMessages & BRIGHTGREEN & dbPlayers(aList(i).dbAIndex).sPlayerName & " has slain " & .mName & "!" & vbCrLf & WHITE
                    AddEXP dbPlayers(aList(i).dbAIndex).iIndex, aList(i).amonIndex, xxx
                    sPAMessage = sPAMessage & BRIGHTWHITE & "Your experience has increased by " & .mEXP & "." & GREEN & vbCrLf
                    If dbPlayers(aList(i).dbAIndex).lFamID <> 0 Then
                        If dbPlayers(aList(i).dbAIndex).sFamCustom <> "0" Then
                            sPAMessage = sPAMessage & BRIGHTWHITE & dbPlayers(aList(i).dbAIndex).sFamCustom & " the " & dbPlayers(aList(i).dbAIndex).sFamName & " gains " & CStr(xxx) & " experience." & GREEN & vbCrLf
                        Else
                            sPAMessage = sPAMessage & BRIGHTWHITE & "Your " & dbPlayers(aList(i).dbAIndex).sFamName & " gains " & CStr(xxx) & " experience." & GREEN & vbCrLf
                        End If
                    End If
                    AddMonsterRgn .mName
            
                    If Not modSC.FastStringComp(pWeapon(dbPlayers(aList(i).dbAIndex).iIndex).wSpellName, "") Then CleanUpSpells dbPlayers(aList(i).dbAIndex).iIndex
                    dbPlayers(aList(i).dbAIndex).dClassPoints = dbPlayers(aList(i).dbAIndex).dClassPoints + 0.1
                    
                    If WhoGoesFirst(PlayerAndMon, aList(i).dbAIndex, aList(i).amonIndex) = Monster Then
                        WrapAndSend dbPlayers(aList(i).dbAIndex).iIndex, sMonMessage & sPAMessage
                        SendToAllInRoom dbPlayers(aList(i).dbAIndex).iIndex, sMessages2 & sMessages, dbPlayers(aList(i).dbAIndex).lLocation
                    Else
                        WrapAndSend dbPlayers(aList(i).dbAIndex).iIndex, sPAMessage
                        SendToAllInRoom dbPlayers(aList(i).dbAIndex).iIndex, sMessages, dbPlayers(aList(i).dbAIndex).lLocation
                    End If
                    ClearOtherAttackers dbPlayers(aList(i).dbAIndex).iIndex, aList(i).amonIndex
                    sScripting dbPlayers(aList(i).dbAIndex).iIndex, , , 0, .mScript
                    ReSetMonsterID aList(i).dbAIndex
                    mRemoveItem aList(i).amonIndex
                    AmountMons = AmountMons - 1
                    GoTo nNext
                Else
                    With dbPlayers(aList(i).dbAIndex)
                        .lHP = .lHP - lMADamage
                    End With
                    If WhoGoesFirst(PlayerAndMon, aList(i).dbAIndex, aList(i).amonIndex) = Monster Then
                        WrapAndSend dbPlayers(aList(i).dbAIndex).iIndex, sMonMessage & sPAMessage
                        SendToAllInRoom dbPlayers(aList(i).dbAIndex).iIndex, sMessages2 & sMessages, dbPlayers(aList(i).dbAIndex).lLocation
                    Else
                        WrapAndSend dbPlayers(aList(i).dbAIndex).iIndex, sPAMessage & sMonMessage
                        SendToAllInRoom dbPlayers(aList(i).dbAIndex).iIndex, sMessages & sMessages2, dbPlayers(aList(i).dbAIndex).lLocation
                    End If
                End If
            End With
            CheckDeath dbPlayers(.dbAIndex).iIndex, lAMONINDEX:=aList(i).amonIndex
        ElseIf .bIs_Room_Spell = True Then
            If .bIs_PvP = False Then
                If .bIs_Player_Attacking = False And .bIs_Mon_Attacking = True Then
                    If aMons(.amonIndex).mLoc <> -1 Then
                        DoMonAttack .dbAIndex, .amonIndex, sMonMessage, sMessages2, lMADamage
                        With dbPlayers(.dbAIndex)
                            If .lHP - lMADamage < 1 Then
                                .lHP = .lHP - lMADamage
                                If CheckDeath(.iIndex, , True, sMonMessage, , sMessages2, bPFlag, , , aList(i).dbAIndex, aList(i).amonIndex) = False Then
                                    WrapAndSend .iIndex, sMonMessage
                                    SendToAllInRoom .iIndex, sMessages2, .lLocation
                                Else
                                    WrapAndSend .iIndex, sMonMessage
                                    SendToAllInRoom .iIndex, sMessages2, aMons(aList(i).amonIndex).mLoc
                                End If
                                If bPFlag Then RemoveFromParty .iIndex
                                GoTo nNext
                            Else
                                .lHP = .lHP - lMADamage
                            End If
                            WrapAndSend .iIndex, sMonMessage
                            SendToAllInRoom .iIndex, sMessages2, aMons(aList(i).amonIndex).mLoc
                        End With
                    End If
                End If
'            ElseIf .bIs_PvP = True And .bVic_Is_Attacking = True Then
'                If dbPlayers(.dbVIndex).lLocation = dbPlayers(.dbAIndex).lLocation Then
'                    DoPlayerAttackPlayer .dbVIndex, .dbAIndex, sPVMessage, sMessages, sPAMessage, lMADamage
'                    With dbPlayers(.dbAIndex)
'                        If .lHP - lPADamage < 1 Then
'                            .lHP = .lHP - lPADamage
'                            WrapAndSend dbPlayers(aList(i).dbVIndex).iIndex, sPVMessage
'                            WrapAndSend .iIndex, sPAMessage
'                            SendToAllInRoom dbPlayers(aList(i).dbVIndex).iIndex, sMessages, dbPlayers(aList(i).dbVIndex).lLocation, .iIndex
'                            CheckDeath .iIndex
'                            GoTo nNext
'                        Else
'                            .lHP = .lHP - lPADamage
'                        End If
'                        WrapAndSend dbPlayers(aList(i).dbVIndex).iIndex, sPVMessage
'                        WrapAndSend .iIndex, sPAMessage
'                        SendToAllInRoom dbPlayers(aList(i).dbVIndex).iIndex, sMessages, dbPlayers(aList(i).dbVIndex).lLocation, .iIndex
'                    End With
'                End If
            End If
            If .bIs_Player_Attacking = True Then
                If dbPlayers(.dbAIndex).dMonsterID = 99998 Then
                    DoRoomAttack .dbAIndex, sPAMessage, sMessages
                    If .bIs_Player_Attacking Then DoFamAttackMon .dbAIndex, .amonIndex, sPAMessage, sMessages, lPADamage
                End If
                WrapAndSend dbPlayers(.dbAIndex).iIndex, sPAMessage
                SendToAllInRoom dbPlayers(.dbAIndex).iIndex, sMessages, dbPlayers(.dbAIndex).lLocation
            End If
            '
            'dbPlayers(.dbAIndex).iHasAttacked = 0
        ElseIf .bIs_Player_Attacking And .bIs_Mon_Attacking = False Then
            DoPlayerAttackMon .dbAIndex, .amonIndex, sPAMessage, sMessages, lPADamage, bDidA
            DoFamAttackMon .dbAIndex, .amonIndex, sPAMessage, sMessages, lPADamage
            With aMons(.amonIndex)
                .mHP = .mHP - lPADamage
                If .mHP <= 0 Then
                    For j = 1 To UBound(aList)
                        With aList(j)
                            If .amonIndex = aList(i).amonIndex Then
                                '.amonIndex = 0
                                .bIs_Mon_Attacking = False
                                .bIs_Player_Attacking = False
                            End If
                        End With
                        If DE Then DoEvents
                    Next
                    DropMonGold aList(i).amonIndex, sS
                    sPAMessage = sPAMessage & sS
                    sMessages = sMessages & sS
                    
                    sS = ""
                    SendDeathText aList(i).amonIndex, sS
                    sPAMessage = sPAMessage & sS
                    sMessages = sMessages & sS
                    
                    sS = ""
                    DropMonItem aList(i).amonIndex, sS
                    sPAMessage = sPAMessage & sS
                    sMessages = sMessages & sS
                    
                    sPAMessage = sPAMessage & WHITE & "You have slain " & .mName & "!" & vbCrLf
                    sMessages = sMessages & BRIGHTGREEN & dbPlayers(aList(i).dbAIndex).sPlayerName & " has slain " & .mName & "!" & vbCrLf & WHITE
                    AddEXP dbPlayers(aList(i).dbAIndex).iIndex, aList(i).amonIndex
                    sPAMessage = sPAMessage & BRIGHTWHITE & "Your experience has increased by " & .mEXP & "." & GREEN & vbCrLf
                    If dbPlayers(aList(i).dbAIndex).lFamID <> 0 Then sPAMessage = sPAMessage & BRIGHTWHITE & "Your " & dbPlayers(aList(i).dbAIndex).sFamName & " gains " & CStr(.mEXP \ RndNumber(3, 15)) & " experience." & GREEN & vbCrLf
                    
                    AddMonsterRgn .mName
            
                    If Not modSC.FastStringComp(pWeapon(Index).wSpellName, "") Then CleanUpSpells dbPlayers(aList(i).dbAIndex).iIndex
                    dbPlayers(aList(i).dbAIndex).dClassPoints = dbPlayers(aList(i).dbAIndex).dClassPoints + 0.1
                    
                    
                    WrapAndSend dbPlayers(aList(i).dbAIndex).iIndex, sPAMessage & sMonMessage
                    SendToAllInRoom dbPlayers(aList(i).dbAIndex).iIndex, sMessages, dbPlayers(aList(i).dbAIndex).lLocation
                    ClearOtherAttackers dbPlayers(aList(i).dbAIndex).iIndex, aList(i).amonIndex
                    sScripting dbPlayers(aList(i).dbAIndex).iIndex, , , 0, .mScript
                    ReSetMonsterID aList(i).dbAIndex
                    mRemoveItem aList(i).amonIndex
                    AmountMons = AmountMons - 1
                    
                Else
                    WrapAndSend dbPlayers(aList(i).dbAIndex).iIndex, sMonMessage & sPAMessage
                    SendToAllInRoom dbPlayers(aList(i).dbAIndex).iIndex, sMessages, dbPlayers(aList(i).dbAIndex).lLocation
                End If
            End With
        ElseIf .bIs_PvP And .bIs_Room_Spell = False Then
            If .bVic_Is_Attacking Then
                If WhoGoesFirst(PlayerVsPlayer, .dbAIndex, , .dbVIndex) = PvPA Then
                    DoPlayerAttackPlayer .dbAIndex, .dbVIndex, sPAMessage, sMessages, sPVMessage, lPADamage
                    With dbPlayers(.dbVIndex)
                        If .lHP - lPADamage < 1 Then
                            .lHP = .lHP - lPADamage
                            WrapAndSend dbPlayers(aList(i).dbAIndex).iIndex, sPAMessage
                            WrapAndSend .iIndex, sPVMessage
                            SendToAllInRoom dbPlayers(aList(i).dbAIndex).iIndex, sMessages, dbPlayers(aList(i).dbAIndex).lLocation, .iIndex
                            CheckDeath .iIndex
                            GoTo nNext
                        End If
                    End With
                    DoPlayerAttackPlayer .dbVIndex, .dbAIndex, sPVMessage, sMessages, sPAMessage, lMADamage
                Else
                    DoPlayerAttackPlayer .dbVIndex, .dbAIndex, sPVMessage, sMessages, sPAMessage, lMADamage
                    With dbPlayers(.dbAIndex)
                        If .lHP - lPADamage < 1 Then
                            .lHP = .lHP - lPADamage
                            WrapAndSend dbPlayers(aList(i).dbVIndex).iIndex, sPAMessage
                            WrapAndSend .iIndex, sPVMessage
                            SendToAllInRoom dbPlayers(aList(i).dbVIndex).iIndex, sMessages, dbPlayers(aList(i).dbVIndex).lLocation, .iIndex
                            CheckDeath .iIndex
                            GoTo nNext
                        End If
                    End With
                    DoPlayerAttackPlayer .dbAIndex, .dbVIndex, sPAMessage, sMessages, sPVMessage, lPADamage
                End If
                With dbPlayers(.dbAIndex)
                    .lHP = .lHP - lMADamage
                    .iHasAttacked = 0
                End With
                With dbPlayers(.dbVIndex)
                    .lHP = .lHP - lPADamage
                    .iHasAttacked = 0
                End With
                WrapAndSend dbPlayers(.dbAIndex).iIndex, sPAMessage
                WrapAndSend dbPlayers(.dbVIndex).iIndex, sPVMessage
                SendToAllInRoom dbPlayers(.dbAIndex).iIndex, sMessages, dbPlayers(.dbAIndex).lLocation, dbPlayers(.dbVIndex).iIndex
                CheckDeath dbPlayers(.dbAIndex).iIndex
                CheckDeath dbPlayers(.dbVIndex).iIndex
            ElseIf .bVic_Is_Attacking = False And .bIs_Room_Spell = False Then
                DoPlayerAttackPlayer .dbAIndex, .dbVIndex, sPAMessage, sMessages, sPVMessage, lPADamage
                WrapAndSend dbPlayers(.dbAIndex).iIndex, sPAMessage
                WrapAndSend dbPlayers(.dbVIndex).iIndex, sPVMessage
                SendToAllInRoom dbPlayers(.dbAIndex).iIndex, sMessages, dbPlayers(.dbAIndex).lLocation, dbPlayers(.dbVIndex).iIndex
                With dbPlayers(.dbVIndex)
                    .lHP = .lHP - lPADamage
                    CheckDeath .iIndex
                    .iHasAttacked = 0
                End With
                dbPlayers(.dbAIndex).iHasAttacked = 0
            End If
        End If
    End With
nNext:
    If DE Then DoEvents
Next
'InitPvP
End Sub

Sub CleanUpAMons()
Dim i As Long
Dim j As Long
For i = LBound(aMons) To UBound(aMons)
    If j > AmountMons Then Exit For
    If aMons(i).mLoc <> -1 Or aMons(i).mLoc <> 0 Then
        If aMons(i).mIsAttacking Then
            CheckIsThere aMons(i).mList(0), i
            j = j + 1
        End If
    End If
    If DE Then DoEvents
Next
'If sIndex <> "" Then
'    SplitFast sIndex, atArr, ";"
'    For i = LBound(atArr) To UBound(atArr)
'        If atArr(i) <> "" Then
'            modUpdateDatabase.RemoveItemFromArray CLng(atArr(i)) - i
'        End If
'        If DE Then DoEvents
'    Next
'    sIndex = ""
'End If
End Sub

Sub InitPvP()
Dim i As Long
If iWhichDir = 0 Then
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If .iIndex <> 0 Then
                .lHasCasted = 0
                If .iPlayerAttacking <> 0 Then
                    If PvPStillThere(CLng(i), .iPlayerAttacking) = True Then
                        'PvPCombat Clng(i), .iPlayerAttacking
                        Combat .iIndex, i, CLng(.iPlayerAttacking)
                    End If
                ElseIf .dMonsterID = 99998 Then
                    If .iHasAttacked <> 1 Then
                        'RoomSpellAttack .iIndex, Clng(i)
                        dbPlayers(j).iHasAttacked = 1
                    Else
                        .iHasAttacked = 0
                    End If
                End If
            End If
        End With
        If DE Then DoEvents
    Next
Else
    For i = UBound(dbPlayers) To LBound(dbPlayers) Step -1
        With dbPlayers(i)
            If .iIndex <> 0 Then
                .lHasCasted = 0
                If .iPlayerAttacking <> 0 Then
                    If PvPStillThere(CLng(i), .iPlayerAttacking) = True Then
                        'PvPCombat Clng(i), .iPlayerAttacking
                        Combat .iIndex, i, CLng(.iPlayerAttacking)
                    End If
                ElseIf .dMonsterID = 99998 Then
                    If .iHasAttacked <> 1 Then
                        'RoomSpellAttack .iIndex, Clng(i)
                        dbPlayers(j).iHasAttacked = 1
                    Else
                        .iHasAttacked = 0
                    End If
                End If
            End If
        End With
        If DE Then DoEvents
    Next
End If
iWhichDir = RndNumber(0, 1)
End Sub

Sub Combat(Index As Long, Optional dbAIndex As Long = 0, Optional dbVIndex As Long = 0, Optional amonIndex As Long = -1, Optional RoomIt As Boolean = False)
Dim sPAMessage As String
Dim lPADamage As Long
Dim lMADamage As Long
Dim sMonMessage As String
Dim sPVMessage As String
Dim sMessages As String
Dim sS As String
Dim bDidA As Boolean
If dbAIndex <> 0 Then
    With dbPlayers(dbAIndex)
        If .iHasAttacked = 0 Then
            bDidA = True
            If dbVIndex <> 0 Then
                dbVIndex = GetPlayerIndexNumber(CLng(dbVIndex))
                DoPlayerAttackPlayer dbAIndex, dbVIndex, sPAMessage, sMessages, sPVMessage, lPADamage
                
            ElseIf amonIndex <> -1 And Not RoomIt Then
                If dbPlayers(dbAIndex).dMonsterID <> 99999 Then
                    
                    DoPlayerAttackMon dbAIndex, amonIndex, sPAMessage, sMessages, lPADamage, bDidA
                End If
                DoFamAttackMon dbAIndex, amonIndex, sPAMessage, sMessages, lPADamage
            ElseIf RoomIt = True Then
                DoRoomAttack dbAIndex, sPAMessage, sMessages
            End If
            If bDidA Then .iHasAttacked = 1
        Else
            .iHasAttacked = 0
        End If
    End With
End If
If dbVIndex <> 0 Then
    dbPlayers(dbVIndex).lHP = dbPlayers(dbVIndex).lHP - lPADamage
    If CheckDeath(Index, True, sPAMessage, dbPlayers(dbAIndex).iIndex, sPVMessage, dbPlayers(dbVIndex).iIndex, sMessages) = True Then
        Exit Sub
    End If
End If
If amonIndex <> -1 And Not RoomIt Then
    With aMons(amonIndex)
        .mHP = .mHP - lPADamage
        If .mHP <= 0 Then
            DropMonGold CLng(amonIndex), sS
            sPAMessage = sPAMessage & sS
            sMessages = sMessages & sS
            
            sS = ""
            SendDeathText CLng(amonIndex), sS
            sPAMessage = sPAMessage & sS
            sMessages = sMessages & sS
            
            sS = ""
            DropMonItem CLng(amonIndex), sS
            sPAMessage = sPAMessage & sS
            sMessages = sMessages & sS
            
            sPAMessage = sPAMessage & WHITE & "You have slain " & .mName & "!" & vbCrLf
            sMessages = sMessages & BRIGHTGREEN & dbPlayers(dbAIndex).sPlayerName & " has slain " & .mName & "!" & vbCrLf & WHITE
            AddEXP Index, CLng(amonIndex)
            sPAMessage = sPAMessage & BRIGHTWHITE & "Your experience has increased by " & .mEXP & "." & GREEN & vbCrLf
            If dbPlayers(dbAIndex).lFamID <> 0 Then sPAMessage = sPAMessage & BRIGHTWHITE & "Your " & dbPlayers(dbAIndex).sFamName & " gains " & CStr(.mEXP \ RndNumber(3, 15)) & " experience." & GREEN & vbCrLf
            
            AddMonsterRgn .mName
    
            If Not modSC.FastStringComp(pWeapon(Index).wSpellName, "") Then CleanUpSpells Index
            dbPlayers(dbAIndex).dClassPoints = dbPlayers(dbAIndex).dClassPoints + 0.1
            
            ClearOtherAttackers Index, CLng(amonIndex)
            ReSetMonsterID CLng(dbAIndex)
            WrapAndSend Index, sPAMessage & sMonMessage
            SendToAllInRoom Index, sMessages, dbPlayers(dbAIndex).lLocation
            sScripting Index, , , 0, aMons(amonIndex).mScript
            mRemoveItem CLng(amonIndex)
            AmountMons = AmountMons - 1
            dbPlayers(dbAIndex).iHasAttacked = 0
            Exit Sub
        End If
    End With
End If
If amonIndex <> -1 Then 'And Not RoomIt Then
    With aMons(amonIndex)
        If .mHasAttacked = 0 Then
            If .mList(0) = 0 Then
                modMonsters.AdjustMonList amonIndex
                If .mList(0) = 0 Then
                    .mIsAttacking = False
                    .mPlayerAttacking = 0
                Else
                    DoMonAttack GetPlayerIndexNumber(, , .mList(0)), amonIndex, sMonMessage, sMessages, lMADamage
                    .mHasAttacked = 1
                End If
            Else
                DoMonAttack GetPlayerIndexNumber(, , .mList(0)), amonIndex, sMonMessage, sMessages, lMADamage
                .mHasAttacked = 1
            End If
        Else
            .mHasAttacked = 0
        End If
    End With
End If
If amonIndex <> -1 And Not RoomIt Then
    With dbPlayers(dbAIndex)
        .lHP = .lHP - lMADamage
    End With
    WrapAndSend Index, Left$(sPAMessage & sMonMessage, Len(sPAMessage & sMonMessage) - 2) & vbCrLf
    SendToAllInRoom Index, Left$(sMessages, Len(sMessages) - 2) & vbCrLf, dbPlayers(dbAIndex).lLocation
    CheckDeath dbPlayers(dbAIndex).iIndex
End If
If dbVIndex <> 0 And Not RoomIt Then
    With dbPlayers(dbVIndex)
        .lHP = .lHP - lMADamage
    End With
    WrapAndSend Index, sPAMessage
    WrapAndSend dbPlayers(dbVIndex).iIndex, sPVMessage
    SendToAllInRoom Index, sMessages, dbPlayers(dbAIndex).lLocation, dbPlayers(dbVIndex).iIndex
    CheckDeath dbPlayers(dbVIndex).iIndex
End If
If RoomIt Then
    WrapAndSend Index, sPAMessage
    WrapAndSend dbPlayers(dbVIndex).iIndex, sPVMessage
    SendToAllInRoom Index, sMessages, dbPlayers(dbAIndex).lLocation, dbPlayers(dbVIndex).iIndex
End If
End Sub

Sub DoMonAttack(dbAIndex As Long, amonIndex As Long, ByRef Messages1 As String, ByRef Messages2 As String, ByRef Damage As Long)
Dim Mes() As String
Dim Mes2() As String
Dim Arr() As String
Dim MaxHit As Long
Dim PlayDodge As Long
Dim MessageID As Long
Dim MessageID2 As Long
Dim Chance As Long
Dim mRnd As Long
Dim lRoundDam As Long
Dim i As Long
Dim AFam As Boolean
Dim s As String
Dim bFound As Boolean
Dim lSpellCount As Long
Dim lEnergyUsed As Long
Dim lAttack As Long
Dim lAttempts As Long
Dim sAppend As String
Dim v As String
Dim q As String
Dim iSwings As Long
MaxHit = modGetData.GetMonsterMaxHit(amonIndex)

If modAttackHelpers.CheckIsThere(dbPlayers(dbAIndex).iIndex, amonIndex) = False Then
    Messages1 = "" '"!@"
    Messages2 = "" '"!@"
    Damage = 0
    Exit Sub
End If
For i = 0 To 4
    With aMons(amonIndex).mSpells(i)
        If .ldbSpellID <> 0 Then
            bFound = True
            lSpellCount = lSpellCount + 1
            .lCurrentCast = 0
        End If
    End With
    If DE Then DoEvents
Next
GoAgain:
Damage = Damage + lRoundDam
lRoundDam = 0
If DE Then DoEvents
If lAttempts >= 4 Then GoTo Done
lAttack = RndNumber(1, lSpellCount + 1)
If lAttack = lSpellCount + 1 Then
    If lEnergyUsed + aMons(amonIndex).mPEnergy > aMons(amonIndex).mEnergy Then
        lAttempts = lAttempts + 1
        GoTo GoAgain
    End If
DoPAttack:
    If aMons(amonIndex).mWeapon.iID = 0 Then
        SplitFast aMons(amonIndex).mMessage, Mes, ":"
        Mes2 = Mes
    Else
        SplitFast aMons(amonIndex).mWeapon.sMessageV, Mes, ":"
        SplitFast aMons(amonIndex).mWeapon.sMessage2, Mes2, ":"
    End If
    iSwings = iSwings + 1
    PlayDodge = modGetData.GetPlayerDodge(dbAIndex)
    MessageID = RndNumber(0, UBound(Mes))
    MessageID2 = RndNumber(0, UBound(Mes2))
    Chance = RndNumber(1, 100)
    If dbPlayers(dbAIndex).iDropped > 0 Then
        Chance = Chance + RndNumber(0, 100)
        PlayDodge = PlayDodge - RndNumber(0, 100)
    End If
    If dbPlayers(dbAIndex).iDebugMode = 1 Then
        Messages1 = Messages1 & BRIGHTWHITE & "[Chance=" & CStr(Chance) & "][PlayDodge=" & CStr(PlayDodge) & "]"
    End If
    If Chance = 1 Then
        lRoundDam = 0
        Messages1 = Messages1 & MAGNETA & aMons(amonIndex).mName & " flinches!" & vbCrLf & WHITE
        Messages2 = Messages2 & MAGNETA & aMons(amonIndex).mName & " flinches!" & vbCrLf & WHITE
    ElseIf Chance <= MaxHit Then
        If aMons(amonIndex).mWeapon.iID = 0 Then
            lRoundDam = RndNumber(CDbl(aMons(amonIndex).mMin), CDbl(aMons(amonIndex).mMax))
            sAppend = ""
        Else
            Erase Arr
            SplitFast aMons(amonIndex).mWeapon.sDamage, Arr, ":"
            lRoundDam = RndNumber(CDbl(Arr(0)), CDbl(Arr(1)))
            sAppend = " with their " & aMons(amonIndex).mWeapon.sItemName
        End If
        mRnd = RndNumber(1, 100)
        If dbPlayers(dbAIndex).iDebugMode = 1 Then
            Messages1 = Messages1 & BRIGHTWHITE & "[mRnd=" & CStr(mRnd) & "]"
        End If
        If mRnd <= Chance \ 8 Then
            Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " misses their attack on you!" & vbCrLf & WHITE
            Messages2 = Messages2 & LIGHTBLUE & aMons(amonIndex).mName & " misses their attack on " & dbPlayers(dbAIndex).sPlayerName & "!" & WHITE & vbCrLf
            lRoundDam = 0
        ElseIf mRnd <= PlayDodge Then
            Messages1 = Messages1 & LIGHTBLUE & "You dodge " & aMons(amonIndex).mName & "'s attack!" & vbCrLf & WHITE
            Messages2 = Messages2 & LIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " dodges " & aMons(amonIndex).mName & "'s attack!" & vbCrLf & WHITE
        ElseIf mRnd > PlayDodge Then
            If ((Chance > (MaxHit - 5)) And Chance < (MaxHit + 5)) Or (Chance > 101) Then
                lRoundDam = lRoundDam * 3
                AFam = False
                If dbPlayers(dbAIndex).sFamName <> "0" Then
                    With dbPlayers(dbAIndex)
                        MaxHit = .lFamCHP + .lFamMHP
                        Chance = RndNumber(1, CDbl(MaxHit))
                        If .iHorse > 0 Then Chance = Chance + RndNumber(1, 700)
                        If Chance - .lFamAcc > MaxHit \ 2 Then
                            AFam = True
                            If .sFamCustom <> "0" Then
                                Messages1 = Messages1 & YELLOW & aMons(amonIndex).mName & " " & Mes(MessageID) & " " & .sFamCustom & " the " & .sFamName & " for " & lRoundDam & " damage" & sAppend & "!" & WHITE & vbCrLf
                                Messages2 = Messages2 & YELLOW & aMons(amonIndex).mName & " " & Mes2(MessageID2) & " " & .sFamCustom & " the " & .sFamName & " for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                            Else
                                Messages1 = Messages1 & YELLOW & aMons(amonIndex).mName & " " & Mes(MessageID) & " your " & .sFamName & " for " & lRoundDam & " damage" & sAppend & "!" & WHITE & vbCrLf
                                Messages2 = Messages2 & YELLOW & aMons(amonIndex).mName & " " & Mes2(MessageID2) & " " & .sPlayerName & "'s " & .sFamName & " for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                            End If
                            modAttackHelpers.SubtractFamHP dbAIndex, lRoundDam, Messages1, Messages2
                            lRoundDam = 0
                        End If
                    End With
                End If
                If AFam = False Then
                    lRoundDam = lRoundDam * 1.5
                    'set the messages
                    If lRoundDam <> 0 Then lRoundDam = lRoundDam - (dbPlayers(dbAIndex).iAC \ 14)
                    If lRoundDam < 1 And lRoundDam <> 0 Then lRoundDam = 1
                    Messages1 = Messages1 & YELLOW & aMons(amonIndex).mName & " critically " & Mes(MessageID) & " you for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                    Messages2 = Messages2 & YELLOW & aMons(amonIndex).mName & " critically " & Mes2(MessageID2) & " " & dbPlayers(dbAIndex).sPlayerName & " for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                    If RndNumber(0, CDbl(dbPlayers(dbAIndex).iCha)) < (dbPlayers(dbAIndex).iCha \ 3) Then
                        s = modGetData.GetUnformatedStringFromID(CLng(dbAIndex), modGetData.GetHitPositionID)
                        If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                    End If
                End If
            Else
                If CheckShield(dbPlayers(dbAIndex).iIndex, dbAIndex) = True Then
                    Messages1 = Messages1 & BRIGHTMAGNETA & aMons(amonIndex).mName & " attempts to hit you, but you block the hit with your shield!" & vbCrLf & WHITE
                    Messages2 = Messages2 & BRIGHTMAGNETA & aMons(amonIndex).mName & "'s swing is blocked by " & dbPlayers(dbAIndex).sPlayerName & "'s shield!" & vbCrLf & WHITE
                    If RndNumber(0, CDbl(dbPlayers(dbAIndex).iCha)) < (dbPlayers(dbAIndex).iCha \ 3) Then
                        s = dbPlayers(dbAIndex).sShield
                        If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                    End If
                Else
                    With dbPlayers(dbAIndex)
                        If .iHorse > 0 Then
                            If RndNumber(1, 100) > 65 And .sFamName <> "0" Then
                                If .sFamCustom <> "0" Then
                                    Messages1 = Messages1 & YELLOW & aMons(amonIndex).mName & " " & Mes(MessageID) & " " & .sFamCustom & " the " & .sFamName & " for " & lRoundDam & " damage" & sAppend & "!" & WHITE & vbCrLf
                                    Messages2 = Messages2 & YELLOW & aMons(amonIndex).mName & " " & Mes2(MessageID2) & " " & .sFamCustom & " the " & .sFamName & " for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                                Else
                                    Messages1 = Messages1 & YELLOW & aMons(amonIndex).mName & " " & Mes(MessageID) & " your " & .sFamName & " for " & lRoundDam & " damage" & sAppend & "!" & WHITE & vbCrLf
                                    Messages2 = Messages2 & YELLOW & aMons(amonIndex).mName & " " & Mes2(MessageID2) & " " & .sPlayerName & "'s " & .sFamName & " for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                                End If
                                modAttackHelpers.SubtractFamHP dbAIndex, lRoundDam, Messages1, Messages2
                                lRoundDam = 0
                            Else
                                If lRoundDam <> 0 Then lRoundDam = lRoundDam - (dbPlayers(dbAIndex).iAC \ 14)
                                If lRoundDam < 1 And lRoundDam <> 0 Then lRoundDam = 1
                                Messages1 = Messages1 & RED & aMons(amonIndex).mName & " " & Mes(MessageID) & " you for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                                Messages2 = Messages2 & RED & aMons(amonIndex).mName & " " & Mes2(MessageID2) & " " & dbPlayers(dbAIndex).sPlayerName & " for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                            End If
                        Else
                            If lRoundDam <> 0 Then lRoundDam = lRoundDam - (dbPlayers(dbAIndex).iAC \ 14)
                            If lRoundDam < 1 And lRoundDam <> 0 Then lRoundDam = 1
                            Messages1 = Messages1 & RED & aMons(amonIndex).mName & " " & Mes(MessageID) & " you for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                            Messages2 = Messages2 & RED & aMons(amonIndex).mName & " " & Mes2(MessageID2) & " " & dbPlayers(dbAIndex).sPlayerName & " for " & lRoundDam & " damage" & sAppend & "!" & vbCrLf & WHITE
                        End If
                        If RndNumber(0, CDbl(.iCha)) < (.iCha \ 3) Then
                            s = modGetData.GetUnformatedStringFromID(CLng(dbAIndex), modGetData.GetHitPositionID)
                            If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                        End If
                    End With
                End If
            End If
        End If
    ElseIf mRnd = PlayDodge Then
        Messages1 = Messages1 & YELLOW & aMons(amonIndex).mName & " severly misses you, and falls to the ground!" & vbCrLf & WHITE
        Messages2 = Messages2 & YELLOW & aMons(amonIndex).mName & " fumbles!" & vbCrLf & WHITE
        GoTo Done
    Else
        lRoundDam = 0
        Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " swings at you, but misses!" & vbCrLf & WHITE
        Messages2 = Messages2 & LIGHTBLUE & aMons(amonIndex).mName & " swings at " & dbPlayers(dbAIndex).sPlayerName & ", but misses!" & vbCrLf & WHITE
    End If
    lEnergyUsed = lEnergyUsed + aMons(amonIndex).mPEnergy
    If lEnergyUsed < aMons(amonIndex).mEnergy Then GoTo GoAgain
Else
    If lEnergyUsed + aMons(amonIndex).mSpells(lAttack - 1).lEnergy > aMons(amonIndex).mEnergy Then
        lAttempts = lAttempts + 1
        GoTo GoAgain
    End If
    If aMons(amonIndex).mSpells(lAttack - 1).lCurrentCast >= aMons(amonIndex).mSpells(lAttack - 1).lMaxCast Then
        lAttempts = lAttempts + 1
        GoTo GoAgain
    ElseIf aMons(amonIndex).mSpells(lAttack - 1).lCurrentCast >= aMons(amonIndex).mSpells(lAttack - 1).lCastPerRound Then
        lAttempts = lAttempts + 1
        GoTo GoAgain
    End If
    iSwings = iSwings + 1
    ReDim Mes(1)
    PlayDodge = modGetData.GetPlayersTotalMR(dbAIndex)
    Chance = RndNumber(1, 100)
    If dbPlayers(dbAIndex).iDropped > 0 Then
        Chance = Chance + RndNumber(0, 100)
        PlayDodge = PlayDodge - RndNumber(0, 100)
    End If
    If dbPlayers(dbAIndex).iDebugMode = 1 Then
        Messages1 = Messages1 & BRIGHTWHITE & "[Chance=" & CStr(Chance) & "][PlayDodge=" & CStr(PlayDodge) & "]"
    End If
    If aMons(amonIndex).mSpells(lAttack - 1).lEnergy <= 0 Then Chance = Chance + 1000: PlayDodge = PlayDodge - 1000
    Mes(0) = dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sMessageV
    Mes(1) = dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sMessage2
    If dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).iUse = 1 Then
        If Chance = 1 Then
            lRoundDam = 0
            Messages1 = Messages1 & MAGNETA & aMons(amonIndex).mName & " fumbles their words!" & vbCrLf & WHITE
            Messages2 = Messages2 & MAGNETA & aMons(amonIndex).mName & " fumbles their words!" & vbCrLf & WHITE
        ElseIf Chance >= MaxHit Then
            lRoundDam = RndNumber(CDbl(dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).lMinDam), CDbl(modGetData.GetSpellDam(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID, aMons(amonIndex).mLevel)))
            If dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).lElement <> -1 Then
                lRoundDam = lRoundDam - modResist.GetResistValue(dbAIndex, dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).lElement)
            End If
            mRnd = RndNumber(1, 100)
            If dbPlayers(dbAIndex).iDebugMode = 1 Then
                Messages1 = Messages1 & BRIGHTWHITE & "[mRnd=" & CStr(mRnd) & "]"
            End If
            If aMons(amonIndex).mSpells(lAttack - 1).lEnergy <= 0 Then mRnd = mRnd + 1000: Chance = Chance - 1000
            If mRnd <= Chance Then
                Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on you!" & vbCrLf & WHITE
                Messages2 = Messages2 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on " & dbPlayers(dbAIndex).sPlayerName & "!" & WHITE & vbCrLf
                lRoundDam = 0
            ElseIf mRnd <= PlayDodge Then
                Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " tries to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on you, but you resist!" & vbCrLf & WHITE
                Messages2 = Messages2 & LIGHTBLUE & aMons(amonIndex).mName & " tries to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on " & dbPlayers(dbAIndex).sPlayerName & ", but " & LCaseFast(modGetData.GetGenderDesc(dbAIndex)) & " resist!" & WHITE & vbCrLf
                lRoundDam = 0
            ElseIf mRnd > PlayDodge Then
                
                
                If (((Chance > (MaxHit - 10)) And Chance < (MaxHit + 10)) Or (Chance > 101)) And aMons(amonIndex).mSpells(lAttack - 1).lEnergy > 0 Then
                    lRoundDam = lRoundDam * 3
                    v = Mes(0)
                    q = Mes(1)
                    v = ReplaceFast(v, "<%v>", "you")
                    v = ReplaceFast(v, "<%c>", aMons(amonIndex).mName)
                    v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                    v = ReplaceFast(v, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
                    q = ReplaceFast(q, "<%v>", dbPlayers(dbAIndex).sPlayerName)
                    q = ReplaceFast(q, "<%c>", aMons(amonIndex).mName)
                    q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                    q = ReplaceFast(q, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
                    AFam = False
                    If dbPlayers(dbAIndex).sFamName <> "0" Then
                        With dbPlayers(dbAIndex)
                            MaxHit = .lFamCHP + .lFamMHP
                            Chance = RndNumber(1, CDbl(MaxHit))
                            If .iHorse > 0 Then Chance = Chance + RndNumber(1, 700)
                            If Chance - .lFamAcc > MaxHit \ 2 Then
                                AFam = True
                                If .sFamCustom <> "0" Then
                                    q = ReplaceFast(q, "you", .sFamCustom & " the " & .sFamName)
                                    v = ReplaceFast(v, .sPlayerName, .sFamCustom & " the " & .sFamName)
                                Else
                                    q = ReplaceFast(q, "you", .sFamName)
                                    v = ReplaceFast(v, .sPlayerName, .sFamName)
                                End If
                                Messages1 = Messages1 & YELLOW & v & WHITE & vbCrLf
                                Messages2 = Messages2 & YELLOW & q & vbCrLf & WHITE
                                modAttackHelpers.SubtractFamHP dbAIndex, lRoundDam, Messages1, Messages2
                                lRoundDam = 0
                            End If
                        End With
                    End If
                    If AFam = False Then
                        lRoundDam = lRoundDam * 1.5
                        v = Mes(0)
                        q = Mes(1)
                        v = ReplaceFast(v, "<%v>", "you")
                        v = ReplaceFast(v, "<%c>", aMons(amonIndex).mName)
                        v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                        v = ReplaceFast(v, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
                        q = ReplaceFast(q, "<%v>", dbPlayers(dbAIndex).sPlayerName)
                        q = ReplaceFast(q, "<%c>", aMons(amonIndex).mName)
                        q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                        q = ReplaceFast(q, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
                        'set the messages
                        Messages1 = Messages1 & YELLOW & "Critically, " & v & vbCrLf & WHITE
                        Messages2 = Messages2 & YELLOW & "Critically " & q & vbCrLf & WHITE
                    End If
                Else
                    v = Mes(0)
                    q = Mes(1)
                    v = ReplaceFast(v, "<%v>", "you")
                    v = ReplaceFast(v, "<%c>", aMons(amonIndex).mName)
                    v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                    v = ReplaceFast(v, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
                    q = ReplaceFast(q, "<%v>", dbPlayers(dbAIndex).sPlayerName)
                    q = ReplaceFast(q, "<%c>", aMons(amonIndex).mName)
                    q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                    q = ReplaceFast(q, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
                    If aMons(amonIndex).mSpells(lAttack - 1).lEnergy > 0 Then
                        With dbPlayers(dbAIndex)
                            If .iHorse > 0 Then
                                If RndNumber(1, 100) > 65 Then
                                    If .sFamCustom <> "0" Then
                                        q = ReplaceFast(q, "you", .sFamCustom & " the " & .sFamName)
                                        v = ReplaceFast(v, .sPlayerName, .sFamCustom & " the " & .sFamName)
                                    Else
                                        q = ReplaceFast(q, "you", .sFamName)
                                        v = ReplaceFast(v, .sPlayerName, .sFamName)
                                    End If
                                    Messages1 = Messages1 & YELLOW & v & WHITE & vbCrLf
                                    Messages2 = Messages2 & YELLOW & q & vbCrLf & WHITE
                                    modAttackHelpers.SubtractFamHP dbAIndex, lRoundDam, Messages1, Messages2
                                    lRoundDam = 0
                                Else
                                    Messages1 = Messages1 & RED & v & vbCrLf & WHITE
                                    Messages2 = Messages2 & RED & q & vbCrLf & WHITE
                                End If
                            Else
                                Messages1 = Messages1 & RED & v & vbCrLf & WHITE
                                Messages2 = Messages2 & RED & q & vbCrLf & WHITE
                            End If
                        End With
                    Else
                        With dbPlayers(dbAIndex)
                            lRoundDam = 0
                            If .iHorse > 0 Then
                                If RndNumber(1, 100) > 65 Then
                                    If .sFamCustom <> "0" Then
                                        q = ReplaceFast(q, "you", .sFamCustom & " the " & .sFamName)
                                        v = ReplaceFast(v, .sPlayerName, .sFamCustom & " the " & .sFamName)
                                    Else
                                        q = ReplaceFast(q, "you", .sFamName)
                                        v = ReplaceFast(v, .sPlayerName, .sFamName)
                                    End If
                                    Messages1 = Messages1 & BRIGHTBLUE & v & WHITE & vbCrLf
                                    Messages2 = Messages2 & BRIGHTBLUE & q & vbCrLf & WHITE
                                    modAttackHelpers.SubtractFamHP dbAIndex, lRoundDam, Messages1, Messages2
                                    lRoundDam = 0
                                Else
                                    Messages1 = Messages1 & BRIGHTBLUE & v & vbCrLf & WHITE
                                    Messages2 = Messages2 & BRIGHTBLUE & q & vbCrLf & WHITE
                                End If
                            Else
                                Messages1 = Messages1 & BRIGHTBLUE & v & vbCrLf & WHITE
                                Messages2 = Messages2 & BRIGHTBLUE & q & vbCrLf & WHITE
                            End If
                        End With
                    End If
                End If
            End If
        ElseIf mRnd = PlayDodge Then
            Messages1 = Messages1 & YELLOW & aMons(amonIndex).mName & " forgets what they were casting!" & vbCrLf & WHITE
            Messages2 = Messages2 & YELLOW & aMons(amonIndex).mName & " looks like they forgot what they were doing!" & vbCrLf & WHITE
            GoTo Done
        Else
            lRoundDam = 0
            Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on you!" & vbCrLf & WHITE
            Messages2 = Message2 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on " & dbPlayers(dbAIndex).sPlayerName & "!" & WHITE & vbCrLf
        End If
    ElseIf dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).iUse = 0 Then
        If Chance >= MaxHit Then
            lRoundDam = RndNumber(CDbl(dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).lMinDam), CDbl(modGetData.GetSpellDam(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID, aMons(amonIndex).mLevel)))
            v = Mes(0)
            q = Mes(1)
            v = ReplaceFast(v, "<%v>", "itself")
            v = ReplaceFast(v, "<%c>", aMons(amonIndex).mName)
            v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
            v = ReplaceFast(v, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
            q = ReplaceFast(q, "<%v>", "itself")
            q = ReplaceFast(q, "<%c>", aMons(amonIndex).mName)
            q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
            q = ReplaceFast(q, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
            aMons(amonIndex).mHP = aMons(amonIndex).mHP + lRoundDam
            If aMons(amonIndex).mHP > aMons(amonIndex).mMaxHP Then aMons(amonIndex).mHP = aMons(amonIndex).mMaxHP
            Messages1 = Messages1 & BRIGHTBLUE & v & vbCrLf & WHITE
            Messages2 = Messages2 & BRIGHTBLUE & q & vbCrLf & WHITE
            lRoundDam = 0
        Else
            lRoundDam = 0
            Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on itself!" & vbCrLf & WHITE
            Messages2 = Messages2 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on itself!" & WHITE & vbCrLf
        End If
'    ElseIf dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).iUse = 3 Then
'        If Chance = 1 Then
'            lRoundDam = 0
'            Messages1 = Messages1 & MAGNETA & aMons(amonIndex).mName & " fumbles their words!" & vbCrLf & WHITE
'            Messages2 = Messages2 & MAGNETA & aMons(amonIndex).mName & " fumbles their words!" & vbCrLf & WHITE
'        ElseIf Chance >= MaxHit Then
'
'
''            lRoundDam = RndNumber(CDbl(dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).lMinDam), CDbl(modgetdata.GetSpellDam(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID, aMons(amonIndex).mLevel)))
''            If dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).lElement <> -1 Then
''                lRoundDam = lRoundDam - modResist.GetResistValue(dbAIndex, dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).lElement)
''            End If
'            mRnd = RndNumber(1, 100)
'            If aMons(amonIndex).mSpells(lAttack - 1).lEnergy <= 0 Then mRnd = mRnd + 1000: Chance = Chance - 1000
'            If mRnd <= Chance Then
'                Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on you!" & vbCrLf & WHITE
'                Messages2 = Messages2 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on " & dbPlayers(dbAIndex).sPlayerName & "!" & WHITE & vbCrLf
'                lRoundDam = 0
'            ElseIf mRnd <= PlayDodge Then
'                Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " tries to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on you, but you resist!" & vbCrLf & WHITE
'                Messages2 = Messages2 & LIGHTBLUE & aMons(amonIndex).mName & " tries to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on " & dbPlayers(dbAIndex).sPlayerName & ", but " & LCaseFast(modgetdata.GetGenderDesc(dbAIndex)) & " resist!" & WHITE & vbCrLf
'                lRoundDam = 0
'            ElseIf mRnd > PlayDodge Then
'                    lRoundDam = 0
'                    v = Mes(0)
'                    q = Mes(1)
'                    v = ReplaceFast(v, "<%v>", "you")
'                    v = ReplaceFast(v, "<%c>", aMons(amonIndex).mName)
'                    v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
'                    v = ReplaceFast(v, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
'                    q = ReplaceFast(q, "<%v>", dbPlayers(dbAIndex).sPlayerName)
'                    q = ReplaceFast(q, "<%c>", aMons(amonIndex).mName)
'                    q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
'                    q = ReplaceFast(q, "<%s>", dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName)
'
'                        With dbPlayers(dbAIndex)
'
'                                Messages1 = Messages1 & RED & v & vbCrLf & WHITE
'                                Messages2 = Messages2 & RED & q & vbCrLf & WHITE
'                                modSpells.DoNonCombatSpell .iIndex, dbAIndex, aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID, , , True, aMons(amonIndex).mName
'                        End With
'                End If
'            End If
'        ElseIf mRnd = PlayDodge Then
'            lRoundDam = 0
'            Messages1 = Messages1 & YELLOW & aMons(amonIndex).mName & " forgets what they were casting!" & vbCrLf & WHITE
'            Messages2 = Messages2 & YELLOW & aMons(amonIndex).mName & " looks like they forgot what they were doing!" & vbCrLf & WHITE
'            GoTo Done
'        Else
'            lRoundDam = 0
'            Messages1 = Messages1 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on itself!" & vbCrLf & WHITE
'            Messages2 = Messages2 & LIGHTBLUE & aMons(amonIndex).mName & " fails to cast " & dbSpells(aMons(amonIndex).mSpells(lAttack - 1).ldbSpellID).sSpellName & " on itself!" & WHITE & vbCrLf
'        End If
    End If
    lEnergyUsed = lEnergyUsed + aMons(amonIndex).mSpells(lAttack - 1).lEnergy
    aMons(amonIndex).mSpells(lAttack - 1).lCurrentCast = aMons(amonIndex).mSpells(lAttack - 1).lCurrentCast + 1
   ' aMons(amonIndex).mSpells(lAttack - 1).lCastPerRound = aMons(amonIndex).mSpells(lAttack - 1).lCastPerRound + 1
    If lEnergyUsed < aMons(amonIndex).mEnergy Then GoTo GoAgain
End If
Done:
Damage = Damage + lRoundDam
If iSwings < 1 Then GoTo DoPAttack
If Damage < 0 Then Damage = 0
End Sub

Sub DoPlayerAttackMon(dbAIndex As Long, amonIndex As Long, ByRef Messages1 As String, ByRef Messages2 As String, ByRef Damage As Long, ByRef bDid As Boolean)
Dim pCrits As Long
Dim pAcc As Long
Dim Mes() As String
Dim Mes2() As String
Dim MaxHit As Long
Dim MonDodge As Long
Dim Arr() As String
Dim Arr2() As String
Dim i As Long
Dim t As Long
Dim Swings As Long
Dim bOutOfMana As Boolean
Dim Index As Long
Dim lRoundDam As Long
Dim MessageID As Long
Dim MessageID2 As Long
Dim Chance As Long
Dim mRnd As Long
Dim bCrit As Boolean
Dim s As String
Dim v As String
Dim q As String
Dim f As Long
Dim fFlag As Boolean
Index = dbPlayers(dbAIndex).iIndex
DUALWIELDGOAGAIN:
SetWeaponStats Index, dbAIndex, fFlag
If dbPlayers(dbAIndex).iDropped > 0 Then Exit Sub
If modSC.FastStringComp(pWeapon(Index).wSpellName, "") Then
    If pWeapon(Index).wBullets = 0 Then
        If pWeapon(Index).wMag = 0 Then
            bOutOfMana = True
            Messages1 = BRIGHTBLUE & "You are out of ammo!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Sub
        Else
            If (dbPlayers(dbAIndex).lMana < pWeapon(Index).wBMana) Then
                bOutOfMana = True
                Messages1 = Messages1 & BRIGHTBLUE & "You are out of mana!" & WHITE & vbCrLf
                Exit Sub
            End If
        End If
    End If
    SplitFast pWeapon(Index).wMessage, Mes, ":"
    SplitFast pWeapon(Index).wMessage2, Mes2, ":"
    pCrits = dbPlayers(dbAIndex).iCrits
    pAcc = dbPlayers(dbAIndex).iAcc
    MaxHit = modGetData.GetPlayerMaxHit(dbAIndex) + pAcc
    If MaxHit > 98 Then MaxHit = 98
    MonDodge = modGetData.GetMonsterDodge(amonIndex)
    Swings = modGetData.GetPlayerSwings(dbAIndex)
Else
    bOutOfMana = False
    If dbPlayers(dbAIndex).lMana < pWeapon(Index).wMana Then
        bOutOfMana = True
        Messages1 = BRIGHTBLUE & "You are out of mana!" & WHITE & vbCrLf
        X(Index) = ""
        dbPlayers(dbAIndex).iCasting = 0
        SpellCombat(Index) = False
        Exit Sub
    Else
        MaxHit& = modGetData.GetSpellChance(Index)
        ReDim Mes(2) As String
        Mes(0) = pWeapon(Index).wMessage
        Mes(1) = pWeapon(Index).wMessage2
        Swings = pWeapon(Index).wCast
    End If
End If
If aMons(amonIndex).mIs_Being_Attacked = True And bOutOfMana = False Then
    If dbPlayers(dbAIndex).iIsBSing = 1 Then
        MessageID = RndNumber(0, CDbl(UBound(Mes)))
        MessageID2 = RndNumber(0, CDbl(UBound(Mes2)))
        Chance = RndNumber(1, 100 - (dbPlayers(dbAIndex).iAcc + dbPlayers(dbAIndex).iLevel))
        If Chance < MaxHit Then
            mRnd = RndNumber(1, 100 + (dbPlayers(dbAIndex).iAcc + dbPlayers(dbAIndex).iLevel))
            If mRnd > MonDodge Then
                'BS Hit
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
                lRoundDam = lRoundDam * 6
                Messages1 = BRIGHTRED & "You " & GREEN & "surprise " & BRIGHTRED & Mes(MessageID) & " " & aMons(amonIndex).mName & " for " & CStr(lRoundDam) & "." & vbCrLf
                Messages2 = BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " surprise " & BRIGHTRED & Mes2(MessageID2) & "s " & aMons(amonIndex).mName & " for " & CStr(lRoundDam) & "." & vbCrLf
            Else
                Messages1 = Messages1 & BRIGHTLIGHTBLUE & aMons(amonIndex).mName & " dodges your backstab!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTLIGHTBLUE & aMons(amonIndex).mName & " dodges " & dbPlayers(dbAIndex).sPlayerName & "'s backstab!" & vbCrLf & WHITE
                lRoundDam = 0
            End If
        Else
            lRoundDam = 0
            Messages1 = Messages1 & BRIGHTLIGHTBLUE & "You miss " & aMons(amonIndex).mName & " with your backstab!" & vbCrLf & WHITE
            Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " misses their backstab on " & aMons(amonIndex).mName & "." & vbCrLf & WHITE
        End If
        dbPlayers(dbAIndex).iIsBSing = 0
        Damage = lRoundDam
    Else
        For i = 1 To Swings
            MessageID = RndNumber(0, 2)
            Chance = RndNumber(1, 100)
            If Chance >= 100 And dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
                lRoundDam = lRoundDam * 5
                Messages1 = BRIGHTRED & "You " & BRIGHTYELLOW & "critically and severly " & BRIGHTRED & Mes(MessageID) & " " & aMons(amonIndex).mName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                Messges2 = BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " critically and severly " & Mes(MessageID) & "s " & aMons(amonIndex).mName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                If pWeapon(Index).wCastSp <> "" Then
                    'If RndNumber(1, 100) > 50 Then
                    Erase Arr
                    Erase Arr2
                    SplitFast pWeapon(Index).wCastSp, Arr, ";"
                    SplitFast pWeapon(Index).wCastSpPer, Arr2, ";"
                    If UBound(Arr2) < UBound(Arr) Then ReDim Preserve Arr2(UBound(Arr))
                    For t = LBound(Arr) To UBound(Arr)
                        If Arr(t) <> "" Then
                            If Arr2(t) <> "" Then
                                If RndNumber(1, 100) < Val(Arr2(t)) Then
                                    With dbSpells(Val(Arr(t)))
                                        f = RndNumber(CDbl(.lMinDam), CDbl(.lMaxDam))
                                        v = .sMessage2
                                        q = .sMessage2
                                        v = ReplaceFast(v, "<%v>", aMons(amonIndex).mName)
                                        v = ReplaceFast(v, "<%c>", pWeapon(Index).wWeaponName)
                                        v = ReplaceFast(v, "<%d>", CStr(f))
                                        v = ReplaceFast(v, "<%s>", .sSpellName)
                                        q = ReplaceFast(q, "<%v>", aMons(amonIndex).mName)
                                        q = ReplaceFast(q, "<%c>", pWeapon(Index).wWeaponName)
                                        q = ReplaceFast(q, "<%d>", CStr(f))
                                        q = ReplaceFast(q, "<%s>", .sSpellName)
                                        Messages1 = Messages1 & BGGREEN & v & vbCrLf & WHITE
                                        Messages2 = Messages2 & BGGREEN & q & vbCrLf & WHITE
                                        lRoundDam = lRoundDam + f
                                    End With
                                End If
                            End If
                        End If
                        If DE Then DoEvents
                    Next
                    'End If
                End If
                If RndNumber(0, CDbl(dbPlayers(dbAIndex).iCha)) < (dbPlayers(dbAIndex).iCha \ 3) Then
                    s = dbPlayers(dbAIndex).sWeapon
                    If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                End If
            ElseIf Chance >= 100 And dbPlayers(dbAIndex).iCasting <> 0 Then
                '===SPELL===
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
                lRoundDam = lRoundDam * 5
                v = Mes(0)
                q = Mes(1)
                v = ReplaceFast(v, "<%v>", aMons(amonIndex).mName)
                v = ReplaceFast(v, "<%c>", "You")
                v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
                q = ReplaceFast(q, "<%v>", aMons(amonIndex).mName)
                q = ReplaceFast(q, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                q = ReplaceFast(q, "<%s>", pWeapon(Index).wSpellName)
                Messages1 = Messages1 & BGBRIGHTYELLOW & "Critically and severly" & BRIGHTRED & ", " & v & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTRED & "Critically and severly, " & q & vbCrLf & WHITE
            ElseIf (Chance <= 1 + (aMons(amonIndex).mAc \ 14)) And dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                lRoundDam = 0
                Messages1 = Messages1 & BRIGHTMAGNETA & "You flinch as you attempt to hit " & aMons(amonIndex).mName & ", and cause no damage!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTMAGNETA & dbPlayers(dbAIndex).sPlayerName & " flinches!" & vbCrLf & WHITE
            ElseIf Chance <= 2 And dbPlayers(dbAIndex).iCasting <> 0 Then
                '===SPELL===
                lRoundDam = 0
                Messages1 = Messages1 & BGPURPLE & "You bite your tounge!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTMAGNETA & dbPlayers(dbAIndex).sPlayerName & " bites their tounge!" & vbCrLf & WHITE
            ElseIf Chance <= (MaxHit) And dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
                mRnd = RndNumber(1, 100)
                If mRnd = MonDodge Then
                    lRoundDam = 0
                    Messages1 = Messages1 & BGYELLOW & "You severly miss " & aMons(amonIndex).mName & ", and fall to the ground!" & vbCrLf & WHITE
                    Messages2 = Messages2 & YELLOW & dbPlayers(dbAIndex).sPlayerName & " fumbles!" & vbCrLf & WHITE
                    Exit For
                ElseIf (mRnd > MonDodge) Or (mRnd <= modGetData.GetPlayerHandicap(dbAIndex)) Then
                    bCrit = False
                    If (Chance > (MaxHit - pCrits)) Then
                        bCrit = True
                        lRoundDam = lRoundDam * 3
                        Messages1 = Messages1 & BRIGHTYELLOW & "You critically " & Mes(MessageID) & " " & aMons(amonIndex).mName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                        Messages2 = Messages2 & BRIGHTYELLOW & dbPlayers(dbAIndex).sPlayerName & " critically " & Mes2(MessageID2) & " " & aMons(amonIndex).mName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                        If pWeapon(Index).wCastSp <> "" Then
                            Erase Arr
                            Erase Arr2
                            SplitFast pWeapon(Index).wCastSp, Arr, ";"
                            SplitFast pWeapon(Index).wCastSpPer, Arr2, ";"
                            If UBound(Arr2) < UBound(Arr) Then ReDim Preserve Arr2(UBound(Arr))
                            For t = LBound(Arr) To UBound(Arr)
                                If Arr(t) <> "" Then
                                    If Arr2(t) <> "" Then
                                        If RndNumber(1, 100) < Val(Arr2(t)) Then
                                            With dbSpells(Val(Arr(t)))
                                                f = RndNumber(CDbl(.lMinDam), CDbl(.lMaxDam))
                                                v = .sMessage2
                                                q = .sMessage2
                                                v = ReplaceFast(v, "<%v>", aMons(amonIndex).mName)
                                                v = ReplaceFast(v, "<%c>", pWeapon(Index).wWeaponName)
                                                v = ReplaceFast(v, "<%d>", CStr(f))
                                                v = ReplaceFast(v, "<%s>", .sSpellName)
                                                q = ReplaceFast(q, "<%v>", aMons(amonIndex).mName)
                                                q = ReplaceFast(q, "<%c>", pWeapon(Index).wWeaponName)
                                                q = ReplaceFast(q, "<%d>", CStr(f))
                                                q = ReplaceFast(q, "<%s>", .sSpellName)
                                                Messages1 = Messages1 & BGGREEN & v & vbCrLf & WHITE
                                                Messages2 = Messages2 & BGGREEN & q & vbCrLf & WHITE
                                                lRoundDam = lRoundDam + f
                                            End With
                                        End If
                                    End If
                                End If
                                If DE Then DoEvents
                            Next
                        End If
                        If RndNumber(0, CDbl(dbPlayers(dbAIndex).iCha)) < (dbPlayers(dbAIndex).iCha \ 3) Then
                            s = dbPlayers(dbAIndex).sWeapon
                            If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                        End If
                    End If
                    If bCrit = False Then
                        Messages1 = Messages1 & BRIGHTRED & "You " & Mes(MessageID) & " " & aMons(amonIndex).mName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                        Messages2 = Messages2 & BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " " & Mes2(MessageID2) & " " & aMons(amonIndex).mName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                        If pWeapon(Index).wCastSp <> "" Then
                            Erase Arr
                            Erase Arr2
                            SplitFast pWeapon(Index).wCastSp, Arr, ";"
                            SplitFast pWeapon(Index).wCastSpPer, Arr2, ";"
                            If UBound(Arr2) < UBound(Arr) Then ReDim Preserve Arr2(UBound(Arr))
                            For t = LBound(Arr) To UBound(Arr)
                                If Arr(t) <> "" Then
                                    If Arr2(t) <> "" Then
                                        If RndNumber(1, 100) < Val(Arr2(t)) Then
                                            With dbSpells(Val(Arr(t)))
                                                f = RndNumber(CDbl(.lMinDam), CDbl(.lMaxDam))
                                                v = .sMessage2
                                                q = .sMessage2
                                                v = ReplaceFast(v, "<%v>", aMons(amonIndex).mName)
                                                v = ReplaceFast(v, "<%c>", pWeapon(Index).wWeaponName)
                                                v = ReplaceFast(v, "<%d>", CStr(f))
                                                v = ReplaceFast(v, "<%s>", .sSpellName)
                                                q = ReplaceFast(q, "<%v>", aMons(amonIndex).mName)
                                                q = ReplaceFast(q, "<%c>", pWeapon(Index).wWeaponName)
                                                q = ReplaceFast(q, "<%d>", CStr(f))
                                                q = ReplaceFast(q, "<%s>", .sSpellName)
                                                Messages1 = Messages1 & BGGREEN & v & vbCrLf & WHITE
                                                Messages2 = Messages2 & BGGREEN & q & vbCrLf & WHITE
                                                lRoundDam = lRoundDam + f
                                            End With
                                        End If
                                    End If
                                End If
                                If DE Then DoEvents
                            Next
                        End If
                        If RndNumber(0, CDbl(dbPlayers(dbAIndex).iCha)) < (dbPlayers(dbAIndex).iCha \ 3) Then
                            s = dbPlayers(dbAIndex).sWeapon
                            If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                        End If
                    End If
                ElseIf mRnd < MonDodge& Then
                    Messages1 = Messages1 & BRIGHTLIGHTBLUE & aMons(amonIndex).mName & " dodges your attack!" & vbCrLf & WHITE
                    Messages2 = Messages2 & BRIGHTLIGHTBLUE & aMons(amonIndex).mName & " dodges " & dbPlayers(dbAIndex).sPlayerName & "'s attack!" & vbCrLf & WHITE
                    lRoundDam = 0
                End If
            ElseIf Chance <= MaxHit And dbPlayers(dbAIndex).iCasting <> 0 Then
                '===SPELL===
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iInt \ 10)
                If (Chance >= MaxHit - 3) And (Chance <= MaxHit + 3) Then
                    lRoundDam = lRoundDam * 3
                    v = Mes(0)
                    q = Mes(1)
                    v = ReplaceFast(v, "<%v>", aMons(amonIndex).mName)
                    v = ReplaceFast(v, "<%c>", "You")
                    v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                    v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
                    q = ReplaceFast(q, "<%v>", aMons(amonIndex).mName)
                    q = ReplaceFast(q, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                    q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                    q = ReplaceFast(q, "<%s>", pWeapon(Index).wSpellName)
                    Messages1 = Messages1 & BRIGHTYELLOW & "Critically, " & v & vbCrLf & WHITE
                    Messages2 = Messages2 & BRIGHTYELLOW & "Critically, " & q & vbCrLf & WHITE
                Else
                    v = Mes(0)
                    q = Mes(1)
                    v = ReplaceFast(v, "<%v>", aMons(amonIndex).mName)
                    v = ReplaceFast(v, "<%c>", "You")
                    v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                    v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
                    q = ReplaceFast(q, "<%v>", aMons(amonIndex).mName)
                    q = ReplaceFast(q, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                    q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                    q = ReplaceFast(q, "<%s>", pWeapon(Index).wSpellName)
                    Messages1 = Messages1 & BRIGHTRED & v & vbCrLf & WHITE
                    Messages2 = Messages2 & BRIGHTRED & q & vbCrLf & WHITE
                End If
            ElseIf (Chance > MaxHit) And dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                lRoundDam = 0
                Messages1 = Messages1 & BRIGHTLIGHTBLUE & "You miss " & aMons(amonIndex).mName & " with your attack!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " misses their attack on " & aMons(amonIndex).mName & "." & vbCrLf & WHITE
            ElseIf (Chance& > MaxHit&) And dbPlayers(dbAIndex).iCasting <> 0 Then
                '===SPELL===
                lRoundDam = 0
                Messages1 = Messages1 & LIGHTBLUE & "You fail to cast " & pWeapon(Index).wSpellName & " on " & aMons(amonIndex).mName & "." & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " fails to cast " & pWeapon(Index).wSpellName & " on " & aMons(amonIndex).mName & "!" & vbCrLf & WHITE
            End If
            If dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                If lRoundDam <> 0 Then lRoundDam = lRoundDam - (aMons(amonIndex).mAc \ 14)
                If lRoundDam < 1 And lRoundDam <> 0 Then lRoundDam = 1
            Else
                '===SPELL===
                dbPlayers(dbAIndex).lMana = dbPlayers(dbAIndex).lMana - pWeapon(Index).wMana
            End If
            Damage = Damage + lRoundDam
            If DE Then DoEvents
            lRoundDam = 0
            If pWeapon(Index).wBullets <> -1 Then
                If pWeapon(Index).wMag = 0 Then
                    pWeapon(Index).wBullets = pWeapon(Index).wBullets - 1
                    dbPlayers(dbAIndex).sWeapon = modItemManip.SetItemBullets(dbPlayers(dbAIndex).sWeapon, modItemManip.SetupItemBullets(dbPlayers(dbAIndex).sWeapon, pWeapon(Index).wBullets, modItemManip.GetItemBulletsID(dbPlayers(dbAIndex).sWeapon), pWeapon(Index).wMag, pWeapon(Index).wBMana))
                    If pWeapon(Index).wBullets = 0 Then
                        bOutOfMana = True
                        Messages1 = Messages1 & BRIGHTBLUE & "You are out of ammo!" & WHITE & vbCrLf
                        Exit For
                    End If
                Else
                    dbPlayers(dbAIndex).lMana = dbPlayers(dbAIndex).lMana - pWeapon(Index).wBMana
                    If (dbPlayers(dbAIndex).lMana < pWeapon(Index).wBMana) Then
                        bOutOfMana = True
                        Messages1 = Messages1 & BRIGHTBLUE & "You are out of mana!" & WHITE & vbCrLf
                        Exit For
                    End If
                End If
            End If
            If (dbPlayers(dbAIndex).lMana < pWeapon(Index).wMana) And dbPlayers(dbAIndex).iCasting <> 0 Then
                bOutOfMana = True
                Messages1 = Messages1 & BRIGHTBLUE & "You are out of mana!" & WHITE & vbCrLf
                dbPlayers(dbAIndex).iCasting = 0
                SpellCombat(Index) = False
                Exit For
            End If
            If aMons(amonIndex).mHP - Damage <= 0 Then Exit For
        Next
    End If
End If
If aMons(amonIndex).mIs_Being_Attacked = False Then bDid = False Else bDid = True
With dbPlayers(dbAIndex)
    .dStamina = .dStamina - RndNumber(0, 1)
    .dHunger = .dHunger - RndNumber(0, 1)
    modMonsters.InsertInMonList amonIndex, .lPlayerID, 0
End With
If Damage < 0 Then Damage = 0
If dbPlayers(dbAIndex).iDualWield = 1 And fFlag = False Then
    fFlag = True
    GoTo DUALWIELDGOAGAIN
End If
End Sub

Sub DoFamAttackMon(dbAIndex As Long, amonIndex As Long, ByRef Messages1 As String, ByRef Messages2 As String, ByRef Damage As Long)
Dim MaxHit As Long
Dim Chance As Long
Dim lRoundDam As Long
Dim dbFamId As Long
Dim s As String
Dim s1 As String
Dim lMin As Long, lMax As Long
If dbPlayers(dbAIndex).lFamID <> 0 Then
    With dbPlayers(dbAIndex)
        dbFamId = GetFamID(.lFamID)
        MaxHit = RndNumber(1, CDbl(.lFamCHP + .lFamMHP))
        modFamiliars.GetFamAttack dbFamId, dbAIndex, lMin, lMax
        For i = 1 To dbFamiliars(dbFamId).lSwings
            Chance = RndNumber(1, CDbl(MaxHit))
            If Chance + .lFamAcc > MaxHit \ 2 Then
                With dbFamiliars(dbFamId)
                    lRoundDam = RndNumber(CDbl(lMin), CDbl(lMax))
                    s = .sAttackMessage
                    s1 = .sMessage2
                End With
                If .sFamCustom <> "0" Then
                    s = ReplaceFast(s, "<%n>", .sFamCustom & " the " & .sFamName)
                    s1 = ReplaceFast(s1, "<%n>", .sFamCustom & " the " & .sFamName)
                Else
                    s = ReplaceFast(s, "<%n>", "Your " & .sFamName)
                    s1 = ReplaceFast(s1, "<%n>", .sPlayerName & "'s " & .sFamName)
                End If
                s = ReplaceFast(s, "<%m>", aMons(amonIndex).mName)
                s = ReplaceFast(s, "<%d>", CStr(lRoundDam))
                Messages1 = Messages1 & BRIGHTBLUE & s & WHITE & vbCrLf
                s1 = ReplaceFast(s1, "<%m>", aMons(amonIndex).mName)
                s1 = ReplaceFast(s1, "<%d>", CStr(lRoundDam))
                Messages2 = Messages2 & BRIGHTBLUE & s1 & WHITE & vbCrLf
                Damage = Damage + lRoundDam
                If aMons(amonIndex).mHP - Damage < 1 Then Exit For
            Else
                With dbFamiliars(dbFamId)
                    s = .sMissMessage
                    s1 = .sMissMessage2
                End With
                If .sFamCustom <> "0" Then
                    s = ReplaceFast(s, "<%n>", .sFamCustom & " the " & .sFamName)
                    s1 = ReplaceFast(s1, "<%n>", .sFamCustom & " the " & .sFamName)
                Else
                    s = ReplaceFast(s, "<%n>", "Your " & .sFamName)
                    s1 = ReplaceFast(s1, "<%n>", .sPlayerName & "'s " & .sFamName)
                End If
                s = ReplaceFast(s, "<%m>", aMons(amonIndex).mName)
                s = ReplaceFast(s, "<%d>", CStr(lRoundDam))
                Messages1 = Messages1 & LIGHTBLUE & s & WHITE & vbCrLf
                s1 = ReplaceFast(s1, "<%n>", .sPlayerName & "'s " & .sFamName)
                s1 = ReplaceFast(s1, "<%m>", aMons(amonIndex).mName)
                s1 = ReplaceFast(s1, "<%d>", CStr(lRoundDam))
                Messages2 = Messages2 & LIGHTBLUE & s1 & WHITE & vbCrLf
            End If
            If DE Then DoEvents
        Next
    End With
End If
End Sub

Sub DoPlayerAttackPlayer(dbAIndex As Long, dbVIndex As Long, ByRef Messages1 As String, ByRef Messages2 As String, ByRef Messages3 As String, ByRef Damage As Long)
Dim pCrits As Long
Dim pAcc As Long
Dim Mes() As String
Dim Mes2() As String
Dim Mes3() As String
Dim MaxHit As Long
Dim MonDodge As Long
Dim i As Long
Dim Swings As Long
Dim bOutOfMana As Boolean
Dim Index As Long
Dim lRoundDam As Long
Dim MessageID As Long
Dim MessageID2 As Long
Dim MessageID3 As Long
Dim Chance As Long
Dim mRnd As Long
Dim bCrit As Boolean
Dim s As String
Dim v As String
Dim q As String
Dim R As String
Index = dbPlayers(dbAIndex).iIndex
SetWeaponStats CLng(Index)
If modSC.FastStringComp(pWeapon(Index).wSpellName, "") Then
    SplitFast pWeapon(Index).wMessage, Mes, ":"
    SplitFast pWeapon(Index).wMessage2, Mes2, ":"
    SplitFast pWeapon(Index).wMessageV, Mes3, ":"
    pCrits = dbPlayers(dbAIndex).iCrits
    pAcc = dbPlayers(dbAIndex).iAcc
    MaxHit = modGetData.GetPlayerMaxHit(dbAIndex) + pAcc
    If MaxHit > 98 Then MaxHit = 98
    MonDodge = modGetData.GetPlayerDodge(dbVIndex)
    Swings = modGetData.GetPlayerSwings(dbVIndex)
Else
    bOutOfMana = False
    If dbPlayers(dbAIndex).lMana < pWeapon(Index).wMana Then
        bOutOfMana = True
        Messages1 = BRIGHTBLUE & "You are out of mana!" & WHITE & vbCrLf
        dbPlayers(dbAIndex).iCasting = 0
        SpellCombat(Index) = False
        Exit Sub
    Else
        MaxHit& = modGetData.GetSpellChance(Index)
        ReDim Mes(2) As String
        ReDim Mes2(2) As String
        ReDim Mes3(2) As String
        Mes(0) = pWeapon(Index).wMessage
        Mes(1) = pWeapon(Index).wMessage2
        Mes(2) = pWeapon(Index).wMessageV
        Swings = pWeapon(Index).wCast
    End If
End If
If dbPlayers(dbAIndex).iPlayerAttacking > 0 And bOutOfMana = False Then
    If dbPlayers(dbAIndex).iIsBSing = 1 Then
        MessageID = RndNumber(0, CDbl(UBound(Mes)))
        MessageID2 = RndNumber(0, CDbl(UBound(Mes2)))
        MessageID3 = RndNumber(0, CDbl(UBound(Mes3)))
        Chance = RndNumber(1, 100 - (dbPlayers(dbAIndex).iAcc + dbPlayers(dbAIndex).iLevel))
        If Chance < MaxHit Then
            mRnd = RndNumber(1, 100 + (dbPlayers(dbAIndex).iAcc + dbPlayers(dbAIndex).iLevel))
            If mRnd > MonDodge Then
                'BS Hit
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
                lRoundDam = lRoundDam * 6
                Messages1 = Messages1 & BRIGHTRED & "You " & GREEN & "surprise " & BRIGHTRED & Mes(MessageID) & " " & dbPlayers(dbVIndex).sPlayerName & " for " & CStr(lRoundDam) & "." & vbCrLf
                Messages2 = Messages2 & BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " surprise " & BRIGHTRED & Mes2(MessageID2) & " " & dbPlayers(dbVIndex).sPlayerName & " for " & CStr(lRoundDam) & "." & vbCrLf
                Messages3 = Messages3 & BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " surprise " & BRIGHTRED & Mes3(MessageID3) & " you for " & CStr(lRoundDam) & "." & vbCrLf
            Else
                Messages1 = Messages1 & BRIGHTLIGHTBLUE & dbPlayers(dbVIndex).sPlayerName & " dodges your backstab!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbVIndex).sPlayerName & " dodges " & dbPlayers(dbAIndex).sPlayerName & "'s backstab!" & vbCrLf & WHITE
                Messages3 = Messages3 & BRIGHTLIGHTBLUE & "You dodge " & dbPlayers(dbAIndex).sPlayerName & "'s backstab!" & vbCrLf & WHITE
                lRoundDam = 0
            End If
        Else
            lRoundDam = 0
            Messages1 = Messages1 & BRIGHTLIGHTBLUE & "You miss " & dbPlayers(dbAIndex).sPlayerName & " with your backstab!" & vbCrLf & WHITE
            Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " misses their backstab on " & dbPlayers(dbAIndex).sPlayerName & "!" & vbCrLf & WHITE
            Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " misses their backstab on you!" & vbCrLf & WHITE
        End If
        dbPlayers(dbAIndex).iIsBSing = 0
        Damage = lRoundDam
    Else
        For i = 1 To Swings
            MessageID = RndNumber(0, CDbl(UBound(Mes)))
            MessageID2 = RndNumber(0, CDbl(UBound(Mes2)))
            MessageID3 = RndNumber(0, CDbl(UBound(Mes3)))
            Chance = RndNumber(1, 100)
            If Chance >= 100 And dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
                lRoundDam = lRoundDam * 3
                Messages1 = Messages1 & BRIGHTRED & "You " & BRIGHTYELLOW & "critically and severly " & BRIGHTRED & Mes(MessageID) & " " & dbPlayers(dbVIndex).sPlayerName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " critically and severly " & Mes2(MessageID2) & "s " & dbPlayers(dbVIndex).sPlayerName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                Messages3 = Messages3 & BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " critically and severly " & Mes3(MessageID3) & "s you for " & lRoundDam & " damage!" & vbCrLf & WHITE
                If RndNumber(0, CDbl(dbPlayers(dbAIndex).iCha)) < (dbPlayers(dbAIndex).iCha \ 3) Then
                    s = dbPlayers(dbAIndex).sWeapon
                    If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                End If
            ElseIf Chance >= 100 And dbPlayers(dbAIndex).iCasting <> 0 Then
                '===SPELL===
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
                If pWeapon(Index).wElement <> -1 Then
                    lRoundDam = lRoundDam - modResist.GetResistValue(dbVIndex, pWeapon(Index).wElement)
                End If
                lRoundDam = lRoundDam - modGetData.GetPlayersMR(dbAIndex, dbVIndex)
                If lRoundDam < 1 Then
                    If RndNumber(1, 100) > 75 Then
                        lRoundDam = 0
                        Messages1 = Messages1 & LIGHTBLUE & dbPlayers(dbVIndex).sPlayerName & " resist your casting of " & pWeapon(Index).wSpellName & "!" & vbCrLf & WHITE
                        Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbVIndex).sPlayerName & " resist " & dbPlayers(dbAIndex).sPlayerName & " cast of " & pWeapon(Index).wSpellName & "!" & vbCrLf & WHITE
                        Messages3 = Messages3 & BRIGHTLIGHTBLUE & "You resist " & dbPlayers(dbAIndex).sPlayerName & " cast of " & pWeapon(Index).wSpellName & "!" & vbCrLf & WHITE
                        GoTo nNext
                    Else
                        lRoundDam = 1
                    End If
                End If
                lRoundDam = lRoundDam * 3
                v = Mes(0)
                v = ReplaceFast(v, "<%v>", dbPlayers(dbVIndex).sPlayerName)
                v = ReplaceFast(v, "<%c>", "You")
                v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
                q = ReplaceFast(q, "<%v>", dbPlayers(dbVIndex).sPlayerName)
                q = ReplaceFast(q, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                q = ReplaceFast(q, "<%s>", pWeapon(Index).wSpellName)
                R = ReplaceFast(R, "<%v>", "you")
                R = ReplaceFast(R, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                R = ReplaceFast(R, "<%d>", CStr(lRoundDam))
                R = ReplaceFast(R, "<%s>", pWeapon(Index).wSpellName)
                Messages1 = Messages1 & BGYELLOW & "Critically and severly" & BRIGHTRED & ", " & v & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTRED & "Critically and severly, " & q & vbCrLf & WHITE
                Messages3 = Messages3 & BRIGHTRED & "Critically and severly, " & R & vbCrLf & WHITE
            ElseIf (Chance <= 1 + (dbPlayers(dbVIndex).iAC \ 14)) And dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                lRoundDam = 0
                Messages1 = Messages1 & BRIGHTMAGNETA & "You flinch as you attempt to hit " & dbPlayers(dbVIndex).sPlayerName & ", and cause no damage!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTMAGNETA & dbPlayers(dbAIndex).sPlayerName & " flinches!" & vbCrLf & WHITE
                Messages3 = Messages3 & BRIGHTMAGNETA & dbPlayers(dbAIndex).sPlayerName & " flinches!" & vbCrLf & WHITE
            ElseIf Chance <= 2 And dbPlayers(dbAIndex).iCasting <> 0 Then
                '===SPELL===
                lRoundDam = 0
                Messages1 = Messages1 & BGPURPLE & "You bite your tounge!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTMAGNETA & dbPlayers(dbAIndex).sPlayerName & " bites their tounge!" & vbCrLf & WHITE
                Messages3 = Messages3 & BRIGHTMAGNETA & dbPlayers(dbAIndex).sPlayerName & " bites their tounge!" & vbCrLf & WHITE
            ElseIf Chance <= (MaxHit) And dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
                mRnd = RndNumber(1, 100)
                If mRnd = MonDodge Then
                    lRoundDam = 0
                    Messages1 = Messages1 & BGYELLOW & "You severly miss " & dbPlayers(dbVIndex).sPlayerName & ", and fall to the ground!" & vbCrLf & WHITE
                    Messages2 = Messages2 & YELLOW & dbPlayers(dbAIndex).sPlayerName & " fumbles!" & vbCrLf & WHITE
                    Messages3 = Messages3 & YELLOW & dbPlayers(dbAIndex).sPlayerName & " fumbles!" & vbCrLf & WHITE
                    Exit For
                ElseIf (mRnd > MonDodge) Or (mRnd <= 10) Then
                    bCrit = False
                    If (Chance > (MaxHit - pCrits)) Then
                        bCrit = True
                        lRoundDam = lRoundDam * 2
                        Messages1 = Messages1 & BRIGHTYELLOW & "You critically " & Mes(MessageID) & " " & dbPlayers(dbVIndex).sPlayerName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                        Messages2 = Messages2 & BRIGHTYELLOW & dbPlayers(dbAIndex).sPlayerName & " critically " & Mes2(MessageID2) & " " & dbPlayers(dbVIndex).sPlayerName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                        Messages3 = Messages3 & BRIGHTYELLOW & dbPlayers(dbAIndex).sPlayerName & " critically " & Mes3(MessageID3) & " you for " & lRoundDam & " damage!" & vbCrLf & WHITE
                        If RndNumber(0, CDbl(dbPlayers(dbAIndex).iCha)) < (dbPlayers(dbAIndex).iCha \ 3) Then
                            s = dbPlayers(dbAIndex).sWeapon
                            If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                        End If
                    End If
                    If bCrit = False Then
                        If CheckShield(dbPlayers(dbVIndex).iIndex, dbVIndex) = True Then
                            Messages1 = Messages1 & BRIGHTMAGNETA & "You attempt to hit " & dbPlayers(dbVIndex).sPlayerName & ", but " & modGetData.GetGenderPronoun(dbVIndex) & " blocks the hit with their shield!" & vbCrLf & WHITE
                            Messages2 = Messages2 & BRIGHTMAGNETA & dbPlayers(dbAIndex).sPlayerName & "'s swing is blocked by " & dbPlayers(dbVIndex).sPlayerName & "'s shield!" & vbCrLf & WHITE
                            Messages2 = Messages2 & BRIGHTMAGNETA & "You block " & dbPlayers(dbAIndex).sPlayerName & "'s swing with your shield!" & vbCrLf & WHITE
                            If RndNumber(0, CDbl(dbPlayers(dbVIndex).iCha)) < (dbPlayers(dbVIndex).iCha \ 3) Then
                                s = dbPlayers(dbVIndex).sShield
                                If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbVIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                            End If
                        Else
                            Messages1 = Messages1 & BRIGHTRED & "You " & Mes(MessageID) & " " & dbPlayers(dbVIndex).sPlayerName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                            Messages2 = Messages2 & BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " " & Mes2(MessageID2) & " " & dbPlayers(dbVIndex).sPlayerName & " for " & lRoundDam & " damage!" & vbCrLf & WHITE
                            Messages3 = Messages3 & BRIGHTRED & dbPlayers(dbAIndex).sPlayerName & " " & Mes3(MessageID3) & " you for " & lRoundDam & " damage!" & vbCrLf & WHITE
                            If RndNumber(0, CDbl(dbPlayers(dbAIndex).iCha)) < (dbPlayers(dbAIndex).iCha \ 3) Then
                                s = dbPlayers(dbAIndex).sWeapon
                                If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbAIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                            End If
                            If RndNumber(0, CDbl(dbPlayers(dbVIndex).iCha)) < (dbPlayers(dbVIndex).iCha \ 3) Then
                                s = modGetData.GetUnformatedStringFromID(CLng(dbVIndex), modGetData.GetHitPositionID)
                                If s <> "0" Then modItemManip.SubtractOneFromItemDUR CLng(dbVIndex), modItemManip.GetItemIDFromUnFormattedString(s), modItemManip.GetItemUsesFromUnFormattedString(s), modItemManip.GetItemDurFromUnFormattedString(s)
                            End If
                        End If
                    End If
                ElseIf mRnd < MonDodge& Then
                    Messages1 = Messages1 & BRIGHTLIGHTBLUE & dbPlayers(dbVIndex).sPlayerName & " dodges your attack!" & vbCrLf & WHITE
                    Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbVIndex).sPlayerName & " dodges " & dbPlayers(dbAIndex).sPlayerName & "'s attack!" & vbCrLf & WHITE
                    Messages3 = Messages3 & BRIGHTLIGHTBLUE & "You dodge " & dbPlayers(dbAIndex).sPlayerName & "'s attack!" & vbCrLf & WHITE
                    lRoundDam = 0
                End If
            ElseIf Chance <= MaxHit And dbPlayers(dbAIndex).iCasting <> 0 Then
                '===SPELL===
                lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iInt \ 10)
                If pWeapon(Index).wElement <> -1 Then
                    lRoundDam = lRoundDam - modResist.GetResistValue(dbVIndex, pWeapon(Index).wElement)
                End If
                lRoundDam = lRoundDam - modGetData.GetPlayersMR(dbAIndex, dbVIndex)
                If lRoundDam < 1 Then
                    If RndNumber(1, 100) > 75 Then
                        lRoundDam = 0
                        Messages1 = Messages1 & LIGHTBLUE & dbPlayers(dbVIndex).sPlayerName & " resist your casting of " & pWeapon(Index).wSpellName & "!" & vbCrLf & WHITE
                        Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbVIndex).sPlayerName & " resist " & dbPlayers(dbAIndex).sPlayerName & " cast of " & pWeapon(Index).wSpellName & "!" & vbCrLf & WHITE
                        Messages3 = Messages3 & BRIGHTLIGHTBLUE & "You resist " & dbPlayers(dbAIndex).sPlayerName & " cast of " & pWeapon(Index).wSpellName & "!" & vbCrLf & WHITE
                        GoTo nNext
                    Else
                        lRoundDam = 1
                    End If
                End If
                v = Mes(0)
                q = Mes(1)
                R = Mes(2)
                If (Chance >= MaxHit - 3) And (Chance <= MaxHit + 3) Then
                    lRoundDam = lRoundDam * 2
                    v = ReplaceFast(v, "<%v>", dbPlayers(dbVIndex).sPlayerName)
                    v = ReplaceFast(v, "<%c>", "You")
                    v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                    v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
                    q = ReplaceFast(q, "<%v>", dbPlayers(dbVIndex).sPlayerName)
                    q = ReplaceFast(q, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                    q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                    q = ReplaceFast(q, "<%s>", pWeapon(Index).wSpellName)
                    R = ReplaceFast(R, "<%v>", "you")
                    R = ReplaceFast(R, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                    R = ReplaceFast(R, "<%d>", CStr(lRoundDam))
                    R = ReplaceFast(R, "<%s>", pWeapon(Index).wSpellName)
                    Messages1 = Messages1 & BRIGHTYELLOW & "Critically, " & v & vbCrLf & WHITE
                    Messages2 = Messages2 & BRIGHTYELLOW & "Critically, " & q & vbCrLf & WHITE
                    Messages3 = Messages3 & BRIGHTYELLOW & "Critically, " & R & vbCrLf & WHITE
                Else
                    v = ReplaceFast(v, "<%v>", dbPlayers(dbVIndex).sPlayerName)
                    v = ReplaceFast(v, "<%c>", "You")
                    v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
                    v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
                    q = ReplaceFast(q, "<%v>", dbPlayers(dbVIndex).sPlayerName)
                    q = ReplaceFast(q, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                    q = ReplaceFast(q, "<%d>", CStr(lRoundDam))
                    q = ReplaceFast(q, "<%s>", pWeapon(Index).wSpellName)
                    R = ReplaceFast(R, "<%v>", "you")
                    R = ReplaceFast(R, "<%c>", dbPlayers(dbAIndex).sPlayerName)
                    R = ReplaceFast(R, "<%d>", CStr(lRoundDam))
                    R = ReplaceFast(R, "<%s>", pWeapon(Index).wSpellName)
                    Messages1 = Messages1 & BRIGHTRED & v & vbCrLf & WHITE
                    Messages2 = Messages2 & BRIGHTRED & q & vbCrLf & WHITE
                    Messages3 = Messages3 & BRIGHTRED & R & vbCrLf & WHITE
                End If
            ElseIf (Chance > MaxHit) And dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                lRoundDam = 0
                Messages1 = Messages1 & BRIGHTLIGHTBLUE & "You miss " & dbPlayers(dbVIndex).sPlayerName & " with your attack!" & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " misses their attack on " & dbPlayers(dbVIndex).sPlayerName & "." & vbCrLf & WHITE
                Messages3 = Messages3 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " misses their attack on you!" & vbCrLf & WHITE
            ElseIf (Chance& > MaxHit&) And dbPlayers(dbAIndex).iCasting <> 0 Then
                '===SPELL===
                lRoundDam = 0
                Messages1 = Messages1 & LIGHTBLUE & "You fail to cast " & pWeapon(Index).wSpellName & " on " & dbPlayers(dbVIndex).sPlayerName & "." & vbCrLf & WHITE
                Messages2 = Messages2 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " fails to cast " & pWeapon(Index).wSpellName & " on " & dbPlayers(dbVIndex).sPlayerName & "!" & vbCrLf & WHITE
                Messages3 = Messages3 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " fails to cast " & pWeapon(Index).wSpellName & " on you!" & vbCrLf & WHITE
            End If
            If dbPlayers(dbAIndex).iCasting = 0 Then
                '===WEAPON===
                If lRoundDam <> 0 Then lRoundDam = lRoundDam - (dbPlayers(dbVIndex).iAC \ 14)
                If lRoundDam < 1 And lRoundDam <> 0 Then lRoundDam = 1
            Else
                '===SPELL===
                dbPlayers(dbAIndex).lMana = dbPlayers(dbAIndex).lMana - pWeapon(Index).wMana
            End If
nNext:
            Damage = Damage + lRoundDam
            If dbPlayers(dbVIndex).lHP - Damage <= lDeath Then Exit For
            If DE Then DoEvents
            lRoundDam = 0
        Next
    End If
End If
With dbPlayers(dbAIndex)
    .dStamina = .dStamina - RndNumber(0, 1)
    .dHunger = .dHunger - RndNumber(0, 1)
End With
If Damage < 0 Then Damage = 0
End Sub
 
Sub DoRoomAttack(dbAIndex As Long, ByRef Messages1 As String, ByRef Messages3 As String)
Dim bOutOfMana As Boolean
Dim Mes() As String
Dim MaxHit As Long
Dim Index As Long
Dim Swings As Long
Dim aListP() As String
Dim aListM() As String
Dim Arr() As String
Dim s As String
Dim R As String
Dim q As String
Dim v As String
Dim m As String
Dim i As Long
Dim Chance As Long
Dim lRoundDam As Long
Dim lIn As Long
Dim sS As String
Index = dbPlayers(dbAIndex).iIndex
bOutOfMana = False
If dbPlayers(dbAIndex).lMana < pWeapon(Index).wMana Then
    bOutOfMana = True
    Messages1 = BRIGHTBLUE & "You are out of mana!" & WHITE & vbCrLf
    dbPlayers(dbAIndex).iCasting = 0
    SpellCombat(Index) = False
    Exit Sub
Else
    MaxHit& = modGetData.GetSpellChance(Index)
    ReDim Mes(2) As String
    Mes(0) = pWeapon(Index).wMessage
    Mes(2) = pWeapon(Index).wMessageV
    Swings = pWeapon(Index).wCast
End If
s = modGetData.GetPlayersIDsHere(dbPlayers(dbAIndex).lLocation)
s = ReplaceFast(s, dbPlayers(dbAIndex).lPlayerID & ";", "")
q = modGetData.GetAllMonstersInRoomATTACKABLE(dbPlayers(dbAIndex).lLocation, dbPlayers(dbAIndex).lDBLocation)
R = dbPlayers(dbAIndex).sParty
If Not modSC.FastStringComp(R, "0") Then
    R = ReplaceFast(R, ":", "")
    SplitFast R, Arr, ";"
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) <> "" Then
            With dbPlayers(GetPlayerIndexNumber(CLng(Arr(i))))
                s = ReplaceFast(s, CStr(.lPlayerID) & ";", "")
            End With
        End If
    Next
End If
If s = "" And q = "" Then
    Messages1 = BRIGHTBLUE & "There is no one left here." & WHITE & vbCrLf
    dbPlayers(dbAIndex).iCasting = 0
    SpellCombat(Index) = False
    dbPlayers(dbAIndex).dMonsterID = 99999
    Exit Sub
End If
If s = "" Then
    ReDim aListP(0)
    aListP(0) = "-1"
Else
    SplitFast s, aListP, ";"
End If
If q = "" Then
    ReDim aListM(0)
    aListM(0) = "-1"
Else
    SplitFast q, aListM, ";"
End If
For i = 1 To Swings
    Chance = RndNumber(1, 100)
    If Chance >= 100 Then
        '===SPELL===
        lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iStr \ 10)
        lRoundDam = lRoundDam * 3
        v = Mes(0)
        m = Mes(2)
        v = ReplaceFast(v, "<%v>", "room")
        v = ReplaceFast(v, "<%c>", "You")
        v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
        v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
        m = ReplaceFast(m, "<%v>", "room")
        m = ReplaceFast(m, "<%c>", dbPlayers(dbAIndex).sPlayerName)
        m = ReplaceFast(m, "<%d>", CStr(lRoundDam))
        m = ReplaceFast(m, "<%s>", pWeapon(Index).wSpellName)
        Messages1 = Messages1 & BGYELLOW & "Critically and severly" & BRIGHTRED & ", " & v & vbCrLf & WHITE
        Messages3 = Messages3 & BRIGHTRED & "Critically and severly, " & m & vbCrLf & WHITE
    ElseIf Chance <= 2 Then
        '===SPELL===
        lRoundDam = 0
        Messages1 = Messages1 & BGPURPLE & "You bite your tounge!" & vbCrLf & WHITE
        Messages3 = Messages3 & BRIGHTMAGNETA & dbPlayers(dbAIndex).sPlayerName & " bites their tounge!" & vbCrLf & WHITE
    ElseIf Chance <= MaxHit Then
        '===SPELL===
        lRoundDam = RndNumber(CDbl(pWeapon(Index).wMin), CDbl(pWeapon(Index).wMax)) + (dbPlayers(dbAIndex).iMaxDamage) + (dbPlayers(dbAIndex).iInt \ 10)
        v = Mes(0)
        m = Mes(2)
        If (Chance >= MaxHit - 3) And (Chance <= MaxHit + 3) Then
            lRoundDam = lRoundDam * 2
            v = ReplaceFast(v, "<%v>", "room")
            v = ReplaceFast(v, "<%c>", "You")
            v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
            v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
            m = ReplaceFast(m, "<%v>", "room")
            m = ReplaceFast(m, "<%c>", dbPlayers(dbAIndex).sPlayerName)
            m = ReplaceFast(m, "<%d>", CStr(lRoundDam))
            m = ReplaceFast(m, "<%s>", pWeapon(Index).wSpellName)
            Messages1 = Messages1 & BRIGHTYELLOW & "Critically, " & v & vbCrLf & WHITE
            Messages3 = Messages3 & BRIGHTYELLOW & "Critically, " & m & vbCrLf & WHITE
        Else
            v = ReplaceFast(v, "<%v>", "room")
            v = ReplaceFast(v, "<%c>", "You")
            v = ReplaceFast(v, "<%d>", CStr(lRoundDam))
            v = ReplaceFast(v, "<%s>", pWeapon(Index).wSpellName)
            m = ReplaceFast(m, "<%v>", "room")
            m = ReplaceFast(m, "<%c>", dbPlayers(dbAIndex).sPlayerName)
            m = ReplaceFast(m, "<%d>", CStr(lRoundDam))
            m = ReplaceFast(m, "<%s>", pWeapon(Index).wSpellName)
            Messages1 = Messages1 & BRIGHTRED & v & vbCrLf & WHITE
            Messages3 = Messages3 & BRIGHTRED & m & vbCrLf & WHITE
        End If
    ElseIf (Chance& > MaxHit&) Then
            '===SPELL===
        lRoundDam = 0
        Messages1 = Messages1 & LIGHTBLUE & "You fail to cast " & pWeapon(Index).wSpellName & " on " & dbPlayers(dbVIndex).sPlayerName & "." & vbCrLf & WHITE
        Messages3 = Messages3 & BRIGHTLIGHTBLUE & dbPlayers(dbAIndex).sPlayerName & " fails to cast " & pWeapon(Index).wSpellName & "!" & vbCrLf & WHITE
        
        
    End If
    dbPlayers(dbAIndex).lMana = dbPlayers(dbAIndex).lMana - pWeapon(Index).wMana
    Damage = Damage + lRoundDam
    If DE Then DoEvents
Next
If aListP(0) <> "-1" Then
    For i = LBound(aListP) To UBound(aListP)
        If aListP(i) <> "" Then
            lIn = GetPlayerIndexNumber(, , CLng(aListP(i)))
            With dbPlayers(lIn)
                If pWeapon(Index).wElement <> -1 Then
                    .lHP = .lHP - Damage + modResist.GetResistValue(lIn, pWeapon(Index).wElement)
                    If .lHP <= 0 Then CheckDeath .iIndex, , True, Messages3, , Messages1
                Else
                    .lHP = .lHP - Damage
                    If .lHP <= 0 Then CheckDeath .iIndex, , True, Messages3, , Messages1
                End If
            End With
        End If
        If DE Then DoEvents
    Next
End If
If aListM(0) <> "-1" Then
    For i = LBound(aListM) To UBound(aListM)
        If aListM(i) <> "" Then
            With aMons(CLng(aListM(i)))
                .mHP = .mHP - Damage
                If .mHP <= 0 Then
                    DropMonGold CLng(aListM(i)), sS
                    Messages1 = Messages1 & sS
                    Messages3 = Messages3 & sS
                    
                    sS = ""
                    SendDeathText CLng(aListM(i)), sS
                    Messages1 = Messages1 & sS
                    Messages3 = Messages3 & sS
                    
                    sS = ""
                    DropMonItem CLng(aListM(i)), sS
                    Messages1 = Messages1 & sS
                    Messages3 = Messages3 & sS
                    
                    Messages1 = Messages1 & WHITE & "You have slain " & .mName & "!" & vbCrLf
                    Messages3 = Messages3 & BRIGHTGREEN & dbPlayers(dbAIndex).sPlayerName & " has slain " & .mName & "!" & vbCrLf & WHITE
                    AddEXP Index, CLng(aListM(i))
                    Messages1 = Messages1 & BRIGHTWHITE & "Your experience has increased by " & .mEXP & "." & GREEN & vbCrLf
                    If dbPlayers(dbAIndex).lFamID <> 0 Then Messages1 = Messages1 & BRIGHTWHITE & "Your " & dbPlayers(dbAIndex).sFamName & " gains " & CStr(.mEXP \ RndNumber(3, 15)) & " experience." & GREEN & vbCrLf
                    
                    AddMonsterRgn .mName
            
                    dbPlayers(dbAIndex).dClassPoints = dbPlayers(dbAIndex).dClassPoints + 0.1
                    
                    ClearOtherAttackers Index, CLng(aListM(i))
                    
                    sScripting Index, , , 0, aMons(CLng(aListM(i))).mScript
                    mRemoveItem CLng(CLng(aListM(i)))
                    AmountMons = AmountMons - 1
                    sS = ""
                End If
            End With
        End If
        If DE Then DoEvents
    Next
End If
dbPlayers(dbAIndex).lRoomSpellFlag = 1
End Sub
