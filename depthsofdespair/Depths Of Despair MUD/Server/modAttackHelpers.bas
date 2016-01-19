Attribute VB_Name = "modAttackHelpers"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modAttackHelpers
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Sub SubtractFamHP(dbAIndex As Long, lDam As Long, ByRef Message1 As String, ByRef Message2 As String)
With dbPlayers(dbAIndex)
    .lFamCHP = .lFamCHP - lDam
    If .lFamCHP <= 0 Then
        If .sFamCustom <> "0" Then
            Message1 = Message1 & BGYELLOW & BRIGHTRED & .sFamCustom & " the " & .sFamName & " is killed!" & WHITE & vbCrLf
            Message2 = Message2 & YELLOW & .sFamCustom & " the " & .sFamName & " is killed!" & WHITE & vbCrLf
        Else
            Message1 = Message1 & BGYELLOW & BRIGHTRED & "Your " & .sFamName & " is killed!" & WHITE & vbCrLf
            Message2 = Message2 & YELLOW & .sPlayerName & "'s " & .sFamName & " is killed!" & WHITE & vbCrLf
        End If
        RemoveStats .iIndex, True
    End If
End With
End Sub





Public Function CheckAttackMonster(Index As Long, dbIndex As Long, s As String) As Boolean
Dim i As Long
Dim j As Long
Dim Arr() As String
Dim n As Long
Dim t As String
Dim R As String
With dbPlayers(dbIndex)
    t = dbMap(.lDBLocation).sAMonIds
    If t <> "" Then
        SplitFast t, Arr, ";"
        R = LCaseFast(s)
        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> "" Then
                n = CLng(Val(Arr(i)))
                If .lLocation = aMons(n).mLoc And R = LCaseFast(aMons(n).mName) Then
                    CheckAttackMonster = True
                    If aMons(n).mAttackable = True Then
                        If dbPlayers(dbIndex).iIsBSing = 0 Then
                            WrapAndSend Index, BRIGHTRED & "You move to attack " & GREEN & aMons(n).mName & BRIGHTRED & "!" & vbCrLf & WHITE
                            SendToAllInRoom Index, BRIGHTRED & .sPlayerName & " moves to attack " & aMons(n).mName & "." & WHITE & vbCrLf, .lLocation
                        Else
                            WrapAndSend Index, BRIGHTRED & "You sneak up behind " & GREEN & aMons(n).mName & BRIGHTRED & "..." & vbCrLf & WHITE
                        End If
                        For j = 0 To 4
                            If j < UBound(aMons(n).mSpells) Then
                                aMons(n).mSpells(j).lCastPerRound = 0
                            End If
                            If DE Then DoEvents
                        Next
                        AddEvil dbIndex, n
                        X(Index) = ""
                        .dMonsterID = n
                        .iPlayerAttacking = 0
                        aMons(n).mIs_Being_Attacked = True
                        aMons(n).mPlayerAttacking = .iIndex
                        modMonsters.InsertInMonList n, dbPlayers(dbIndex).lPlayerID, 0
                        Exit For
                    Else
                        WrapAndSend Index, RED & "You seem unable to attack that being." & WHITE & vbCrLf
                        X(Index) = ""
                        CheckAttackMonster = True
                        dbPlayers(dbIndex).iIsBSing = 0
                        Exit For
                    End If
                End If
            End If
            If DE Then DoEvents
        Next
    End If
End With
End Function

Public Function CheckPlayerAttack(Index As Long, dbIndex As Long, s As String) As Boolean
Dim dbVIndex As Long
Dim l&, u&, a&
dbVIndex = GetPlayerIndexNumber(, s)
If dbVIndex = 0 Then
    WrapAndSend Index, RED & "You do not see them here!" & WHITE & vbCrLf
    X(Index) = ""
    CheckPlayerAttack = True
    Exit Function
End If
If dbPlayers(dbVIndex).iGhostMode = 1 Then
    WrapAndSend Index, RED & "You do not see them here!" & WHITE & vbCrLf
    X(Index) = ""
    CheckPlayerAttack = True
    Exit Function
End If
a& = dbPlayers(dbIndex).iLevel
l& = dbPlayers(dbVIndex).iLevel - lPvPLevel
u& = dbPlayers(dbVIndex).iLevel + lPvPLevel
If (a& < l&) Or (a& > u&) Then
    WrapAndSend Index, RED & "You are about to attack " & dbPlayers(dbVIndex).sPlayerName & ", but decide not to." & WHITE & vbCrLf
    X(Index) = ""
    CheckPlayerAttack = True
    Exit Function
End If
If InStr(1, LCaseFast(modGetData.GetPlayersHereWithoutRiding(CLng(dbPlayers(dbIndex).lLocation), dbIndex)), LCaseFast(s)) = 0 Then
    WrapAndSend Index, RED & "You do not see them here!" & WHITE & vbCrLf
    X(Index) = ""
    CheckPlayerAttack = True
    Exit Function
End If
If LCaseFast(s) = LCaseFast(dbPlayers(dbIndex).sPlayerName) Then
    WrapAndSend Index, RED & "You would look pretty stupid attacking yourself." & WHITE & vbCrLf
    X(Index) = ""
    CheckPlayerAttack = True
    Exit Function
End If
With dbPlayers(dbVIndex)
    SendToAllInRoom Index, BRIGHTRED & dbPlayers(dbIndex).sPlayerName & " moves to attack " & .sPlayerName & "." & WHITE & vbCrLf, .lLocation, .iIndex
    WrapAndSend .iIndex, BRIGHTRED & dbPlayers(dbIndex).sPlayerName & " moves to intiate combat with you!" & WHITE & vbCrLf
    dbPlayers(dbIndex).iPlayerAttacking = .iIndex
    AddPvPEvil Index, .iIndex
    WrapAndSend Index, BRIGHTRED & "You move to initiate combat with " & .sPlayerName & "!" & vbCrLf & WHITE
End With
X(Index) = ""
CheckPlayerAttack = True
End Function

Public Function IsSpell(sSpell As String, sShorts As String, ByRef iID As Long) As Boolean
Dim i As Long
Dim tArr() As String
sShorts = LCaseFast(sShorts)
SplitFast Left$(sShorts, Len(sShorts) - 1), tArr, ";"

For i = 0 To UBound(tArr())
    If tArr(i) = sSpell Then
        iID& = i
        IsSpell = True
        Exit For
    End If
    If DE Then DoEvents
Next
End Function

Public Function DoSpell(Index As Long, dbIndex As Long, s$, sIDs As String, iID&, ByRef SpellName$, ByRef bRoom As Boolean) As Boolean
Dim i As Long
Dim tArr2() As String
Dim dbSpell As Long
Dim dbT As Long
SplitFast Left$(sIDs, Len(sIDs) - 1), tArr2, ";"
dbSpell = GetSpellID(, CLng(tArr2(iID&)))
With dbSpells(dbSpell)
    SpellName$ = .sSpellName
    '0 - Healing
    '1 - Combat
    '2 - Teleport
    '3 - Bless
    '4 - Room Spell
    '~5 - Party Spell
    Select Case .iUse
        Case 0, 2, 3
            If InStr(ReplaceFast(LCaseFast(X(Index)), "cast ", ""), " ") > 0 Then
                X(Index) = Mid$(s$, InStr(5, s$, " ") + 1)
                X(Index) = SmartFind(Index, X(Index), Player_In_Room)
                dbT = GetPlayerIndexNumber(, X(Index))
                If dbT = 0 Then
                    WrapAndSend Index, RED & "You are unable to locate " & X(Index) & "." & WHITE & vbCrLf
                    X(Index) = ""
                    DoSpell = True
                    Exit Function
                End If
                'HealingSpell False, dbIndex, dbSpell, x(Index)
                modSpells.DoNonCombatSpell Index, dbIndex, dbSpell, True, dbT
                
                X(Index) = ""
                SpellCombat(Index) = False
                DoSpell = True
                Exit Function
            Else
                'HealingSpell True, dbIndex, dbSpell
                modSpells.DoNonCombatSpell Index, dbIndex, dbSpell
                X(Index) = ""
                SpellCombat(Index) = False
                DoSpell = True
                Exit Function
            End If
        Case 1
            If dbMap(GetMapIndex(dbPlayers(dbIndex).lLocation)).iSafeRoom = 1 Then
                WrapAndSend Index, RED & "You are unable to intiate combat here." & vbCrLf & WHITE
                SpellCombat(Index) = False
                X(Index) = ""
                DoSpell = True
                Exit Function
            End If
        Case 4
            If dbMap(GetMapIndex(dbPlayers(dbIndex).lLocation)).iSafeRoom = 1 Then
                WrapAndSend Index, RED & "You are unable to intiate combat here." & vbCrLf & WHITE
                SpellCombat(Index) = False
                X(Index) = ""
                DoSpell = True
                Exit Function
            Else
                X(Index) = ""
                bRoom = True
            End If
            
'        Case 2
'            TeleportSpell Index, CLng(tArr2(iID&))
'            x(Index) = ""
'            DoSpell = True
'            Exit Function
'        Case 3
'            If InStr(s, " ") > 0 Then
'                s$ = Mid$(s, InStr(1, s$, " ") + 1)
'                s$ = SmartFind(Index, s$, Player_In_Room)
'                s$ = GetPlayerIndexNumber(, s$)
'                If s$ = "0" Then
'                    WrapAndSend Index, RED & "You cannot seem to find them here." & WHITE & vbCrLf
'                    x(Index) = ""
'                    DoSpell = True
'                    Exit Function
'                End If
'                BlessOthers Index, CLng(i), CLng(s)
'                x(Index) = ""
'                DoSpell = True
'                Exit Function
'            Else
'                BlessMe CLng(i), dbIndex
'                x(Index) = ""
'                DoSpell = True
'                Exit Function
'            End If
'        Case 4
'            bRoom = True
'            DoSpell = False
        Case 5
            dbT = GetPlayerIndexNumber(Index)
'            With dbPlayers(dbT)
'                If Not modSC.FastStringComp(.sParty, "0") Then
'                    SplitFast ReplaceFast(.sParty, ":", ""), tArr2, ";"
'                    For i = LBound(tArr2) To UBound(tArr2)
'                        If Not modSC.FastStringComp(tArr2(i), "") Then
'                            If CLng(tArr2(i)) <> .iIndex Then
'                                modSpells.DoNonCombatSpell Index, dbT, dbSpell, True, GetPlayerIndexNumber(CLng(tArr2(i)))
'                            Else
                                modSpells.DoNonCombatSpell Index, dbT, dbSpell
                                DoSpell = True
                                Exit Function
'                            End If
'                        End If
'                        If DE Then DoEvents
'                    Next
'                End If
'            End With
    End Select
End With
DoSpell = False
End Function

Public Function SpellMon(Index As Long, dbIndex As Long, s As String, sSpellName As String) As Boolean
Dim i As Long
Dim j As Long
Dim R As String
Dim t As String
Dim n As Long
Dim Arr() As String
With dbPlayers(dbIndex)
    t = dbMap(.lDBLocation).sAMonIds
    If t <> "" Then
        SplitFast t, Arr, ";"
        R = LCaseFast(s)
        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> "" Then
                n = CLng(Val(Arr(i)))
                If .lLocation = aMons(n).mLoc And modSC.FastStringComp(R, LCaseFast(aMons(n).mName)) Then
                    SpellMon = True
                    If aMons(i).mAttackable = True Then
                        .iCasting = dbSpells(GetSpellID(sSpellName)).lID
                        .dMonsterID = n
                        aMons(n).mIs_Being_Attacked = True
                        aMons(n).mPlayerAttacking = .iIndex
                        For j = 0 To 4
                            aMons(n).mSpells(j).lCastPerRound = 0
                            DoEvents
                        Next
                        modMonsters.InsertInMonList n, .lPlayerID, 0
                        AddEvil dbIndex, n
                        WrapAndSend Index, BRIGHTRED & "You move to cast " & sSpellName & " on " & aMons(n).mName & "!" & vbCrLf & WHITE
                        SendToAllInRoom Index, BRIGHTRED & .sPlayerName & " moves to cast " & sSpellName & " on " & aMons(n).mName & "." & WHITE & vbCrLf, .lLocation
                        X(Index) = ""
                        Exit Function
                    Else
                        WrapAndSend Index, RED & "You seem unable to attack that being." & WHITE & vbCrLf
                        X(Index) = ""
                        Exit Function
                    End If
                End If
            End If
            If DE Then DoEvents
        Next
    End If
End With
SpellMon = False
End Function

Sub SpellPlayer(Index As Long, dbIndex As Long, s As String, SpellName As String)
Dim dbVIndex As Long
Dim a&
Dim l&
Dim u&
dbVIndex = GetPlayerIndexNumber(, s)
If dbVIndex = 0 Then
    WrapAndSend Index, RED & "You do not see them here!" & WHITE & vbCrLf
    X(Index) = ""
    Exit Sub
End If
a& = dbPlayers(dbIndex).iLevel
l& = dbPlayers(dbVIndex).iLevel - lPvPLevel
u& = dbPlayers(dbVIndex).iLevel + lPvPLevel
If (a& < l&) Or (a& > u&) Then
    WrapAndSend Index, RED & "You are about to attack " & dbPlayers(dbVIndex).sPlayerName & ", but decide not to." & WHITE & vbCrLf
    X(Index) = ""
    Exit Sub
End If
If InStr(1, LCaseFast(modGetData.GetPlayersHereWithoutRiding(CLng(dbPlayers(dbIndex).lLocation), dbIndex)), s) = 0 Then
    WrapAndSend Index, RED & "You do not see them here!" & WHITE & vbCrLf
    X(Index) = ""
    Exit Sub
End If
With dbPlayers(dbVIndex)
    WrapAndSend .iIndex, BRIGHTRED & dbPlayers(dbIndex).sPlayerName & " moves to cast " & SpellName & " on you!" & WHITE & vbCrLf
    dbPlayers(dbIndex).iPlayerAttacking = .iIndex
    AddPvPEvil Index, .iIndex
End With
With dbPlayers(dbIndex)
    .iCasting = CLng(dbSpells(GetSpellID(SpellName)).lID)
    SendToAllInRoom Index, BRIGHTRED & .sPlayerName & " moves to cast " & SpellName & " on " & s & "." & WHITE & vbCrLf, .lLocation, dbPlayers(dbVIndex).iIndex
End With
WrapAndSend Index, BRIGHTRED & "You move to cast " & SpellName & " on " & s & "!" & vbCrLf & WHITE
X(Index) = ""
End Sub

Sub DoTheRoom(iPID As Long, SpellID As Long, SpellName As String)
Dim s As String
Dim sP As String
Dim tArr() As String
Dim bFound As Boolean
Dim Index As Long
Index = dbPlayers(iPID).iIndex
s = modGetData.GetAllMonstersInRoom(dbPlayers(iPID).lLocation)
sP = modGetData.GetPlayersDBIndexesHere(dbPlayers(iPID).lLocation)
If (modSC.FastStringComp(s, "")) And (modMain.DCount(sP, ";") <= 1) Then
    WrapAndSend CLng(Index), RED & "There is nothing here." & WHITE & vbCrLf
    SpellCombat(Index) = False
    dbPlayers(iPID).iCasting = 0
    X(Index) = ""
    Exit Sub
End If
Erase tArr
If Not modSC.FastStringComp(s, "") Then
    SplitFast Left$(s, Len(s) - 1), tArr, ";"
Else
    ReDim tArr(0) As String
    tArr(0) = "-1"
End If
With dbPlayers(iPID)
    bFound = False
    For i = LBound(tArr) To UBound(tArr)
        If tArr(i) = "-1" Then Exit For
        If aMons(tArr(i)).mAttackable = True Then
            bFound = True
            aMons(tArr(i)).mIs_Being_Attacked = True
            aMons(tArr(i)).mPlayerAttacking = .iIndex
            AddEvil iPID, Val(tArr(i))
            'send messages
        Else
            'do nothing
        End If
        If DE Then DoEvents
    Next
    Erase tArr
    bFound = False
    If Not modSC.FastStringComp(sP, .iIndex & ";") Then
        SplitFast Left$(sP, Len(sP) - 1), tArr, ";"
        For i = LBound(tArr) To UBound(tArr)
            If CLng(tArr(i)) <> .iIndex Then
                bFound = True
                'send messages
            Else
                'do nothing
            End If
            If DE Then DoEvents
        Next
    End If
    If bFound Or Not modSC.FastStringComp(sP, "") Then
        .dMonsterID = 99998
        SpellCombat(Index) = True
        .iCasting = SpellID
        SetWeaponStats Index
        .dStamina = .dStamina - RndNumber(0, 4)
        .dHunger = .dHunger - RndNumber(0, 2)
        WrapAndSend Index, BRIGHTRED & "You move to cast " & SpellName$ & " on " & "the room!" & vbCrLf & WHITE
        SendToAllInRoom Index, BRIGHTRED & .sPlayerName & " moves to cast " & SpellName$ & " on the room." & WHITE & vbCrLf, .lLocation
        X(Index) = ""
        Exit Sub
    End If

End With
End Sub

Public Function Attack(Index As Long, Optional dbIndex As Long) As Boolean

Dim Short$, MayBeSpell As String, IDs$
Dim tArr() As String, CastingID&
Dim SpellName$
Dim iPID As Long
Dim bFound As Boolean
Dim bRoom As Boolean
Dim s As String
Dim i As Long
Dim bResult As Boolean
bRoom = False
bFound = False
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 1)), "a") Then
    If Mid$(LCaseFast(X(Index)), 2, 1) <> " " And Mid$(LCaseFast(X(Index)), 2, 1) <> "t" Then
        Exit Function
    End If
    Attack = True
    s = X(Index)
    s = Mid$(s, InStr(1, s, " ") + 1)
    s = SmartFind(Index, s, Monster_In_Room)
    If dbIndex <> 0 Then
        iPID = dbIndex
    Else
        iPID = GetPlayerIndexNumber(Index)
    End If
    If dbPlayers(iPID).iGhostMode = 1 Then
        WrapAndSend Index, RED & "You may not attack while in ghost mode." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If modMiscFlag.GetMiscFlag(iPID, [Can Attack]) = 1 Then
        WrapAndSend Index, RED & "Something is stopping you from attacking." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    dbPlayers(iPID).iSneaking = 0
    With dbPlayers(iPID)
        If dbMap(GetMapIndex(dbPlayers(iPID).lLocation)).iSafeRoom = 1 Then
            WrapAndSend Index, RED & "You are unable to intiate combat here." & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        bResult = CheckAttackMonster(Index, iPID, s)
    End With
    If lIsPvP = 1 Then
        If Not bResult And GetMonsterID(LCaseFast(s)) = 0 Then
            s = SmartFind(Index, s, Player_In_Room)
            bResult = CheckPlayerAttack(Index, iPID, s)
        End If
    End If
    
    If Not bResult Then
        WrapAndSend Index, RED & "You do not see that here!" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbPlayers(iPID).iResting = 0
    dbPlayers(iPID).iMeditating = 0
ElseIf modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "bs ") Then
    Attack = True
    s = X(Index)
    s = Mid$(s, InStr(1, s, " ") + 1)
    s = SmartFind(Index, s, Monster_In_Room)
    If dbIndex <> 0 Then
        iPID = dbIndex
    Else
        iPID = GetPlayerIndexNumber(Index)
    End If
    If dbPlayers(iPID).iGhostMode = 1 Then
        WrapAndSend Index, RED & "You may not attack while in ghost mode." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If modMiscFlag.GetMiscFlag(iPID, [Can Attack]) = 1 Then
        WrapAndSend Index, RED & "Something is stopping you from attacking." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    If dbPlayers(iPID).iSneaking = 0 Or modMiscFlag.GetMiscFlag(iPID, [Can Backstab]) <> 1 Then
        With dbPlayers(iPID)
            .iSneaking = 0
            If dbMap(dbPlayers(iPID).lDBLocation).iSafeRoom = 1 Then
                WrapAndSend Index, RED & "You are unable to intiate combat here." & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
            End If
            bResult = CheckAttackMonster(Index, iPID, s)
        End With
        
        If lIsPvP = 1 Then
            If Not bResult And GetMonsterID(LCaseFast(s)) = 0 Then
                s = SmartFind(Index, s, Player_In_Room)
                bResult = CheckPlayerAttack(Index, iPID, s)
            End If
        End If
        
        If Not bResult Then
            WrapAndSend Index, RED & "You do not see that here!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        dbPlayers(iPID).iResting = 0
        dbPlayers(iPID).iMeditating = 0
    Else
        With dbPlayers(iPID)
            .iSneaking = 0
            If dbMap(dbPlayers(iPID).lDBLocation).iSafeRoom = 1 Then
                WrapAndSend Index, RED & "You are unable to intiate combat here." & vbCrLf & WHITE
                X(Index) = ""
                Exit Function
            End If
            .iIsBSing = 1
            bResult = CheckAttackMonster(Index, iPID, s)
        End With
        
        If lIsPvP = 1 Then
            If Not bResult And GetMonsterID(LCaseFast(s)) = 0 Then
                s = SmartFind(Index, s, Player_In_Room)
                bResult = CheckPlayerAttack(Index, iPID, s)
            End If
        End If
        
        If Not bResult Then
            WrapAndSend Index, RED & "You do not see that here!" & WHITE & vbCrLf
            X(Index) = ""
            dbPlayers(iPID).iIsBSing = 0
            Exit Function
        End If
        dbPlayers(iPID).iResting = 0
        dbPlayers(iPID).iMeditating = 0
    End If
Else
    bFound = False
    If dbIndex <> 0 Then
        iPID = dbIndex
    Else
        iPID = GetPlayerIndexNumber(Index)
    End If
    If iPID <> 0 Then
        With dbPlayers(iPID)
            If .sSpells = "0" Then Exit Function
            SplitFast .sSpellShorts, tArr, ";"
            For i = LBound(tArr) To UBound(tArr)
                If Not modSC.FastStringComp(tArr(i), "") Then
                    If modSC.FastStringComp(Left$(X(Index), 4), tArr(i)) Then
                        bFound = True
                        Erase tArr
                        Exit For
                    End If
                End If
                If DE Then DoEvents
            Next
        End With
    Else
        Exit Function
    End If
    If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "cast ") Or bFound Then    'if they are casting a spell
        If dbPlayers(iPID).iGhostMode = 1 Then
            WrapAndSend Index, RED & "You may not attack while in ghost mode." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        Attack = True
        If modMiscFlag.GetMiscFlag(iPID, [Can Cast Spell]) = 1 Then
            WrapAndSend Index, RED & "Something is stopping you from attacking." & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
        s = X(Index)
        If Not bFound Then s = Mid$(s, 6, Len(s) - 5)
        bFound = False
        MayBeSpell = LCaseFast(TrimIt(Mid$(s, 1, InStr(1, s, " "))))
        If modSC.FastStringComp(MayBeSpell, "") Then MayBeSpell = LCaseFast(TrimIt(s))
        Short$ = modGetData.GetPlayersSpellShorts(Index)
        IDs$ = modGetData.GetPlayersSpellIds(Index)
        bFound = False
        If Short$ <> "0" Then
            bResult = IsSpell(MayBeSpell, Short$, CastingID&)
            If Not bResult Then
                WrapAndSend Index, RED & "You currently don't know that spell." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
            SpellCombat(Index) = True
            IDs$ = ReplaceFast(IDs$, ":", "")
            dbPlayers(iPID).iSneaking = 0
            bResult = DoSpell(Index, iPID, s, IDs, CastingID&, SpellName$, bRoom)
            If bResult = False Then
                If Not bRoom Then
                    s = Mid$(s, InStr(1, s, " ") + 1)
                    s = SmartFind(Index, s, Monster_In_Room)
                    bResult = SpellMon(Index, iPID, s, SpellName)
                    If lIsPvP = 1 Then
                        If Not bResult And GetMonsterID(LCaseFast(s)) = 0 Then
                            With dbPlayers(iPID)
                                .dStamina = .dStamina - RndNumber(0, 2)
                                .dHunger = .dHunger - RndNumber(0, 1)
                            End With
                            s = SmartFind(Index, s, Player_In_Room)
                            SpellPlayer Index, iPID, s, SpellName$
                        End If
                    End If
                Else
                    Erase tArr
                    SplitFast IDs, tArr, ";"
                    DoTheRoom iPID, Val(tArr(CastingID)), SpellName
                End If
            Else
                With dbPlayers(iPID)
                    .dStamina = .dStamina - RndNumber(0, 2)
                    .dHunger = .dHunger - RndNumber(0, 1)
                End With
            End If
            dbPlayers(iPID).iResting = 0
            dbPlayers(iPID).iMeditating = 0
            X(Index) = ""
        End If
    End If
End If
End Function

Public Function AttackCommands(Index As Long, Optional dbIndex As Long) As Boolean
If Break(Index) = True Then AttackCommands = True: Exit Function
If Attack(Index, dbIndex) = True Then AttackCommands = True: Exit Function 'check for the 'a' or 'cast' command
AttackCommands = False
End Function

Public Function CheckDeath(Index As Long, Optional IsVictim As Boolean = False, Optional CombatMessage As Boolean = False, Optional ByRef Messages1 As String = "", Optional ByRef Messages2 As String = "", Optional ByRef Messages3 As String = "", Optional ByRef PartyFlag As Boolean = False, Optional SendStatline As Boolean = True, Optional ByRef IsDropped As Boolean, Optional dbIndex As Long = 0, Optional lAMONINDEX As Long) As Boolean
'checking player death
Dim tArr1() As String, tArr2() As String, aFlgs() As String
Dim lDeath As Long
Dim i As Long
Dim j As Long
Dim s As String
Dim sEQ As String

CheckDeath = False
If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    If .lHP <= 0 Then
        IsDropped = True
        .iDropped = 1
        .iHorse = 0
'        If CombatMessage = True Then
'            BreakOffCombat dbIndex, True, True, IsVictim, Messages1, Messages2, Messages3
'        Else
        BreakOffCombat dbIndex, True
        'End If
        .iPlayerAttacking = 0
        If .lRegain <> 1 Then .lRegain = -1
        'If CombatMessage And .lHP > lDeath Then
        '    WrapAndSend ATTACKINGINDEX, CombatMessagesATTACKER, SendStatline
        '    If VICTIMINDEX <> 0 Then WrapAndSend VICTIMINDEX, CombatMessagesVICTIM, SendStatline
        '    SendToAllInRoom ATTACKINGINDEX, CombatMessagesROOM & vbCrLf, .lLocation, VICTIMINDEX, SendStatline
        'End If
        If .iHasSentDropped <> 1 And CombatMessage = False Then
            WrapAndSend Index, BGRED & "You fall to the ground!" & WHITE & vbCrLf
            SendToAllInRoom Index, BGRED & .sPlayerName & BRIGHTRED & " falls to the ground!" & WHITE & vbCrLf, .lLocation
            .iHasSentDropped = 1
        ElseIf .iHasSentDropped <> 1 And CombatMessage = True Then
            Messages1 = Messages1 & BGRED & "You fall to the ground!" & WHITE & vbCrLf
            Messages3 = Messages3 & BGRED & .sPlayerName & BRIGHTRED & " falls to the ground!" * WHITE & vbCrLf
            .iHasSentDropped = 1
        End If
        If .lHP <= lDeath And CombatMessage = True Then
            PartyFlag = True
            GoTo DIENOW
        ElseIf .lHP > lDeath And CombatMessage = True Then
            PartyFlag = True
        Else
            RemoveFromParty Index
        End If
    Else
        .iDropped = 0
        .iHasSentDropped = 0
    End If
End With
If dbPlayers(dbIndex).lHP <= lDeath Then  'get the death level
DIENOW:
    CheckDeath = True
    i = GetMapIndex(CLng(dbPlayers(dbIndex).lLocation))
    With dbMap(i)
        If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
        If dbPlayers(dbIndex).sInventory <> "0" Then .sItems = .sItems & dbPlayers(dbIndex).sInventory
        If modSC.FastStringComp(.sItems, "") Then .sItems = "0"
        '.dGold = .dGold + dbPlayers(dbIndex).dGold
        lDeath = .lDeathRoom
    End With
    With dbPlayers(dbIndex)
        If CombatMessage = True Then
            Messages1 = Messages1 & BGGREEN & "You have been killed!" & WHITE & vbCrLf
            If IsVictim = True Then Messages2 = Messages2 & bgreen & "You have killed " & .sPlayerName & "!" & WHITE & vbCrLf
            Messages3 = Messages3 & bgreen & .sPlayerName & " has been killed!" & WHITE & vbCrLf
        Else
            SendToAllInRoom Index, BGLIGHTBLUE & .sPlayerName & " has just been killed!" & WHITE & vbCrLf, .lLocation
            WrapAndSend Index, BGGREEN & "You have been killed!" & WHITE & vbCrLf
        End If
    
        sEQ = modGetData.GetPlayersEq(Index, dbIndex)
        If lAMONINDEX <> -1 Then
            aMons(lAMONINDEX).mPEQ = aMons(lAMONINDEX).mPEQ & sEQ
            aMons(lAMONINDEX).mMoney = .dGold
            If CombatMessage = True Then
                If sEQ <> "" Then Messages3 = Messages3 & RED & aMons(lAMONINDEX).mName & GREEN & " takes " & ReplaceFast(modGetData.GetPlayersEqFromNums(Index, True, dbIndex), ",", ", ") & " from " & vbRed & .sPlayerName & GREEN & "." & WHITE & vbCrLf
                If .dGold <> "" Then Messages3 = Messages3 & RED & aMons(lAMONINDEX).mName & GREEN & " takes " & LIGHTBLUE & CStr(.dGold) & YELLOW & " gold " & GREEN & "from " & RED & .sPlayerName & GREEN & "." & WHITE & vbCrLf
            End If
            With dbMap(i)
                If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
                'If dbPlayers(dbIndex).sInventory <> "0" Then .sItems = .sItems & dbPlayers(dbIndex).sInventory
                If modSC.FastStringComp(.sItems, "") Then .sItems = "0"
            End With
        Else
            With dbMap(i)
                If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
                .sItems = .sItems & sEQ
                .dGold = .dGold + dbPlayers(dbIndex).dGold
                If modSC.FastStringComp(.sItems, "") Then .sItems = "0"
            End With
        End If
        .dGold = 0
        .lPaper = 0
        .sInventory = "0"
        If .sArms <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sArms), 0
        If .sBack <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sBack), 0
        If .sBody <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sBody), 0
        If .sEars <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sEars), 0
        If .sFace <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sFace), 0
        If .sFeet <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sFeet), 0
        If .sHands <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sHands), 0
        If .sHead <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sHead), 0
        If .sLegs <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sLegs), 0
        If .sNeck <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sNeck), 0
        For j = 0 To 5
            If .sRings(j) <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sRings(j)), 0
            If DE Then DoEvents
        Next
        If .sWaist <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sWaist), 0
        If .sWeapon <> "0" Then modItemManip.AdjustStats dbIndex, modItemManip.GetItemIDFromUnFormattedString(.sWeapon), 0
        
        .sArms = "0"
        .sBack = "0"
        .sBody = "0"
        .sEars = "0"
        .sFace = "0"
        .sFeet = "0"
        .sHands = "0"
        .sHead = "0"
        .sLegs = "0"
        .sNeck = "0"
        .sRings(0) = "0"
        .sRings(1) = "0"
        .sRings(2) = "0"
        .sRings(3) = "0"
        .sRings(4) = "0"
        .sRings(5) = "0"
        .sShield = "0"
        .sWaist = "0"
        .sWeapon = "0"
'        If CombatMessage = True Then
'            BreakOffCombat dbIndex, False, True, IsVictim, Messages1, Messages2, Messages3
'        Else
            BreakOffCombat dbIndex
        'End If
        .lHP = .lMaxHP \ 2
        If .lHP < 1 Then .lHP = 1
        .lMana = .lMana \ 3
        If .lMana < 1 Then .lMana = 1
        .lRegain = 0
        'If .iPlayerAttacking <> 0 Then
        '    If dbPlayers(GetPlayerIndexNumber(.iPlayerAttacking)).iPlayerAttacking = dbIndex Then
        '        BreakOffCombat GetPlayerIndexNumber(.iPlayerAttacking)
        '        .iPlayerAttacking = 0
        '    End If
        'End If
        .iLives = .iLives - 1
        .iDropped = 0
        .iHorse = 0
        .iHasSentDropped = 0
        .dHunger = 100
        .dStamina = CDbl(RndNumber(50, 100))
        If .sBlessSpells <> "0" Then
            SplitFast Left$(.sBlessSpells, Len(.sBlessSpells) - 1), tArr1, "Œ"
            'Timeout~roll~spellname~dbspellid
            For j = LBound(tArr1) To UBound(tArr1)
                Erase tArr2
                SplitFast tArr1(j), tArr2, "~"
                tArr2(0) = CLng(tArr2(0)) - 1
                If CLng(tArr2(0)) <= 0 Then
                    If Not modSC.FastStringComp(dbSpells(Val(tArr2(3))).sFlags, "0") Then
                        modUseItems.DoFlags dbIndex, dbSpells(Val(tArr2(3))).sFlags, lRoll:=CLng(tArr2(1)), Inverse:=True
                    End If
                    .sBlessSpells = ReplaceFast(.sBlessSpells, CStr((Val(tArr2(0)) + 1)) & "~" & tArr2(1) & "~" & tArr2(2) & "~" & tArr2(3) & "Œ", "", 1, 1)
                    If modSC.FastStringComp(.sBlessSpells, "") Then .sBlessSpells = "0"
                    sSend Index, LIGHTBLUE & dbSpells(Val(tArr2(3))).sRunOutMessage, , CombatMessage, Messages1
                Else
                    s = CStr((CLng(tArr2(0)) + 1)) & "~" & tArr2(1) & "~" & tArr2(2) & "~" & tArr2(3) & "Œ"
                    
                    .sBlessSpells = ReplaceFast(.sBlessSpells, s, Join(tArr2, "~") & "Œ", 1, 1)
                End If
                If DE Then DoEvents
            Next j
        End If
        If .iLives <= 0 Then
            WrapAndSend Index, Messages1
            ReRollEm Index
        Else
            SendToAllInRoom Index, BGLIGHTBLUE & .sPlayerName & " appears in the room in a flash of white light!" & WHITE & vbCrLf, lDeath
            RemoveStats Index, CombatMessage
        End If
        .lLocation = lDeath
        .lDBLocation = GetMapIndex(lDeath)
        If CombatMessage = True Then
            PartyFlag = True
        Else
            RemoveFromParty Index
        End If
    End With
End If
End Function

Public Function CheckIsThere(Index As Long, MonsterID As Long) As Boolean
'makeing sure the player is still with the monster
Dim a As Long
On Error GoTo eh1
CheckIsThere = True
a = GetPlayerIndexNumber(Index)
If dbPlayers(a).lLocation <> aMons(MonsterID).mLoc Then   'finding out if the player
                'is at the location of the monster
    CheckIsThere = False
    If RndNumber(0, 100) > 55 Then
        If RoamMonsters(MonsterID, dbPlayers(a).lLocation, modGetData.GetRoomExitFrom2Points(aMons(MonsterID).mdbMapID, dbPlayers(a).lLocation), dbPlayers(a).lDBLocation) = True Then
            CheckIsThere = True
            Exit Function
        End If
    End If
    If aMons(MonsterID).mPlayerAttacking = Index Then 'if not, stop the monster from attacking
        aMons(MonsterID).mPlayerAttacking = 0
        aMons(MonsterID).mIsAttacking = False
        aMons(MonsterID).mIs_Being_Attacked = False
        modMonsters.RemoveFromMonList MonsterID, dbPlayers(a).lPlayerID
        dbPlayers(a).dMonsterID = 99999
    End If
    SpellCombat(Index) = False
    dbPlayers(a).iCasting = 0
    WrapAndSend Index, YELLOW & "You have left combat." & WHITE & vbCrLf
    Exit Function
'ElseIf dbPlayers(a).dMonsterID = 99998 Then
'    CheckIsThere = False
End If
eh1:
End Function

Public Function CheckShield(Index As Long, Optional dbIndex As Long) As Boolean
Dim psShield As Long
Dim piStr As Long
Dim piLevel As Long
Dim Chance As Double
If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    If modSC.FastStringComp(.sShield, "0") Then
        CheckShield = False
        Exit Function
    End If
    psShield = modItemManip.GetItemIDFromUnFormattedString(.sShield)
    piStr = .iStr
    piLevel = .iLevel
End With
With dbItems(GetItemID(, psShield))
    If .sWorn = "shield" Then
        Chance = RndNumber(0, .iAC + .iArmorType + piStr + piLevel)
        If Chance > (.iAC + .iArmorType + piStr + piLevel) / 1.1 Then
            CheckShield = True
        Else
            CheckShield = False
        End If
    End If
End With
End Function

Public Function PvPAcc(AttackingDBIndex As Long) As Long
    Dim CurVal As Long
    With dbPlayers(AttackingDBIndex)
        CurVal = .iAcc
    End With
    PvPAcc = CurVal
End Function

Public Function PvPCleanUp(Index1 As Long, Index2 As Long)
With dbPlayers(GetPlayerIndexNumber(Index2))
    .iPlayerAttacking = 0
    .iCasting = 0
    SpellCombat(.iIndex) = False
End With
With dbPlayers(Index1)
    .iPlayerAttacking = 0
    .iCasting = 0
    SpellCombat(.iIndex) = False
End With
End Function

Public Function PvPCrits(AttackingDBIndex As Long) As Long
    Dim CurVal As Long
    With dbPlayers(AttackingDBIndex)
        CurVal = .iCrits
    End With
    PvPCrits = CurVal
End Function

Public Function PvPDodge(AttackingDBIndex As Long, IndexNumber As Long) As Long
    Dim d As Double
    With dbPlayers(AttackingDBIndex)
        d = (.iAgil + .iDodge + .iAC) / 100
        d = RoundFast(d, 2)
        d = d * 100
        If d < 0 Then d = 0
        If d > 62 Then d = 62
        If .iPartyRank = 2 Then d = d + 10
        If d > 70 Then d = 70
    End With
    PvPDodge = CLng(d)
    
End Function

'Public Function PvPSpellMax(AttackingDBIndex As Long) As Long
'    Dim CurVal As Long
'    With dbPlayers(AttackingDBIndex)
'        CurVal = .iSC
'    End With
'    PvPSpellMax = CurVal
'End Function

'Public Function PvPSpellResist(AttackingDBIndex As Long) As Long
'    Dim CurVal As Long
'    CurVal = 100 - PvPSpellMax(AttackingDBIndex)
'    PvPSpellResist = CurVal
'End Function



Public Function PvPStillThere(Index1 As Long, Index2 As Long) As Boolean
If Index2 = 0 Then Exit Function
If dbPlayers(Index1).lLocation <> dbPlayers(GetPlayerIndexNumber(Index2)).lLocation Then
    WrapAndSend dbPlayers(Index1).iIndex, YELLOW & "You have dis-engaged combat." & WHITE & vbCrLf
    WrapAndSend Index2, YELLOW & "You have dis-engaged combat." & WHITE & vbCrLf
    PvPCleanUp Index1, Index2
    PvPStillThere = False
Else
    PvPStillThere = True
End If
End Function

Public Function PvPWeaponMax(AttackingDBIndex As Long) As Long
Dim d As Double
With dbPlayers(AttackingDBIndex)
    d = (.iDex + .iStr) / 100
    d = RoundFast(d, 2)
    d = d * 100
    If d > 98 Then d = 98
    If d < 0 Then d = 0
    PvPWeaponMax = CLng(d)
End With
    
End Function

Public Function ReSetMonsterID(dbIndex As Long)
With dbPlayers(dbIndex)
    .dMonsterID = 99999
End With
End Function

Sub AddEXP(Index As Long, MonsterID As Long, Optional ByRef famexp As Double)
Dim d As Double
Dim dbIndex As Long
Dim CurrentFamLevel As Long
Dim NowFamLevel As Long
Dim dbFamId As Long
dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    .dEXP = .dEXP + aMons(MonsterID).mEXP
    .dTotalEXP = .dTotalEXP + aMons(MonsterID).mEXP
    If .lFamID <> 0 Then
        d = aMons(MonsterID).mEXP \ RndNumber(3, 100)
        If d < 1 Then d = 1
        If d > .dFamEXPN Then d = .dFamEXPN
        famexp = d
        dbFamId = GetFamID(.lFamID)
        .dFamCEXP = .dFamCEXP + d
        .dFamTEXP = .dFamTEXP + d
        If .dFamCEXP >= .dFamEXPN Then
            .lFamLevel = .lFamLevel + 1
            .lFamMHP = .lFamMHP + RndNumber(CDbl(dbFamiliars(dbFamId).lLevelMod), CDbl(.lFamLevel))
            .lFamAcc = .lFamAcc + RndNumber(0, 1)
            .dFamCEXP = 0
            .dFamEXPN = .dFamEXPN + RndNumber(1, CDbl(.lFamLevel))
        End If
    End If
End With
End Sub

Sub CleanUpSpells(Index As Long)
'clean up the spell combat
'make it so they aren't casting
With dbPlayers(GetPlayerIndexNumber(Index))
    .iCasting = 0
End With
SpellCombat(Index) = False 'disable the value
pWeapon(Index).wSpellName = "" 'dump the spellname
End Sub

Sub ClearOtherAttackers(Index As Long, MonsterID As Long)
'after the monster is dead, all the ones attacking it must stop
Dim dVal As Double
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        If Index <> .iIndex Then
            If .dMonsterID = CDbl(MonsterID) Then
                .dMonsterID = 99999
                dVal = RoundFast(CDbl(aMons(MonsterID).mEXP / 4), 0)
                .dEXP = .dEXP + dVal
                .dTotalEXP = .dTotalEXP + dVal
                .dClassPoints = .dClassPoints + 0.05
                WrapAndSend .iIndex, BRIGHTWHITE & "You gain " & CStr(dVal) & " experience." & WHITE & vbCrLf
            End If
        End If
    End With
    If DE Then DoEvents
Next
End Sub

Sub DropMonGold(MonsterID As Long, ByRef Message As String)
'drops the monster gold to the floor
On Error GoTo eh1
Dim a As Long
With dbMap(aMons(MonsterID).mdbMapID)
    a = RndNumber(0, CDbl(aMons(MonsterID).mMoney)) + aMons(MonsterID).mPMoney
    .dGold = .dGold + a
    If a <> 0 Then Message = GREEN & aMons(MonsterID).mName & LIGHTBLUE & " drops " & CStr(a) & " gold on the ground!" & vbCrLf
End With
eh1:
End Sub

Sub DropMonItem(MonsterID As Long, ByRef Message As String)
'drops an item if the monster drops one
On Error GoTo eh1
Dim dbItemID As Long
Dim dbCorpseID As Long
Dim dbMapIndex As Long
Dim i As Long
Dim Arr() As String
Dim Arr2() As String
With dbMonsters(aMons(MonsterID).mdbMonID)
    If .sDropItem = "0" And .iDropCorpse = 0 Then Exit Sub
    dbMapIndex = aMons(MonsterID).mdbMapID
    If ReplaceFast(aMons(MonsterID).mPEQ, "0", "") <> "" Then
        With dbMap(dbMapIndex)
            If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
            SplitFast aMons(MonsterID).mPEQ, Arr, ";"
            For i = LBound(Arr) To UBound(Arr)
                If Arr(i) <> "" And Arr(i) <> "0" Then
                    .sItems = .sItems & Arr(i)
                    With dbItems(GetItemID(, modItemManip.GetItemIDFromUnFormattedString(Arr(i))))
                        Message = Message & GREEN & aMons( _
                            MonsterID).mName & LIGHTBLUE & " drops " & BRIGHTRED & _
                            ReplaceFast( _
                            modItemManip.GetItemAdjectivesFromUnFormattedString(Arr( _
                            i)), "|", _
                            " ") & .sItemName & LIGHTBLUE & " to the ground!" & _
                            vbCrLf
                    End With
                End If
                If DE Then DoEvents
            Next
            aMons(MonsterID).mPEQ = ""
            Erase Arr
        End With
    End If
    If .sDropItem <> "0" Then
        SplitFast .sDropItem, Arr, ";"
        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> "" Then
                SplitFast Arr(i), Arr2, "/"
                dbItemID = GetItemID(, Val(Arr2(0)))
                If dbItemID = 0 Then GoTo nNext
                If RndNumber(1, 100) > Val(Arr2(1)) Then GoTo nNext 'random chance
                If dbItems(dbItemID).iInGame >= dbItems(dbItemID).iLimit And dbItems(dbItemID).iLimit <> 0 Then
                
                Else
                    If dbItems(dbItemID).iLimit <> 0 Then dbItems(dbItemID).iInGame = dbItems(dbItemID).iInGame + 1
                    With dbMap(dbMapIndex)
                        If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
                        .sItems = .sItems & ":" & dbItems(dbItemID).iID & "/" & dbItems(dbItemID).lDurability & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(dbItemID).iUses & ";"
                        Message = Message & GREEN & aMons(MonsterID).mName & LIGHTBLUE & " drops " & BRIGHTRED & dbItems(dbItemID).sItemName & LIGHTBLUE & " to the ground!" & vbCrLf
                    End With
                End If
            End If
nNext:
            If DE Then DoEvents
        Next
    End If
    If .iDropCorpse <> 0 Then
        dbCorpseID = GetItemID(, CLng(.iDropCorpse))
        If dbItems(dbCorpseID).iInGame >= dbItems(dbCorpseID).iLimit And dbItems(dbCorpseID).iLimit <> 0 Then
        Else
            If dbItems(dbCorpseID).iLimit <> 0 Then dbItems(dbCorpseID).iInGame = dbItems(dbCorpseID).iInGame + 1
            With dbMap(dbMapIndex)
                If modSC.FastStringComp(.sItems, "0") Then .sItems = ""
                If dbCorpseID <> 0 Then
                    .sItems = .sItems & ":" & dbItems(dbCorpseID).iID & "/" & dbItems(dbCorpseID).lDurability & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(dbCorpseID).iUses & ";"
                    Message = Message & GREEN & aMons(MonsterID).mName & LIGHTBLUE & " falls to the ground, dead." & vbCrLf
                End If
            End With
        End If
    End If
End With
eh1:
End Sub

Sub ReRollEm(Index As Long)
Dim tVar As String, lLoc As Long, tdGold As Double
Dim dbIndex As Long
dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    If Not modSC.FastStringComp(.sInventory, "0") Then tVar = tVar & .sInventory
    If Not modSC.FastStringComp(.sArms, "0") Then tVar = tVar & .sArms & ";"
    If Not modSC.FastStringComp(.sBody, "0") Then tVar = tVar & .sBody & ";"
    If Not modSC.FastStringComp(.sFeet, "0") Then tVar = tVar & .sFeet & ";"
    If Not modSC.FastStringComp(.sHands, "0") Then tVar = tVar & .sHands & ";"
    If Not modSC.FastStringComp(.sHead, "0") Then tVar = tVar & .sHead & ";"
    If Not modSC.FastStringComp(.sLegs, "0") Then tVar = tVar & .sLegs & ";"
    If Not modSC.FastStringComp(.sWaist, "0") Then tVar = tVar & .sWaist & ";"
    If Not modSC.FastStringComp(.sWeapon, "0") Then tVar = tVar & .sWeapon & ";"
    If Not modSC.FastStringComp(.sFace, "0") Then tVar = tVar & .sFace & ";"
    If Not modSC.FastStringComp(.sShield, "0") Then tVar = tVar & .sShield & ";"
    If Not modSC.FastStringComp(.sEars, "0") Then tVar = tVar & .sEars & ";"
    If Not modSC.FastStringComp(.sBack, "0") Then tVar = tVar & .sBack & ";"
    If Not modSC.FastStringComp(.sRings(0), "0") Then tVar = tVar & .sRings(0) & ";"
    If Not modSC.FastStringComp(.sRings(1), "0") Then tVar = tVar & .sRings(1) & ";"
    If Not modSC.FastStringComp(.sRings(2), "0") Then tVar = tVar & .sRings(2) & ";"
    If Not modSC.FastStringComp(.sRings(3), "0") Then tVar = tVar & .sRings(3) & ";"
    If Not modSC.FastStringComp(.sRings(4), "0") Then tVar = tVar & .sRings(4) & ";"
    If Not modSC.FastStringComp(.sRings(5), "0") Then tVar = tVar & .sRings(5) & ";"
    .dBank = 0
    .dGold = 0
    .dMonsterID = 99999
    .iAC = 0
    .iAcc = 0
    .iCasting = 1
    .iCrits = 0
    .iDodge = 0
    .lFamID = 0
    .iHorse = 0
    .iInvitedBy = 0
    .iLeadingParty = 0
    .iMaxDamage = 0
    .iPartyLeader = 0
    .iPlayerAttacking = 0
    .iResting = 0
    .iMeditating = 0
    .iStun = 0
    .lLocation = 1
    .sBlessSpells = "0"
    .sFamName = "0"
    .sFamCustom = "0"
    .lFamAcc = 0
    .lFamCHP = 0
    .lFamID = 0
    .lFamLevel = 0
    .lFamMax = 0
    .lFamMHP = 0
    .lFamMin = 0
    .dFamCEXP = 0
    .dFamEXPN = 0
    .dFamTEXP = 0
    .sInventory = "0"
    .sArms = "0"
    .sBody = "0"
    .sFeet = "0"
    .sHands = "0"
    .sHead = "0"
    .sLegs = "0"
    .sWaist = "0"
    .sBack = "0"
    .sFace = "0"
    .sEars = "0"
    .sShield = "0"
    .sNeck = "0"
    .sWeapon = "0"
    .sSpells = "0"
    .sSpellShorts = "0"
    .sQuest1 = "0"
    .sQuest2 = "0"
    .sQuest3 = "0"
    .sQuest4 = "0"
    .sRings(0) = "0"
    .sRings(1) = "0"
    .sRings(2) = "0"
    .sRings(3) = "0"
    .sRings(4) = "0"
    .sRings(5) = "0"
    .dEXP = 0
    .dEXPNeeded = 0
    .dTotalEXP = 0
    .iIsReadyToTrain = 0
    .lMaxMana = 0
    .dClassPoints = 0
    .iClassBonusLevel = 0
    .lClassChanges = 0
    .lMana = 0
    .lMaxHP = 0
    .lMaxMana = 0
    If modSysopCommands.IsSysop(Index) = True Then
        .sStatsPlus = "0/0/0/0/0/0/1/0/0/0/0/0/0/0/0/0"
    Else
        .sStatsPlus = "0/0/0/0/0/0/0/0/0/0/0/0/0/0/0/0"
    End If
    .sElements = "0/0/0/0/0/0/0/0/0"
    .sTrainStats = "0/0/0/0/0/0/0"
    If modMiscFlag.GetMiscFlag(dbIndex, [Can Be De-Sysed]) = 1 Then
        .sMiscFlag = "0000000000000000000100000000000"
    Else
        .sMiscFlag = "0000000000000000000000000000000"
    End If
    tdGold = .dGold
    .dGold = 0
    lLoc = .lLocation
End With
With dbMap(GetMapIndex(lLoc))
    .sItems = .sItems & tVar
    .dGold = .dGold + tdGold
End With
pLogOn(Index) = True
pPoint(Index) = 7
LogOnSequence Index
End Sub

Sub SendDeathText(MonsterID As Long, ByRef Message As String)
'send the monsters death text if they have any
If aMons(MonsterID).mDeathText <> "0" Then Message = GREEN & aMons(MonsterID).mDeathText & WHITE & vbCrLf
End Sub
