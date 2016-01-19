Attribute VB_Name = "modStats"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modStats
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Function EXP(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "exp") Then
    EXP = True
    Dim dEXP As Double, dNEXP As Double
    Dim dPercentage As Double
    Dim l As Long
    Dim d As Double
    Dim ToSend$
    Dim s As String
    With dbPlayers(GetPlayerIndexNumber(Index))
        dEXP = .dEXP
        dNEXP = .dEXPNeeded
        ToSend$ = LIGHTBLUE & "Level: " & MAGNETA & .iLevel & LIGHTBLUE & " Current EXP: " & MAGNETA & FormatNumber(dEXP, 0, vbUseDefault, vbFalse, vbTrue) & LIGHTBLUE & " (" & MAGNETA & CStr(FormatNumber(.dTotalEXP, 0, vbUseDefault, vbFalse, vbTrue)) & LIGHTBLUE & ") EXP Needed: " & MAGNETA & FormatNumber(dNEXP, 0, vbUseDefault, vbFalse, vbTrue) & LIGHTBLUE
    End With
    dPercentage = dEXP / dNEXP
    dPercentage = FormatNumber(dPercentage, 2)
    d = dPercentage
    If d > 1 Then d = 1
    l = 40 * d
    If l < 0 Then l = 0
    If dPercentage < 1 Then
        dPercentage = dPercentage * 100
        ToSend$ = ToSend$ & " (" & MAGNETA & dPercentage & "%" & LIGHTBLUE & ")"
    Else
        Dim tVal$
        tVal$ = (dPercentage * 100) & "%"
        ToSend$ = ToSend$ & " (" & MAGNETA & tVal$ & LIGHTBLUE & ")"
    End If
    s = YELLOW & "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿" & vbCrLf
    s = s & "³ 0%" & BRIGHTRED & "<" & GREEN & String$(l, "Û") & YELLOW & Space$(40 - l) & BRIGHTRED & ">" & YELLOW & "100% ³" & vbCrLf
    s = s & "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
    ToSend$ = ToSend$ & vbCrLf & s & vbCrLf & WHITE
    WrapAndSend Index, ToSend$
    X(Index) = ""
End If
End Function


Sub SetWeaponStats(Index As Long, Optional dbIndex As Long, Optional OffHand As Boolean = False)
Dim Item As String
Dim iItemID As Long
Dim s As String
Dim i As Long
Dim t As String
Dim Arr() As String
If SpellCombat(Index) = False Then
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    If OffHand Then
        Item = dbPlayers(dbIndex).sShield
    Else
        Item = dbPlayers(dbIndex).sWeapon
    End If
    If Item = "0" Then
        pWeapon(Index).wMin = 1
        pWeapon(Index).wMax = 2 + dbPlayers(dbIndex).iMaxDamage
        pWeapon(Index).wSpeed = 5
        pWeapon(Index).wMessage = "punch:slap:punch"
        pWeapon(Index).wMessage2 = "punches:slaps:punches"
        pWeapon(Index).wMessageV = "punches:slaps:punches"
        pWeapon(Index).wCast = 0
        pWeapon(Index).wSpellName = ""
        pWeapon(Index).wMana = 0
        pWeapon(Index).wElement = -1
        pWeapon(Index).wWeaponName = "fist"
        pWeapon(Index).wBullets = -1
        pWeapon(Index).wMag = -1
        pWeapon(Index).wBMana = -1
        Exit Sub
    End If
    iItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(Item))
    With dbItems(GetItemID(, iItemID))
        If .iType <> 3 Then
            pWeapon(Index).wMin = Val(Mid$(.sDamage, 1, InStr(1, .sDamage, ":")))
            pWeapon(Index).wMax = Val(Mid$(.sDamage, InStr(1, .sDamage, ":") + 1, Len(.sDamage) - (InStr(1, .sDamage, ":") - 1))) + dbPlayers(GetPlayerIndexNumber(Index)).iMaxDamage
            pWeapon(Index).wSpeed = .iSpeed
            pWeapon(Index).wMessage = .sSwings
            pWeapon(Index).wMessage2 = .sMessage2
            pWeapon(Index).wMessageV = .sMessageV
            pWeapon(Index).wCast = 0
            pWeapon(Index).wSpellName = ""
            pWeapon(Index).wMana = 0
            pWeapon(Index).wMAXDAM = 0
            pWeapon(Index).wMINDAM = 0
            pWeapon(Index).wWeaponName = .sItemName
            pWeapon(Index).wSB = 0
            pWeapon(Index).wCastSp = ""
            pWeapon(Index).wBullets = -1
            pWeapon(Index).wMag = -1
            pWeapon(Index).wBMana = -1
        Else
            If modItemManip.GetItemBulletsID(Item) <> 0 Then
                pWeapon(Index).wSpeed = .iSpeed
                With dbItems(GetItemID(, modItemManip.GetItemBulletsID(Item)))
                    pWeapon(Index).wMin = Val(Mid$(.sDamage, 1, InStr(1, .sDamage, ":")))
                    pWeapon(Index).wMax = Val(Mid$(.sDamage, InStr(1, .sDamage, ":") + 1, Len(.sDamage) - (InStr(1, .sDamage, ":") - 1))) + dbPlayers(GetPlayerIndexNumber(Index)).iMaxDamage
                    pWeapon(Index).wMessage = .sSwings
                    pWeapon(Index).wMessage2 = .sMessage2
                    pWeapon(Index).wMessageV = .sMessageV
                    pWeapon(Index).wCast = 0
                    pWeapon(Index).wSpellName = ""
                    pWeapon(Index).wMana = 0
                    pWeapon(Index).wMAXDAM = 0
                    pWeapon(Index).wMINDAM = 0
                    pWeapon(Index).wWeaponName = .sItemName
                    pWeapon(Index).wSB = 0
                    pWeapon(Index).wCastSp = ""
                    pWeapon(Index).wBullets = modItemManip.GetItemBulletsLeft(Item)
                    pWeapon(Index).wMag = modItemManip.GetItemBulletsMagical(Item)
                    pWeapon(Index).wBMana = modItemManip.GetItemBulletsMana(Item)
                End With
            Else
                pWeapon(Index).wBullets = 0
            End If
        End If
        s = modItemManip.GetItemEnchantsFromUnFormattedString(Item)
        If s <> "" Then
            SplitFast s, Arr, "|"
            For i = LBound(Arr) To UBound(Arr)
                If Arr(i) <> "" Then
                    t = Mid$(Arr(i), 4)
                    Select Case Left$(Arr(i), 3)
                        Case "swi"
                            pWeapon(Index).wSB = pWeapon(Index).wSB + Val(t)
                        Case "mab"
                            pWeapon(Index).wMAXDAM = pWeapon(Index).wMAXDAM + Val(t)
                        Case "mib"
                            pWeapon(Index).wMINDAM = pWeapon(Index).wMINDAM + Val(t)
                        Case "csp"
                            If IsNumeric(t) Then
                                pWeapon(Index).wCastSp = pWeapon(Index).wCastSp & GetSpellID(, Val(t)) & ";"
                            Else
                                pWeapon(Index).wCastSp = pWeapon(Index).wCastSp & GetSpellID(t) & ";"
                            End If
                        Case "cs%"
                            pWeapon(Index).wCastSpPer = pWeapon(Index).wCastSpPer & t & ";"
                    End Select
                End If
                If DE Then DoEvents
            Next
        End If
        pWeapon(Index).wMax = pWeapon(Index).wMax + pWeapon(Index).wMAXDAM
        pWeapon(Index).wMin = pWeapon(Index).wMin + pWeapon(Index).wMINDAM
    End With
Else
    Dim SpellID&, Level&
    With dbPlayers(GetPlayerIndexNumber(Index))
        SpellID& = .iCasting
        Level& = .iLevel
    End With
    If SpellID& = 0 Then
        SpellCombat(Index) = False
        Exit Sub
    End If
    With dbSpells(GetSpellID(, SpellID&))
        pWeapon(Index).wCast = .iCast
        pWeapon(Index).wSpellName = .sSpellName
        pWeapon(Index).wMessage = .sMessage
        pWeapon(Index).wMessage2 = .sMessage2
        pWeapon(Index).wMessageV = .sMessageV
        pWeapon(Index).wSpeed = 0
        pWeapon(Index).wMin = .lMinDam
        pWeapon(Index).wMana = .lMana
        pWeapon(Index).wElement = .lElement
        pWeapon(Index).wMAXDAM = 0
        pWeapon(Index).wMINDAM = 0
        pWeapon(Index).wSB = 0
        pWeapon(Index).wCastSp = ""
        pWeapon(Index).wBullets = -1
        pWeapon(Index).wMag = -1
        pWeapon(Index).wBMana = -1
        Dim Damage&
        If Level& > CLng(.iLevelMax) Then Level& = CLng(.iLevelMax)
        Damage& = CLng(.lMaxDam) + (Level& * CLng(.iLevelModify))
        pWeapon(Index).wMax = Damage&
    End With
End If
End Sub

Public Function PlayerStats(Index As Long) As Boolean
Dim tArr() As String
Dim i As Long
Dim dbIndex As Long
Dim ToSend$
If modSC.FastStringComp(LCaseFast(X(Index)), "stat") Then
    PlayerStats = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        ToSend$ = YELLOW & "Name:  " & LIGHTBLUE & .sPlayerName & Space(14 - Len( _
            .sPlayerName)) & YELLOW & "Level: " & LIGHTBLUE & .iLevel
        ToSend$ = ToSend$ & vbCrLf & YELLOW & "Race:  " & LIGHTBLUE & .sRace & _
            Space(14 - Len(.sRace)) & YELLOW & "EXP: " & LIGHTBLUE & FormatNumber(.dEXP, 0, vbUseDefault, vbFalse, vbTrue) & YELLOW & "/" & LIGHTBLUE & FormatNumber(.dTotalEXP, 0, vbUseDefault, vbFalse, vbTrue)
        ToSend$ = ToSend$ & vbCrLf & YELLOW & "Class: " & LIGHTBLUE & .sClass
        ToSend$ = ToSend$ & vbCrLf & YELLOW & "Lives: " & LIGHTBLUE & .iLives
        ToSend$ = ToSend$ & YELLOW & Space(14 - Len(CStr( _
            .iLives))) & "Enc: " & LIGHTBLUE
        ToSend$ = ToSend$ & modGetData.GetPlayersTotalItems( _
            Index, dbIndex) & YELLOW & "/" & LIGHTBLUE & modGetData.GetPlayersMaxItems(Index, dbIndex)
        ToSend$ = ToSend$ & vbCrLf & YELLOW & "Strength:  " & LIGHTBLUE & .iStr _
            & Space(10 - Len(CStr(.iStr))) & YELLOW & "Agility: " & LIGHTBLUE & .iAgil
        ToSend$ = ToSend$ & vbCrLf & YELLOW & "Intelect:  " & LIGHTBLUE & .iInt _
            & YELLOW & Space(10 - Len( _
            CStr(.iInt))) & "Charm:   " & LIGHTBLUE & .iCha & YELLOW
        ToSend$ = ToSend$ & vbCrLf & "Dexterity: " & LIGHTBLUE & .iDex
        ToSend$ = ToSend$ & vbCrLf & YELLOW & "AC: " & LIGHTBLUE & CStr(.iAC) & _
            YELLOW & "/" & LIGHTBLUE & CStr(CLng(.iAC / 14)) & Space$(16 - Len(CStr(.iAC)) - Len( _
            CStr(CLng(.iAC / 14)))) & YELLOW
        
        ToSend$ = ToSend$ & "HP: " & LIGHTBLUE & .lHP & YELLOW & "/" & _
            LIGHTBLUE & .lMaxHP & vbCrLf & YELLOW
        ToSend$ = ToSend$ & "SC: " & LIGHTBLUE & modMiscFlag.GetStatsPlusTotal(dbIndex, [Spell Casting]) & Space$(21 - Len("SC: " & CStr(((.iLevel + .iInt + .iAgil + .iDex) \ 3)))) & YELLOW
        If .lMaxMana <> 0 Then
            ToSend$ = ToSend$ & "MA: " & LIGHTBLUE & .lMana & YELLOW & "/" & _
                LIGHTBLUE & .lMaxMana & vbCrLf & WHITE
        Else
            ToSend$ = ToSend$ & vbCrLf & WHITE
        End If
        If Not modSC.FastStringComp(.sBlessSpells, "0") Then
            SplitFast Left$(.sBlessSpells, Len(.sBlessSpells) - 1), tArr, "Œ"
            For i = LBound(tArr) To UBound(tArr)
                ToSend$ = ToSend$ & LIGHTBLUE & dbSpells(CLng(Mid$(tArr(i), _
                    InStrRev(tArr(i), "~") + 1, (Len(tArr(i)) - InStrRev(tArr(i), _
                    "~"))))).sStatMessage & WHITE & vbCrLf
                If DE Then DoEvents
            Next
        End If
        Erase tArr
        SplitFast ToSend$, tArr, vbCrLf
        Dim lLen As Long
        For i = LBound(tArr) To UBound(tArr)
            If lLen < Len(tArr(i)) Then lLen = Len(tArr(i))
            If DE Then DoEvents
        Next
        For i = LBound(tArr) To UBound(tArr)
            If tArr(i) <> "æ" And tArr(i) <> "" Then tArr(i) = YELLOW & "º " & tArr(i) & Space$(lLen - (Len(tArr(i)) - modGetData.GetANSIColorChanges(tArr(i)))) & YELLOW & " º"
            If DE Then DoEvents
        Next
        ToSend$ = Join(tArr, vbCrLf)
        ToSend$ = LIGHTBLUE & "Your statistics:" & vbCrLf & YELLOW & "ÉÍ" & String(lLen, "Í") & "Í»" & vbCrLf & ToSend$
        ToSend$ = ToSend$ & YELLOW & "ÈÍ" & String(lLen, "Í") & "Í¼" & vbCrLf
        WrapAndSend Index, ToSend$
        X(Index) = ""
    End With
End If
End Function

Public Function StatsExtended(Index As Long) As Boolean
Dim s As String
Dim dbIndex As Long
If LCaseFast(X(Index)) = "stats+" Then
    StatsExtended = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        s = GREEN & "Magic Resistence: " & LIGHTBLUE & modMiscFlag.GetStatsPlusTotal(dbIndex, [Magic Resistance]) & vbCrLf
        s = s & GREEN & "Dodge:            " & LIGHTBLUE & modGetData.GetPlayerDodge(dbIndex) & vbCrLf
        s = s & GREEN & "Crits:            " & LIGHTBLUE & dbPlayers(dbIndex).iCrits & vbCrLf
        s = s & GREEN & "Accuracy:         " & LIGHTBLUE & dbPlayers(dbIndex).iAcc & vbCrLf
        s = s & GREEN & "Spell Casting:    " & LIGHTBLUE & modMiscFlag.GetStatsPlusTotal(dbIndex, [Spell Casting]) & vbCrLf
        s = s & GREEN & "Stealth:          " & LIGHTBLUE & CStr(modMiscFlag.GetStatsPlusTotal(dbIndex, Steath)) & vbCrLf
        s = s & GREEN & "Hunger:           " & LIGHTBLUE & CStr(.dHunger) & vbCrLf
        s = s & GREEN & "Stamina:          " & LIGHTBLUE & CStr(.dStamina) & vbCrLf
        s = s & GREEN & "Perception:       " & LIGHTBLUE & modMiscFlag.GetStatsPlusTotal(dbIndex, Perception) & vbCrLf
        s = s & GREEN & "Animal Relations: " & LIGHTBLUE & modMiscFlag.GetStatsPlusTotal(dbIndex, [Animal Relations]) & vbCrLf
        s = s & GREEN & "Reputation:       " & GetRep(Index) & vbCrLf
        s = s & GREEN & "Wealth:           " & LIGHTBLUE & CStr((.dBank + .dGold)) & GREEN & " gold" & vbCrLf & WHITE
        WrapAndSend Index, s
        X(Index) = ""
    End With
End If
End Function

Public Function Stamina(Index As Long, Optional dbIndex As Long) As Boolean
Dim s As String
Dim d As Double
Dim l As Long
If LCaseFast(X(Index)) = "stamina" Then
    Stamina = True
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    d = dbPlayers(dbIndex).dStamina
    If d > 100 Then d = 100
    l = 40 * (d / 100)
    If l < 0 Then l = 0
    s = YELLOW & "         ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿" & vbCrLf
    s = s & "         ³    " & LIGHTBLUE & "Stamina Level" & YELLOW & "    ³" & vbCrLf
    s = s & "         ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ" & vbCrLf
    s = s & GREEN & "0%                  50%                100%" & YELLOW & vbCrLf
    s = s & "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿" & vbCrLf
    s = s & "³" & GetColor(l) & String$(l, "Û") & YELLOW & Space$(40 - l) & "³" & vbCrLf
    s = s & "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ" & vbCrLf
    WrapAndSend Index, s
    X(Index) = ""
End If
End Function

Public Function Hunger(Index As Long, Optional dbIndex As Long) As Boolean
Dim s As String
Dim d As Double
Dim l As Long
If LCaseFast(X(Index)) = "hunger" Then
    Hunger = True
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    d = dbPlayers(dbIndex).dHunger
    If d > 100 Then d = 100
    l = 40 * (d / 100)
    If l < 0 Then l = 0
    s = YELLOW & "         ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿" & vbCrLf
    s = s & "         ³    " & LIGHTBLUE & "Hunger  Level" & YELLOW & "    ³" & vbCrLf
    s = s & "         ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ" & vbCrLf
    s = s & GREEN & "0%                  50%                100%" & YELLOW & vbCrLf
    s = s & "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿" & vbCrLf
    s = s & "³" & GetColor(l) & String$(l, "Û") & YELLOW & Space$(40 - l) & "³" & vbCrLf
    s = s & "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ" & vbCrLf
    WrapAndSend Index, s
    X(Index) = ""
End If
End Function

Private Function GetColor(lVal As Long) As String
Select Case lVal
    Case Is < 5
        GetColor = BRIGHTRED
    Case Is < 10
        GetColor = RED
    Case Is < 15
        GetColor = BRIGHTMAGNETA
    Case Is < 20
        GetColor = BRIGHTYELLOW
    Case Is < 25
        GetColor = YELLOW
    Case Is < 30
        GetColor = LIGHTBLUE
    Case Is < 35
        GetColor = BLUE
    Case Is < 41
        GetColor = GREEN
    Case Else
        GetColor = GREEN
End Select
End Function
