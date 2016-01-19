Attribute VB_Name = "modTrain"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modTrain
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Type ClassRank
    dBegin As Double
    dInter As Double
    dMaste As Double
    dGuru  As Double
    sBBonus As String
    sIBonus As String
    sMBonus As String
    sGBonus As String
End Type

Public Function Train(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(X(Index)), "train") Then
    Train = True
    Dim iEXP As Double, iNEXP As Double
    Dim sLoc$, ToSend$
    Dim dbIndex As Long
    Dim cRank As ClassRank
    Dim dbClassID As Long
    Dim l As Long
    Dim Arr() As String
    Dim dbRaceID As Long
    Dim m As Long
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        iEXP = .dEXP
        iNEXP = .dEXPNeeded
        sLoc$ = .lLocation
        dbClassID = GetClassID(.sClass)
        dbRaceID = GetRaceID(.sRace)
    End With
    With dbMap(GetMapIndex(CLng(sLoc$)))
        If .iType = 2 Then
            If iEXP >= iNEXP Then
                ToSend$ = MAGNETA & "The trainer chants some strange words, and you feel energy surround you." & vbCrLf
                
            Else
                WrapAndSend Index, RED & "If you want to get to the next level, you must first gain expierence." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        Else
            WrapAndSend Index, RED & "You know you can only train at a trainer, right?" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    With dbPlayers(dbIndex)
        With dbClass(GetClassID(.sClass))
            cRank.dBegin = .dBeginnerMax
            cRank.dInter = .dIntermediateMax
            cRank.dMaste = .dMasterMax
            cRank.dGuru = .dGuru
            cRank.sBBonus = .sBBonus
            cRank.sGBonus = .sGBonus
            cRank.sIBonus = .sIBonus
            cRank.sMBonus = .sMBonus
        End With

        If .iClassBonusLevel = 0 And .dClassPoints > cRank.dBegin Then
            .iClassBonusLevel = 1
            modTrain.AddBonusStats dbIndex, cRank.sBBonus
        ElseIf .dClassPoints > cRank.dInter Then
            If .iClassBonusLevel = 1 Then
                .iClassBonusLevel = 2
                modTrain.AddBonusStats dbIndex, cRank.sIBonus
            End If
        ElseIf .dClassPoints > cRank.dMaste Then
            If .iClassBonusLevel = 2 Then
                .iClassBonusLevel = 3
                modTrain.AddBonusStats dbIndex, cRank.sMBonus
            End If
        ElseIf .dClassPoints > cRank.dGuru Then
            If .iClassBonusLevel = 3 Then
                .iClassBonusLevel = 4
                modTrain.AddBonusStats dbIndex, cRank.sGBonus
            End If
        End If
        .dEXP = .dEXP - .dEXPNeeded
        .dEXPNeeded = .dEXPNeeded + .iLevel
        .iLevel = .iLevel + 1
        l = RndNumber(1, CDbl(.iDex + (.iInt \ 3) + (.iCha \ 4) + (.iAgil \ 5)))
        SplitFast dbRaces(dbRaceID).sHP, Arr, ":"
        If l > (Val(Arr(1)) - Val(Arr(0))) Then l = (Val(Arr(1)) - Val(Arr(0)))
        m = RndNumber(1, CDbl(.iInt + (.iCha \ 3) + (.iDex \ 4) + (.iAgil \ 5) + (.iStr \ 6)))
        If m > (dbClass(dbClassID).iMaxMana - dbClass(dbClassID).iMinMana) Then m = (dbClass(dbClassID).iMaxMana - dbClass(dbClassID).iMinMana)
        IncreaseTrain dbIndex, 5, l
        IncreaseTrain dbIndex, 6, m
        .lMaxHP = .lMaxHP + l + dbClass(dbClassID).lHPBonus
        .lMaxMana = .lMaxMana + m + dbClass(dbClassID).lMABonus
        ToSend$ = ToSend$ & LIGHTBLUE & "Your level has increased to " & .iLevel & "." & vbCrLf
        ToSend$ = ToSend$ & LIGHTBLUE & "Your HP has increased by " & l & "." & vbCrLf
        ToSend$ = ToSend$ & LIGHTBLUE & "Your Mana has increased by " & m & "." & vbCrLf
        .iIsReadyToTrain = .iIsReadyToTrain + 1 + dbClass(dbClassID).lCPBonus
        WrapAndSend Index, ToSend$
        X(Index) = ""
    End With
End If
End Function

Sub AddBonusStats(dbIndex As Long, sBonuses As String)
Dim tArr() As String
SplitFast sBonuses, tArr, ":"
For i = LBound(tArr) To UBound(tArr)
    With dbPlayers(dbIndex)
        If tArr(i) <> "" Then
            Select Case Left$(tArr(i), 3)
                Case "mhp" 'Max Hitpoints
                    .lMaxHP = .lMaxHP + CLng(ReplaceFast(tArr(i), "mhp", ""))
                Case "str" 'Strength
                    .iStr = .iStr + CLng(ReplaceFast(tArr(i), "str", ""))
                Case "agi" 'Agility
                    .iAgil = .iAgil + CLng(ReplaceFast(tArr(i), "agi", ""))
                Case "int" 'Intellect
                    .iInt = .iInt + CLng(ReplaceFast(tArr(i), "int", ""))
                Case "cha" 'Charm
                    .iCha = .iCha + CLng(ReplaceFast(tArr(i), "cha", ""))
                Case "dex" 'Dexterity
                    .iDex = .iDex + CLng(ReplaceFast(tArr(i), "dex", ""))
                Case "pac" 'Armor Class
                    .iAC = .iAC + CLng(ReplaceFast(tArr(i), "pac", ""))
                Case "acc" 'Accurracy
                    .iAcc = .iAcc + CLng(ReplaceFast(tArr(i), "acc", ""))
                Case "cri" 'Crits
                    .iCrits = .iCrits + CLng(ReplaceFast(tArr(i), "cri", ""))
                Case "mma" 'Max Mana
                    .lMaxMana = .lMaxMana + CLng(ReplaceFast(tArr(i), "mma", ""))
                Case "dam" 'damage bonus
                    .iMaxDamage = .iMaxDamage + CLng(ReplaceFast(tArr(i), "dam", ""))
                Case "dod" 'dodge
                    .iDodge = .iDodge + CLng(ReplaceFast(tArr(i), "dod", ""))
                Case "sne"
                    modMiscFlag.SetMiscFlag dbIndex, [Can Sneak], CLng(ReplaceFast(tArr(i), "sne", ""))
            End Select
        End If
    End With
    If DE Then DoEvents
Next
End Sub

Public Function TrainStats(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(X(Index)), "train stats") Then
    TrainStats = True
    pPoint(Index) = 15
    X(Index) = ""
    TrainMode Index
End If
End Function

Sub TrainMode(Index As Long, Optional WithError As Boolean = False, Optional ErrMsg As String = "", Optional Position As Long = 0)
Dim dbIndex As Long
Dim s As String
If pPoint(Index) = 15 Then
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        '.iTrainSlot = 0
        s = ANSICLS & BLACK & "ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ" & vbCrLf
        s = s & BLACK & "Û" & WHITE & "ÉÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ»" & BLACK & "Û" & vbCrLf
        s = s & BLACK & "Û" & WHITE & "º          Train your character's stats            º" & BLACK & "Û" & vbCrLf
        s = s & BLACK & "Û" & WHITE & "ÌÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹" & BLACK & "Û" & vbCrLf
        s = s & BLACK & "Û" & WHITE & "º  Name : "
        s = s & BRIGHTWHITE & .sPlayerName & WHITE
        s = s & String$(41 - Len(.sPlayerName), " ") & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º  Level: "
        s = s & BRIGHTWHITE & CStr(.iLevel) & WHITE
        s = s & String$(41 - Len(CStr(.iLevel)), " ") & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º  Race : "
        s = s & BRIGHTWHITE & .sRace & WHITE
        s = s & String$(41 - Len(.sRace), " ") & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º  Class: "
        s = s & BRIGHTWHITE & .sClass & WHITE
        s = s & String$(41 - Len(.sClass), " ") & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º  EXP  : "
        s = s & BRIGHTWHITE & CStr(.dEXP) & WHITE
        s = s & String$(41 - Len(CStr(.dEXP)), " ") & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "ÌÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º          Change your characters stats            º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "ÌÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º Strength : "
        s = s & BRIGHTWHITE & CStr(.iStr) & WHITE
        s = s & String$(38 - Len(CStr(.iStr)), " ") & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º Intelect : "
        s = s & BRIGHTWHITE & CStr(.iInt) & WHITE
        s = s & String$(38 - Len(CStr(.iInt)), " ") & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º Dexterity: "
        s = s & BRIGHTWHITE & CStr(.iDex) & WHITE
        s = s & String$(38 - Len(CStr(.iDex)), " ") & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º Agility  : "
        s = s & BRIGHTWHITE & CStr(.iAgil) & WHITE
        s = s & String$(22 - Len(CStr(.iAgil)), " ") & "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¶" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "º Charm    : "
        s = s & BRIGHTWHITE & CStr(.iCha) & WHITE
        s = s & String$(22 - Len(CStr(.iCha)), " ") & "³Points: "
        s = s & BRIGHTWHITE & CStr(.iIsReadyToTrain) & WHITE & String$(7 - Len(CStr(.iIsReadyToTrain)), " ")
        s = s & "º" & BLACK & "Û" & vbCrLf
        s = s & "Û" & WHITE & "ÈÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍËÍÍÍÍÍÍËÍÍ¼" & BLACK & "Û" & vbCrLf
        If Not WithError Then
            s = s & "ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ" & WHITE & "º Done º" & BLACK & "ÛÛÛÛ" & vbCrLf
        Else
            s = s & "Û" & BRIGHTRED & ErrMsg & BLACK & String$(41 - Len(ErrMsg), "Û") & WHITE & "º Done º" & BLACK & "ÛÛÛÛ" & vbCrLf
        End If
        s = s & "ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ" & WHITE & "ÈÍÍÍÍÍÍ¼" & BLACK & "ÛÛÛÛ" & vbCrLf
        s = s & "ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ" & WHITE
        s = s & SetMoveCursor(13 + Position, 15)
        WrapAndSend Index, s
    End With
End If
End Sub

Sub MOVECURSOR(Index As Long)
With dbPlayers(GetPlayerIndexNumber(Index))
    Select Case X(Index)
        Case UP_ARROW
            If .iTrainSlot > 0 Then .iTrainSlot = .iTrainSlot - 1: SendCursorMovement Index, .iTrainSlot
            
        Case DOWN_ARROW
            If .iTrainSlot < 5 Then .iTrainSlot = .iTrainSlot + 1: SendCursorMovement Index, .iTrainSlot
            
        Case RIGHT_ARROW, LEFT_ARROW
            
        End Select
End With

End Sub

Sub SendCursorMovement(Index As Long, TrainSlot As Long)
Dim s As String
Select Case TrainSlot
    Case 0
        s = SetMoveCursor(13, 15)
    Case 1
        s = SetMoveCursor(14, 15)
    Case 2
        s = SetMoveCursor(15, 15)
    Case 3
        s = SetMoveCursor(16, 15)
    Case 4
        s = SetMoveCursor(17, 15)
    Case 5
        s = SetMoveCursor(19, 45)
End Select
WrapAndSend Index, s
X(Index) = ""
End Sub

Sub ValidateAndAdjust(Index As Long)
Dim dbIndex As Long
Dim i As Double
Dim s As String

dbIndex = GetPlayerIndexNumber(Index)
If X(Index) <> "" Then
    If IsNumeric(X(Index)) Then
        i = CDbl(X(Index))
    Else
    
    End If
Else
    With dbPlayers(dbIndex)
        If .iTrainSlot <> 5 Then
            modTrain.SendCursorMovement Index, .iTrainSlot + 1
            .iTrainSlot = .iTrainSlot + 1
            X(Index) = ""
            Exit Sub
        End If
    End With
End If
With dbPlayers(dbIndex)
    If .iTrainSlot <> 5 And Not IsNumeric(X(Index)) Then
        TrainMode Index, True, "Cannot set, not a number", dbPlayers(dbIndex).iTrainSlot
        X(Index) = ""
        Exit Sub
    End If

    Select Case .iTrainSlot
        Case 0
            If i >= .iStr Then
                If i - .iStr <= .iIsReadyToTrain Then
                    .iIsReadyToTrain = .iIsReadyToTrain - (i - .iStr)
                    .iStr = i
                    IncreaseTrain dbIndex, 0, i - .iStr
                    i = .iTrainSlot
                    TrainMode Index, , , CLng(i)
                Else
                    TrainMode Index, True, "You do not have enough points", .iTrainSlot
                    X(Index) = ""
                    Exit Sub
                End If
            Else
                GoTo LowValue
            End If
            s = SetMoveCursor(14, 15)
            
        Case 1
            If i >= .iInt Then
                If i - .iInt <= .iIsReadyToTrain Then
                    .iIsReadyToTrain = .iIsReadyToTrain - (i - .iInt)
                    .iInt = i
                    IncreaseTrain dbIndex, 1, i - .iInt
                    i = .iTrainSlot
                    TrainMode Index, , , CLng(i)
                Else
                    TrainMode Index, True, "You do not have enough points", .iTrainSlot
                    X(Index) = ""
                    Exit Sub
                End If
            Else
                GoTo LowValue
            End If
            s = SetMoveCursor(15, 15)
        Case 2
            If i >= .iDex Then
                If i - .iDex <= .iIsReadyToTrain Then
                    .iIsReadyToTrain = .iIsReadyToTrain - (i - .iDex)
                    .iDex = i
                    IncreaseTrain dbIndex, 2, i - .iDex
                    i = .iTrainSlot
                    TrainMode Index, , , CLng(i)
                Else
                    TrainMode Index, True, "You do not have enough points", .iTrainSlot
                    X(Index) = ""
                    Exit Sub
                End If
            Else
                GoTo LowValue
            End If
            s = SetMoveCursor(16, 15)
        Case 3
            If i >= .iAgil Then
                If i - .iAgil <= .iIsReadyToTrain Then
                    .iIsReadyToTrain = .iIsReadyToTrain - (i - .iAgil)
                    .iAgil = i
                    IncreaseTrain dbIndex, 3, i - .iAgil
                    i = .iTrainSlot
                    TrainMode Index, , , CLng(i)
                Else
                    TrainMode Index, True, "You do not have enough points", .iTrainSlot
                    X(Index) = ""
                    Exit Sub
                End If
            Else
                GoTo LowValue
            End If
            s = SetMoveCursor(17, 15)
        Case 4
            If i >= .iCha Then
                If i - .iCha <= .iIsReadyToTrain Then
                    .iIsReadyToTrain = .iIsReadyToTrain - (i - .iCha)
                    .iCha = i
                    IncreaseTrain dbIndex, 4, i - .iCha
                    i = .iTrainSlot
                    TrainMode Index, , , CLng(i)
                Else
                    TrainMode Index, True, "You do not have enough points", .iTrainSlot
                    X(Index) = ""
                    Exit Sub
                End If
            Else
                GoTo LowValue
            End If
            s = SetMoveCursor(19, 44)
        Case 5
            pPoint(Index) = 0
            .iTrainSlot = 0
            WrapAndSend Index, ANSICLS
'            .iSC = (.iLevel + .iInt + .iAgil + .iDex) \ 3
'            If .iSC >= 100 Then .iSC = 99
            modMiscFlag.RedoStatsPlus dbIndex
            modMiscFlag.SetStatsPlus dbIndex, [Max Items Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Max Items Bonus]) + dbClass(GetClassID(.sClass)).lMaxItemsBonus
            WrapAndSend Index, modgetdata.GetRoomDescription(dbIndex, CLng(.lLocation))
            Exit Sub
    End Select
    .iTrainSlot = .iTrainSlot + 1
End With
If s <> "" Then
    WrapAndSend Index, s
    X(Index) = ""
Else
LowValue:
    i = dbPlayers(dbIndex).iTrainSlot
    TrainMode Index, True, "Cannot set value: Too Low", CLng(i)
    X(Index) = ""
End If
End Sub

Public Sub IncreaseTrain(dbIndex As Long, StatIndex As Long, Value As Long)
Dim Arr() As String
Dim i As Long
SplitFast dbPlayers(dbIndex).sTrainStats, Arr, "/"
Arr(StatIndex) = Val(Arr(StatIndex)) + Value
dbPlayers(dbIndex).sTrainStats = ""
For i = LBound(Arr) To UBound(Arr)
    With dbPlayers(dbIndex)
        .sTrainStats = .sTrainStats & Arr(i) & "/"
    End With
    If DE Then DoEvents
Next
With dbPlayers(dbIndex)
    .sTrainStats = Left$(.sTrainStats, Len(.sTrainStats) - 1)
End With
End Sub
