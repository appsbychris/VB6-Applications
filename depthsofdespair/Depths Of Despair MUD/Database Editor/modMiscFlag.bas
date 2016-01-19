Attribute VB_Name = "modMiscFlag"
'0   Can Attack  0 or 1
'1   Can Cast Spell  0 or 1
'2   Can Sneak   0 or 1
'3   Gibberish Talk  0 or 1
'4   Guild Rank  0 to 5 (Leader, General, Lieutenant, Soldier, Normal, Scrub
'5   Invisible 0 Or 1
'6   Can Eq Head 0 or 1
'7   Can Eq Face 0 or 1
'8   Can Eq Ears 0 or 1
'9   Can Eq Neck 0 or 1
'10  Can Eq Body 0 or 1
'11  Can Eq Back 0 or 1
'12  Can Eq Arms 0 or 1
'13  Can Eq Shield   0 or 1
'14  Can Eq Hands    0 or 1
'15  Can Eq Legs 0 or 1
'16  Can Eq Feet 0 or 1
'17  Can Eq Waist    0 or 1
'18  Can Eq Weapon   0 or 1
'19  Can Be De-Sysed 0 or 1
Public Enum MiscFlag
    [Can Attack] = 0
    [Can Cast Spell] = 1
    [Can Sneak] = 2
    [Gibberish Talk] = 3
    [Guild Rank] = 4
    [Invisible] = 5
    [Can Eq Head] = 6
    [Can Eq Face] = 7
    [Can Eq Ears] = 8
    [Can Eq Neck] = 9
    [Can Eq Body] = 10
    [Can Eq Back] = 11
    [Can Eq Arms] = 12
    [Can Eq Shield] = 13
    [Can Eq Hands] = 14
    [Can Eq Legs] = 15
    [Can Eq Feet] = 16
    [Can Eq Waist] = 17
    [Can Eq Weapon] = 18
    [Can Be De-Sysed] = 19
    [See Invisible] = 20
    [See Hidden] = 21
    [Can Eq Ring 0] = 22
    [Can Eq Ring 1] = 23
    [Can Eq Ring 2] = 24
    [Can Eq Ring 3] = 25
    [Can Eq Ring 4] = 26
    [Can Eq Ring 5] = 27
End Enum

Public Enum StatsPlusList
    [Spell Casting Base] = 0
    [Spell Casting Bonus] = 1
    [Magic Resistance Base] = 2
    [Magic Resistance Bonus] = 3
    [Perception Base] = 4
    [Perception Bonus] = 5
    [Is A Sysop] = 6
    [Pallete Number] = 7
    [Max Items Base] = 8
    [Max Items Bonus] = 9
    [Stealth Base] = 10
    [Stealth Bonus] = 11
    [Animal Relations Base] = 12
    [Animal Relations Bonus] = 13
    [Thieving Base] = 14
    [Thieving Bonus] = 15
End Enum

Public Enum StatsPlusTotal
    [Spell Casting] = 0
    [Magic Resistance] = 1
    [Perception] = 2
    [Max Items] = 3
    [Steath] = 4
    [Animal Relations] = 5
    [Thieving] = 6
End Enum

Public Function GetMiscFlag(dbIndex As Long, WhichFlag As MiscFlag) As Long
Dim s As String
If dbIndex = 0 Then Exit Function
s = Mid$(dbPlayers(dbIndex).sMiscFlag, WhichFlag + 1, 1)
GetMiscFlag = CLng(Val(s))
End Function

Public Sub SetMiscFlag(dbIndex As Long, WhichFlag As MiscFlag, SetAs As Long)
Dim i As Long
Dim l As Long
Dim s As Long
Mid$(dbPlayers(dbIndex).sMiscFlag, WhichFlag + 1, 1) = CStr(SetAs)
For i = 6 To 18
    l = GetMiscFlag(dbIndex, i)
    With dbPlayers(dbIndex)
        Select Case i
            Case 6
                If .sHead <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sHead)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
'                    WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 7
                If .sFace <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sFace)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                   ' WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 8
                If .sEars <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sEars)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                   ' WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 9
                If .sNeck <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sNeck)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                   ' WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 10
                If .sBody <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sBody)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                   ' WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 11
                If .sBack <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sBack)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                    'WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 12
                If .sArms <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sArms)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                   ' WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 13
                If .sShield <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sShield)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                    'WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 14
                If .sHands <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sHands)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                   ' WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 15
                If .sLegs <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sLegs)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                  '  WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 16
                If .sFeet <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sFeet)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                  '  WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 17
                If .sWaist <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sWaist)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                  '  WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 18
                If .sWeapon <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sWeapon)
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                '    WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
            Case 22 To 27
                If .sRings(i - 22) <> "0" And l = 1 Then
                    s = modItemManip.GetItemIDFromUnFormattedString(.sRings(i - 22))
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, s
                    modItemManip.TakeItemFromInvAndPutOnGround dbIndex, s
                 '   WrapAndSend dbPlayers(dbIndex).iIndex, LIGHTBLUE & "Your " & clsGetData.GetItemNameWithS(GetItemID(, s)) & "to the ground." & WHITE & vbCrLf, False
                End If
        End Select
    End With
    If DE Then DoEvents
Next
End Sub

Public Function GetStatsPlus(dbIndex As Long, WhichOne As StatsPlusList) As Long
Dim Arr() As String
With dbPlayers(dbIndex)
    SplitFast .sStatsPlus, Arr, "/"
    GetStatsPlus = CLng(Val(Arr(WhichOne)))
End With
End Function

Public Function GetStatsPlusTotal(dbIndex As Long, WhichOne As StatsPlusTotal) As Long
Dim Arr() As String
With dbPlayers(dbIndex)
    SplitFast .sStatsPlus, Arr, "/"
    Select Case WhichOne
        Case 0
            GetStatsPlusTotal = GetStatsPlus(dbIndex, [Spell Casting Base]) + GetStatsPlus(dbIndex, [Spell Casting Bonus])
        Case 1
            GetStatsPlusTotal = GetStatsPlus(dbIndex, [Magic Resistance Base]) + GetStatsPlus(dbIndex, [Magic Resistance Bonus])
        Case 2
            GetStatsPlusTotal = GetStatsPlus(dbIndex, [Perception Base]) + GetStatsPlus(dbIndex, [Perception Bonus])
        Case 3
            GetStatsPlusTotal = GetStatsPlus(dbIndex, [Max Items Base]) + GetStatsPlus(dbIndex, [Max Items Bonus])
        Case 4
            GetStatsPlusTotal = GetStatsPlus(dbIndex, [Stealth Base]) + GetStatsPlus(dbIndex, [Stealth Bonus])
        Case 5
            GetStatsPlusTotal = GetStatsPlus(dbIndex, [Animal Relations Base]) + GetStatsPlus(dbIndex, [Animal Relations Bonus])
        Case 6
            GetStatsPlusTotal = GetStatsPlus(dbIndex, [Thieving Base]) + GetStatsPlus(dbIndex, [Thieving Bonus])
    End Select
End With
End Function

Public Sub SetStatsPlus(dbIndex As Long, WhichOne As StatsPlusList, Value As Long)
Dim Arr() As String
Dim i As Long
With dbPlayers(dbIndex)
    SplitFast .sStatsPlus, Arr, "/"
    Arr(WhichOne) = CStr(Value)
    .sStatsPlus = ""
    For i = LBound(Arr) To UBound(Arr)
        .sStatsPlus = .sStatsPlus & Arr(i) & "/"
        If DE Then DoEvents
    Next
    .sStatsPlus = Left$(.sStatsPlus, Len(.sStatsPlus) - 1)
End With
End Sub

Public Sub RedoStatsPlus(dbIndex As Long)
'Public Enum StatsPlusList
'    [Spell Casting Base] = 0
'    [Spell Casting Bonus] = 1
'    [Magic Resistance Base] = 2
'    [Magic Resistance Bonus] = 3
'    [Perception Base] = 4
'    [Perception Bonus] = 5
'    [Is A Sysop] = 6
'    [Pallete Number] = 7
'    [Max Items Base] = 8
'    [Max Items Bonus] = 9
'    [Stealth Base] = 10
'    [Stealth Bonus] = 11
'    [Animal Relations Base] = 12
'    [Animal Relations Bonus] = 13
'End Enum
Dim Arr() As String
Dim i As Long
With dbPlayers(dbIndex)
    SplitFast .sStatsPlus, Arr, "/"
    Arr(0) = CStr(CLng((.iLevel + .iInt + .iAgil + .iDex) \ 3))
    Arr(2) = CStr(clsGetData.GetpMR(dbIndex))
    Arr(4) = CStr(clsGetData.GetPlayersPerceptionBase(dbIndex, False))
    Arr(8) = CStr(CLng(((.iStr \ 2) + (.iDex \ 3) + (.iCha \ 4) + (.iInt \ 6) + (.iAgil \ 5))))
    Arr(10) = CStr(CLng((.iAgil + (.iDex) + (.iCha / 1.2685) + (.iLevel / 1.012545))))
    Arr(12) = CStr(CLng((.iCha + (.dClassPoints / 2.245845) + (.iDex / 4.21544) + (.iAgil / 5.84545) + (.iInt / 2.124584))))
    If CLng(Arr(8)) > 20 Then Arr(8) = "20"
    If CLng(Arr(8)) < 4 Then Arr(8) = "4"
    Arr(8) = CStr(CLng(CLng(Arr(8)) + (.iStr \ 10)))
    Arr(14) = CStr(Val(.iInt) + Val(.iAgil) + Val(.iCha * 2))
    Arr(14) = CStr(Val(Arr(14)) * 0.14873)
    Arr(14) = CStr(Val(Arr(14) + (.iDex * 4.5)))
    Arr(14) = CLng(Val(Arr(14)))
    If Val(Arr(14)) > 93 Then Arr(14) = "94"
    .sStatsPlus = ""
    For i = LBound(Arr) To UBound(Arr)
        .sStatsPlus = .sStatsPlus & Arr(i) & "/"
        If DE Then DoEvents
    Next
    .sStatsPlus = Left$(.sStatsPlus, Len(.sStatsPlus) - 1)
End With
End Sub

