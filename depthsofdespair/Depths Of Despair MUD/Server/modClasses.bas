Attribute VB_Name = "modClasses"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modClasses
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function TrainClass(Index As Long) As Boolean
Dim dbIndex As Long
Dim dbLoc As Long
Dim dbClassIndex As Long
Dim dbCurClass As Long

Dim cRank As ClassRank
Dim i As Long
Dim j As Long
Dim s As String
Dim s2 As String
Dim tArr() As String

Dim bOn As Boolean
Dim dVal As Double
Dim dVal2 As Double
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 7)), "train c") Then
    TrainClass = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        dbLoc = .lDBLocation
    End With
    With dbMap(dbLoc)
        If .iType <> 6 Then
            WrapAndSend Index, RED & "You cannot train in a profession here." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        dbClassIndex = GetClassID(, .iTrainClass)
        If dbClassIndex = 0 Then
            WrapAndSend Index, RED & "You cannot train in any profession here." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    With dbPlayers(dbIndex)
        If .lClassChanges <> 0 Then
            dbCurClass = GetClassID(.sClass)
            With dbClass(dbCurClass)
                cRank.dBegin = .dBeginnerMax
                cRank.dInter = .dIntermediateMax
                cRank.dMaste = .dMasterMax
                cRank.dGuru = .dGuru
                cRank.sBBonus = .sBBonus
                cRank.sGBonus = .sGBonus
                cRank.sIBonus = .sIBonus
                cRank.sMBonus = .sMBonus
            End With
            If .dClassPoints > cRank.dBegin Then AddBonusStatsInverse dbIndex, cRank.sBBonus
            If .dClassPoints > cRank.dInter Then AddBonusStatsInverse dbIndex, cRank.sIBonus
            If .dClassPoints > cRank.dMaste Then AddBonusStatsInverse dbIndex, cRank.sMBonus
            If .dClassPoints > cRank.dGuru Then AddBonusStatsInverse dbIndex, cRank.sGBonus
            .dClassPoints = 0
            SplitFast ReplaceFast(.sSpells, ":", ""), tArr, ";"
            For i = LBound(tArr) To UBound(tArr)
                If tArr(i) <> "" And tArr(i) <> "0" Then
                    With dbSpells(GetSpellID(, CLng(tArr(i))))
                        If .iType = dbClass(dbCurClass).iSpellType Then
                            s = s & CStr(.lID) & ";"
                            s2 = s2 & .sShort & ";"
                        End If
                    End With
                End If
                If DE Then DoEvents
            Next
            If s <> "" Then
                Erase tArr
                SplitFast s, tArr, ";"
                s = ""
                For i = LBound(tArr) To UBound(tArr)
                    If Not modSC.FastStringComp(tArr(i), "") Then j = j + 1
                    If DE Then DoEvents
                Next
                If j > 1 Then
                    If RndNumber(0, 1) = 0 Then
                        bOn = False
                        For i = LBound(tArr) To UBound(tArr)
                            If Not modSC.FastStringComp(tArr(i), "") Then
                                If bOn Then
                                    s = s & tArr(i) & ";"
                                    bOn = False
                                    GoTo nNexT1
                                End If
                                If Not bOn Then bOn = True
nNexT1:
                            End If
                            If DE Then DoEvents
                        Next
                    Else
                        bOn = True
                        For i = LBound(tArr) To UBound(tArr)
                            If Not modSC.FastStringComp(tArr(i), "") Then
                                If bOn Then
                                    s = s & tArr(i) & ";"
                                    bOn = False
                                    GoTo nNexT2
                                End If
                                If Not bOn Then bOn = True
nNexT2:
                            End If
                            If DE Then DoEvents
                        Next
                    End If
                Else
                    For i = LBound(tArr) To UBound(tArr)
                        If tArr(i) <> "" Then
                            s = s & tArr(i) & ";"
                            tArr(i) = ""
                        End If
                        If DE Then DoEvents
                    Next
                End If
                Erase tArr
                SplitFast s, tArr, ";"
                s = ""
                For i = LBound(tArr) To UBound(tArr)
                    If Not modSC.FastStringComp(tArr(i), "") Then
                        .sSpells = ReplaceFast(.sSpells, ":" & tArr(i) & ";", "")
                        If modSC.FastStringComp(.sSpells, "") Then .sSpells = "0"
                        .sSpellShorts = ReplaceFast(.sSpellShorts, dbSpells(GetSpellID(, CLng(tArr(i)))).sShort & ";", "")
                        If modSC.FastStringComp(.sSpellShorts, "") Then .sSpellShorts = "0"
                    End If
                    If DE Then DoEvents
                Next
            End If
            .iClassBonusLevel = 0
            .lClassChanges = .lClassChanges + 1
            .sClass = dbClass(dbClassIndex).sName
            .dEXPNeeded = .dEXPNeeded - dbClass(dbCurClass).dEXP
            .dEXPNeeded = .dEXPNeeded + dbClass(dbClassIndex).dEXP
            .iAcc = .iAcc + dbClass(dbClassIndex).iAcc
            .iCrits = .iCrits + dbClass(dbClassIndex).iCrits
            .iArmorType = dbClass(dbClassIndex).iArmorType
            .iWeapons = dbClass(dbClassIndex).iWeapon
            .iAcc = .iAcc - dbClass(dbCurClass).iAcc
            .iCrits = .iCrits - dbClass(dbCurClass).iCrits
            .iSpellLevel = dbClass(dbClassIndex).iSpellLevel
            .iSpellType = dbClass(dbClassIndex).iSpellType
            
            .iAC = .iAC + dbClass(dbClassIndex).lACBonus
            .iMaxDamage = .iMaxDamage + dbClass(dbClassIndex).lDamBonus
            .iDodge = .iDodge + dbClass(dbClassIndex).lDodgeBonus
            .iVision = .iVision + dbClass(dbClassIndex).lVisionBonus
                
            .iAC = .iAC - dbClass(dbCurClass).lACBonus
            .iMaxDamage = .iMaxDamage - dbClass(dbCurClass).lDamBonus
            .iDodge = .iDodge - dbClass(dbCurClass).lDodgeBonus
            .iVision = .iVision - dbClass(dbCurClass).lVisionBonus
            
            If dbClass(dbClassIndex).iMaxMana = 0 Then
                .lMaxMana = 0
                .lMana = 0
            End If
            
            If dbClass(dbClassIndex).iMaxMana <> 0 And .lMaxMana = 0 Then
                .lMaxMana = RndNumber(CDbl(dbClass(dbClassIndex).iMinMana), CDbl(dbClass(dbClassIndex).iMaxMana))
                .lMana = .lMaxMana
                
            End If
            .dTotalEXP = .dTotalEXP - .dEXP
            .dEXP = 0
            dVal = .dTotalEXP
            dVal2 = .dEXPNeeded * .iLevel
            If dVal2 > dVal Then
                Do Until dVal2 > dVal
                    .iLevel = .iLevel - 1
                    dVal2 = .dEXPNeeded * .iLevel
                    If DE Then DoEvents
                Loop
            End If
        Else
            modTrain.AddBonusStats dbIndex, dbClass(dbClassIndex).sBaseBonus
            .sClass = dbClass(dbClassIndex).sName
            .lClassChanges = .lClassChanges + 1
            .iClassBonusLevel = 0
            .dClassPoints = 0
            If dbClass(dbClassIndex).iMaxMana <> 0 And .lMaxMana = 0 Then
                .lMaxMana = RndNumber(CDbl(dbClass(dbClassIndex).iMinMana), CDbl(dbClass(dbClassIndex).iMaxMana))
                .lMana = .lMaxMana
            End If
            .iAcc = dbClass(dbClassIndex).iAcc
            .iCrits = dbClass(dbClassIndex).iCrits
            .iArmorType = dbClass(dbClassIndex).iArmorType
            .iWeapons = dbClass(dbClassIndex).iWeapon
            .iSpellLevel = dbClass(dbClassIndex).iSpellLevel
            .iSpellType = dbClass(dbClassIndex).iSpellType
            .dEXPNeeded = .dEXPNeeded + dbClass(dbClassIndex).dEXP
            
            .iAC = .iAC + dbClass(dbClassIndex).lACBonus
            .iMaxDamage = .iMaxDamage + dbClass(dbClassIndex).lDamBonus
            .iDodge = .iDodge + dbClass(dbClassIndex).lDodgeBonus
            .iVision = .iVision + dbClass(dbClassIndex).lVisionBonus
            
            modMiscFlag.SetMiscFlag dbIndex, [Can Dual Wield], dbClass(dbClassIndex).lCanDualWield
            modMiscFlag.SetMiscFlag dbIndex, [Can Sneak], dbClass(dbClassIndex).lCanSneak
            modMiscFlag.SetMiscFlag dbIndex, [Can Steal], dbClass(dbClassIndex).lCanSteal
            modMiscFlag.SetMiscFlag dbIndex, [Can Backstab], dbClass(dbClassIndex).lCanBS
            dVal = .dTotalEXP
            dVal2 = .dEXPNeeded * .iLevel
            If dVal2 > dVal Then
                Do Until dVal2 > dVal
                    .iLevel = .iLevel - 1
                    dVal2 = .dEXPNeeded * .iLevel
                    If DE Then DoEvents
                Loop
            End If
        End If
        Dim dbItemID As Long
        s = ""
        If Not modSC.FastStringComp(.sWeapon, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sWeapon))
            If modWeaponsAndArmor.PlayerCanUseWeapon(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sArms, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sArms))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sBack, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sBack))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sBody, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sBody))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sEars, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sEars))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sFace, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sFace))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sFeet, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sFeet))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sHands, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sHands))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sHead, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sHead))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sLegs, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sLegs))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sNeck, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sNeck))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sShield, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sShield))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        If Not modSC.FastStringComp(.sWaist, "0") Then
            dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sWaist))
            If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
                modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
                modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
                s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
            End If
        End If
        WrapAndSend Index, LIGHTBLUE & s & "The trainer takes you to a back room to train you." & vbCrLf & "After all is done, you are now trained as an Apprentice " & dbClass(dbClassIndex).sName & "." & WHITE & vbCrLf
        SendToAllInRoom Index, LIGHTBLUE & ReplaceFast(s, "Your", .sPlayerName & "'s") & .sPlayerName & " trains as a " & dbClass(dbClassIndex).sName & "." & WHITE & vbCrLf, .lLocation
    End With
    X(Index) = ""
End If
End Function

Sub JustChangeClass(dbIndex As Long, iClassID As Long)
Dim dbClassIndex As Long
Dim dbItemID As Long
dbClassIndex = GetClassID(, iClassID)
With dbPlayers(dbIndex)
    .sClass = dbClass(dbClassIndex).sName
    .iClassBonusLevel = 0
    .dClassPoints = 0
    If dbClass(dbClassIndex).iMaxMana <> 0 And .lMaxMana = 0 Then
        .lMaxMana = RndNumber(CDbl(dbClass(dbClassIndex).iMinMana), CDbl(dbClass(dbClassIndex).iMaxMana))
        .lMana = .lMaxMana
    End If
    .iAcc = dbClass(dbClassIndex).iAcc
    .iCrits = dbClass(dbClassIndex).iCrits
    .iArmorType = dbClass(dbClassIndex).iArmorType
    .iWeapons = dbClass(dbClassIndex).iWeapon
    .iSpellLevel = dbClass(dbClassIndex).iSpellLevel
    .iSpellType = dbClass(dbClassIndex).iSpellType
    .dEXPNeeded = .dEXPNeeded + dbClass(dbClassIndex).dEXP
    
    .iAC = .iAC + dbClass(dbClassIndex).lACBonus
    .iMaxDamage = .iMaxDamage + dbClass(dbClassIndex).lDamBonus
    .iDodge = .iDodge + dbClass(dbClassIndex).lDodgeBonus
    .iVision = .iVision + dbClass(dbClassIndex).lVisionBonus
    
    modMiscFlag.SetMiscFlag dbIndex, [Can Dual Wield], dbClass(dbClassIndex).lCanDualWield
    modMiscFlag.SetMiscFlag dbIndex, [Can Sneak], dbClass(dbClassIndex).lCanSneak
    modMiscFlag.SetMiscFlag dbIndex, [Can Steal], dbClass(dbClassIndex).lCanSteal
    modMiscFlag.SetMiscFlag dbIndex, [Can Backstab], dbClass(dbClassIndex).lCanBS
    dVal = .dTotalEXP
    dVal2 = .dEXPNeeded * .iLevel
    If dVal2 > dVal Then
        Do Until dVal2 > dVal
            .iLevel = .iLevel - 1
            dVal2 = .dEXPNeeded * .iLevel
            If DE Then DoEvents
        Loop
    End If
    
    s = ""
    If Not modSC.FastStringComp(.sWeapon, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sWeapon))
        If modWeaponsAndArmor.PlayerCanUseWeapon(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sArms, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sArms))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sBack, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sBack))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sBody, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sBody))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sEars, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sEars))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sFace, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sFace))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sFeet, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sFeet))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sHands, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sHands))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sHead, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sHead))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sLegs, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sLegs))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sNeck, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sNeck))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sShield, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sShield))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
    If Not modSC.FastStringComp(.sWaist, "0") Then
        dbItemID = GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sWaist))
        If modWeaponsAndArmor.PlayerCanUseArmor(dbIndex, dbItemID) = False Then
            modItemManip.TakeEqItemAndPlaceInInv CLng(dbIndex), dbItems(dbItemID).iID
            modItemManip.TakeItemFromInvAndPutOnGround CLng(dbIndex), dbItems(dbItemID).iID
            s = LIGHTBLUE & "Your " & dbItems(dbItemID).sItemName & " falls to the ground." & vbCrLf
        End If
    End If
End With
End Sub

Sub AddBonusStatsInverse(dbIndex As Long, sBonuses As String)
Dim tArr() As String
SplitFast sBonuses, tArr, ":"
For i = LBound(tArr) To UBound(tArr)
    With dbPlayers(dbIndex)
        If tArr(i) <> "" Then
            Select Case Left$(tArr(i), 3)
                Case "mhp" 'Max Hitpoints
                    .lMaxHP = .lMaxHP - CLng(ReplaceFast(tArr(i), "mhp", ""))
                Case "str" 'Strength
                    .iStr = .iStr - CLng(ReplaceFast(tArr(i), "str", ""))
                Case "agi" 'Agility
                    .iAgil = .iAgil - CLng(ReplaceFast(tArr(i), "agi", ""))
                Case "int" 'Intellect
                    .iInt = .iInt - CLng(ReplaceFast(tArr(i), "int", ""))
                Case "cha" 'Charm
                    .iCha = .iCha - CLng(ReplaceFast(tArr(i), "cha", ""))
                Case "dex" 'Dexterity
                    .iDex = .iDex - CLng(ReplaceFast(tArr(i), "dex", ""))
                Case "pac" 'Armor Class
                    .iAC = .iAC - CLng(ReplaceFast(tArr(i), "pac", ""))
                Case "acc" 'Accurracy
                    .iAcc = .iAcc - CLng(ReplaceFast(tArr(i), "acc", ""))
                Case "cri" 'Crits
                    .iCrits = .iCrits - CLng(ReplaceFast(tArr(i), "cri", ""))
                Case "mma" 'Max Mana
                    .lMaxMana = .lMaxMana - CLng(ReplaceFast(tArr(i), "mma", ""))
                Case "dam" 'damage bonus
                    .iMaxDamage = .iMaxDamage - CLng(ReplaceFast(tArr(i), "dam", ""))
                Case "dod" 'dodge
                    .iDodge = .iDodge - CLng(ReplaceFast(tArr(i), "dod", ""))
            End Select
        End If
    End With
    If DE Then DoEvents
Next
End Sub
