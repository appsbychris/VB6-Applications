Attribute VB_Name = "modDebug"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modDebug
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function ParseDebug(Index As Long) As Boolean
Dim s As String
Dim a As String
Dim sC As String
Dim sS As String
Dim i As Long
Dim j As Long
Dim Arr() As String
   On Error GoTo ParseDebug_Error

If Left$(X(Index), 6) = "debug." Then
    If dbPlayers(GetPlayerIndexNumber(Index)).iDebugMode = 1 Then
        ParseDebug = True
        s = Mid$(X(Index), 7)
        Select Case Left$(s, 5)
            Case "view."
                s = Mid$(s, 6)
                If InStr(1, s, ".") Then sC = Left$(s, InStr(1, s, ".") - 1) Else sC = s
                Select Case sC
                    Case "roomitems"
                        a = BRIGHTWHITE & "Items In room (including hidden)" & vbCrLf
                        With dbMap(GetMapIndex(dbPlayers(GetPlayerIndexNumber(Index)).lLocation))
                            a = a & .sItems & vbCrLf & .sHidden & vbCrLf & .sLetters & vbCrLf
                        End With
                        WrapAndSend Index, a
                    Case "pinv"
                        a = BRIGHTWHITE & "Your inventory:" & vbCrLf
                        a = a & dbPlayers(GetPlayerIndexNumber(Index)).sInventory & vbCrLf & modGetData.GetPlayersEq(Index) & vbCrLf
                        WrapAndSend Index, a
                    Case "inv"
                        i = InStr(1, s, ".")
                        s = Mid$(s, i + 1)
                        s = modSmartFind.SmartFind(Index, s, All_Players)
                        i = GetPlayerIndexNumber(, s)
                        If i = 0 Then
                            X(Index) = ""
                            Exit Function
                        End If
                        With dbPlayers(i)
                            a = BRIGHTWHITE & .sPlayerName & "'s inventory:" & vbCrLf
                            a = a & .sInventory & vbCrLf & modGetData.GetPlayersEq(.iIndex) & vbCrLf
                        End With
                        WrapAndSend Index, a
                    Case "remoteroomitems"
                        i = InStr(1, s, ".")
                        s = Mid$(s, i + 1)
                        a = BRIGHTWHITE & "Items In room [" & Val(s) & "] (including hidden)" & vbCrLf
                        With dbMap(GetMapIndex(Val(s)))
                            a = a & .sItems & vbCrLf & .sHidden & vbCrLf & .sLetters & vbCrLf
                        End With
                        WrapAndSend Index, a
                    Case "pquest"
                        With dbPlayers(GetPlayerIndexNumber(Index))
                            a = BRIGHTWHITE & "Quest and Flag status-" & vbCrLf
                            a = a & "QUEST1=" & .sQuest1 & vbCrLf
                            a = a & "QUEST2=" & .sQuest2 & vbCrLf
                            a = a & "QUEST3=" & .sQuest3 & vbCrLf
                            a = a & "QUEST4=" & .sQuest4 & vbCrLf
                            a = a & "Flag1=" & .iFlag1 & vbCrLf
                            a = a & "Flag2=" & .iFlag2 & vbCrLf
                            a = a & "Flag3=" & .iFlag3 & vbCrLf
                            a = a & "Flag4=" & .iFlag4 & vbCrLf
                            WrapAndSend Index, a
                        End With
                    Case "classpts"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        sC = SmartFind(Index, sC, All_Players)
                        i = GetPlayerIndexNumber(sPlayerName:=sC)
                        If i = 0 Then X(Index) = "": Exit Function
                        With dbPlayers(i)
                            a = BRIGHTWHITE & .sPlayerName & "'s CLASSPTS: " & CStr(.dClassPoints)
                            WrapAndSend Index, a & WHITE & vbCrLf
                        End With
                    Case "quest"
                        i = InStr(1, s, ".")
                        s = Mid$(s, i + 1)
                        s = modSmartFind.SmartFind(Index, s, All_Players)
                        i = GetPlayerIndexNumber(, s)
                        If i = 0 Then
                            X(Index) = ""
                            Exit Function
                        End If
                        With dbPlayers(i)
                            a = BRIGHTWHITE & "Quest and Flag status (" & .sPlayerName & ")-" & vbCrLf
                            a = a & "QUEST1=" & .sQuest1 & vbCrLf
                            a = a & "QUEST2=" & .sQuest2 & vbCrLf
                            a = a & "QUEST3=" & .sQuest3 & vbCrLf
                            a = a & "QUEST4=" & .sQuest4 & vbCrLf
                            a = a & "Flag1=" & .iFlag1 & vbCrLf
                            a = a & "Flag2=" & .iFlag2 & vbCrLf
                            a = a & "Flag3=" & .iFlag3 & vbCrLf
                            a = a & "Flag4=" & .iFlag4 & vbCrLf
                            WrapAndSend Index, a
                        End With
                    Case "date"
                        a = BRIGHTWHITE & "Date Info-" & vbCrLf
                        a = a & modTime.TimeOfDay & vbCrLf
                        a = a & modTime.CurYear & vbCrLf
                        a = a & modTime.udtMonths(modTime.MonthOfYear).MonthName & vbCrLf
                        a = a & modTime.udtMonths(modTime.MonthOfYear).CurDay & "/"
                        a = a & modTime.udtMonths(modTime.MonthOfYear).DaysAMonth & vbCrLf
                        a = a & modTime.GetDayName(udtDays) & WHITE & vbCrLf
                        WrapAndSend Index, a
                    Case "eq"
                        With dbPlayers(GetPlayerIndexNumber(Index))
                            a = .sArms & ";" & .sBack & ";" & .sBody & ";" & .sEars & ";" & .sFace & _
                                ";" & .sFeet & ";" & .sHands & ";" & .sHead & ";" & .sLegs & ";" & .sNeck & _
                                ";" & .sShield & ";" & .sWaist & ";" & .sWeapon & ";" & .sRings( _
                                0) & ";" & .sRings(1) & ";" & .sRings(2) & ";" & .sRings(3) & ";" & .sRings( _
                                4) & ";" & .sRings(5) & vbCrLf
                        End With
                        WrapAndSend Index, a
                    Case "mons"
                        a = BRIGHTWHITE & "Current monsters:"
                        j = 0
                        For i = LBound(aMons) To UBound(aMons)
                            If aMons(i).miID <> 0 Then j = j + 1
                            If DE Then DoEvents
                        Next
                        a = a & " " & j & "/" & MaxMonsters & WHITE & vbCrLf
                        WrapAndSend Index, a
                    Case "roomswithmons"
                        j = 0
                        For i = 1 To UBound(dbMap)
                            With dbMap(i)
                                If .sMonsters <> "0" Then
                                    j = j + 1
                                End If
                            End With
                            If DE Then DoEvents
                        Next
                        a = BRIGHTWHITE & "Current Rooms With Monsters: " & j & "/" & UBound(dbMap) & WHITE & vbCrLf
                        WrapAndSend Index, a
                    Case "script"
                        WrapAndSend Index, dbMap(dbPlayers(GetPlayerIndexNumber(Index)).lDBLocation).sScript
                    Case "roomlight"
                        a = BRIGHTWHITE & "CURRENT ROOM LIGHT: " & dbMap(GetMapIndex(dbPlayers(GetPlayerIndexNumber(Index)).lLocation)).lLight & WHITE & vbCrLf
                        WrapAndSend Index, a
                    Case "perftimer"
                        With dbPlayers(GetPlayerIndexNumber(Index))
                            If .lQueryTimer = 1 Then
                                .lQueryTimer = 0
                            Else
                                .lQueryTimer = 1
                            End If
                        End With
                    Case "hit%"
                        Dim MaxHit As Long
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        sC = SmartFind(Index, sC, All_Players)
                        i = GetPlayerIndexNumber(sPlayerName:=sC)
                        If i = 0 Then X(Index) = "": Exit Function
                        With dbPlayers(i)
                            MaxHit = modGetData.GetPlayerMaxHit(i) + .iAcc
                            If MaxHit > 98 Then MaxHit = 98
                            a = BRIGHTWHITE & .sPlayerName & "'s Hit %: " & CStr(MaxHit)
                            WrapAndSend Index, a & WHITE & vbCrLf
                        End With
                    Case "swings"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        sC = SmartFind(Index, sC, All_Players)
                        i = GetPlayerIndexNumber(sPlayerName:=sC)
                        If i = 0 Then X(Index) = "": Exit Function
                        With dbPlayers(i)
                            SetWeaponStats .iIndex, i
                            a = BRIGHTWHITE & .sPlayerName & "'s Swings: " & CStr(modGetData.GetPlayerSwings(i))
                            WrapAndSend Index, a & WHITE & vbCrLf
                        End With
                    Case "misc"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        sC = SmartFind(Index, sC, All_Players)
                        i = GetPlayerIndexNumber(sPlayerName:=sC)
                        If i = 0 Then i = GetPlayerIndexNumber(Index)
                        With dbPlayers(i)
                            a = BRIGHTWHITE & .sPlayerName & "'s MiscFlag: " & .sMiscFlag
                            WrapAndSend Index, a & WHITE & vbCrLf
                        End With
                End Select
            Case "edit."
                s = Mid$(s, 6)
                If InStr(1, s, ".") Then sC = Left$(s, InStr(1, s, ".") - 1) Else sC = s
                Select Case sC
                    Case "time"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        SplitFast sC, Arr, "."
                        modTime.AddTime Val(Arr(0)), Val(Arr(1)), Val(Arr(2))
                    Case "classpts"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        SplitFast sC, Arr, "."
                        If UBound(Arr) < 1 Then X(Index) = "": Exit Function
                        Arr(0) = SmartFind(Index, Arr(0), All_Players)
                        i = GetPlayerIndexNumber(sPlayerName:=Arr(0))
                        If i = 0 Then X(Index) = "": Exit Function
                        With dbPlayers(i)
                            a = BRIGHTWHITE & .sPlayerName & "'s CLASSPTS modified by " & Arr(1) & ". Was:" & CStr(.dClassPoints) & " Now:"
                            .dClassPoints = .dClassPoints + Val(Arr(1))
                            a = a & .dClassPoints & "." & WHITE & vbCrLf
                            WrapAndSend Index, a
                        End With
                    Case "limited"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        SplitFast sC, Arr, "."
                        If UBound(Arr) < 1 Then X(Index) = "": Exit Function
                        Arr(0) = SmartFind(Index, Arr(0), All_Items)
                        If GetItemID(Arr(0)) = 0 Then X(Index) = "": Exit Function
                        With dbItems(GetItemID(Arr(0)))
                            a = BRIGHTWHITE & "LIMITED VALUE CHANGED BY " & CStr(Val(Arr(1))) & ". WAS: " & .iInGame & " NOW: "
                            .iInGame = .iInGame + Val(Arr(1))
                            a = a & .iInGame & WHITE & vbCrLf
                            WrapAndSend Index, a
                        End With
                    Case "ac"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        SplitFast sC, Arr, "."
                        If UBound(Arr) < 1 Then X(Index) = "": Exit Function
                        Arr(0) = SmartFind(Index, Arr(0), All_Players)
                        i = GetPlayerIndexNumber(sPlayerName:=Arr(0))
                        If i = 0 Then X(Index) = "": Exit Function
                        With dbPlayers(i)
                            a = BRIGHTWHITE & .sPlayerName & "'s AC modified by " & Arr(1) & ". Was:" & .iAC & " Now:"
                            .iAC = .iAC + Val(Arr(1))
                            a = a & .iAC & "." & WHITE & vbCrLf
                            WrapAndSend Index, a
                        End With
                    Case "destroyitem"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        modItemManip.RemoveItemFromInv GetPlayerIndexNumber(Index), CLng(Val(sC))
                    Case "roomlight"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        dbMap(GetMapIndex(dbPlayers(GetPlayerIndexNumber(Index)).lLocation)).lLight = CLng(Val(sC))
                    Case "saferoom"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        If Val(sC) > 1 Then sC = "1"
                        If Val(sC) < 0 Then sC = "0"
                        dbMap(GetMapIndex(dbPlayers(GetPlayerIndexNumber(Index)).lLocation)).iSafeRoom = CLng(Val(sC))
                    Case "monattack"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        SplitFast sC, Arr, "."
                        If UBound(Arr) < 1 Then X(Index) = "": Exit Function
                        If Val(Arr(1)) > 1 Then Arr(1) = "1"
                        If Val(Arr(1)) < 0 Then Arr(1) = "0"
                        Arr(0) = SmartFind(Index, Arr(0), Monster_In_Room)
                        a = BRIGHTWHITE & "Couldn't find monster " & Arr(0) & "."
                        With dbPlayers(GetPlayerIndexNumber(Index))
                            For i = LBound(aMons) To UBound(aMons)
                                If .lLocation = aMons(i).mLoc And LCaseFast(Arr(0)) = LCaseFast(aMons(i).mName) Then
                                    aMons(i).mAttackable = True
                                    a = BRIGHTWHITE & aMons(i).mName & " attackable changed to " & Arr(1) & "."
                                    Exit For
                                End If
                                If DE Then DoEvents
                            Next
                        End With
                        WrapAndSend Index, a & WHITE & vbCrLf
                    Case "killmon"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        sC = SmartFind(Index, sC, Monster_In_Room)
                        a = BRIGHTWHITE & "Couldn't find monster " & sC & "."
                        With dbPlayers(GetPlayerIndexNumber(Index))
                            For i = LBound(aMons) To UBound(aMons)
                                If .lLocation = aMons(i).mLoc And LCaseFast(sC) = LCaseFast(aMons(i).mName) Then
                                    With aMons(i)
                                        .mHP = 0
                                        If .mHP <= 0 Then
                                            a = ""
                                            s = ""
                                            DropMonGold i, sS
                                            a = a & sS
                                            s = s & sS
                                            
                                            sS = ""
                                            SendDeathText i, sS
                                            a = a & sS
                                            s = s & sS
                                            
                                            sS = ""
                                            DropMonItem i, sS
                                            a = a & sS
                                            s = s & sS
                                            
                                            a = a & WHITE & "You have slain " & .mName & "!" & vbCrLf
                                            s = s & BRIGHTGREEN & dbPlayers(GetPlayerIndexNumber(Index)).sPlayerName & " has slain " & .mName & "!" & vbCrLf & WHITE
                                            AddEXP Index, i
                                            a = a & BRIGHTWHITE & "Your experience has increased by " & .mEXP & "." & GREEN & vbCrLf
                                            If dbPlayers(GetPlayerIndexNumber(Index)).lFamID <> 0 Then a = a & BRIGHTWHITE & "Your " & dbPlayers(GetPlayerIndexNumber(Index)).sFamName & " gains " & CStr(.mEXP \ RndNumber(3, 15)) & " experience." & GREEN & vbCrLf
                                            
                                            AddMonsterRgn .mName
                                    
                                            If Not modSC.FastStringComp(pWeapon(Index).wSpellName, "") Then CleanUpSpells Index
                                            dbPlayers(GetPlayerIndexNumber(Index)).dClassPoints = dbPlayers(GetPlayerIndexNumber(Index)).dClassPoints + 0.1
                                            
                                                
                                            ClearOtherAttackers Index, i
                                            sScripting Index, , , 0, .mScript
                                            ReSetMonsterID GetPlayerIndexNumber(Index)
                                            mRemoveItem i
                                            AmountMons = AmountMons - 1
                                        End If
                                    End With
                                    Exit For
                                End If
                                If DE Then DoEvents
                            Next
                        End With
                        WrapAndSend Index, a
                        If s <> "" Then SendToAllInRoom Index, s, dbPlayers(GetPlayerIndexNumber(Index)).lLocation
                    Case "clearroomitems"
                        With dbMap(dbPlayers(GetPlayerIndexNumber(Index)).lDBLocation)
                            .sItems = ""
                        End With
                    Case "clearroommonsters"
                        With dbMap(dbPlayers(GetPlayerIndexNumber(Index)).lDBLocation)
                            .sMonsters = "0"
                        End With
                End Select
            Case "crte."
                s = Mid$(s, 6)
                If InStr(1, s, ".") Then sC = Left$(s, InStr(1, s, ".") - 1) Else sC = s
                Select Case sC
                    Case "item"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        If Not sC Like "#*/#*/E{*}F{*}A{*}B{#*|#*|*|*}/#*" Then
                            a = BRIGHTWHITE & "FORMAT: #/#/E{}F{}A{}B{0|0|0|0}/#" & WHITE & vbCrLf
                            WrapAndSend Index, a
                            X(Index) = ""
                            Exit Function
                        End If
                        With dbPlayers(GetPlayerIndexNumber(Index))
                            If .sInventory = "0" Then .sInventory = ""
                            .sInventory = .sInventory & ":" & sC & ";"
                        End With
                    Case "door"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        SplitFast sC, Arr, "."
                        If UBound(Arr) < 1 Then X(Index) = "": Exit Function
                        i = GetMapIndex(dbPlayers(GetPlayerIndexNumber(Index)).lLocation)
                        j = Val(Arr(1))
                        If j = 2 Then j = 1
                        If j < 0 Then j = 0
                        Select Case LCaseFast(Arr(0))
                            Case "n"
                                With dbMap(i)
                                    If .lNorth <> 0 Then
                                        .lDN = j
                                        With dbMap(GetMapIndex(.lNorth))
                                            .lDS = j
                                        End With
                                    End If
                                End With
                            Case "s"
                                With dbMap(i)
                                    If .lSouth <> 0 Then
                                        .lDS = j
                                        With dbMap(GetMapIndex(.lSouth))
                                            .lDN = j
                                        End With
                                    End If
                                End With
                            Case "e"
                                With dbMap(i)
                                    If .lEast <> 0 Then
                                        .lDE = j
                                        With dbMap(GetMapIndex(.lEast))
                                            .lDW = j
                                        End With
                                    End If
                                End With
                            Case "w"
                                With dbMap(i)
                                    If .lWest <> 0 Then
                                        .lDW = j
                                        With dbMap(GetMapIndex(.lWest))
                                            .lDE = j
                                        End With
                                    End If
                                End With
                            Case "nw"
                                With dbMap(i)
                                    If .lNorthWest <> 0 Then
                                        .lDNW = j
                                        With dbMap(GetMapIndex(.lNorthWest))
                                            .lDSE = j
                                        End With
                                    End If
                                End With
                            Case "ne"
                                With dbMap(i)
                                    If .lNorthEast <> 0 Then
                                        .lDNE = j
                                        With dbMap(GetMapIndex(.lNorthEast))
                                            .lDSW = j
                                        End With
                                    End If
                                End With
                            Case "sw"
                                With dbMap(i)
                                    If .lSouthWest <> 0 Then
                                        .lDSW = j
                                        With dbMap(GetMapIndex(.lSouthWest))
                                            .lDNE = j
                                        End With
                                    End If
                                End With
                            Case "se"
                                With dbMap(i)
                                    If .lSouthEast <> 0 Then
                                        .lDSE = j
                                        With dbMap(GetMapIndex(.lSouthEast))
                                            .lDNW = j
                                        End With
                                    End If
                                End With
                            Case "u"
                                With dbMap(i)
                                    If .lUp <> 0 Then
                                        .lDU = j
                                        With dbMap(GetMapIndex(.lUp))
                                            .lDD = j
                                        End With
                                    End If
                                End With
                            Case "d"
                                With dbMap(i)
                                    If .lDown <> 0 Then
                                        .lDD = j
                                        With dbMap(GetMapIndex(.lDown))
                                            .lDU = j
                                        End With
                                    End If
                                End With
                        End Select
                End Select
            Case "runn."
                s = Mid$(s, 6)
                If InStr(1, s, ".") Then sC = Left$(s, InStr(1, s, ".") - 1) Else sC = s
                Select Case sC
                    Case "script"
                        sC = Mid$(s, InStr(1, s, ".") + 1)
                        sC = ReplaceFast(sC, "]", vbCrLf)
                        sScripting Index, , , , , , , , , , sC
                End Select
        End Select
        X(Index) = ""
    End If
End If

   On Error GoTo 0
   Exit Function

ParseDebug_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ParseDebug of Module modDebug"
End Function
