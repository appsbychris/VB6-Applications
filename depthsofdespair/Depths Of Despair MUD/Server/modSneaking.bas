Attribute VB_Name = "modSneaking"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modSneaking
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function Sneak(Index As Long) As Boolean
Dim s As String
Dim dbIndex As Long
Dim lMax As Long
Dim iChance As Long
Dim sMes As String
Dim sOth As String

If Left$(LCaseFast(X(Index)), 2) = "sn" Then
    s = X(Index)
    If Len(s) > 2 Then
        If Mid$(s, 3, 1) <> "e" Then
            Exit Function
        End If
    End If
    Sneak = True
    dbIndex = GetPlayerIndexNumber(Index)
    If modMiscFlag.GetMiscFlag(dbIndex, [Can Sneak]) = 1 Then
        WrapAndSend Index, RED & "Something is stopping you from sneaking." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    If dbPlayers(dbIndex).iGhostMode = 1 Then
        WrapAndSend Index, RED & "Something is stopping you from sneaking." & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    s = ReplaceFast(modGetData.GetPlayersHereWithoutRiding(dbPlayers(dbIndex).lLocation, dbIndex), dbPlayers(dbIndex).sPlayerName & ";", "")
    s = s & modGetData.GetAllMonsNamesFromRoom(dbPlayers(dbIndex).lLocation, dbPlayers(dbIndex).lDBLocation)
    If s <> "" Then
        WrapAndSend Index, RED & "There are things in here watching you..." & whtie & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(dbIndex)
        If .iHorse > 0 Then
            WrapAndSend Index, RED & "You can't sneak around while riding your " & .sFamName & "." & whtie & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        lMax = modMiscFlag.GetStatsPlusTotal(dbIndex, Steath)
        If lMax > 96 Then lMax = 96
        sMes = BRIGHTWHITE & "Attempting to sneak..." & WHITE & vbCrLf
        WaitFor 100
        iChance = RndNumber(1, 100)
        If dbClass(GetClassID(.sClass)).lCanSneak = 0 Then iChance = iChance + RndNumber(1, 400)
        If iChance <= lMax Then
            sMes = sMes & BRIGHTWHITE & "You feel that you are hidden..." & WHITE & vbCrLf
            .iSneaking = 1
        Else
            Select Case RndNumber(0, 4)
                Case 0
                    sMes = sMes & BRIGHTRED & "You stumble over your own feet!" & WHITE & vbCrLf
                    sOth = LIGHTBLUE & .sPlayerName & " stumbles over their own feet." & WHITE & vbCrLf
                Case 1
                    sMes = sMes & BRIGHTRED & "You sneeze!" & WHITE & vbCrLf
                    sOth = LIGHTBLUE & .sPlayerName & " sneezes." & WHITE & vbCrLf
                Case 2
                    sMes = sMes & BRIGHTRED & "You notice something looking at you!" & WHITE & vbCrLf
                    sOth = LIGHTBLUE & .sPlayerName & " looks around as if something is watching them." & WHITE & vbCrLf
                Case 3
                    If RndNumber(0, 1) = 0 Then
                        If .sWeapon <> "0" Then
                            If .iAgil < 20 Then
                                s = .sWeapon
                                sMes = sMes & BRIGHTRED & "You drop your weapon!" & WHITE & vbCrLf
                                modItemManip.TakeEqItemAndPlaceInInv dbIndex, modItemManip.GetItemIDFromUnFormattedString(s)
                                modItemManip.TakeItemFromInvAndPutOnGround dbIndex, modItemManip.GetItemIDFromUnFormattedString(s)
                                sOth = LIGHTBLUE & .sPlayerName & "'s weapon falls out of their hands." & WHITE & vbCrLf
                            Else
                                sMes = sMes & BRIGHTRED & "You almost drop your weapon!" & WHITE & vbCrLf
                                sOth = LIGHTBLUE & .sPlayerName & " fumbles with their weapon." & WHITE & vbCrLf
                            End If
                        Else
                            sMes = sMes & BRIGHTRED & "You hiccup!" & WHITE & vbCrLf
                            sOth = LIGHTBLUE & .sPlayerName & " hiccups." & WHITE & vbCrLf
                        End If
                    Else
                        sMes = sMes & BRIGHTRED & "You get a chill up your spine!" & WHITE & vbCrLf
                        sOth = LIGHTBLUE & .sPlayerName & " shivers." & WHITE & vbCrLf
                    End If
                Case 4
                    sMes = sMes & BRIGHTRED & "You cough!" & WHITE & vbCrLf
                    sOth = LIGHTBLUE & .sPlayerName & " coughs." & WHITE & vbCrLf
            End Select
        End If
        X(Index) = ""
        WrapAndSend Index, sMes
        If sOth <> "" Then SendToAllInRoom Index, sOth, CStr(.lLocation)
    End With
End If
End Function
