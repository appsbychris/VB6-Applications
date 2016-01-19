Attribute VB_Name = "modTame"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modTame
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function Tame(Index As Long) As Boolean
Dim s As String
Dim dbIndex As Long
Dim dbMonID As Long
Dim dbMapId As Long
Dim lMax As Long
Dim iChance As Long
Dim sMons As String
Dim sSe As String
Dim sOt As String
Dim i As Long
Dim b As Boolean
Dim amonIndex As Long
Dim FamID As Long
Dim dbFamId As Long
Dim Arr() As String
Dim n As Long
If Left$(LCaseFast(X(Index)), 5) = "tame " Then
    Tame = True
    s = ReplaceFast(X(Index), "tame ", "")
    s = SmartFind(Index, s, Monster_In_Room)
    dbMonID = GetMonsterID(s)
    If dbMonID = 0 Then
        WrapAndSend Index, RED & "You don't see that here!" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If .iHorse > 0 Then
            WrapAndSend Index, RED & "You can't tame this monster while atop your " & .sFamName & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    dbMapId = dbPlayers(dbIndex).lDBLocation
    sMons = modgetdata.GetAllMonsNamesFromRoom(dbMap(dbMapId).lRoomID, dbMapId)
    If InStr(1, sMons, s & ";") = 0 Then
        WrapAndSend Index, RED & "You don't see that here!" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbMonsters(dbMonID)
        If .iTameToFam <> 0 Then
            lMax = lMax - .iAC - .dHP
        Else
            WrapAndSend Index, GREEN & .sMonsterName & " seems to be ignoring you!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    With dbPlayers(dbIndex)
        lMax = lMax + modMiscFlag.GetStatsPlusTotal(dbIndex, [Animal Relations])
        If lMax < 2 Then lMax = 2
        If lMax > 96 Then lMax = 96
        iChance = RndNumber(1, 100)
        If dbMap(dbMapId).sAMonIds <> "" Then
            SplitFast dbMap(dbMapId).sAMonIds, Arr, ";"
            For i = LBound(Arr) To UBound(Arr)
                If Arr(i) <> "" Then
                    n = CLng(Val(Arr(i)))
                    With aMons(n)
                        If (.mName = dbMonsters(dbMonID).sMonsterName) And ( _
                            .mLoc = dbPlayers(dbIndex).lLocation) And ( _
                            .mHostile = False) And (.mIs_Being_Attacked = False) Then
                            '==================================================='
                            amonIndex = n
                            b = True
                            Exit For
                        End If
                    End With
                End If
                If DE Then DoEvents
            Next
        End If
        If Not b Then
            WrapAndSend Index, GREEN & aMons(amonIndex).mName & " seems to be ignoring you!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If iChance <= lMax Then
            FamID = CLng(dbMonsters(dbMonID).iTameToFam)
            dbFamId = GetFamID(FamID)
            RemoveStats Index
            .sFamName = dbFamiliars(dbFamId).sFamName
            .lFamID = FamID
            .lFamMHP = RndNumber(CDbl(dbFamiliars(dbFamId).lStartHPMin), CDbl(dbFamiliars(dbFamId).lStartHPMax))
            .lFamCHP = .lFamMHP
            .dFamCEXP = 0
            .dFamEXPN = dbFamiliars(dbFamId).dEXPPerLevel
            .dFamTEXP = 0
            .lFamLevel = 1
            .lFamAcc = 0
            .lFamMin = dbFamiliars(dbFamId).lMinDam
            .lFamMax = dbFamiliars(dbFamId).lMaxDam
            AddStats Index
            modUpdateDatabase.mRemoveItem CLng(amonIndex)
            sSe = LIGHTBLUE & "You tame the " & dbMonsters(dbMonID).sMonsterName & "!" & WHITE & vbCrLf
            sOt = LIGHTBLUE & .sPlayerName & " tames " & dbMonsters(dbMonID).sMonsterName & "." & WHITE & vbCrLf
            WrapAndSend Index, sSe
            SendToAllInRoom Index, sOt, .lLocation
        Else
            aMons(amonIndex).mIsAttacking = True
            aMons(amonIndex).mPlayerAttacking = Index
            modMonsters.InsertInMonList amonIndex, .lPlayerID, 0
            WrapAndSend .iIndex, BRIGHTRED & aMons(amonIndex).mName & " moves to attack you!" & WHITE & vbCrLf
            SendToAllInRoom .iIndex, BRIGHTRED & aMons(amonIndex).mName & " moves to attack " & .sPlayerName & "." & WHITE & vbCrLf, .lLocation
'            With dbMonsters(dbMonID)
'
'                SplitFast .sAttack, tArr, ":"
'                iDam = RndNumber(CDbl(tArr(0)), CDbl(tArr(1)))
'                With dbPlayers(dbIndex)
'                    .lHP = .lHP - iDam
'                End With
'                WrapAndSend Index, BRIGHTRED & .sMonsterName & " attacks you for " & CStr(iDam) & " damage!" & WHITE & vbCrLf
'                SendToAllInRoom Index, BRIGHTRED & .sMonsterName & " attacks " & dbPlayers(dbIndex).sPlayerName & "!" & WHITE & vbCrLf, CStr(dbPlayers(dbIndex).lLocation)
'                CheckDeath Index
'            End With
        End If
    End With
    X(Index) = ""
End If
End Function
