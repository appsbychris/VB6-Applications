Attribute VB_Name = "modGuild"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modGuild
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function GuildCommands(Index As Long) As Boolean
If TopGuild(Index) = True Then GuildCommands = True: Exit Function
If LeaveGuild(Index) = True Then GuildCommands = True: Exit Function
If AddMember(Index) = True Then GuildCommands = True: Exit Function
If JoinGuild(Index) = True Then GuildCommands = True: Exit Function
If ChangeRank(Index) = True Then GuildCommands = True: Exit Function
If Guild(Index) = True Then GuildCommands = True: Exit Function
If RemoveMember(Index) = True Then GuildCommands = True: Exit Function
If CreateGuild(Index) = True Then GuildCommands = True: Exit Function
If DisbandGuild(Index) = True Then GuildCommands = True: Exit Function
If GuildTalk(Index) = True Then GuildCommands = True: Exit Function
GuildCommands = False
End Function

Public Function CreateGuild(Index As Long) As Boolean
Dim sG As String
Dim i As Long
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 13)), "create guild ") Then
    CreateGuild = True
    'X(Index) = ReplaceFast(X(Index), "create guild ", "", 1, 1)
    sG = Mid$(X(Index), InStr(1, X(Index), " ") + 1)
    sG = Mid$(sG, InStr(1, sG, " ") + 1)
    If Len(sG) < 2 Then
        WrapAndSend Index, RED & "The minimum length for a guild's name is 2 characters." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If Len(sG) > 10 Then
        WrapAndSend Index, RED & "The maximum length for a guild's name is 10 characters." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    If dbPlayers(dbIndex).sGuild <> "0" Then
        WrapAndSend Index, RED & "You already have a guild." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        If modSC.FastStringComp(LCaseFast(dbPlayers(i).sGuild), LCaseFast(sG)) Then
            WrapAndSend Index, RED & "Someone else as already taken that name." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
            Exit For
        End If
        If DE Then DoEvents
    Next
    With dbPlayers(dbIndex)
        .sGuild = sG
        .iGuildLeader = 1
    End With
    modMiscFlag.SetMiscFlag dbIndex, [Guild Rank], 5
    WrapAndSend Index, YELLOW & "You create the guild " & sG & "." & vbCrLf & WHITE
    X(Index) = ""
End If
End Function

Public Function GuildTalk(Index As Long) As Boolean
Dim a As Long
Dim sT As String
Dim sG As String
Dim sN As String
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 2)), "gu") Then
    If InStr(1, X(Index), " ") = 0 Then Exit Function
    If Len(X(Index)) < 3 Then Exit Function
    If Len(X(Index)) > 2 Then
        If Not Mid$(X(Index), 3, 1) Like " " And Not Mid$(X(Index), 3, 1) Like "i" Then
            Exit Function
        End If
    End If
    GuildTalk = True
    For a = 1 To InStr(1, X(Index), " ")
        X(Index) = Mid$(X(Index), 2)
    Next a
    sT = X(Index)
    If modSC.FastStringComp(sT, "") Then
        
        WrapAndSend Index, RED & "It would help if you had a message." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
        
    End If
    With dbPlayers(GetPlayerIndexNumber(Index))
        If modSC.FastStringComp(.sGuild, "0") Then
            WrapAndSend Index, RED & "You aren't currently in a guild." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sG = .sGuild
        sN = .sPlayerName
    End With
    For a = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(a)
            If .sGuild = sG Then
                If .iIndex <> 0 Then
                    WrapAndSend .iIndex, BRIGHTGREEN & sN & " writes: " & GREEN & sT & WHITE & vbCrLf
                End If
            End If
        End With
        If DE Then DoEvents
    Next
    X(Index) = ""
End If
                
End Function

Public Function LeaveGuild(Index As Long) As Boolean
Dim sN As String
Dim sG As String
Dim bGL As Boolean
Dim i As Long
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 11)), "leave guild") Then
    LeaveGuild = True
    bGL = False
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        sN = .sPlayerName
        sG = .sGuild
        If modSC.FastStringComp(sG, "0") Then
            WrapAndSend Index, RED & "You aren't in a guild." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .sGuild = "0"
        If .iGuildLeader = 1 Then bGL = True
        .iGuildLeader = 0
    End With
    modMiscFlag.SetMiscFlag dbIndex, [Guild Rank], 0
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If .sGuild = sG Then
                
                If .iIndex <> 0 Then WrapAndSend .iIndex, BRIGHTRED & sN & GREEN & " leaves the guild." & WHITE & vbCrLf
                If bGL = True Then
                    .sGuild = 0
                    modMiscFlag.SetMiscFlag i, [Guild Rank], 0
                    If .iIndex <> 0 Then WrapAndSend .iIndex, BRIGHTRED & sN & GREEN & " disbands the guild." & WHITE & vbCrLf
                End If
                
            End If
        End With
        If DE Then DoEvents
    Next
    
    sN = YELLOW & "You leave the guild " & sG & "." & vbCrLf & WHITE
    If bGL = True Then sN = sN & YELLOW & "You are no longer the leader of the guild " & sG & WHITE & vbCrLf
    WrapAndSend Index, sN
    X(Index) = ""
End If
End Function

Public Function DisbandGuild(Index As Long) As Boolean
Dim sN As String
Dim sG As String
Dim i As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 13)), "disband guild") Then
    DisbandGuild = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        If .iGuildLeader <> 1 Then
            WrapAndSend Index, RED & "You have to be the leader of a guild to disband one." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sN = .sPlayerName
        sG = .sGuild
        If modSC.FastStringComp(sG, "0") Then
            WrapAndSend Index, RED & "You aren't in a guild." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .sGuild = "0"
        .iGuildLeader = 0
    End With
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If modSC.FastStringComp(.sGuild, sG) Then
                If bGL = True Then
                    .sGuild = 0
                    If .iIndex <> 0 Then WrapAndSend .iIndex, BRIGHTRED & sN & GREEN & " disbands the guild." & WHITE & vbCrLf
                End If
            End If
        End With
        If DE Then DoEvents
    Next
    
    sN = YELLOW & "You disband the guild " & sG & "." & vbCrLf & WHITE
    
    WrapAndSend Index, sN
    X(Index) = ""
End If
End Function

Public Function AddMember(Index As Long) As Boolean
Dim sN      As String
Dim sG      As String
Dim sI      As String
Dim iID     As Long
Dim i       As Long
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 11)), "add member ") Then
    AddMember = True
    sI = ReplaceFast(X(Index), "add member ", "", 1, 1)
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        i = modMiscFlag.GetMiscFlag(dbIndex, [Guild Rank])
        If .iGuildLeader <> 1 And i <> 4 Then
            WrapAndSend Index, RED & "You have to be a guild leader to invite members." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sN = .sPlayerName
        sG = .sGuild
        If modSC.FastStringComp(sG, "0") Then
            WrapAndSend Index, RED & "You aren't in a guild." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If modSC.FastStringComp(sN, "") Then
            WrapAndSend Index, RED & "You have to specify someone." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    sI = SmartFind(Index, sI, All_Players)
    iID = GetPlayerIndexNumber(, sI)
    If iID = 0 Then
        WrapAndSend Index, RED & "You could not find that person." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(iID)
        If .iIndex = 0 Then
            WrapAndSend Index, RED & "You could not find that person." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .sInvitedToGuild = sG
        WrapAndSend .iIndex, BRIGHTGREEN & sN & GREEN & " invites you to join their guild " & BRIGHTGREEN & sG & GREEN & "." & WHITE & vbCrLf
        sN = .sPlayerName
    End With
    sN = YELLOW & "You invite " & sN & " to join your guild." & vbCrLf & WHITE
    WrapAndSend Index, sN
    X(Index) = ""
End If
End Function

Public Function JoinGuild(Index As Long) As Boolean
Dim sN      As String
Dim sG      As String
Dim i       As Long
Dim dbIndex As Long
If modSC.FastStringComp(TrimIt(LCaseFast(X(Index))), "join guild") Then
    JoinGuild = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If modSC.FastStringComp(.sInvitedToGuild, "0") Then
            WrapAndSend Index, RED & "You haven't been invited to a guild." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If Not modSC.FastStringComp(.sGuild, "0") Then
            WrapAndSend Index, RED & "You must leave the guild you are currently appart of before joining another." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sN = .sPlayerName
        .sGuild = .sInvitedToGuild
        sG = .sGuild
        .sInvitedToGuild = "0"
        modMiscFlag.SetMiscFlag dbIndex, [Guild Rank], 1
    End With
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If modSC.FastStringComp(.sGuild, sG) Then
                If .iIndex <> 0 Then WrapAndSend .iIndex, BRIGHTRED & sN & BLUE & " joins the guild." & WHITE & vbCrLf
            End If
        End With
        If DE Then DoEvents
    Next
    WrapAndSend Index, BLUE & "You join the guild " & sG & "." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function RemoveMember(Index As Long) As Boolean
Dim sN      As String
Dim sG      As String
Dim sI      As String
Dim iID     As Long
Dim i       As Long
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 14)), "remove member ") Then
    RemoveMember = True
    sI = ReplaceFast(X(Index), "remove member ", "", 1, 1)
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        i = modMiscFlag.GetMiscFlag(dbIndex, [Guild Rank])
        If .iGuildLeader <> 1 And i <> 4 Then
            WrapAndSend Index, RED & "You have to be a guild leader to remove members." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        sN = .sPlayerName
        sG = .sGuild
        If modSC.FastStringComp(sG, "0") Then
            WrapAndSend Index, RED & "You aren't in a guild." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If modSC.FastStringComp(sN, "") Then
            WrapAndSend Index, RED & "You have to specify someone." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    sI = SmartFind(Index, sI, All_Players)
    iID = GetPlayerIndexNumber(, sI)
    If iID = 0 Then
        WrapAndSend Index, RED & "You could not find that person." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(iID)
        If .iIndex = 0 Then
            WrapAndSend Index, RED & "You could not find that person." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .sGuild = "0"
        WrapAndSend .iIndex, BRIGHTGREEN & sN & GREEN & " removes you from their guild " & BRIGHTGREEN & sG & GREEN & "." & WHITE & vbCrLf
        sN = .sPlayerName
    End With
    sN = YELLOW & "You remove " & sN & " from your guild." & vbCrLf & WHITE
    WrapAndSend Index, sN
    X(Index) = ""
End If
End Function

Public Function ChangeRank(Index As Long) As Boolean
Dim sTarget     As String
Dim sRank       As String
Dim s           As String
Dim dbIndex     As Long
Dim dbTIndex    As Long
Dim i           As Long
Dim j           As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 8)), "promote ") Or modSC.FastStringComp(LCaseFast(Left$(X(Index), 7)), "demote ") Then
    ChangeRank = True
    s = LCaseFast(Mid$(X(Index), InStr(1, X(Index), " ") + 1))
    sTarget = Left$(s, InStr(1, s, " ") - 1)
    sRank = Mid$(s, InStr(1, s, " ") + 1)
    sTarget = SmartFind(Index, sTarget, All_Players)
    dbTIndex = GetPlayerIndexNumber(, sTarget)
    If dbTIndex = 0 Then
        WrapAndSend Index, RED & sTarget & " is not in your guild." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    If Not modSC.FastStringComp(dbPlayers(dbIndex).sGuild, dbPlayers(dbTIndex).sGuild) Then
        WrapAndSend Index, RED & sTarget & " is not in your guild." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If modSC.FastStringComp(dbPlayers(dbIndex).sGuild, "0") Then
        WrapAndSend Index, RED & "You aren't currently in a guild." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If dbPlayers(dbTIndex).iGuildLeader = 1 Then
        WrapAndSend Index, RED & "You do not have the power to change the leader's rank." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    i = modMiscFlag.GetMiscFlag(dbIndex, [Guild Rank])
    If i <> 5 And i <> 4 Then
        WrapAndSend Index, RED & "You do not have the power to change ranks." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    If IsNumeric(sRank) Then
        If Val(sRank) < 5 And Val(sRank) > -1 Then
            j = Val(sRank)
        Else
            WrapAndSend Index, RED & "That rank does not exsist." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        Select Case j
            Case 4
                sRank = "General"
            Case 3
                sRank = "Lieutenant"
            Case 2
                sRank = "Soldier"
            Case 1
                sRank = "Normal"
            Case 0
                sRank = "Scrub"
        End Select
    Else
        If Len(sRank) < 3 Then
            WrapAndSend Index, RED & "That rank does not exsist." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        Select Case Left$(sRank, 3)
            Case "gen"
                j = 4
                sRank = "General"
            Case "lie"
                j = 3
                sRank = "Lieutenant"
            Case "sol"
                j = 2
                sRank = "Soldier"
            Case "nor"
                j = 1
                sRank = "Normal"
            Case "scr"
                j = 0
                sRank = "Scrub"
            Case Else
                WrapAndSend Index, RED & "That rank does not exsist." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
        End Select
    End If
    modMiscFlag.SetMiscFlag dbTIndex, [Guild Rank], j
    WrapAndSend Index, YELLOW & "You " & GREEN & "change " & YELLOW & dbPlayers(dbTIndex).sPlayerName & "'s" & GREEN & " guild rank to " & YELLOW & sRank & GREEN & "." & WHITE & vbCrLf
    WrapAndSend dbPlayers(dbTIndex).iIndex, YELLOW & dbPlayers(dbIndex).sPlayerName & " " & GREEN & "changes " & YELLOW & "your " & GREEN & "guild rank to " & YELLOW & sRank & GREEN & "." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function Guild(Index As Long) As Boolean
Dim sG As String
Dim sN As String
Dim i As Long
Dim s As String
Dim LL As Long
If modSC.FastStringComp(TrimIt(LCaseFast(X(Index))), "guild") Then
    Guild = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        sG = .sGuild
        If modSC.FastStringComp(sG, "0") Then
            WrapAndSend Index, RED & "You aren't in a guild." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If modSC.FastStringComp(.sGuild, sG) Then
                If .iGuildLeader = 1 Then
                    sN = sN & GREEN & .sPlayerName & BRIGHTGREEN & " (Leader)" & GREEN & vbCrLf
                Else
                    LL = modMiscFlag.GetMiscFlag(i, [Guild Rank])
                    Select Case LL
                        Case 4
                            s = " (" & BRIGHTBLUE & "General" & GREEN & ")"
                        Case 3
                            s = " (" & BLUE & "Lieutenant" & GREEN & ")"
                        Case 2
                            s = " (" & BRIGHTYELLOW & "Soldier" & GREEN & ")"
                        Case 1
                            s = " (" & YELLOW & "Normal" & GREEN & ")"
                        Case 0
                            s = " (" & RED & "scrub" & GREEN & ")"
                    End Select
                    sN = sN & GREEN & .sPlayerName & s & vbCrLf
                End If
            End If
        End With
        If DE Then DoEvents
    Next
'    SplitFast sN, tArr, vbCrLf
'    For i = LBound(tArr) To UBound(tArr) - 1
'        If Len(tArr(i)) > ll Then ll = Len(tArr(i))
'        tArr(i) = "º" & tArr(i)
'        If DE Then DoEvents
'    Next
'    For i = LBound(tArr) To UBound(tArr) - 1
'        tArr(i) = tArr(i) & Space(IIf(InStr(1, tArr(i), "(Leader)"), 30, 28) - Len(tArr(i))) & "º"
'
'        If DE Then DoEvents
'    Next
'    sN = Join(tArr, vbCrLf)
    sN = GREEN & "Your guild members (" & BRIGHTGREEN & sG & GREEN & ")" & vbCrLf & sN
'    & _
'        vbCrLf & "É" & String(26, "Í") & "»" & vbCrLf & sN & "È" & String(26, _
'        "Í") & "¼" & vbCrLf & WHITE
    WrapAndSend Index, sN
    X(Index) = ""
End If
End Function

