Attribute VB_Name = "modPartyCommands"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modPartyCommands
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'


Public Function PartyCommands(Index As Long) As Boolean
'Function to check for party commands
'Inviting
If InviteToParty(Index) = True Then PartyCommands = True: Exit Function
'joining
If FollowLeader(Index) = True Then PartyCommands = True: Exit Function
'listing party
If ListParty(Index) = True Then PartyCommands = True: Exit Function
'leaving party
If LeaveParty(Index) = True Then PartyCommands = True: Exit Function
If UnInviteParty(Index) = True Then PartyCommands = True: Exit Function
If FrontRank(Index) = True Then PartyCommands = True: Exit Function
If BackRank(Index) = True Then PartyCommands = True: Exit Function
End Function

Public Function UnInviteParty(Index As Long) As Boolean
Dim s As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 9)), "uninvite ") Then
    UnInviteParty = True
    X(Index) = TrimIt(ReplaceFast(X(Index), "uninvite ", ""))
    X(Index) = modSmartFind.SmartFind(Index, X(Index), All_Players)
    s = GetPlayerIndexNumber(, X(Index))
    If modPartyCommands.PlayerIsInParty(GetPlayerIndexNumber(dbPlayers(s).iPartyLeader), s) Then
        With dbPlayers(s)
            modPartyCommands.RemoveFromParty .iIndex, True
            WrapAndSend .iIndex, LIGHTBLUE & "You are removed from the party." & WHITE & vbCrLf
        End With
    Else
        WrapAndSend Index, RED & "They are not in your travel party." & WHITE & vbCrLf
    End If
    X(Index) = ""
End If
End Function

Public Function FollowLeader(Index As Long) As Boolean
'Function to have the player follow the person who invited them
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 4)), "join") Then  'keyword
    FollowLeader = True
    Dim LeaderIndex As String, Follower As String 'strings to hold names and locations
    Dim FollowerLoc As String, LeaderName As String
    Dim lSParty As String, tArr() As String
    Dim sFinalP As String
    With dbPlayers(GetPlayerIndexNumber(Index))
        LeaderIndex = .iInvitedBy
        Follower = .sPlayerName
        FollowerLoc = .lLocation
    End With
    If LeaderIndex = 0 Then
        WrapAndSend Index, RED & "You can't seem to find the person who invited you." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(GetPlayerIndexNumber(CLng(LeaderIndex)))
        If .lLocation = CLng(FollowerLoc) Then
            LeaderName = .sPlayerName
            If modSC.FastStringComp(.sParty, "0") Then .sParty = ""
            lSParty = .sParty
            .sParty = .sParty & ":" & Index & ";"
            .iLeadingParty = 1
            .iPartyRank = 1
            .iPartyLeader = CLng(LeaderIndex)
        Else
            WrapAndSend Index, RED & "You can't seem to find the person who invited you." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    sFinalP = lSParty
    If Not modSC.FastStringComp(lSParty, "") Then
        lSParty = ReplaceFast(lSParty, ":", "")
        If DCount(lSParty, ";") > 1 Then 'see how many members there are
            lSParty = Left$(lSParty, Len(lSParty) - 1) 'split to an array
            'tArr = Split(lSParty, ";")
            SplitFast lSParty, tArr, ";"
        Else
            lSParty = Left$(lSParty, Len(lSParty) - 1)
            ReDim tArr(0) As String
            tArr(0) = lSParty 'if only 1 member, you only need 1 spot in an array
        End If
        Dim i As Long, j As Long
        For i = LBound(tArr) To UBound(tArr)
            For j = LBound(dbPlayers) To UBound(dbPlayers)
                With dbPlayers(j)
                    If CLng(tArr(i)) = .iIndex Then
                        .sParty = .sParty & ":" & Index & ";"
                        Exit For
                    End If
                End With
                If DE Then DoEvents
            Next j
            If DE Then DoEvents
        Next i
    End If
    sFinalP = sFinalP & ":" & LeaderIndex & ";"
    With dbPlayers(GetPlayerIndexNumber(Index))
        If modSC.FastStringComp(.sParty, "0") Then .sParty = ""
        .sParty = sFinalP
        .iLeadingParty = 0
        .iInvitedBy = 0
        .iPartyLeader = LeaderIndex
        .iPartyRank = 1
    End With
    'send out messages
    WrapAndSend Index, LIGHTBLUE & "You gather with " & LeaderName & "'s travel party." & WHITE & vbCrLf
    WrapAndSend CLng(LeaderIndex), LIGHTBLUE & Follower & " gathers your travel party." & WHITE & vbCrLf
    SendToAllInRoom Index, LIGHTBLUE & Follower & " joins " & LeaderName & "'s travel party." & WHITE & vbCrLf, CLng(FollowerLoc), CLng(LeaderIndex)
    X(Index) = ""
End If
End Function

Public Function InviteToParty(Index As Long) As Boolean
'function to invite people to join your party
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 7)), "invite ") Then 'keyword
    InviteToParty = True
    Dim WhoInvite As String, InvitedIndex As String
    Dim Leader As String, LeaderLoc As String
    Dim Found As Boolean 'flag
    Found = False
    'get the players name
    WhoInvite = SmartFind(Index, Mid$(X(Index), InStr(1, X(Index), " ") + 1, Len(X(Index)) - InStr(1, X(Index), " ")), Player_In_Room)
    With dbPlayers(GetPlayerIndexNumber(Index))
        If Not modSC.FastStringComp(.sParty, "0") And .iLeadingParty = 0 Then
            WrapAndSend Index, RED & "You can't lead a party if your in one." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        Leader = .sPlayerName
        LeaderLoc = CStr(.lLocation)
        
    End With
    If GetPlayerIndexNumber(, WhoInvite) <> 0 Then
        With dbPlayers(GetPlayerIndexNumber(, WhoInvite))
            If .iIndex <> 0 Then
                InvitedIndex = .iIndex
                WhoInvite = .sPlayerName
                .iInvitedBy = Index
            Else
                WrapAndSend Index, RED & "You can't seem to find " & WhoInvite & " around here." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        End With
    Else
        WrapAndSend Index, RED & "You can't seem to find " & WhoInvite & " around here." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    'send out messages
    WrapAndSend CLng(InvitedIndex), LIGHTBLUE & Leader & " has invited you to join their travel party." & WHITE & vbCrLf
    WrapAndSend Index, LIGHTBLUE & "You invite " & WhoInvite & " to join your travel party." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function LeaveParty(Index As Long) As Boolean
'function for players to leave the party
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "leave") Then 'keyword
    LeaveParty = True
    X(Index) = ""
    RemoveFromParty Index 'call the remove from party sub
    'send message
    WrapAndSend Index, LIGHTBLUE & "You leave the travel party" & WHITE & vbCrLf
End If
End Function

Public Function ListParty(Index As Long) As Boolean
'function to list all the people in the party
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "party") Then  'keyword
    ListParty = True
    Dim ToSend$
    Dim LeaderIndex As String, PartyMembers As String
    Dim tArr() As String 'temp array
    Dim dPercent As Double
    Dim dMana As Double
    Dim sMana As String
    With dbPlayers(GetPlayerIndexNumber(Index))
        LeaderIndex = .iPartyLeader
        PartyMembers = .sParty
    End With
    If modSC.FastStringComp(PartyMembers, "0") Then  'make sure there is a party
        WrapAndSend Index, RED & "You realize that you can not see who's in your party, because your aren't in one." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    PartyMembers = ReplaceFast(PartyMembers, ":", "")
    If DCount(PartyMembers, ";") > 1 Then 'see how many members there are
        'PartyMembers = Left$(PartyMembers, Len(PartyMembers) - 1) 'split to an array
        'tArr = Split(PartyMembers, ";")
        SplitFast PartyMembers, tArr, ";"
    Else
        PartyMembers = Left$(PartyMembers, Len(PartyMembers) - 1)
        ReDim tArr(0) As String
        tArr(0) = PartyMembers 'if only 1 member, you only need 1 spot in an array
    End If
    
    For i = 0 To UBound(tArr) 'loop the array
        For j = LBound(dbPlayers) To UBound(dbPlayers)
            With dbPlayers(j)
                If Not modSC.FastStringComp(tArr(i), "") Then
                    If CLng(tArr(i)) = .iIndex Then
                        dPercent = .lHP / .lMaxHP
                        dPercent = RoundFast(dPercent, 2)
                        dPercent = dPercent * 100
                        If .lMaxMana <> 0 Then
                            dMana = .lMana / .lMaxMana
                            dMana = RoundFast(dMana, 2)
                            dMana = dMana * 100
                            sMana = dMana
                        Else
                            sMana = "N/A"
                        End If
                        ToSend$ = ToSend$ & GREEN & .sPlayerName & Space(20 - Len(.sPlayerName)) & dPercent & "% HP" & Space(12 - Len(CStr(dPercent) & "&&&&")) & sMana & "% MA" & Space(11 - Len(sMana)) & modgetdata.GetPlayerRankFromNum(.iPartyRank) & vbCrLf
                        Exit For
                    End If
                End If
            End With
            If DE Then DoEvents
        Next j
        If DE Then DoEvents
    Next i
    With dbPlayers(GetPlayerIndexNumber(Index))
        'put you in the message
        dPercent = .lHP / .lMaxHP
        dPercent = FormatNumber(dPercent, 2)
        dPercent = dPercent * 100
        If .lMaxMana <> 0 Then
            dMana = .lMana / .lMaxMana
            dMana = RoundFast(dMana, 2)
            dMana = dMana * 100
            sMana = dMana
        Else
            sMana = "N/A"
        End If
        ToSend$ = GREEN & .sPlayerName & Space(20 - Len(.sPlayerName)) & dPercent & "% HP" & Space(12 - Len(CStr(dPercent) & "&&&&")) & sMana & "% MA" & Space(8) & modgetdata.GetPlayerRankFromNum(.iPartyRank) & vbCrLf & ToSend$
    End With
    With dbPlayers(GetPlayerIndexNumber(CLng(LeaderIndex)))
        ToSend$ = YELLOW & "People you see in your travel party (" & .sPlayerName & "'s party):" & vbCrLf & ToSend$ & WHITE & vbCrLf
    End With
    'Send out the message
    WrapAndSend Index, ToSend$
    X(Index) = ""
    
End If
End Function

Public Function PlayerIsInParty(dbLeaderIndex As Long, dbPlayerInQuestionIndex) As Boolean
With dbPlayers(dbLeaderIndex)
    If .sParty <> "0" Then
        If InStr(1, .sParty, ":" & dbPlayers(dbPlayerInQuestionIndex).iIndex & ";") Then
            PlayerIsInParty = True
        Else
            PlayerIsInParty = False
        End If
    End If
End With
End Function

Sub RemoveFromParty(Index As Long, Optional RemovedFrom As Boolean = False)
'Sub to remove a person from everyone's party
Dim tArr() As String
Dim tVar As String
Dim sName$
Dim PartyLeaderIndex As Long
If GetPlayerIndexNumber(Index) = 0 Then Exit Sub
With dbPlayers(GetPlayerIndexNumber(Index))
    tVar = .sParty
    sName$ = .sPlayerName
    .sParty = "0"
    If .iInvitedBy = .iPartyLeader Then .iInvitedBy = 0
    .iLeadingParty = 0
    .iPartyLeader = 0
    .iPartyRank = 0
End With
If modSC.FastStringComp(tVar, "0") Then Exit Sub  'make sure there is a party
tVar = ReplaceFast(tVar, ":", "") 'get rid of the colons
If DCount(tVar, ";") > 1 Then 'count the semi-colons
   ' tVar = Left$(tVar, Len(tVar) - 1) 'split to an array
    'tArr = Split(tVar, ";")
    SplitFast tVar, tArr, ";"
Else
    tVar = Left$(tVar, Len(tVar) - 1) 'if there is only 1 member,
    ReDim tArr(0) As String 'there only needs to be 1 spot in
    tArr(0) = tVar 'the array
End If
For i = 0 To UBound(tArr) 'loop the array
    If Not modSC.FastStringComp(tArr(i), "") Then
        If GetPlayerIndexNumber(CLng(tArr(i))) <> 0 Then
            With dbPlayers(GetPlayerIndexNumber(CLng(tArr(i))))
                .sParty = ReplaceFast(.sParty, ":" & Index & ";", "", 1, 1)
                PartyLeaderIndex = .iPartyLeader
                If modSC.FastStringComp(.sParty, "") Then .sParty = "0"
                If modSC.FastStringComp(.sParty, "0") Then
                    .iInvitedBy = 0
                    .iLeadingParty = 0
                    .iPartyLeader = 0
                End If
                If Not RemovedFrom Then
                    WrapAndSend .iIndex, LIGHTBLUE & sName$ & " leaves the travel party." & WHITE & vbCrLf
                Else
                    WrapAndSend .iIndex, LIGHTBLUE & sName$ & " was removed from the travel party." & WHITE & vbCrLf
                End If
            End With
        End If
    End If
    If DE Then DoEvents
Next
'send out message to the leader
If Not RemovedFrom Then
    WrapAndSend PartyLeaderIndex, LIGHTBLUE & sName$ & " leaves your travel party." & WHITE & vbCrLf
Else
    WrapAndSend PartyLeaderIndex, LIGHTBLUE & "You remove " & sName$ & " from your travel party." & WHITE & vbCrLf
End If
End Sub

Public Function BackRank(Index As Long) As Boolean
Dim tArr() As String, tVar As String
Dim i As Long, sName$
If modSC.FastStringComp(LCaseFast(X(Index)), "backrank") Then
    BackRank = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        If Not modSC.FastStringComp(.sParty, "0") Then
            If .iLeadingParty <> 1 Then
                tVar = .sParty
                sName$ = .sPlayerName
                .iPartyRank = 2
            Else
                WrapAndSend Index, RED & "You are leading the party." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        Else
            WrapAndSend Index, RED & "You aren't in a party." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    tVar = ReplaceFast(tVar, ":", "")
    If DCount(tVar, ";") > 1 Then 'count the semi-colons
        SplitFast tVar, tArr, ";"
    Else
        tVar = Left$(tVar, Len(tVar) - 1) 'if there is only 1 member,
        ReDim tArr(0) As String 'there only needs to be 1 spot in
        tArr(0) = tVar 'the array
    End If
    For i = 0 To UBound(tArr) 'loop the array
        If Not modSC.FastStringComp(tArr(i), "") Then
            If GetPlayerIndexNumber(CLng(tArr(i))) <> 0 Then
                With dbPlayers(GetPlayerIndexNumber(CLng(tArr(i))))
                    If .iIndex <> 0 Then
                        WrapAndSend .iIndex, LIGHTBLUE & sName$ & " moves to the backranks." & WHITE & vbCrLf
                    End If
                End With
            End If
        End If
        If DE Then DoEvents
    Next
    WrapAndSend Index, LIGHTBLUE & "You move to the backranks of the party." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function FrontRank(Index As Long) As Boolean
Dim tArr() As String, tVar As String
Dim i As Long, sName$
If modSC.FastStringComp(LCaseFast(X(Index)), "frontrank") Then
    FrontRank = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        If Not modSC.FastStringComp(.sParty, "0") Then
            If .iLeadingParty <> 1 Then
                tVar = .sParty
                sName$ = .sPlayerName
                .iPartyRank = 1
            Else
                WrapAndSend Index, RED & "You are leading the party." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        Else
            WrapAndSend Index, RED & "You aren't in a party." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    tVar = ReplaceFast(tVar, ":", "")
    If DCount(tVar, ";") > 1 Then 'count the semi-colons
        SplitFast tVar, tArr, ";"
    Else
        tVar = Left$(tVar, Len(tVar) - 1) 'if there is only 1 member,
        ReDim tArr(0) As String 'there only needs to be 1 spot in
        tArr(0) = tVar 'the array
    End If
    For i = 0 To UBound(tArr) 'loop the array
        If Not modSC.FastStringComp(tArr(i), "") Then
            If GetPlayerIndexNumber(CLng(tArr(i))) <> 0 Then
                With dbPlayers(GetPlayerIndexNumber(CLng(tArr(i))))
                    If .iIndex <> 0 Then
                        WrapAndSend .iIndex, LIGHTBLUE & sName$ & " moves to the frontranks." & WHITE & vbCrLf
                    End If
                End With
            End If
        End If
        If DE Then DoEvents
    Next
    WrapAndSend Index, LIGHTBLUE & "You move to the frontranks of the party." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function
