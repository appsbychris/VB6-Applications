Attribute VB_Name = "modBreak"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modBreak
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function Break(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "bre") Then
    If Len(X(Index)) > 3 Then
        If InStr(1, X(Index), " ") > 0 Then
            Exit Function
        End If
        If Mid$(X(Index), 4, 1) Like "[a]" Then
            '
        Else
            Exit Function
        End If
    End If
    Break = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        If .dMonsterID <> 99999 And .dMonsterID <> 99998 Then aMons(.dMonsterID).mIs_Being_Attacked = False
        .dMonsterID = 99999
        .iCasting = 0
        .iPlayerAttacking = 0
        .iSneaking = 0
        .iResting = 0
        .iMeditating = 0
        SendToAllInRoom Index, YELLOW & .sPlayerName & " breaks off combat." & WHITE & vbCrLf, .lLocation
    End With
    WrapAndSend Index, YELLOW & "You break off combat." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Sub BreakOffCombat(dbIndex As Long, Optional KeepMonAttack As Boolean = False) ', Optional bJUSTCHECK As Boolean = False, Optional IsVictim As Boolean = False, Optional ByRef Messages1 As String = "", Optional ByRef Messages2 As String = "", Optional ByRef Messages3 As String = "")
With dbPlayers(dbIndex)
    If .dMonsterID = 99999 And .iPlayerAttacking = 0 Then Exit Sub
    If .dMonsterID <> 99999 And .dMonsterID <> 99998 Then
        If KeepMonAttack = False Then
            aMons(.dMonsterID).mIs_Being_Attacked = False
            aMons(.dMonsterID).mIsAttacking = False
            aMons(.dMonsterID).mHasAttacked = 0
            aMons(.dMonsterID).mPlayerAttacking = -1
        End If
    End If
    .dMonsterID = 99999
    .iCasting = 0
    .iPlayerAttacking = 0
    .iSneaking = 0
    .iResting = 0
    .iMeditating = 0
    'If bJUSTCHECK Then
    '    Messages1 = Messages1 & YELLOW & "You break off combat." & WHITE & vbCrLf
    '    If IsVictim Then Messages2 = Messages2 & YELLOW & .sPlayerName & " breaks off combat with you." & WHITE & vbCrLf
    '    Messages3 = Messages3 & YELLOW & .sPlayerName & " breaks off combat." & WHITE & vbCrLf
    'Else
        'If KeepMonAttack = True Then
        '    SendToAllInRoom .iIndex, YELLOW & .sPlayerName & " breaks off combat." & WHITE & vbCrLf, .lLocation
        '    WrapAndSend .iIndex, YELLOW & "You break off combat." & WHITE & vbCrLf
        'End If
    'End If
End With
End Sub
