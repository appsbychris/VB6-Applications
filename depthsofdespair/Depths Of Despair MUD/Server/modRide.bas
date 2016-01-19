Attribute VB_Name = "modRide"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modRide
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function RideFam(Index As Long) As Boolean
Dim dbIndex As Long
Dim dbFamId As Long
If modSC.FastStringComp(LCaseFast(X(Index)), "ride") Then
    RideFam = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If .lFamID = 0 Then
            WrapAndSend Index, RED & "You don't have a familiar to ride!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        dbFamId = GetFamID(.lFamID)
        With dbFamiliars(dbFamId)
            If .lRidable < 1 Then
                WrapAndSend Index, RED & "You can't ride your familiar." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
        End With
        .iSneaking = 0
        .iHorse = dbFamiliars(dbFamId).lSpeed
        WrapAndSend Index, LIGHTBLUE & "You mount your " & .sFamName & "." & WHITE & vbCrLf
        SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " mounts " & modgetdata.GetGenderPronoun(dbIndex, True) & " " & dbFamiliars(dbFamId).sFamName & "." & WHITE & vbCrLf, .lLocation
        X(Index) = ""
    End With
End If
End Function

Public Function GetOffFam(Index As Long) As Boolean
Dim dbIndex As Long
If modSC.FastStringComp(LCaseFast(X(Index)), "get off") Then
    GetOffFam = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If .iHorse < 1 Then
            WrapAndSend Index, RED & "You aren't riding anything." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .iSneaking = 0
        .iHorse = 0
        WrapAndSend Index, LIGHTBLUE & "You dis-mount from your " & .sFamName & "." & WHITE & vbCrLf
        SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " dis-mounts " & modgetdata.GetGenderPronoun(dbIndex, True) & " " & .sFamName & "." & WHITE & vbCrLf, .lLocation
        X(Index) = ""
    End With
End If
End Function
