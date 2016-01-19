Attribute VB_Name = "modFamiliars"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modFamiliars
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Enum FAMFLGS
    famIDNumber = 0
    famCustom = 1
    famName = 2
    famLevel = 3
    famTEXP = 4
    famCEXP = 5
    famEXPN = 6
    famCHP = 7
    famMHP = 8
    famMin = 9
    famMax = 10
    famAcc = 11
End Enum

Public Function GetFamFlag(dbIndex As Long, w As FAMFLGS) As String
Dim Arr() As String
Dim s As String
With dbPlayers(dbIndex)
    If .sFamFlags = "0" Then .sFamFlags = "0/0/0/0/0/0/0/0/0/0/0/0"
    SplitFast .sFamFlags, Arr, "/"
    GetFamFlag = Arr(w)
End With
End Function

Public Sub SetFamFlag(dbIndex As Long, w As FAMFLGS, sSet As String)
Dim Arr() As String
Dim s As String
With dbPlayers(dbIndex)
    SplitFast .sFamFlags, Arr, "/"
    Arr(w) = sSet
    .sFamFlags = Join(Arr, "/")
End With
End Sub

Public Sub LoadFamFlags(dbIndex As Long)
With dbPlayers(dbIndex)
    .lFamID = CLng(Val(modFamiliars.GetFamFlag(dbIndex, famIDNumber)))
    .sFamCustom = modFamiliars.GetFamFlag(dbIndex, famCustom)
    .sFamName = modFamiliars.GetFamFlag(dbIndex, famName)
    .lFamLevel = CLng(Val(modFamiliars.GetFamFlag(dbIndex, famLevel)))
    .dFamTEXP = Val(modFamiliars.GetFamFlag(dbIndex, famTEXP))
    .dFamCEXP = Val(modFamiliars.GetFamFlag(dbIndex, famCEXP))
    .dFamEXPN = Val(modFamiliars.GetFamFlag(dbIndex, famEXPN))
    .lFamCHP = CLng(Val(modFamiliars.GetFamFlag(dbIndex, famCHP)))
    .lFamMHP = CLng(Val(modFamiliars.GetFamFlag(dbIndex, famMHP)))
    .lFamMin = CLng(Val(modFamiliars.GetFamFlag(dbIndex, famMin)))
    .lFamMax = CLng(Val(modFamiliars.GetFamFlag(dbIndex, famMax)))
    .lFamAcc = CLng(Val(modFamiliars.GetFamFlag(dbIndex, famAcc)))
End With
End Sub

Public Sub UpdateFamFlags(dbIndex As Long)
Dim s As String
With dbPlayers(dbIndex)
    s = CStr(.lFamID) & "/" & .sFamCustom & "/" & .sFamName & "/" & CStr(.lFamLevel) & "/" & CStr(.dFamTEXP) & "/" & CStr(.dFamCEXP) & "/" & _
        CStr(.dFamEXPN) & "/" & CStr(.lFamCHP) & "/" & CStr(.lFamMHP) & "/" & CStr(.lFamMin) & "/" & _
        CStr(.lFamMax) & "/" & CStr(.lFamAcc)
    .sFamFlags = s
End With
End Sub


Sub AddStats(Index As Long, Optional dbIndex As Long)
Dim FamID As Long
If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
FamID = dbPlayers(dbIndex).lFamID
If FamID = 0 Then Exit Sub
With dbFamiliars(GetFamID(FamID))
    modUseItems.DoFlags dbIndex, .sFlags
End With
End Sub

Sub RemoveStats(Index As Long, Optional DontSend As Boolean = False, Optional dbIndex As Long)
Dim FamID As Long
If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
FamID = dbPlayers(dbIndex).lFamID
If FamID = 0 Then Exit Sub
With dbFamiliars(GetFamID(FamID))
    modUseItems.DoFlags dbIndex, .sFlags, Inverse:=True
    With dbPlayers(dbIndex)
        If DontSend = False Then
            If .sFamCustom <> "0" Then
                WrapAndSend Index, BGPURPLE & .sFamCustom & " the " & .sFamName & " leaves you." & WHITE & vbCrLf
            Else
                WrapAndSend Index, BGPURPLE & "Your " & .sFamName & " leaves you." & WHITE & vbCrLf
            End If
        End If
        .lFamID = 0
        .sFamName = "0"
        .sFamCustom = "0"
        .lFamAcc = 0
        .lFamCHP = 0
        .lFamID = 0
        .lFamLevel = 0
        .lFamMax = 0
        .lFamMHP = 0
        .lFamMin = 0
        .iHorse = 0
    End With
End With
End Sub

Public Function NameFam(Index As Long) As Boolean
Dim s As String
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 8)), "name fam") Then
    NameFam = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        If .lFamID = 0 Then
            WrapAndSend Index, RED & "You don't have a familiar!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        s = Mid$(X(Index), 10)
        If modValidate.ValidateName(s) = True Then
            WrapAndSend Index, RED & "Invalid name: " & s & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        .sFamCustom = s
        WrapAndSend Index, LIGHTBLUE & "You name your " & .sFamName & " " & .sFamCustom & "." & WHITE & vbCrLf
        X(Index) = ""
    End With
End If
End Function

Public Function KillFam(Index As Long) As Boolean
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 8)), "kill fam") Then
    KillFam = True
    With dbPlayers(GetPlayerIndexNumber(Index))
        If .sFamName = "0" Then KillFam = False: Exit Function
        If .sFamCustom <> "0" Then
            WrapAndSend Index, BGRED & "You brutially kill " & .sFamCustom & " the " & .sFamName & WHITE & vbCrLf
            SendToAllInRoom Index, BGRED & .sPlayerName & " brutially kills " & .sFamCustom & " the " & .sFamName & "." & vbCrLf & WHITE, .lLocation
        Else
            WrapAndSend Index, BGRED & "You brutially kill your " & .sFamName & WHITE & vbCrLf
            SendToAllInRoom Index, BGRED & .sPlayerName & " brutially kills their " & .sFamName & "." & vbCrLf & WHITE, .lLocation
        End If
        X(Index) = ""
        RemoveStats Index, True
    End With
End If
End Function

Public Function FamStats(Index As Long) As Boolean
Dim dbIndex As Long
Dim dbFamID As Long
Dim d As Double
Dim s As String
Dim lMin As Long
Dim lMax As Long
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "fam") Then
    dbIndex = GetPlayerIndexNumber(Index)
    FamStats = True
    With dbPlayers(dbIndex)
        If .lFamID <> 0 Then
            s = LIGHTBLUE & "Familiar: " & YELLOW & .sFamName & vbCrLf
            If .sFamCustom <> "0" Then s = s & LIGHTBLUE & "Name    : " & YELLOW & .sFamCustom & vbCrLf
            s = s & LIGHTBLUE & "Level   : " & YELLOW & .lFamLevel & vbCrLf
            s = s & LIGHTBLUE & "EXP     : Total: " & YELLOW & .dFamTEXP & LIGHTBLUE & " Current: " & YELLOW & .dFamCEXP & LIGHTBLUE & " TNL: " & YELLOW & .dFamEXPN
            d = .dFamCEXP / .dFamEXPN
            d = RoundFast(d, 2)
            d = d * 100
            s = s & LIGHTBLUE & " (" & YELLOW & CStr(d) & "%" & LIGHTBLUE & ")" & vbCrLf
            s = s & LIGHTBLUE & "HP      : " & YELLOW & .lFamCHP & LIGHTBLUE & "/" & YELLOW & .lFamMHP & vbCrLf
            s = s & LIGHTBLUE & "Damage  : " & YELLOW & .lFamMin & LIGHTBLUE & "-" & YELLOW & .lFamMax & vbCrLf
            s = s & LIGHTBLUE & "Accuracy: " & YELLOW & .lFamAcc & WHITE & vbCrLf
            WrapAndSend Index, s
        Else
            WrapAndSend Index, RED & "You don't have a familiar!" & WHITE & vbCrLf
        End If
        X(Index) = ""
    End With
End If
End Function

Public Sub GetFamAttack(dbFamID As Long, dbIndex As Long, ByRef Min As Long, ByRef Max As Long)
Dim lMod As Long
With dbPlayers(dbIndex)
    If .lFamLevel > dbFamiliars(dbFamID).lLevelMax Then
        lMod = dbFamiliars(dbFamID).lLevelMax * dbFamiliars(dbFamID).lLevelMod
    Else
        lMod = dbFamiliars(dbFamID).lLevelMod * .lFamLevel
    End If
    Min = .lFamMin + RndNumber(0, CDbl(lMod))
    Max = .lFamMax + RndNumber(0, CDbl(lMod))
End With
End Sub
