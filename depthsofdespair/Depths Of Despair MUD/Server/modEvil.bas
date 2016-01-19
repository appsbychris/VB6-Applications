Attribute VB_Name = "modEvil"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modEvil
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Sub AddEvil(dbIndex As Long, amonIndex As Long)
Dim i As Long
Dim b As Boolean
b = False
If amonIndex > UBound(aMons) Then Exit Sub
i = aMons(amonIndex).mEvil
With dbPlayers(dbIndex)
    Select Case .iEvil
        Case Is > 999
            If i < 40 Then .iEvil = .iEvil + (i \ 15): b = True Else .iEvil = .iEvil - 1
        Case 793 To 999
            If i < 1 Then .iEvil = .iEvil + (i \ 14): b = True Else .iEvil = .iEvil - 2
            
        Case 626 To 792
            If i < 1 Then .iEvil = .iEvil + (i \ 13): b = True Else .iEvil = .iEvil - 3
        Case 356 To 625
            If i < 1 Then .iEvil = .iEvil + (i \ 12): b = True Else .iEvil = .iEvil - 4
        Case 256 To 355
            If i < 1 Then .iEvil = .iEvil + (i \ 11): b = True Else .iEvil = .iEvil - 5
        Case 176 To 235
            If i < 1 Then .iEvil = .iEvil + (i \ 9): b = True Else .iEvil = .iEvil - 6
        Case 141 To 175
            If i < 1 Then .iEvil = .iEvil + (i \ 8): b = True Else .iEvil = .iEvil - 7
        Case 121 To 140
            If i < 1 Then .iEvil = .iEvil + (i \ 7): b = True Else .iEvil = .iEvil - 8
        Case 100 To 120
            'GetRep = BRIGHTRED & "Chaos Lord"
            If i < 1 Then .iEvil = .iEvil + (i \ 6): b = True Else .iEvil = .iEvil - 9
        Case 70 To 99
            'GetRep = RED & "Vile Outlaw"
            If i < 1 Then .iEvil = .iEvil + (i \ 5): b = True Else .iEvil = .iEvil - 10
        Case 40 To 69
            'GetRep = BRIGHTYELLOW & "Outlawed Scum"
            If i < 1 Then .iEvil = .iEvil + (i \ 4): b = True Else .iEvil = .iEvil - 11
        Case 10 To 39
            'GetRep = YELLOW & "Petty Thief"
            If i < 1 Then .iEvil = .iEvil + (i): b = True Else .iEvil = .iEvil - 10
        Case -10 To 9
            'GetRep = WHITE & "Citizen"
            If i = 0 Then i = -1
            If i < 1 Then .iEvil = .iEvil + (i * 1.25): b = True Else .iEvil = .iEvil - 9
        Case -40 To -11
            'GetRep = GREEN & "Law-abiding Citizen"
            If i < 1 Then .iEvil = .iEvil + (i * 1.75): b = True Else .iEvil = .iEvil - 8
        Case -70 To -41
            'GetRep = BRIGHTGREEN & "Peace Keeper"
            If i < 1 Then .iEvil = .iEvil + (i * 2.25): b = True Else .iEvil = .iEvil - 7
        Case -99 To -71
            'GetRep = BRIGHTBLUE & "Law Enforcer"
            If i < 1 Then .iEvil = .iEvil + (i * 3.25): b = True Else .iEvil = .iEvil - 6
        Case -175 To -100
            'GetRep = BRIGHTWHITE & "Peace Lord"
            If i < 1 Then .iEvil = .iEvil + (i * 3.5): b = True Else .iEvil = .iEvil - 5
        Case -256 To -176
            If i < 1 Then .iEvil = .iEvil + (i * 3.75): b = True Else .iEvil = .iEvil - 4
        Case -344 To -257
            If i < 1 Then .iEvil = .iEvil + (i * 4): b = True Else .iEvil = .iEvil - 3
        Case -488 To -345
            If i < 1 Then .iEvil = .iEvil + (i * 4.22): b = True Else .iEvil = .iEvil - 2
        Case -625 To -489
            If i < 1 Then .iEvil = .iEvil + (i * 4.68): b = True Else .iEvil = .iEvil - 1
        Case -827 To -626
            If i < 1 Then .iEvil = .iEvil + (i * 5.6): b = True Else .iEvil = .iEvil - 1
        Case -999 To -828
            If i < 1 Then .iEvil = .iEvil + (i * 7.25): b = True Else .iEvil = .iEvil - 1
        Case Is < -999
            If i < 1 Then .iEvil = .iEvil + (i * 10): b = True Else .iEvil = .iEvil - 1
    End Select
    If b Then WrapAndSend dbPlayers(dbIndex).iIndex, MAGNETA & "Your reputation drops." & WHITE & vbCrLf
End With
End Sub

Sub AddPvPEvil(Index1 As Long, Index2 As Long)
Dim piEvil As Long
Dim b As Boolean
b = False
With dbPlayers(GetPlayerIndexNumber(Index2))
    piEvil = .iEvil
End With
With dbPlayers(GetPlayerIndexNumber(Index1))
    Select Case .iEvil
        Case Is >= 100
            'GetRep = BRIGHTRED & "Chaos Lord"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 2): b = True
        Case 70 To 99
            'GetRep = RED & "Vile Outlaw"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 2.5): b = True
        Case 40 To 69
            'GetRep = BRIGHTYELLOW & "Outlawed Scum"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 2.8): b = True
        Case 10 To 39
            'GetRep = YELLOW & "Petty Thief"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 3): b = True
        Case -10 To 9
            'GetRep = WHITE & "Citizen"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 3.2): b = True
        Case -40 To -11
            'GetRep = GREEN & "Law-abiding Citizen"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 3.3): b = True
        Case -70 To -41
            'GetRep = BRIGHTGREEN & "Peace Keeper"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 3.5): b = True
        Case -99 To -71
            'GetRep = BRIGHTBLUE & "Law Enforcer"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 3.7): b = True
        Case Is <= -100
            'GetRep = BRIGHTWHITE & "Peace Lord"
            If piEvil < 10 Then .iEvil = .iEvil + GetX(piEvil, 3.9): b = True
    End Select
End With
If b = True Then WrapAndSend Index1, MAGNETA & "Your reputation drops." & WHITE & vbCrLf
End Sub

Public Function GetX(Num As Long, DivBy As Single) As Long
Dim X As Long
X = CLng(Num / DivBy)
If X < 0 Then X = -1 * X
GetX = X
End Function

Public Function GetRep(Index As Long) As String
Select Case dbPlayers(GetPlayerIndexNumber(Index)).iEvil
    Case Is >= 100
        GetRep = BRIGHTRED & "Chaos Lord"
    Case 70 To 99
        GetRep = RED & "Vile Outlaw"
    Case 40 To 69
        GetRep = BRIGHTYELLOW & "Outlawed Scum"
    Case 10 To 39
        GetRep = YELLOW & "Petty Thief"
    Case -10 To 9
        GetRep = WHITE & "Citizen"
    Case -40 To -11
        GetRep = GREEN & "Respected Member"
    Case -70 To -41
        GetRep = BRIGHTGREEN & "Peace Keeper"
    Case -99 To -71
        GetRep = BRIGHTBLUE & "Law Enforcer"
    Case Is <= -100
        GetRep = BRIGHTWHITE & "Peace Lord"
End Select
End Function
