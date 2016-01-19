Attribute VB_Name = "modValidate"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modValidate
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function ValidateName(ByVal s As String) As Boolean
Dim i As Long
s = LCaseFast$(s)
For i = 1 To UBound(dbPlayers)
    With dbPlayers(i)
        If modSC.FastStringComp(s, LCaseFast(.sPlayerName)) Then
            ValidateName = True
            Exit Function
        ElseIf modSC.FastStringComp(s, LCaseFast(.sFamCustom)) Then
            ValidateName = True
            Exit Function
        End If
    End With
    If DE Then DoEvents
Next
For i = 1 To UBound(dbFamiliars)
    With dbFamiliars(i)
        If modSC.FastStringComp(s, LCaseFast(.sFamName)) Then
            ValidateName = True
            Exit Function
        End If
    End With
    If DE Then DoEvents
Next
For i = 1 To UBound(dbMonsters)
    With dbMonsters(i)
        If modSC.FastStringComp(s, LCaseFast(.sMonsterName)) Then
            ValidateName = True
            Exit Function
        End If
    End With
    If DE Then DoEvents
Next
For i = 1 To UBound(dbItems)
    With dbItems(i)
        If modSC.FastStringComp(s, LCaseFast(.sItemName)) Then
            ValidateName = True
            Exit Function
        End If
    End With
    If DE Then DoEvents
Next
For i = 1 To UBound(dbClass)
    With dbClass(i)
        If modSC.FastStringComp(s, LCaseFast(.sName)) Then
            ValidateName = True
            Exit Function
        End If
    End With
    If DE Then DoEvents
Next
For i = 1 To UBound(dbRaces)
    With dbRaces(i)
        If modSC.FastStringComp(s, LCaseFast(.sName)) Then
            ValidateName = True
            Exit Function
        End If
    End With
    If DE Then DoEvents
Next
For i = 1 To UBound(dbSpells)
    With dbSpells(i)
        If modSC.FastStringComp(s, LCaseFast(.sSpellName)) Then
            ValidateName = True
            Exit Function
        End If
        If modSC.FastStringComp(s, LCaseFast(.sShort)) Then
            ValidateName = True
            Exit Function
        End If
    End With
    If DE Then DoEvents
Next
For i = 1 To UBound(dbEmotions)
    With dbEmotions(i)
        If modSC.FastStringComp(s, LCaseFast(.sSyntax)) Then
            ValidateName = True
            Exit Function
        End If
    End With
    If DE Then DoEvents
Next
'97-122
For i = 1 To Len(s)
    Select Case Asc(Mid$(s, i, 1))
        Case 97 To 122
        
        Case Else
            ValidateName = True
            Exit Function
    End Select
    If DE Then DoEvents
Next
End Function
