Attribute VB_Name = "modSet"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modSet
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function SetCommand(Index As Long) As Boolean
Dim s As String
Dim m As Long
If LCaseFast(Left$(X(Index), 4)) = "set " Then
    SetCommand = True
    s = Mid$(X(Index), 5)
    m = InStr(1, s, " ")
    If m = 0 Then SetCommand = False: Exit Function
    Mid$(s, 1, m - 1) = LCaseFast(Mid$(s, 1, m - 1))
    Select Case Mid$(s, 1, m - 1)
        Case "statline"
            With dbPlayers(GetPlayerIndexNumber(Index))
                s = Mid$(s, m + 1)
                Select Case s
                    Case "0"
                        If .lMaxMana < 1 Then
                            s = "HP=;hp/;mhp"
                        Else
                            s = "HP=;hp/;mhp,MA=;ma/;mma"
                        End If
                    Case "1"
                        If .lMaxMana < 1 Then
                            s = "HP=;hp/;mhp,XPtnl=;cexp/;nexp"
                        Else
                            s = "HP=;hp/;mhp,MA=;ma/;mma,XPtnl=;cexp/;nexp"
                        End If
                    Case "2"
                        If .lMaxMana < 1 Then
                            s = "HP=;hp,XPtnl=;%exp,Hunger=;%hunger,Stamina=;%stamina"
                        Else
                            s = "HP=;hp,MA=;ma,XPtnl=;%exp,Hunger=;%hunger,Stamina=;%stamina"
                        End If
                    Case "3"
                        If .lMaxMana < 1 Then
                            s = "HP=;hp,Hunger=;%hunger,Stamina=;%stamina"
                        Else
                            s = "HP=;hp,MA=;ma,Hunger=;%hunger,Stamina=;%stamina"
                        End If
                    Case "4"
                        If .lMaxMana < 1 Then
                            s = "HP=;hp/;mhp,Hun=;%hunger,Sta=;%stamina"
                        Else
                            s = "HP=;hp/;mhp,MA=;ma/;mma,Hun=;%hunger,Sta=;%stamina"
                        End If
                    Case "5"
                        If .lMaxMana < 1 Then
                            s = "HP=;hp,%EXP=;%exp,H=;%hunger,S=;%stamina"
                        Else
                            s = "HP=;hp,MA=;ma,%EXP=;%exp,H=;%hunger,S=;%stamina"
                        End If
                    Case "?"
                        WrapAndSend Index, BRIGHTWHITE & "Stateline syntax:" & vbCrLf & _
                                          LIGHTBLUE & ";hp" & BRIGHTWHITE & " , " & LIGHTBLUE & ";mhp" & BRIGHTWHITE & "    : " & LIGHTBLUE & "Current hit points" & BRIGHTWHITE & " , " & LIGHTBLUE & "Max hit points" & vbCrLf & _
                                          LIGHTBLUE & ";ma" & BRIGHTWHITE & " , " & LIGHTBLUE & ";mma" & BRIGHTWHITE & "    : " & LIGHTBLUE & "Current mana" & BRIGHTWHITE & " , " & LIGHTBLUE & "Max mana" & vbCrLf & _
                                          LIGHTBLUE & ";%hunger" & BRIGHTWHITE & "      : " & LIGHTBLUE & "Current hunger level percent" & vbCrLf & _
                                          LIGHTBLUE & ";%stamina" & BRIGHTWHITE & "     : " & LIGHTBLUE & "Current stamina level percent" & vbCrLf & _
                                          LIGHTBLUE & ";%exp" & BRIGHTWHITE & "         : " & LIGHTBLUE & "Current EXP % until next level up" & vbCrLf & _
                                          LIGHTBLUE & ";cexp" & BRIGHTWHITE & " , " & LIGHTBLUE & ";nexp" & BRIGHTWHITE & " : " & LIGHTBLUE & "Current EXP" & BRIGHTWHITE & " , " & LIGHTBLUE & "EXP required to level up" & vbCrLf & _
                                          BRIGHTWHITE & vbCrLf & "Current statline : " & LIGHTBLUE & .sStatline & WHITE & vbCrLf & vbCrLf
                        X(Index) = ""
                        Exit Function
                    'Case Else
                        
'                        If .lMaxMana < 1 Then
'                            s = "HP=;hp/;mhp"
'                        Else
'                            s = "HP=;hp/;mhp,MA=;ma/;mma"
'                        End If
                End Select
                .sStatline = s
            End With
            WrapAndSend Index, BRIGHTWHITE & "Done." & WHITE & vbCrLf
        Case "pal"
            s = Mid$(s, m + 1)
            modMiscFlag.SetStatsPlus GetPlayerIndexNumber(Index), [Pallete Number], CLng(Val(s))
            WrapAndSend Index, BRIGHTWHITE & "Done." & WHITE & vbCrLf
    End Select
    X(Index) = ""
End If
End Function
