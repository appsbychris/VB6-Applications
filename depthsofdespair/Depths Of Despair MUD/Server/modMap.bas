Attribute VB_Name = "modMap"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modMap
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'ษอออออออออออออออออออออออออออออออออป
'บ\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/บ
' -0--0--0--0--0--0--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
' \|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/
' -0--0--0--0--0--0--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
' \|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/
' -0--0--0--0--0--*--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
' \|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/
' -0--0--0--0--0--0--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
' \|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/
' -0--0--0--0--0--0--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
'ศอออออออออออออออออออออออออออออออออผ
''5x11
''11x23
'
'2,5 start

Private Type mpArea
    lRealID As Long
    dN As Long
    dS As Long
    ddE As Long
    dW As Long
    dNE As Long
    dNW As Long
    dSE As Long
    dSW As Long
    cN As Long
    CS As Long
    cE As Long
    cW As Long
    cNE As Long
    cNW As Long
    cSE As Long
    cSW As Long
    sIsRoom As Boolean
End Type
Dim udtMaparea(0 To 4, 0 To 10) As mpArea

Private Function CheckBound(curID As Long, MoveID As Long) As Boolean
Dim i As Long
CheckBound = True
i = curID
Do Until i <= 0
    i = i - 11
Loop
If i = -1 Then
    If MoveID - 1 = curID Then
        CheckBound = False
    ElseIf MoveID + 10 = curID Then
        CheckBound = False
    ElseIf MoveID - 12 = curID Then
        CheckBound = False
    End If
ElseIf i = 0 Then
    If MoveID + 1 = curID Then
        CheckBound = False
    ElseIf MoveID - 10 = curID Then
        CheckBound = False
    ElseIf MoveID + 12 = curID Then
        CheckBound = False
    End If
'Else
'    CheckBound = False
End If

End Function

Public Function Map(Index As Long) As Boolean
Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long
Dim u As Long
Dim n As Long
Dim b As Boolean
Dim c As Boolean
Dim s As String
Dim m As Long
Dim dup As String
Dim sT As Long
Dim sBlackBall As String
Dim xx As Long
Dim yy As Long
Dim Arr() As String
If LCaseFast(X(Index)) = "map" Then
    Map = True
    X(Index) = ""
    xx = 5
    yy = 2
    dbIndex = GetPlayerIndexNumber(Index)
    n = dbPlayers(dbIndex).lLocation
    sT = n
    i = 0
    s = ""
    u = -1
    Erase udtMaparea
    sBlackBall = ""
    Do Until b
        With udtMaparea(yy, xx)
            If .sIsRoom = False Then
                k = GetMapIndex(n)
                .lRealID = dbMap(k).lRoomID
                .dN = dbMap(k).lNorth
                .dS = dbMap(k).lSouth
                .ddE = dbMap(k).lEast
                .dW = dbMap(k).lWest
                .dNE = dbMap(k).lNorthEast
                .dNW = dbMap(k).lNorthWest
                .dSE = dbMap(k).lSouthEast
                .dSW = dbMap(k).lSouthWest
                .cN = IIf(.dN = 0, 0, 1)
                .CS = IIf(.dS = 0, 0, 1)
                .cE = IIf(.ddE = 0, 0, 1)
                .cW = IIf(.dW = 0, 0, 1)
                .cNE = IIf(.dNE = 0, 0, 1)
                .cNW = IIf(.dNW = 0, 0, 1)
                .cSE = IIf(.dSE = 0, 0, 1)
                .cSW = IIf(.dSW = 0, 0, 1)
                .sIsRoom = True
            End If
            If Len(dup) > 7 Then
                If Left$(dup, 4) = Right$(dup, 4) Or Left$(dup, 5) = Right$(dup, 5) Then
                    i = 1
                    If InStr(1, sBlackBall, ":" & .lRealID & ":") = 0 Then
                        sBlackBall = sBlackBall & ":" & .lRealID & ":"
                    Else
                        b = True
                    End If
                    'dup = ""
                End If
                If Len(dup) > 10 Then dup = ""
                'Else
                    'dup = ""
                'End If
            End If
            If i = 0 Then
                If .cN = 1 And yy <> 0 And u <> 0 And InStr(1, sBlackBall, ":" & .dN & ":") = 0 And .dN <> sT Then
                    n = .dN
                    s = s & "n;"
                    yy = yy - 1
                    u = 1
                    dup = dup & "0"
                ElseIf .CS = 1 And yy <> 4 And u <> 1 And InStr(1, sBlackBall, ":" & .dS & ":") = 0 And .dS <> sT Then
                    n = .dS
                    s = s & "s;"
                    yy = yy + 1
                    u = 0
                    dup = dup & "1"
                ElseIf .cE = 1 And xx <> 10 And u <> 2 And InStr(1, sBlackBall, ":" & .ddE & ":") = 0 And .ddE <> sT Then
                    n = .ddE
                    s = s & "e;"
                    xx = xx + 1
                    u = 3
                    dup = dup & "2"
                ElseIf .cW = 1 And xx <> 0 And u <> 3 And InStr(1, sBlackBall, ":" & .dW & ":") = 0 And .dW <> sT Then
                    n = .dW
                    s = s & "w;"
                    xx = xx - 1
                    u = 2
                    dup = dup & "3"
                ElseIf .cNE = 1 And xx <> 10 And yy <> 0 And u <> 4 And InStr(1, sBlackBall, ":" & .dNE & ":") = 0 And .dNE <> sT Then
                    n = .dNE
                    s = s & "ne;"
                    xx = xx + 1
                    yy = yy - 1
                    u = 7
                    dup = dup & "4"
                ElseIf .cNW = 1 And xx <> 0 And yy <> 0 And u <> 5 And InStr(1, sBlackBall, ":" & .dNW & ":") = 0 And .dNW <> sT Then
                    n = .dNW
                    s = s & "nw;"
                    xx = xx - 1
                    yy = yy - 1
                    u = 6
                    dup = dup & "5"
                ElseIf .cSE = 1 And xx <> 10 And yy <> 4 And u <> 6 And InStr(1, sBlackBall, ":" & .dSE & ":") = 0 And .dSE <> sT Then
                    n = .dSE
                    s = s & "se;"
                    xx = xx + 1
                    yy = yy + 1
                    u = 5
                    dup = dup & "6"
                ElseIf .cSW = 1 And xx <> 0 And yy <> 4 And u <> 7 And InStr(1, sBlackBall, ":" & .dSW & ":") = 0 And .dSW <> sT Then
                    n = .dSW
                    s = s & "sw;"
                    xx = xx - 1
                    yy = yy + 1
                    u = 4
                    dup = dup & "7"
                Else
                    If s = "" Then
                        b = True
                    Else
                        i = 1
                        If InStr(1, sBlackBall, ":" & .lRealID & ":") = 0 Then
                            sBlackBall = sBlackBall & ":" & .lRealID & ":"
                        Else
                            b = True
                        End If
                    End If
                End If
            Else
                If s <> "" Then
                    s = Left$(s, Len(s) - 1)
                    SplitFast s, Arr, ";"
                    For j = UBound(Arr) To LBound(Arr) Step -1
                        Select Case modGetData.GetShortDir(modGetData.GetOppositeDirection(modGetData.GetLongDir(Arr(j))))
                            Case "n"
                                yy = yy - 1
                                'n = .dN
                                'u = 1
                                'i = 0
                            Case "s"
                                yy = yy + 1
                                'n = .dS
                                'u = 0
                                'i = 0
                            Case "e"
                                xx = xx + 1
                                'n = .ddE
                                'u = 3
                                'i = 0
                            Case "w"
                                xx = xx - 1
                                'n = .dW
                               ' u = 2
                                'i = 0
                            Case "ne"
                                yy = yy - 1
                                xx = xx + 1
                                'n = .dNE
                                'u = 7
                                'i = 0
                            Case "nw"
                                yy = yy - 1
                                xx = xx - 1
                                'n = .dNW
                                'u = 6
                                'i = 0
                            Case "se"
                                yy = yy + 1
                                xx = xx + 1
                                'n = .dSE
                                'u = 5
                                'i = 0
                            Case "sw"
                                yy = yy + 1
                                xx = xx - 1
                                'n = .dSW
                                'u = 4
                                'i = 0
                        End Select
                        If DE Then DoEvents
                    Next
                    i = 0
                    u = -1
                    s = ""
                    dup = ""
                Else
                    u = -1
                    i = 0
                    dup = ""
                End If
            End If
       End With
       If DE Then DoEvents
    Loop

    ReDim Arr(14) As String
    '23
    u = 0
    k = 0
    h = 0
    For i = 0 To 4
        For j = 0 To 10
            With udtMaparea(i, j)
                If .sIsRoom Then
                    If .dNW <> 0 Then Arr(h) = Arr(h) & "\" Else Arr(h) = Arr(h) & " "
                    If .dN <> 0 Then Arr(h) = Arr(h) & "บ" Else Arr(h) = Arr(h) & " "
                    If .dNE <> 0 Then Arr(h) = Arr(h) & "/" Else Arr(h) = Arr(h) & " "
                    If .dW <> 0 Then Arr(h + 1) = Arr(h + 1) & "อ" Else Arr(h + 1) = Arr(h + 1) & " "
                    If .lRealID = dbPlayers(dbIndex).lLocation Then
                        Arr(h + 1) = Arr(h + 1) & BRIGHTRED & "ฒ" & BRIGHTWHITE
                    Else
                        Arr(h + 1) = Arr(h + 1) & "ฒ"
                    End If
                    If .ddE <> 0 Then Arr(h + 1) = Arr(h + 1) & "อ" Else Arr(h + 1) = Arr(h + 1) & " "
                    If .dSW <> 0 Then Arr(h + 2) = Arr(h + 2) & "/" Else Arr(h + 2) = Arr(h + 2) & " "
                    If .dS <> 0 Then Arr(h + 2) = Arr(h + 2) & "บ" Else Arr(h + 2) = Arr(h + 2) & " "
                    If .dSE <> 0 Then Arr(h + 2) = Arr(h + 2) & "\" Else Arr(h + 2) = Arr(h + 2) & " "
                Else
                    Arr(h) = Arr(h) & "   "
                    Arr(h + 1) = Arr(h + 1) & "   "
                    Arr(h + 2) = Arr(h + 2) & "   "
                End If
            End With
            If DE Then DoEvents
        Next
        h = h + 3
        If DE Then DoEvents
    Next
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = YELLOW & "บ " & BRIGHTWHITE & Arr(i) & YELLOW & " บ"
    Next
'ษอออออออออออออออออออออออออออออออออป
'บ                                 บ
'บ\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/บ
' -0--0--0--0--0--0--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
' \|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/
' -0--0--0--0--0--0--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
' \|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/
' -0--0--0--0--0--*--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
' \|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/
' -0--0--0--0--0--0--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
' \|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/
' -0--0--0--0--0--0--0--0--0--0--0-
' /|\/|\/|\/|\/|\/|\/|\/|\/|\/|\/|\
'ศอออออออออออออออออออออออออออออออออผ
    s = Join(Arr, vbCrLf)
    s = YELLOW & "ษอออออออออออออออออออออออออออออออออออป" & vbCrLf & _
        "บ                                   บ" & vbCrLf & _
        s & vbCrLf & _
        "บ                                   บ" & vbCrLf & _
        "ศอออออออออออออออออออออออออออออออออออผ"

    WrapAndSend Index, s & WHITE & vbCrLf
End If
End Function
'Dim dbIndex As Long
'Dim b As Boolean
'Dim bBoo As Boolean
'Dim s As String
'Dim T As String
'Dim i As Long
'Dim j As Long
'Dim k As Long
'Dim m As Long
'If LCaseFast(X(Index)) = "map" Then
'    Map = True
'    X(Index) = ""
'    dbIndex = GetPlayerIndexNumber(Index)
'    InitArr
'    s = ""
'    k = GetMapIndex(dbPlayers(dbIndex).lLocation)
'    AddRm k
'    j = UBound(tRooms)
'    m = j
'    Do Until b
'        bBoo = True
'        i = 0
'        Do Until Not bBoo Or i > 7
'            bBoo = GoDir(i, j)
'            i = i + 1
'            If DE Then DoEvents
'        Loop
'        j = UBound(tRooms)
'        If m = j Then b = True Else m = j
'        If DE Then DoEvents
'    Loop
'
''    b = False
''    For i = 1 To 11
''        s = s & Space$(23) & vbCrLf
''    Next
''    i = dbPlayers(dbIndex).lLocation
''    j = 136
''    Do Until b
''        If InStr(1, s, vbCrLf) = 0 Then
''            b = True
''        End If
''        With dbMap(i)
''            Mid$(s, j, 1) = "0"
''            If .lNorth <> 0 Then
''                If j - 25 > 0 Then
''                    If Mid$(s, j - 25, 1) <> "|" Then
''                        b = True
''                        Mid$(s, j - 25, 1) = "|"
''                        If j - 50 > 0 Then
''                            b = False
''                            j = j - 50
''                            GoTo DoLoop
''                        End If
''                    Else
''                        j = j + 100
''                    End If
''                End If
''            End If
''            If .lSouth <> 0 Then
''                If j + 25 < 276 Then
''                    If Mid$(s, j + 25, 1) <> "|" Then
''                        b = True
''                        Mid$(s, j + 25, 1) = "|"
''                        If j + 50 < 276 Then
''                            b = False
''                            j = j + 50
''                            GoTo DoLoop
''                        End If
''                    Else
''                        j = j - 100
''                    End If
''                End If
''            End If
'''            If .lEast <> 0 Then
'''                Select Case j + 1
'''                    Case 24, 74, 124, 174, 224, 274
'''                    Case Else
'''                        If Mid$(s, j + 1, 1) <> "-" Then
'''                            b = True
'''                            Mid$(s, j + 1, 1) = "-"
'''                            If j + 2 < 276 Then
'''                                b = False
'''                                j = j + 2
'''                                GoTo DoLoop
'''                            End If
'''                        End If
'''                End Select
'''            End If
'''            If .lWest <> 0 Then
'''                Select Case j - 1
'''                    Case 1, 25, 50, 75, 100, 125, 150, 175, 200, 225
'''                    Case Else
'''                        If Mid$(s, j - 1, 1) <> "-" Then
'''                            b = True
'''                            Mid$(s, j - 1, 1) = "-"
'''                            If j - 2 > 0 Then
'''                                b = False
'''                                j = j - 2
'''                                GoTo DoLoop
'''                            End If
'''                        End If
'''                End Select
'''            End If
'''            If .lNorthWest <> 0 Then
'''                If j - 26 > 0 Then
'''                    If Mid$(s, j - 26, 1) <> "\" And Mid$(s, j - 26, 1) <> "X" Then
'''                        b = True
'''                        If Mid$(s, j - 26, 1) <> "/" Then
'''                            Mid$(s, j - 26, 1) = "X"
'''                        Else
'''                            Mid$(s, j - 26, 1) = "\"
'''                        End If
'''                        If j - 52 > 0 Then
'''                            b = False
'''                            j = j - 52
'''                            GoTo DoLoop
'''                        End If
'''                    End If
'''                End If
'''            End If
'''            If .lNorthEast <> 0 Then
'''                If j - 24 > 0 Then
'''                    If Mid$(s, j - 24, 1) <> "/" And Mid$(s, j - 24, 1) <> "X" Then
'''                        b = True
'''                        If Mid$(s, j - 24, 1) <> "\" Then
'''                            Mid$(s, j - 24, 1) = "X"
'''                        Else
'''                            Mid$(s, j - 24, 1) = "/"
'''                        End If
'''                        If j - 48 > 0 Then
'''                            b = False
'''                            j = j - 48
'''                            GoTo DoLoop
'''                        End If
'''                    End If
'''                End If
'''            End If
'''            If .lSouthEast <> 0 Then
'''                If j + 26 < 276 Then
'''                    If Mid$(s, j + 26, 1) <> "\" And Mid$(s, j + 26, 1) <> "X" Then
'''                        b = True
'''                        If Mid$(s, j + 26, 1) <> "/" Then
'''                            Mid$(s, j + 26, 1) = "X"
'''                        Else
'''                            Mid$(s, j + 26, 1) = "\"
'''                        End If
'''                        If j + 52 < 276 Then
'''                            b = False
'''                            j = j + 52
'''                            GoTo DoLoop
'''                        End If
'''                    End If
'''                End If
'''            End If
'''            If .lSouthWest <> 0 Then
'''                If j + 24 < 276 Then
'''                    If Mid$(s, j + 24, 1) <> "/" And Mid$(s, j + 24, 1) <> "X" Then
'''                        b = True
'''                        If Mid$(s, j + 24, 1) <> "\" Then
'''                            Mid$(s, j + 24, 1) = "X"
'''                        Else
'''                            Mid$(s, j + 24, 1) = "/"
'''                        End If
'''                        If j + 48 < 276 Then
'''                            b = False
'''                            j = j + 48
'''                            GoTo DoLoop
'''                        End If
'''                    End If
'''                End If
'''            End If
''            If j = k Then b = True
''        End With
''DoLoop:
''    k = j
''    Loop
'''    j = 1
'''    For i = 1 To 11
'''        j = j + 25
'''        t = t & Mid$(s, j, 25) & vbCrLf
'''    Next
'    s = Join(Arr, vbCrLf)
'    WrapAndSend Index, BRIGHTWHITE & s & WHITE
'End If
'End Function
