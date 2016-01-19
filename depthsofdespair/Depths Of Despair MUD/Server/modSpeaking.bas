Attribute VB_Name = "modSpeaking"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modSpeaking
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Sub SendToAll(SendWhat$)
Dim a As Long
On Error Resume Next
For a = 1 To UBound(dbPlayers)
    With dbPlayers(a)
        If .iIndex <> 0 Then
            If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False Then
                WrapAndSend .iIndex, ReplaceFakeANSI(a, Wrap(SendWhat$, 80))
            End If
        End If
    End With
    If DE Then DoEvents
Next a
End Sub

Sub SendToAllInRoom(Index As Long, ByVal SendWhat As String, Location As Long, Optional SecondIndex As Long = 0, Optional SendStatline As Boolean = True)
If SecondIndex = 0 Then
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If .iIndex <> 0 Then
                If .iIndex <> Index Then
                    If .lLocation = Location Then
                        If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False Then
                            WrapAndSend .iIndex, SendWhat, SendStatline
                        End If
                    End If
                End If
            End If
        End With
        If DE Then DoEvents
    Next
Else
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If .iIndex <> 0 Then
                If .iIndex <> Index Then
                    If .iIndex <> SecondIndex Then
                        If .lLocation = Location Then
                            If pPoint(.iIndex) = 0 And pLogOn(.iIndex) = False And pLogOnPW(.iIndex) = False Then
                                WrapAndSend .iIndex, SendWhat, SendStatline
                            End If
                        End If
                    End If
                End If
            End If
        End With
        If DE Then DoEvents
    Next
End If
End Sub

Public Function Speaking(Index As Long, Optional dbIndex As Long = 0) As Boolean
'////////SPEAKING////////

Dim b As Boolean
Dim iPlayerIndex As Long
Dim s As String
   On Error GoTo Speaking_Error

If pLogOn(Index) = False And pLogOnPW(Index) = False And pPoint(Index) = 0 Then
    Speaking = True
    dbIndex = GetPlayerIndexNumber(Index)
    If dbIndex <> 0 Then
        b = sScripting(Index, dbPlayers(dbIndex).lLocation)
    Else
        dbIndex = GetPlayerIndexNumber(Index)
        b = sScripting(Index, dbPlayers(dbIndex).lLocation)
    End If
    If dbPlayers(dbIndex).lCanClear = 0 Then
        If b = True Then
            X(Index) = ""
            Exit Function
        End If
    End If
    If modSC.FastStringComp(LCaseFast(Left$(X(Index), 5)), "brod ") Then
        X(Index) = Mid$(X(Index), InStr(1, X(Index), " ") + 1)
        s = X(Index)
        If modMiscFlag.GetMiscFlag(dbIndex, [Gibberish Talk]) <> 0 Then s = modGetData.GetGibberish(s, modMiscFlag.GetMiscFlag(dbIndex, [Gibberish Talk]))
        SendToAll BRIGHTMAGNETA & dbPlayers(dbIndex).sPlayerName & " broadcast: " & s & WHITE & vbCrLf
        X(Index) = ""
    ElseIf modSC.FastStringComp(Left$(X(Index), 1), "/") Then
        Dim sPlayer As String, sMessage As String
        sPlayer = TrimIt(Mid$(X(Index), 2, InStr(1, X(Index), " ") - 1))
        For a = 1 To InStr(1, X(Index), " ")
            X(Index) = Mid$(X(Index), 2)
            If DE Then DoEvents
        Next
        s = X(Index)
        sPlayer = SmartFind(Index, sPlayer, All_Players)
        iPlayerIndex = GetPlayerIndexNumber(, sPlayer)
        If iPlayerIndex = 0 Then
            WrapAndSend Index, RED & "You can't seem to get a hold of " & sPlayer & "." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If modMiscFlag.GetMiscFlag(dbIndex, [Gibberish Talk]) <> 0 Then s = modGetData.GetGibberish(s, modMiscFlag.GetMiscFlag(dbIndex, [Gibberish Talk]))
        With dbPlayers(iPlayerIndex)
            If pPoint(.iIndex) = 0 Then
                WrapAndSend .iIndex, BGBLUE & dbPlayers(dbIndex).sPlayerName & " telepaths: " & s & WHITE & vbCrLf
                WrapAndSend Index, BRIGHTWHITE & "Your message has been sent." & WHITE & vbCrLf
            Else
                WrapAndSend Index, RED & "You can't seem to get a hold of " & sPlayer & "." & WHITE & vbCrLf
            End If
            X(Index) = ""
            Exit Function
        End With
    ElseIf modSC.FastStringComp(X(Index), "") Then
        ToSend$ = modGetData.GetRoomDescription(dbIndex, CLng(dbPlayers(dbIndex).lLocation), False)
        WrapAndSend Index, ToSend$ & WHITE & vbCrLf
        X(Index) = ""
    Else
        With dbPlayers(dbIndex)
            s = X(Index)
            If .iGhostMode = 1 Then
                WrapAndSend Index, RED & "You may not talk to the room in ghost mode." & WHITE & vbCrLf
                X(Index) = ""
                Exit Function
            End If
            If modMiscFlag.GetMiscFlag(dbIndex, [Gibberish Talk]) <> 0 Then s = modGetData.GetGibberish(s, modMiscFlag.GetMiscFlag(dbIndex, [Gibberish Talk]))
            SendToAllInRoom Index, GREEN & .sPlayerName & " said """ & s & """" & vbCrLf & WHITE, .lLocation
            WrapAndSend Index, GREEN & "You said """ & s & """" & vbCrLf & WHITE
            X(Index) = ""
        End With
    End If
End If
'////////END////////

   On Error GoTo 0
   Exit Function

Speaking_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Speaking of Module modSpeaking"
End Function

Sub WrapAndSend(Index As Long, ByVal SendWhat As String, Optional SendStatline As Boolean = True)
Dim dbIndex As Long
dbIndex = GetPlayerIndexNumber(Index)

SendWhat = Wrap(SendWhat, 80)
SendWhat = ReplaceFakeANSI(dbIndex, SendWhat)
If SendStatline Then
    If pPoint(Index) = 0 And pLogOn(Index) = False And pLogOnPW(Index) = False Then
        'frmMain.ws(Index).SendData "[6n"
        'DoEvents
        'If pWeapon(Index).lCol = 1 Then
        '    SendWhat = "[1A[80D[0J" & vbCrLf & SendWhat & modgetdata.GetStatLine(dbIndex, lLen)
        'Else
        If dbIndex = 0 Then
            If pWeapon(Index).lCol = 0 Then
            '"[1A[80D" & vbCrLf & "[0J[80D"
                SendWhat = "[1A[80D" & vbCrLf & "[0J[80D" & SendWhat & modGetData.GetStatLine(dbIndex) & X(Index)
            Else
            '"[1A[80D" & vbCrLf &
                SendWhat = "[1A[80D" & vbCrLf & "[0J[80D" & SendWhat & modGetData.GetStatLine(dbIndex)
            End If
        Else
            If dbPlayers(dbIndex).lQueryTimer = 1 Then
                If pWeapon(Index).lCol = 0 Then
                '"[1A[80D" & vbCrLf & "[0J[80D"
                    SendWhat = "[1A[80D" & vbCrLf & "[0J[80D" & SendWhat & modGetData.GetStatLine(dbIndex) & "(" & CStr(FinishTimer) & "s)" & X(Index)
                Else
                '"[1A[80D" & vbCrLf &
                    SendWhat = "[1A[80D" & vbCrLf & "[0J[80D" & SendWhat & modGetData.GetStatLine(dbIndex) & "(" & CStr(FinishTimer) & "s)"
                End If
            Else
                If pWeapon(Index).lCol = 0 Then
                '"[1A[80D" & vbCrLf & "[0J[80D"
                    SendWhat = "[1A[80D" & vbCrLf & "[0J[80D" & SendWhat & modGetData.GetStatLine(dbIndex) & X(Index)
                Else
                '"[1A[80D" & vbCrLf &
                    SendWhat = "[1A[80D" & vbCrLf & "[0J[80D" & SendWhat & modGetData.GetStatLine(dbIndex)
                End If
            End If
        End If
        'SendWhat = SendWhat
    End If
Else
    SendWhat = vbCrLf & SendWhat
End If
If frmMain.ws(Index).State = sckConnected And Not modSC.FastStringComp(SendWhat, "") Then
    frmMain.ws(Index).SendData SendWhat    '& x(Index)
    DoEvents
End If
'SendWhat = ""
End Sub

Public Function ReplaceFakeANSI(dbIndex As Long, sText As String) As String
If dbIndex <> 0 Then
    With dbPlayers(dbIndex)
        Select Case modMiscFlag.GetStatsPlus(dbIndex, [Pallete Number])
            Case 0
                sText = ReplaceFast(sText, RED, rRED)
                sText = ReplaceFast(sText, GREEN, rGREEN)
                sText = ReplaceFast(sText, YELLOW, rYELLOW)
                sText = ReplaceFast(sText, BLUE, rBLUE)
                sText = ReplaceFast(sText, MAGNETA, rMAGNETA)
                sText = ReplaceFast(sText, LIGHTBLUE, rLIGHTBLUE)
                sText = ReplaceFast(sText, WHITE, rWHITE)
                sText = ReplaceFast(sText, BGRED, rBGRED)
                sText = ReplaceFast(sText, BGGREEN, rBGGREEN)
                sText = ReplaceFast(sText, BGYELLOW, rBGYELLOW)
                sText = ReplaceFast(sText, BGBLUE, rBGBLUE)
                sText = ReplaceFast(sText, BGPURPLE, rBGPURPLE)
                sText = ReplaceFast(sText, BGLIGHTBLUE, rBGLIGHTBLUE)
                sText = ReplaceFast(sText, BRIGHTYELLOW, rbYELLOW)
                sText = ReplaceFast(sText, BRIGHTGREEN, rbGREEN)
                sText = ReplaceFast(sText, BRIGHTRED, rbRED)
                sText = ReplaceFast(sText, BRIGHTBLUE, rbBLUE)
                sText = ReplaceFast(sText, BRIGHTMAGNETA, rbMAGNETA)
                sText = ReplaceFast(sText, BRIGHTLIGHTBLUE, rbLIGHTBLUE)
                sText = ReplaceFast(sText, BRIGHTWHITE, rbWHITE)
            Case 1
                sText = ReplaceFast(sText, RED, rGREEN)
                sText = ReplaceFast(sText, GREEN, rRED)
                sText = ReplaceFast(sText, YELLOW, rBLUE)
                sText = ReplaceFast(sText, BLUE, rYELLOW)
                sText = ReplaceFast(sText, MAGNETA, rLIGHTBLUE)
                sText = ReplaceFast(sText, LIGHTBLUE, rMAGNETA)
                sText = ReplaceFast(sText, WHITE, rWHITE)
                sText = ReplaceFast(sText, BGRED, rBGGREEN)
                sText = ReplaceFast(sText, BGGREEN, rBGRED)
                sText = ReplaceFast(sText, BGYELLOW, rBGBLUE)
                sText = ReplaceFast(sText, BGBLUE, rBGYELLOW)
                sText = ReplaceFast(sText, BGPURPLE, rBGLIGHTBLUE)
                sText = ReplaceFast(sText, BGLIGHTBLUE, rBGPURPLE)
                sText = ReplaceFast(sText, BRIGHTYELLOW, rbGREEN)
                sText = ReplaceFast(sText, BRIGHTGREEN, rbYELLOW)
                sText = ReplaceFast(sText, BRIGHTRED, rbBLUE)
                sText = ReplaceFast(sText, BRIGHTBLUE, rbRED)
                sText = ReplaceFast(sText, BRIGHTMAGNETA, rbLIGHTBLUE)
                sText = ReplaceFast(sText, BRIGHTLIGHTBLUE, rbMAGNETA)
                sText = ReplaceFast(sText, BRIGHTWHITE, rbWHITE)
            Case 2
                sText = ReplaceFast(sText, RED, rbRED)
                sText = ReplaceFast(sText, GREEN, rbGREEN)
                sText = ReplaceFast(sText, YELLOW, rbYELLOW)
                sText = ReplaceFast(sText, BLUE, rbBLUE)
                sText = ReplaceFast(sText, MAGNETA, rbMAGNETA)
                sText = ReplaceFast(sText, LIGHTBLUE, rbLIGHTBLUE)
                sText = ReplaceFast(sText, WHITE, rbWHITE)
                sText = ReplaceFast(sText, BGRED, rBGRED)
                sText = ReplaceFast(sText, BGGREEN, rBGGREEN)
                sText = ReplaceFast(sText, BGYELLOW, rBGYELLOW)
                sText = ReplaceFast(sText, BGBLUE, rBGBLUE)
                sText = ReplaceFast(sText, BGPURPLE, rBGPURPLE)
                sText = ReplaceFast(sText, BGLIGHTBLUE, rBGLIGHTBLUE)
                sText = ReplaceFast(sText, BRIGHTYELLOW, rYELLOW)
                sText = ReplaceFast(sText, BRIGHTGREEN, rGREEN)
                sText = ReplaceFast(sText, BRIGHTRED, rRED)
                sText = ReplaceFast(sText, BRIGHTBLUE, rBLUE)
                sText = ReplaceFast(sText, BRIGHTMAGNETA, rMAGNETA)
                sText = ReplaceFast(sText, BRIGHTLIGHTBLUE, rLIGHTBLUE)
                sText = ReplaceFast(sText, BRIGHTWHITE, rWHITE)
            Case 3
                sText = ReplaceFast(sText, RED, rRED)
                sText = ReplaceFast(sText, GREEN, rLIGHTBLUE)
                sText = ReplaceFast(sText, YELLOW, rGREEN)
                sText = ReplaceFast(sText, BLUE, rMAGNETA)
                sText = ReplaceFast(sText, MAGNETA, rbRED)
                sText = ReplaceFast(sText, LIGHTBLUE, rLIGHTBLUE)
                sText = ReplaceFast(sText, WHITE, rWHITE)
                sText = ReplaceFast(sText, BGRED, rBGRED)
                sText = ReplaceFast(sText, BGGREEN, rBGGREEN)
                sText = ReplaceFast(sText, BGYELLOW, rBGYELLOW)
                sText = ReplaceFast(sText, BGBLUE, rBGBLUE)
                sText = ReplaceFast(sText, BGPURPLE, rBGPURPLE)
                sText = ReplaceFast(sText, BGLIGHTBLUE, rBGLIGHTBLUE)
                sText = ReplaceFast(sText, BRIGHTYELLOW, rbBLUE)
                sText = ReplaceFast(sText, BRIGHTGREEN, rbYELLOW)
                sText = ReplaceFast(sText, BRIGHTRED, rbMAGNETA)
                sText = ReplaceFast(sText, BRIGHTBLUE, rbBLUE)
                sText = ReplaceFast(sText, BRIGHTMAGNETA, rbYELLOW)
                sText = ReplaceFast(sText, BRIGHTLIGHTBLUE, rYELLOW)
                sText = ReplaceFast(sText, BRIGHTWHITE, rbWHITE)
            Case 4
                sText = ReplaceFast(sText, RED, "")
                sText = ReplaceFast(sText, GREEN, "")
                sText = ReplaceFast(sText, YELLOW, "")
                sText = ReplaceFast(sText, BLUE, "")
                sText = ReplaceFast(sText, MAGNETA, "")
                sText = ReplaceFast(sText, LIGHTBLUE, "")
                sText = ReplaceFast(sText, WHITE, "")
                sText = ReplaceFast(sText, BGRED, "")
                sText = ReplaceFast(sText, BGGREEN, "")
                sText = ReplaceFast(sText, BGYELLOW, "")
                sText = ReplaceFast(sText, BGBLUE, "")
                sText = ReplaceFast(sText, BGPURPLE, "")
                sText = ReplaceFast(sText, BGLIGHTBLUE, "")
                sText = ReplaceFast(sText, BRIGHTYELLOW, "")
                sText = ReplaceFast(sText, BRIGHTGREEN, "")
                sText = ReplaceFast(sText, BRIGHTRED, "")
                sText = ReplaceFast(sText, BRIGHTBLUE, "")
                sText = ReplaceFast(sText, BRIGHTMAGNETA, "")
                sText = ReplaceFast(sText, BRIGHTLIGHTBLUE, "")
                sText = ReplaceFast(sText, BRIGHTWHITE, "")
        End Select
    End With
Else
    sText = ReplaceFast(sText, RED, rRED)
    sText = ReplaceFast(sText, GREEN, rGREEN)
    sText = ReplaceFast(sText, YELLOW, rYELLOW)
    sText = ReplaceFast(sText, BLUE, rBLUE)
    sText = ReplaceFast(sText, MAGNETA, rMAGNETA)
    sText = ReplaceFast(sText, LIGHTBLUE, rLIGHTBLUE)
    sText = ReplaceFast(sText, WHITE, rWHITE)
    sText = ReplaceFast(sText, BGRED, rBGRED)
    sText = ReplaceFast(sText, BGGREEN, rBGGREEN)
    sText = ReplaceFast(sText, BGYELLOW, rBGYELLOW)
    sText = ReplaceFast(sText, BGBLUE, rBGBLUE)
    sText = ReplaceFast(sText, BGPURPLE, rBGPURPLE)
    sText = ReplaceFast(sText, BGLIGHTBLUE, rBGLIGHTBLUE)
    sText = ReplaceFast(sText, BRIGHTYELLOW, rbYELLOW)
    sText = ReplaceFast(sText, BRIGHTGREEN, rbGREEN)
    sText = ReplaceFast(sText, BRIGHTRED, rbRED)
    sText = ReplaceFast(sText, BRIGHTBLUE, rbBLUE)
    sText = ReplaceFast(sText, BRIGHTMAGNETA, rbMAGNETA)
    sText = ReplaceFast(sText, BRIGHTLIGHTBLUE, rbLIGHTBLUE)
    sText = ReplaceFast(sText, BRIGHTWHITE, rbWHITE)
End If
ReplaceFakeANSI = sText
End Function

Public Sub StripAndStoreANSI(sIn As String, ByRef sOut As String, ByRef sNums As String)
Dim i As Long
For i = 1 To Len(sIn)
    Select Case Mid$(sIn, i, 1)
        Case RED, GREEN, YELLOW, BLUE, MAGNETA, LIGHTBLUE, WHITE, BGRED, BGYELLOW, BGBLUE, _
             BGPURPLE, BGLIGHTBLUE, BRIGHTYELLOW, BRIGHTGREEN, BRIGHTRED, BRIGHTBLUE, _
             BRIGHTMAGNETA, BRIGHTLIGHTBLUE, BRIGHTWHITE, Chr$(0)
             sNums = sNums & i & ";"
        Case Else
            sOut = sOut & Mid$(sIn, i, 1)
    End Select
Next
End Sub

Public Function Wrap(ByVal sInText As String, iSize As Long) As String
On Error Resume Next
Dim iPos As Long
Dim lLen As Long
Dim sClean As String
Dim aNums As String
Dim Arr() As String
Dim lCount As Long
Dim i As Long
Dim j As Long
Dim tArr() As String
Dim WasGreater As Boolean
Dim tVal$, FinalVal$, HoldVal$
WasGreater = False
lLen = Len(sInText)
If lLen > iSize Then
    WasGreater = True
    SplitFast sInText, tArr, vbCrLf
    For i = 0 To UBound(tArr)
        sClean = ""
        aNums = ""
        StripAndStoreANSI tArr(i), sClean, aNums
        If Len(sClean) > iSize Then
            HoldVal$ = tArr(i)
            tArr(i) = ""
            FinalVal$ = ""
            Do Until Len(sClean) < iSize
                tVal$ = Left$(sClean$, iSize)
                If Len(sClean) > iSize + 1 Then
                    If Mid$(sClean, iSize + 1, 1) = " " Then
                        iPos = iSize
                    Else
                        iPos = InStrRev(tVal$, " ")
                        If Right$(tVal, 1) = " " Then iPos = iPos - 1
                    End If
                Else
                    iPos = InStrRev(tVal$, " ")
                End If
                If iPos = 0 Then
                    Do Until Len(HoldVal) < iSize
                        FinalVal = FinalVal & Left$(HoldVal, iSize) & vbCrLf
                        HoldVal = Mid$(HoldVal, iSize + 1)
                        If DE Then DoEvents
                    Loop
                    If HoldVal <> "" Then
                        FinalVal = FinalVal & HoldVal
                        HoldVal = ""
                    Else
                        FinalVal = Mid$(FinalVal, 1, Len(FinalVal) - 2)
                    End If
                    Exit Do
                End If
                lCount = 0
                If aNums <> "" Then
                    Erase Arr
                    SplitFast aNums, Arr, ";"
                    For j = LBound(Arr) To UBound(Arr)
                        If Arr(j) <> "" Then
                            If iPos + lCount > Val(Arr(j)) Then
                                lCount = lCount + 1
                            Else
                                Exit For
                            End If
                        End If
                        If DE Then DoEvents
                    Next
                End If
                tVal = Left$(HoldVal, iSize + lCount)
                FinalVal$ = FinalVal$ & Left$(tVal$, iPos + lCount) & vbCrLf
                HoldVal$ = Mid$(HoldVal$, iPos + lCount + 1)
                If Left$(HoldVal$, 1) = " " Then HoldVal = Mid$(HoldVal, 2)
                sClean = ""
                aNums = ""
                StripAndStoreANSI HoldVal, sClean, aNums
                If DE Then DoEvents
            Loop
            FinalVal$ = FinalVal$ & HoldVal$
            tArr(i) = FinalVal$
        End If
        If DE Then DoEvents
    Next
End If
If WasGreater = True Then
    Wrap = Join(tArr, vbCrLf)
Else
    Wrap = sInText
End If
End Function


