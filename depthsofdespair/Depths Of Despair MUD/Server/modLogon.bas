Attribute VB_Name = "modLogon"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modLogon
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Sub LogOnSequence(Index As Long)
'////////LOG ON SEQUENCE////////
Dim intNum&, ToSend$
Dim bFound As Boolean
If pLogOn(Index) = True Then
    If Not modSC.FastStringComp(LCaseFast(X(Index)), "new") And pPoint(Index) = 0 Then
        bFound = False
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            If modSC.FastStringComp(LCaseFast(dbPlayers(i).sSeenAs), LCaseFast(X(Index))) Then
                bFound = True
                Exit For
            End If
            If DE Then DoEvents
        Next
        If Not bFound Then
            sSend Index, "color.brightred & ;There appears to be no lifeform that exsist by that name.; & newline & ;Did you mis-spell it? Or do you wish to create it.; & newline & ;If you wish to create it, type ""; & color.brightyellow & ;new; & color.brightred & ;""; & newline & ;Otherwise, type the name of your account.: ;", ""
            X(Index) = ""
            Exit Sub
        End If
        bFound = False
        For i = LBound(dbPlayers) To UBound(dbPlayers)
            If modSC.FastStringComp(LCaseFast(dbPlayers(i).sSeenAs), LCaseFast(X(Index))) Then
                If dbPlayers(i).iIndex <> 0 Then
                    sSend Index, "color.brightred & ;That player is already online; & newline & ;Choose a different account, or create a new one.; & newline & ;If you wish to create a new one, type ""; & color.brightyellow & ;new; & color.brightred & ;""; & newline & ;Otherwise, type the name of another account.: ;", ""
                    X(Index) = ""
                    Exit Sub
                End If
                Exit For
            End If
            If DE Then DoEvents
        Next
        UpdateList Format$(X(Index), vbProperCase) & " has signed on at " & Time & "."
        sSend Index, "newline & color.brightyellow & ;Input your secrect code: ;", ""
        pLogOn(Index) = False
        PNAME(Index) = X(Index)
        X(Index) = ""
        pLogOnPW(Index) = True
    Else
        Select Case pPoint(Index)
            Case 0:
                UpdateList "New player is signing up at " & Time & "."
                WrapAndSend Index, MAGNETA & "What do you wish to be called?: " & WHITE
                pPoint(Index) = 1
                X(Index) = ""
                Exit Sub
            Case 1:
                If modValidate.ValidateName(X(Index)) = True Then
                    WrapAndSend Index, MAGNETA & "That name is already in use." & vbCrLf & "What do you wish to be called?:" & WHITE
                    X(Index) = ""
                    Exit Sub
                End If
                If Len(X(Index)) < 5 Then
                    WrapAndSend Index, MAGNETA & "Your name must be at least 5 characters long." & vbCrLf & "What do you wish to be called?:" & WHITE
                    X(Index) = ""
                    Exit Sub
                ElseIf Len(X(Index)) > 12 Then
                    WrapAndSend Index, MAGNETA & "Your name cannot be longer then 12 characters." & vbCrLf & "What do you wish to be called?:" & WHITE
                    X(Index) = ""
                    Exit Sub
                End If
                If InStr(1, X(Index), " ") Then
                    WrapAndSend Index, MAGNETA & "Your name cannot contain a space." & vbCrLf & "What do you wish to be called?:" & WHITE
                    X(Index) = ""
                    Exit Sub
                End If
                If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "new") Then X(Index) = Right$(X(Index), Len(X(Index)) - 3)
                If X(Index) = "" Then
                    WrapAndSend Index, MAGNETA & "Your name can not be nothing." & WHITE
                    X(Index) = ""
                    Exit Sub
                End If
                UpdateList "New user has chosen the name " & X(Index) & "."
                PNAME(Index) = X(Index)
                WrapAndSend Index, MAGNETA & "Pick a secret word that no one can guess: " & WHITE
                pPoint(Index) = 55
                X(Index) = ""
                Exit Sub
            Case 2:
                intNum& = 1
                For i = LBound(dbRaces) To UBound(dbRaces)
                    ToSend$ = ToSend$ & intNum& & ": " & dbRaces(i).sName & vbCrLf
                    intNum& = intNum& + 1
                    If DE Then DoEvents
                Next
                WrapAndSend Index, ANSICLS & MAGNETA & "What race will you live by?" & vbCrLf & ToSend$ & _
                    vbCrLf & "Choose a number that corresponds with your race (1-" & intNum& - 1 & "): " & WHITE
                pPoint(Index) = 3
                X(Index) = ""
                Exit Sub
            Case 3:
                ChooseRace X(Index), Index
                Exit Sub
            Case 7:
                WrapAndSend Index, BRIGHTRED & "You have been annilated!" & vbCrLf & "Please wait..." & WHITE & vbCrLf
                intNum& = 1
                For i = LBound(dbRaces) To UBound(dbRaces)
                    ToSend$ = ToSend$ & intNum& & ": " & dbRaces(i).sName & vbCrLf
                    intNum& = intNum& + 1
                    If DE Then DoEvents
                Next
                WrapAndSend Index, MAGNETA & "What race will you live by?" & vbCrLf & ToSend$ & _
                    vbCrLf & "Choose a number that corresponds with your race (1-" & intNum& - 1 & "): " & WHITE
                pPoint(Index) = 8
                X(Index) = ""
                Exit Sub
            Case 8:
                ChooseRace X(Index), Index
                Exit Sub
            Case 55:
                If Not modSC.FastStringComp(X(Index), "") Then
                    WrapAndSend Index, BRIGHTRED & "Please wait..." & WHITE & vbCrLf
                    RollCharacter PNAME(Index), X(Index), Index
                    WrapAndSend Index, MAGNETA & "What is your gender? (M)ale/(F)emale/(I)t?: " & WHITE
                    pPoint(Index) = 56
                    X(Index) = ""
                    Exit Sub
                Else
                    WrapAndSend Index, RED & "Pick a secret word that no one can guess: " & WHITE
                    X(Index) = ""
                    Exit Sub
                End If
            Case 56:
                If X(Index) = "" Then
                    WrapAndSend Index, MAGNETA & "What is your gender? (M)ale/(F)emale/(I)t?: " & WHITE
                    Exit Sub
                End If
                Select Case Left$(LCaseFast(X(Index)), 1)
                    Case "m", "i", "f"
                        modNewChar.ChooseGender Index
                        sSend Index, ";" & ANSICLS & "; & color.yellow & ;Character developemnt:; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your hair style; & newline & color.bgred & ;None; & newline & color.white & ;Short; & newline & ;Ear-Length; & newline & ;Shoulder Length; & newline & ;Mid-Back Length; & newline & ;Lower-Back Length; & newline & ;Thigh Length" & SetMoveCursor(4, 1) & ";", ""
                    Case Else
                        X(Index) = ""
                        WrapAndSend Index, MAGNETA & "What is your gender? (M)ale/(F)emale/(I)t?: " & WHITE
                        Exit Sub
                End Select

                pPoint(Index) = 2
        End Select
    End If
ElseIf pLogOnPW(Index) = True Then
    For i = LBound(dbPlayers) To UBound(dbPlayers)
        With dbPlayers(i)
            If modSC.FastStringComp(LCaseFast(.sSeenAs), LCaseFast(PNAME(Index))) Then
                If modSC.FastStringComp(LCaseFast(.sPlayerPW), LCaseFast(X(Index))) Then
                    If .iLives = 0 Then
                        pPoint(Index) = 7
                        pLogOnPW(Index) = False
                        pLogOn(Index) = True
                        .iIndex = Index
                        X(Index) = ""
                        LogOnSequence Index
                        Exit Sub
                    End If
                    WrapAndSend Index, BRIGHTBLUE & "Welcome back " & BRIGHTRED & .sSeenAs & "." & WHITE & vbCrLf
                    pLogOnPW(Index) = False
                    SendToAll BRIGHTGREEN & .sSeenAs & " has joined the world." & WHITE & vbCrLf
                    .iIndex = Index
                    X(Index) = ""
                    Exit Sub
                Else
                    WrapAndSend Index, RED & "That is not the secret word..." & vbCrLf & "Maybe you mis-typed it?: " & WHITE
                    X(Index) = ""
                    
                End If
            End If
        End With
        If DE Then DoEvents
    Next
End If
'////////END////////
End Sub

