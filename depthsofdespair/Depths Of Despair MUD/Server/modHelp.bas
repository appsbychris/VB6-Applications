Attribute VB_Name = "modHelp"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modHelp
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Function MUDHelp(Index As Long) As Boolean
'Function for the help command in the game
'so people know how to do stuff
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 4)), "help") Then   'if the commands
    MUDHelp = True
    Dim HelpFile As String, Temp As String 'variables to hold values
    Dim tArr() As String, sTopic As String
    Open App.Path & "\help.dat" For Input As #1 'load the help data
        While Not EOF(1)
            Line Input #1, Temp
            HelpFile = HelpFile & Temp & vbCrLf 'build the file to a variable
            If DE Then DoEvents
        Wend
    Close #1
    'tArr() = Split(HelpFile, vbCrLf) 'split it by returns
    SplitFast HelpFile, tArr, vbCrLf
    If modSC.FastStringComp(LCaseFast(X(Index)), "help") Then   'if they just want a list of the topics
        Temp = BRIGHTWHITE & "Help topics-" & vbCrLf
        For i = 0 To UBound(tArr()) 'get all the topic headers
            If Not modSC.FastStringComp(tArr(i), "") Then
                'build to a string
                Temp = Temp & Left$(tArr(i), InStr(1, tArr(i), "|") - 1) & ", "
            End If
            If DE Then DoEvents
        Next
        'build the end of the message
        Temp = Left$(Temp, Len(Temp) - 2) & WHITE & vbCrLf
        WrapAndSend Index, Temp 'send it to the user
        X(Index) = ""
    Else
        'if they want a specific topic
        'get the topic
        On Error GoTo eh1
        sTopic = TrimIt(Mid$(X(Index), InStr(1, X(Index), " "), Len(X(Index)) - InStr(1, X(Index), " ") + 1))
        'find the topic in the array
        For i = 0 To UBound(tArr())
            If Not modSC.FastStringComp(tArr(i), "") Then
                If modSC.FastStringComp(LCaseFast(Left$(tArr(i), InStr(1, tArr(i), "|") - 1)), LCaseFast(sTopic)) Then
                    'send out the topic
                    WrapAndSend Index, BRIGHTWHITE & sTopic & " help-" & vbCrLf & Mid$(tArr(i), InStr(1, tArr(i), "|") + 1, Len(tArr(i)) - InStr(1, tArr(i), "|")) & WHITE & vbCrLf
                    Exit For
                End If
            End If
            If DE Then DoEvents
        Next
        X(Index) = ""
    End If
End If
eh1:
End Function
