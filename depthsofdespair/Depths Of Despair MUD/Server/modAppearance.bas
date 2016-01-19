Attribute VB_Name = "modAppearance"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modAppearance
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Enum Appear
    [Hair Length] = 0
    [Hair Color] = 1
    [Hair Style] = 2
    [Eye Color] = 3
    [moustache] = 4
    [beard] = 5
End Enum
    
Sub MoveAppearance(Index As Long, Msg As String)
Dim dbIndex As Long
Dim i As Long
dbIndex = GetPlayerIndexNumber(Index)
Select Case pPoint(Index)
    Case 57
        i = UBound(HairLen)
    Case 58
        i = UBound(ColorLst)
    Case 59
        i = UBound(HairStyle)
    Case 60
        i = UBound(ColorLst)
    Case 61
        i = 6
    Case 62
        i = 6
End Select
Select Case Msg
    Case UP_ARROW, LEFT_ARROW
        With dbPlayers(dbIndex)
            If .lAppStep > 0 Then .lAppStep = .lAppStep - 1 Else .lAppStep = i
        End With
    Case DOWN_ARROW, RIGHT_ARROW
        With dbPlayers(dbIndex)
            If .lAppStep < i Then .lAppStep = .lAppStep + 1 Else .lAppStep = 0
        End With
End Select
Select Case pPoint(Index)
    Case 57
        SendAppearanceScreen dbIndex, [Hair Length]
    Case 58
        SendAppearanceScreen dbIndex, [Hair Color]
    Case 59
        SendAppearanceScreen dbIndex, [Hair Style]
    Case 60
        SendAppearanceScreen dbIndex, [Eye Color]
    Case 61
        SendAppearanceScreen dbIndex, moustache
    Case 62
        SendAppearanceScreen dbIndex, beard
End Select
End Sub

Sub SendAppearanceScreen(dbIndex As Long, WhichOne As Appear)
Select Case WhichOne
    Case 0
        sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your hair length; & newline & " & ";" & HairLen(dbPlayers(dbIndex).lAppStep) & ";", ""
    Case 1
        sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your hair color; & newline & " & ";" & ColorLst(dbPlayers(dbIndex).lAppStep) & ";", ""
    Case 2
        sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your hair style; & newline & " & ";" & HairStyle(dbPlayers(dbIndex).lAppStep) & ";", ""
    Case 3
        sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your eye color; & newline & " & ";" & ColorLst(dbPlayers(dbIndex).lAppStep) & ";", ""
    Case 4
        Select Case dbPlayers(dbIndex).lAppStep
            Case 0
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your moustache style; & newline & color.bgred & ;None; & newline & color.white & ;Normal; & newline & ;Box Car; & newline & ;Bullet Heads; & newline & ;Horse Shoe; & newline & ;Regent; & newline & ;Shermanic" & SetMoveCursor(4, 1) & ";", ""
            Case 1
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your moustache style; & newline & color.white & ;None; & newline & color.bgred & ;Normal; & color.white & newline & ;Box Car; & newline & ;Bullet Heads; & newline & ;Horse Shoe; & newline & ;Regent; & newline & ;Shermanic" & SetMoveCursor(5, 1) & ";", ""
            Case 2
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your moustache style; & newline & color.white & ;None; & newline & ;Normal; & newline & color.bgred & ;Box Car; & color.white & newline & ;Bullet Heads; & newline & ;Horse Shoe; & newline & ;Regent; & newline & ;Shermanic" & SetMoveCursor(6, 1) & ";", ""
            Case 3
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your moustache style; & newline & color.white & ;None; & newline & ;Normal; & newline &  ;Box Car; & newline & color.bgred & ;Bullet Heads; & color.white & newline & ;Horse Shoe; & newline & ;Regent; & newline & ;Shermanic" & SetMoveCursor(7, 1) & ";", ""
            Case 4
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your moustache style; & newline & color.white & ;None; & newline & ;Normal; & newline &  ;Box Car; & newline & ;Bullet Heads; & newline & color.bgred & ;Horse Shoe; & color.white & newline & ;Regent; & newline & ;Shermanic" & SetMoveCursor(8, 1) & ";", ""
            Case 5
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your moustache style; & newline & color.white & ;None; & newline & ;Normal; & newline &  ;Box Car; & newline & ;Bullet Heads; & newline & ;Horse Shoe; & newline & color.bgred & ;Regent; & color.white & newline & ;Shermanic" & SetMoveCursor(9, 1) & ";", ""
            Case 6
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your moustache style; & newline & color.white & ;None; & newline & ;Normal; & newline &  ;Box Car; & newline & ;Bullet Heads; & newline & ;Horse Shoe; & newline & ;Regent; & newline &  color.bgred & ;Shermanic; & color.white & ;" & SetMoveCursor(10, 1) & ";", ""
        End Select
    Case 5
        Select Case dbPlayers(dbIndex).lAppStep
            Case 0
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your beard style; & newline & color.bgred & ;None; & newline & color.white & ;Short Stubble; & newline & ;Bushy; & newline & ;Medium Length and Straight; & newline & ;Long and Curly; & newline & ;Long and Raspy; & newline & ;Medium Length and Curly" & SetMoveCursor(4, 1) & ";", ""
            Case 1
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your beard style; & newline & color.white & ;None; & newline & color.bgred & ;Short Stubble; & color.white & newline & ;Bushy; & newline & ;Medium Length and Straight; & newline & ;Long and Curly; & newline & ;Long and Raspy; & newline & ;Medium Length and Curly" & SetMoveCursor(5, 1) & ";", ""
            Case 2
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your beard style; & newline & color.white & ;None; & newline & ;Short Stubble; & newline & color.bgred & ;Bushy; & color.white & newline & ;Medium Length and Straight; & newline & ;Long and Curly; & newline & ;Long and Raspy; & newline & ;Medium Length and Curly" & SetMoveCursor(6, 1) & ";", ""
            Case 3
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your beard style; & newline & color.white & ;None; & newline & ;Short Stubble; & newline &  ;Bushy; & newline & color.bgred & ;Medium Length and Straight; & color.white & newline & ;Long and Curly; & newline & ;Long and Raspy; & newline & ;Medium Length and Curly" & SetMoveCursor(7, 1) & ";", ""
            Case 4
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your beard style; & newline & color.white & ;None; & newline & ;Short Stubble; & newline &  ;Bushy; & newline & ;Medium Length and Straight; & newline & color.bgred & ;Long and Curly; & color.white & newline & ;Long and Raspy; & newline & ;Medium Length and Curly" & SetMoveCursor(8, 1) & ";", ""
            Case 5
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your beard style; & newline & color.white & ;None; & newline & ;Short Stubble; & newline &  ;Bushy; & newline & ;Medium Length and Straight; & newline & ;Long and Curly; & newline & color.bgred & ;Long and Raspy; & color.white & newline & ;Medium Length and Curly" & SetMoveCursor(9, 1) & ";", ""
            Case 6
                sSend dbPlayers(dbIndex).iIndex, ";" & ANSICLS & "; color.yellow & ;Character development (Use arrow keys to navigate):; & newline & color.lightblue & ;====================================================================; & newline & ;Choose your beard style; & newline & color.white & ;None; & newline & ;Short Stubble; & newline &  ;Bushy; & newline & ;Medium Length and Straight; & newline & ;Long and Curly; & newline & ;Long and Raspy; & newline &  color.bgred & ;Medium Length and Curly; & color.white & ;" & SetMoveCursor(10, 1) & ";", ""
        End Select
End Select
End Sub

Public Function GetPlayerAppearanceNumber(dbIndex As Long, WhichOne As Appear) As Long
Dim Arr() As String
With dbPlayers(dbIndex)
    SplitFast .sAppearance, Arr, ":"
    GetPlayerAppearanceNumber = Val(Arr(WhichOne))
End With
End Function

Public Sub SetPlayerAppearanceNumber(dbIndex As Long, WhichOne As Appear, lNewV As Long)
Dim Arr() As String
Dim i As Long
With dbPlayers(dbIndex)
    SplitFast .sAppearance, Arr, ":"
    Arr(WhichOne) = CStr(lNewV)
    .sAppearance = ""
    For i = LBound(Arr) To UBound(Arr)
        .sAppearance = .sAppearance & Arr(i) & ":"
        If DE Then DoEvents
    Next
    .sAppearance = Left$(.sAppearance, Len(.sAppearance) - 1)
End With
End Sub

Public Sub GetPlayerAppearance(dbIndex As Long, ByRef HairC As String, ByRef HairL As String, ByRef HairS As String, ByRef EyeC As String, ByRef MStyle As String, ByRef BStyle As String)
Dim Arr() As String
With dbPlayers(dbIndex)
    SplitFast .sAppearance, Arr, ":"
    HairL = HairLen(CLng(Arr(0)))
        
    HairC = ColorLst(CLng(Arr(1)))
        
    HairS = HairStyle(CLng(Arr(2)))
        
    EyeC = ColorLst(CLng(Arr(3)))
        
End With
End Sub
