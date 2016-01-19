Attribute VB_Name = "modTime"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modTime
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
'9 days a week
'Tiinaday
'Lyrarday
'Smeriday
'Nanolday
'Juionday
'Piaddday
'Thyaiday
'Layayday
'Nercyday
'8 months a year
'Dumleee
'Celeergo
'Fawrel
'Pilfle
'Beconied
'Nehieve
'Hoiiarni
'Moaltbay
'162 days a year
'24 hour days
'24:59:59
'4 seasons
'Spring
'Summer
'Winter
'Fall
Public TimeOfDay As String
Public DayOfMonth As Long
Public CurYear As Long
Public Season As Long
Public aDays(8) As String
Public aMonths(7) As String
Public Enum MonthOrder
    Dumleee = 0
    Celeergo = 1
    Fawrel = 2
    Pilfle = 3
    Beconied = 4
    Nehieve = 5
    Hoiiarni = 6
    Moaltbay = 7
End Enum
Public Enum DayOrder
    Tiinaday = 0
    Lyrarday = 1
    Smeriday = 2
    Nanolday = 3
    Juionday = 4
    Piaddday = 5
    Thyaiday = 6
    Layayday = 7
    Nercyday = 8
End Enum
Public Enum TimeMess
    SunRise = 0
    SunFull = 1
    SunSet = 2
    SunGone = 3
End Enum
Public Enum Season
    Spring = 0
    Summer = 1
    Fall = 2
    Winter = 3
End Enum
Private Type MessageSent
    SunRise As Boolean
    SunFull As Boolean
    SunSet As Boolean
    SunGone As Boolean
End Type
Public Type udtMonth
    MonthName As String
    DaysAMonth As Long
    CurDay As Long
End Type
Dim udtMS As MessageSent
Public udtDays As DayOrder
Public udtMonths(7) As udtMonth
Public MonthOfYear As MonthOrder
Public SeasonOfYear As Season
Private bSentMonth As Boolean
Private bSentDay As Boolean

Public Sub MoveTime()
modTime.AddTime 0, 0, 10
For i = LBound(dbEvents) To UBound(dbEvents)
    With dbEvents(i)
        If .lPlayerID <> 0 Then
            If modScripts.CheckDateDif("<=", Mid$(.sEndTime, InStr(1, .sEndTime, "/") + 1)) = True Then
                If modScripts.CheckTimeDif(">=", Mid$(.sEndTime, 1, InStr(1, .sEndTime, "/") - 1)) = True Then
                    .lIsComplete = 1
                End If
            End If
            If modScripts.CheckDateDif("<=", Mid$(.sExpire, InStr(1, .sExpire, "/") + 1)) = True Then
                If modScripts.CheckTimeDif("<=", Mid$(.sExpire, 1, InStr(1, .sExpire, "/") - 1)) = True Then
                    .lIsComplete = 0
                    .lPlayerID = 0
                    .sCustomID = "0"
                    .sEndTime = "0"
                    .sExpire = "0"
                    .sStartTime = "0"
                End If
            End If
        End If
    End With
Next
End Sub

Public Sub LoadMonths()
Dim i As Long
For i = 0 To 7
    With udtMonths(i)
        .MonthName = modTime.GetMonthName(i)
        .DaysAMonth = modTime.GetDaysInAMonth(i)
        .CurDay = 1
    End With
    If DE Then DoEvents
Next
End Sub

Sub AddTime(Hours As Long, Minutes As Long, Seconds As Long)
Dim s As String
Dim t As String
Dim u As String
s = Mid$(TimeOfDay, 1, 2)
s = CLng(s) + Hours
u = Mid$(TimeOfDay, 4, 2)
u = CLng(u) + Minutes
t = Mid$(TimeOfDay, 7, 2)
t = CLng(t) + Seconds
If CLng(t) >= 60 Then
    Do While CLng(t) >= 60
        u = CLng(u) + 1
        t = CLng(t) - 60
        If DE Then DoEvents
    Loop
End If
If CLng(u) >= 60 Then
    Do While CLng(u) >= 60
        s = CLng(s) + 1
        u = CLng(u) - 60
        If DE Then DoEvents
    Loop
End If
If CLng(s) > 24 Then
    modTime.ChangeDay udtDays
    Do While CLng(s) > 24
        s = CLng(s) - 24
        If DE Then DoEvents
    Loop
End If
If Len(t) < 2 Then t = "0" & t
If Len(u) < 2 Then u = "0" & u
If Len(s) < 2 Then s = "0" & s
TimeOfDay = s & ":" & u & ":" & t
Select Case CLng(Mid$(TimeOfDay, 1, 2))
    Case 6
        If Not udtMS.SunRise Then
            modTime.SendTimeMessage SunRise
            With udtMS
                .SunFull = False
                .SunGone = False
                .SunRise = True
                .SunSet = False
            End With
        End If
    Case 11
        If Not udtMS.SunFull Then
            modTime.SendTimeMessage SunFull
            With udtMS
                .SunFull = True
                .SunGone = False
                .SunRise = False
                .SunSet = False
            End With
        End If
    Case 18
        If Not udtMS.SunSet Then
            modTime.SendTimeMessage SunSet
            With udtMS
                .SunFull = False
                .SunGone = False
                .SunRise = False
                .SunSet = True
            End With
        End If
    Case 21
        If Not udtMS.SunGone Then
            modTime.SendTimeMessage SunGone
            With udtMS
                .SunFull = False
                .SunGone = True
                .SunRise = False
                .SunSet = False
            End With
        End If
End Select
End Sub

Public Function TimeSubs(Index As Long) As Boolean
If ShowTime(Index) = True Then TimeSubs = True: Exit Function
If ShowDate(Index) = True Then TimeSubs = True: Exit Function
If ShowMonth(Index) = True Then TimeSubs = True: Exit Function
If ShowYear(Index) = True Then TimeSubs = True: Exit Function
If ShowCalendar(Index) = True Then TimeSubs = True: Exit Function
End Function

Public Function ShowTime(Index As Long) As Boolean
Dim s As String
Dim m As String
Dim n As String
'Dim clsNtW As clsNumsToWords
If LCaseFast(X(Index)) = "time" Then
    ShowTime = True
    'Set clsNtW = New clsNumsToWords
    s = s & Mid$(TimeOfDay, 1, 2)
    If CLng(Mid$(TimeOfDay, 4, 2)) > 29 Then s = s & ":30"
    If Val(Left$(s, 2)) > 12 Then
        m = CStr(Val(Left$(s, 2)) - 12)
        If Len(m) < 2 Then m = "0" & m
        Mid$(s, 1, 2) = m
    End If
    m = LCaseFast(modNumsToWords.ConvertNumberToText(Mid$(s, 4, 2)))
    n = LCaseFast(modNumsToWords.ConvertNumberToText(Mid$(s, 1, 2)))
    s = n & " " & m & GREEN & modTime.GetDayNight
    If modSC.FastStringComp(Left$(s, 1), "0") Then Mid$(s, 1, 1) = " "
    's = s & m
    WrapAndSend Index, GREEN & "You determine the time is around " & LIGHTBLUE & s & GREEN & "." & WHITE & vbCrLf
    X(Index) = ""
End If
Set clsNtW = Nothing
End Function

Public Function ShowDate(Index As Long) As Boolean
If LCaseFast(X(Index)) = "day" Then
    ShowDate = True
    WrapAndSend Index, GREEN & "It is " & LIGHTBLUE & GetDayName(udtDays) & GREEN & "." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function ShowMonth(Index As Long) As Boolean
If LCaseFast(X(Index)) = "month" Then
    ShowMonth = True
    WrapAndSend Index, GREEN & "It is the " & LIGHTBLUE & udtMonths(MonthOfYear).CurDay & modTime.GetSuffix(CLng(udtMonths(MonthOfYear).CurDay)) & GREEN & " of the month " & LIGHTBLUE & udtMonths(MonthOfYear).MonthName & GREEN & "." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function ShowYear(Index As Long) As Boolean
If LCaseFast(X(Index)) = "year" Then
    ShowYear = True
    WrapAndSend Index, GREEN & "It is the year " & LIGHTBLUE & CStr(CurYear) & GREEN & "." & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function GetSuffix(lNum As Long) As String
If lNum < 10 Then
    Select Case lNum
        Case 4 To 9
            GetSuffix = "th"
        Case 2
            GetSuffix = "nd"
        Case 1
            GetSuffix = "st"
        Case 3
            GetSuffix = "rd"
    End Select
ElseIf lNum >= 10 And lNum < 19 Then
    GetSuffix = "th"
Else
    Select Case CLng(Right$(CStr(lNum), 1))
        Case 1
            GetSuffix = "st"
        Case 2
            GetSuffix = "nd"
        Case 3
            GetSuffix = "rd"
        Case Else
            GetSuffix = "th"
    End Select
End If
End Function

Public Function GetDayNight() As String
Dim s As String
Select Case CLng(Mid$(TimeOfDay, 1, 2))
    Case 1 To 7
        s = s & " early in the morning"
    Case 8 To 11
        s = s & " in the morning"
    Case 12 To 15
        s = s & " in the afternoon"
    Case 16 To 19
        s = s & " in the evening"
    Case 20 To 22
        s = s & " at night"
    Case 23 To 24
        s = s & " late at night"
End Select
GetDayNight = s
End Function

Sub WriteTimeToFile()
WriteINI "TimeOfDay", TimeOfDay
End Sub

Sub WriteYearToFile()
WriteINI "Year", CStr(CurYear)
End Sub

Sub WriteDayToFile()
WriteINI "DayOfWeek", CStr(udtDays)
End Sub

Sub WriteMonthOfYearToFile()
WriteINI "MonthOfYear", CStr(MonthOfYear)
WriteINI "DayOfMonth", CStr(DayOfMonth)
End Sub

Sub SetTimeOfDay()
Dim s As String
s = GetINI("TimeOfDay")
If s = "Error" Then s = "24:59:59"
TimeOfDay = s
End Sub

Sub SetYear()
Dim s As String
s = GetINI("Year")
If s = "Error" Then s = "1257"
CurYear = CLng(s)
End Sub

Sub SetMonthOfYear()
Dim s As String
s = GetINI("MonthOfYear")
If s = "Error" Then s = "0"
MonthOfYear = CLng(s)
s = GetINI("DayOfMonth")
If s = "Error" Then s = "1"
DayOfMonth = CLng(s)
udtMonths(MonthOfYear).CurDay = DayOfMonth
End Sub

Sub SetDayOfWeek()
Dim s As String
s = GetINI("DayOfWeek")
If s = "Error" Then s = "0"
udtDays = CLng(s)
End Sub

Public Sub SetNameArrays()
Dim f As Long
Dim s As String
Dim Arr() As String
Dim i As Long
Dim j As Long
f = FreeFile
Open App.Path & "\days.aimg" For Binary As #f
    s = Input$(LOF(f), f)
Close #f
SplitFast s, Arr, vbCrLf
j = 0
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" And j < 9 Then
        aDays(j) = Arr(i)
        j = j + 1
    End If
Next
If j < 9 Then
    For i = j To 8
        aDays(j) = "NO NAME GIVEN"
    Next
End If

f = FreeFile
Open App.Path & "\months.aimg" For Binary As #f
    s = Input$(LOF(f), f)
Close #f
SplitFast s, Arr, vbCrLf
j = 0
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" And j < 7 Then
        aMonths(j) = Arr(i)
        j = j + 1
    End If
Next
If j < 7 Then
    For i = j To 7
        aMonths(j) = "NO NAME GIVEN"
    Next
End If
End Sub

Sub SendTimeMessage(Choose As TimeMess)
Select Case Choose
    Case 0
        SendToAll YELLOW & "The sun begins to rise." & WHITE & vbCrLf
        ChangeMons True
    Case 1
        SendToAll BRIGHTYELLOW & "The sun is full over the land." & WHITE & vbCrLf
    Case 2
        SendToAll LIGHTBLUE & "The sun begins to set." & WHITE & vbCrLf
        ChangeMons False
    Case 3
        SendToAll BLUE & "The sun has set." & WHITE & vbCrLf
End Select
End Sub

Sub ChangeMons(bDayTime As Boolean)
Dim i As Long
   'On Error GoTo ChangeMons_Error

If bDayTime Then
    For i = LBound(aMons) To UBound(aMons)
        If aMons(i).mAtDayMonster <> 0 And (aMons(i).mLoc <> 0 And aMons(i).mLoc <> -1) Then
            With dbMonsters(GetMonsterID(, aMons(i).mAtDayMonster))
                If (Not aMons(i).mIs_Being_Attacked) And (Not aMons(i).mIsAttacking) Then
                    SendToAllInRoom 0, BRIGHTRED & aMons(i).mName & YELLOW & " changes into a " & BRIGHTRED & .sMonsterName & YELLOW & " as the sun rises!" & WHITE & vbCrLf, aMons(i).mLoc
                    aMons(i).mName = .sMonsterName
                    aMons(i).mHP = .dHP
                    aMons(i).mMessage = .sMessage
                    aMons(i).mAc = .iAC
                    aMons(i).mEXP = .dEXP
                    aMons(i).mMin = Val(Mid$(.sAttack, 1, InStr(1, .sAttack, ":")))  'its min damage
                    aMons(i).mMax = Val(Mid$(.sAttack, InStr(1, .sAttack, ":") + 1, Len(.sAttack) - (InStr(1, .sAttack, ":") - 1))) 'its max damage
                    aMons(i).mDesc = .sDesc
                    If .lWeapon <> 0 Then aMons(i).mWeapon = dbItems(GetItemID(, .lWeapon))
                    aMons(i).mEnergy = .lEnergy
                    aMons(i).mPEnergy = .lPEnergy
                    SetUpAmonSpells i, .sSpells
                    aMons(i).mLevel = .lLevel
                    aMons(i).mMoney = .dMoney
                    aMons(i).mDeathText = .sDeathText
                    aMons(i).mHostile = IIf(.iHostile = 1, True, False)
                    aMons(i).mAttackable = IIf(.iAttackable = 0, True, False)
                    aMons(i).mRoams = .iRoams
                    aMons(i).mDontAttackIfItem = .iDontAttackIfItem
                    aMons(i).mMaxHP = .dHP
                    aMons(i).miID = .lID
                    
                End If
            End With
        End If
        If DE Then DoEvents
    Next
Else
    For i = LBound(aMons) To UBound(aMons)
        If aMons(i).mAtNightMonster <> 0 And (aMons(i).mLoc <> 0 And aMons(i).mLoc <> -1) Then
            With dbMonsters(GetMonsterID(, aMons(i).mAtNightMonster))
                If (Not aMons(i).mIs_Being_Attacked) And (Not aMons(i).mIsAttacking) Then
                    SendToAllInRoom 0, BRIGHTRED & aMons(i).mName & YELLOW & " changes into a " & BRIGHTRED & .sMonsterName & YELLOW & " as the sun sets!" & WHITE & vbCrLf, aMons(i).mLoc
                    aMons(i).mName = .sMonsterName
                    aMons(i).mHP = .dHP
                    aMons(i).mMessage = .sMessage
                    aMons(i).mAc = .iAC
                    aMons(i).mEXP = .dEXP
                    aMons(i).mMin = Val(Mid$(.sAttack, 1, InStr(1, .sAttack, ":")))  'its min damage
                    aMons(i).mMax = Val(Mid$(.sAttack, InStr(1, .sAttack, ":") + 1, Len(.sAttack) - (InStr(1, .sAttack, ":") - 1))) 'its max damage
                    aMons(i).mDesc = .sDesc
                    aMons(i).mEnergy = .lEnergy
                    aMons(i).mPEnergy = .lPEnergy
                    If .lWeapon <> 0 Then aMons(i).mWeapon = dbItems(GetItemID(, .lWeapon))
                    SetUpAmonSpells i, .sSpells
                    aMons(i).mLevel = .lLevel
                    aMons(i).mMoney = .dMoney
                    aMons(i).mDeathText = .sDeathText
                    aMons(i).mHostile = IIf(.iHostile = 1, True, False)
                    aMons(i).mAttackable = IIf(.iAttackable = 0, True, False)
                    aMons(i).mRoams = .iRoams
                    aMons(i).mDontAttackIfItem = .iDontAttackIfItem
                    aMons(i).mMaxHP = .dHP
                    aMons(i).miID = .lID
                End If
            End With
        End If
        If DE Then DoEvents
    Next
End If

   On Error GoTo 0
   Exit Sub

ChangeMons_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ChangeMons of Module modTime"
End Sub

Sub ChangeDay(ToDay As DayOrder)
With udtMonths(MonthOfYear)
    .CurDay = .CurDay + 1
    If .CurDay > .DaysAMonth Then
        bSentMonth = False
        ChangeMonth MonthOfYear 'GetMonthIndex(.MonthName)
        .CurDay = 1
    End If
End With
If ToDay + 1 <= 8 Then
    bSentDay = False
    udtDays = ToDay + 1
Else
    bSentDay = False
    udtDays = 0
End If
CheckBDays
If Not bSentDay Then
    SendToAll GREEN & "It is now " & LIGHTBLUE & GetDayName(udtDays) & GREEN & "." & WHITE & vbCrLf
    bSentDay = True
End If
End Sub

Sub CheckBDays()
Dim i As Long
Dim m As Long
Dim n As Long
Dim lMonth As Long
Dim lDay As Long
For i = LBound(dbPlayers) To UBound(dbPlayers)
    With dbPlayers(i)
        m = InStr(1, .sBirthDay, "/")
        If m <> 0 Then
            lMonth = CLng(Left$(.sBirthDay, m - 1))
            n = InStr(m + 1, .sBirthDay, "/")
            lDay = CLng(Mid$(.sBirthDay, m + 1, n - m - 1))
            If udtMonths(MonthOfYear).CurDay = lDay Then
                If MonthOfYear = lMonth - 1 Then
                    .lAge = .lAge + 1
                    SendToAll GREEN & "It is " & LIGHTBLUE & .sPlayerName & "'s" & GREEN & " birthday today!" & WHITE & vbCrLf
                    If .lAge = dbRaces(GetRaceID(.sRace)).lMaxAge Then
                        If lAgeD = 2 Then
                            .iLives = 0
                            .lHP = lDeath - 100
                            CheckDeath .iIndex
                            WrapAndSend .iIndex, BRIGHTRED & "You have a heart attack!" & WHITE & vbCrLf
                        ElseIf lAgeD = 1 Then
                            .iStr = .iStr - RndNumber(0, 1)
                            .iAgil = .iAgil - RndNumber(0, 1)
                            .iDex = .iDex - RndNumber(0, 1)
                            .iInt = .iInt - RndNumber(0, 1)
                            .iCha = .iCha - RndNumber(0, 2)
                            .lMaxHP = .lMaxHP - RndNumber(0, 1)
                            If .lMaxHP < 50 Then .lMaxHP = 50
                            If .lHP > .lMaxHP Then .lHP = .lMaxHP
                            WrapAndSend .iIndex, BRIGHTRED & "You feel weaker!" & WHITE & vbCrLf
                        End If
                    End If
                    
                End If
            End If
        End If
    End With
    If DE Then DoEvents
Next
End Sub

Sub ChangeMonth(CurMonth As MonthOrder)
If CurMonth + 1 <= 7 Then
    CurMonth = CurMonth + 1
    MonthOfYear = CurMonth
Else
    CurMonth = 0
    MonthOfYear = CurMonth
    CurYear = CurYear + 1
End If
Select Case MonthOfYear
    Case 0, 1
        SeasonOfYear = Spring
    Case 2, 3
        SeasonOfYear = Summer
    Case 4, 5
        SeasonOfYear = Fall
    Case 6, 7
        SeasonOfYear = Winter
End Select
If Not bSentMonth Then
    SendToAll GREEN & "You are now in the month of " & LIGHTBLUE & modTime.GetMonthName(CurMonth) & GREEN & "."
    bSentMonth = True
End If
End Sub

Public Function GetDayName(CurDay As DayOrder) As String
'Tiinaday
'Lyrarday
'Smeriday
'Nanolday
'Juionday
'Piaddday
'Thyaiday
'Layayday
'Nercyday
GetDayName = aDays(CurDay)
'Select Case CurDay
'    Case 0
'        GetDayName = "Tiinaday"
'    Case 1
'        GetDayName = "Lyrarday"
'    Case 2
'        GetDayName = "Smeriday"
'    Case 3
'        GetDayName = "Nanolday"
'    Case 4
'        GetDayName = "Juionday"
'    Case 5
'        GetDayName = "Piaddday"
'    Case 6
'        GetDayName = "Thyaiday"
'    Case 7
'        GetDayName = "Layayday"
'    Case 8
'        GetDayName = "Nercyday"
'End Select
End Function

Public Function GetMonthIndex(CurMonth As String) As String
'Dumleee
'Celeergo
'Fawrel
'Pilfle
'Beconied
'Nehieve
'Hoiiarni
'Moaltbay
Select Case CurMonth
    Case "Dumleee"
        GetMonthIndex = 0
    Case "Celeergo"
        GetMonthIndex = 1
    Case "Fawrel"
        GetMonthIndex = 2
    Case "Pilfle"
        GetMonthIndex = 3
    Case "Beconied"
        GetMonthIndex = 4
    Case "Nehieve"
        GetMonthIndex = 5
    Case "Hoiiarni"
        GetMonthIndex = 6
    Case "Moaltbay"
        GetMonthIndex = 7
End Select
End Function

Public Function GetMonthName(CurMonth As MonthOrder) As String
'Dumleee
'Celeergo
'Fawrel
'Pilfle
'Beconied
'Nehieve
'Hoiiarni
'Moaltbay
GetMonthName = aMonths(CurMonth)
'Select Case CurMonth
'    Case 0
'        GetMonthName = "Dumleee"
'    Case 1
'        GetMonthName = "Celeergo"
'    Case 2
'        GetMonthName = "Fawrel"
'    Case 3
'        GetMonthName = "Pilfle"
'    Case 4
'        GetMonthName = "Beconied"
'    Case 5
'        GetMonthName = "Nehieve"
'    Case 6
'        GetMonthName = "Hoiiarni"
'    Case 7
'        GetMonthName = "Moaltbay"
'End Select
End Function

Public Function GetDaysInAMonth(CurMonth As MonthOrder) As Long
'Dumleee
'Celeergo
'Fawrel
'Pilfle
'Beconied
'Nehieve
'Hoiiarni
'Moaltbay
Select Case CurMonth
    Case 0
        GetDaysInAMonth = 18
    Case 1
        GetDaysInAMonth = 27
    Case 2
        GetDaysInAMonth = 9
    Case 3
        GetDaysInAMonth = 36
    Case 4
        GetDaysInAMonth = 18
    Case 5
        GetDaysInAMonth = 9
    Case 6
        GetDaysInAMonth = 18
    Case 7
        GetDaysInAMonth = 27
End Select
End Function

Public Function ShowCalendar(Index As Long) As Boolean
Dim s As String
Dim sM As String
Dim i As Long
Dim d As Long
Dim j As Long
Dim m As Long
Dim b As Boolean
Dim k As Long
Dim z As Long
If LCaseFast(X(Index)) = "calendar" Then
    ShowCalendar = True
    s = WHITE & "ÉÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ»" & vbCrLf & "º "
    sM = modTime.udtMonths(modTime.MonthOfYear).MonthName
    i = 26 - Len(sM)
    If (i And 1) = 1 Then
        s = s & Space$(i \ 2) & sM & Space$((i \ 2) + 1) & WHITE & " º"
    Else
        s = s & Space$(i \ 2) & sM & Space$(i \ 2) & WHITE & " º"
    End If
    s = s & vbCrLf & "º "
    For z = 0 To 8
        s = s & UCase$(Left$(aDays(z), 1)) & "  "
        If DE Then DoEvents
    Next
    s = s & "º" & vbCrLf & "ÇÄÄÄÂÄÄÂÄÄÂÄÄÂÄÄÂÄÄÂÄÄÂÄÄÂÄÄÄ¶" & vbCrLf & "º "
    d = modTime.GetDaysInAMonth(modTime.MonthOfYear) + udtDays + 1
    j = 1
    k = 0
    For i = 1 To d
        If k <= udtDays And Not b Then
            s = s & "  " & WHITE & "³"
            
        Else
            If Not b Then k = 1
            b = True
            If k = modTime.udtMonths(modTime.MonthOfYear).CurDay Then s = s & BRIGHTRED
            If Len(CStr(k)) = 1 Then
                s = s & "0" & CStr(k) & WHITE & "³"
            Else
                s = s & CStr(k) & WHITE & "³"
            End If
        End If
        If j = 9 And i <> d Then
            s = Left$(s, Len(s) - 1) & " º" & vbCrLf & "ÇÄÄÄÅÄÄÅÄÄÅÄÄÅÄÄÅÄÄÅÄÄÅÄÄÅÄÄÄ¶" & vbCrLf & "º "
            j = 0
        ElseIf i = d And j <> 9 Then
            For m = 1 To (9 - j)
                s = s & "  ³"
                If DE Then DoEvents
            Next
            s = Left$(s, Len(s) - 1) & " º" & vbCrLf
        ElseIf j = 9 And i = d Then
            s = Left$(s, Len(s) - 1) & " º" & vbCrLf
            
        End If
        j = j + 1
        k = k + 1
        If DE Then DoEvents
    Next
    s = s & "ÇÄÄÄÁÄÄÁÄÄÁÄÄÁÄÄÁÄÄÁÄÄÁÄÄÁÄÄÄ¶" & vbCrLf & "ÈÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼"
    WrapAndSend Index, s & WHITE & vbCrLf
    X(Index) = ""
End If
End Function

Public Function AddTimeNotReal(Hours As Long, Minutes As Long, Seconds As Long, Months As Long, Days As Long, Years As Long) As String
Dim s As String
Dim t As String
Dim u As String
Dim R As Long
Dim v As Long
Dim w As Long
s = Mid$(TimeOfDay, 1, 2)
s = CLng(s) + Hours
u = Mid$(TimeOfDay, 4, 2)
u = CLng(u) + Minutes
t = Mid$(TimeOfDay, 7, 2)
t = CLng(t) + Seconds
R = MonthOfYear
R = R + Months
v = udtMonths(MonthOfYear).CurDay
v = v + Days
w = CurYear
w = w + Years
If CLng(t) >= 60 Then
    Do While CLng(t) >= 60
        u = CLng(u) + 1
        t = CLng(t) - 60
        If DE Then DoEvents
    Loop
End If
If CLng(u) >= 60 Then
    Do While CLng(u) >= 60
        s = CLng(s) + 1
        u = CLng(u) - 60
        If DE Then DoEvents
    Loop
End If
If CLng(s) > 24 Then
    Do While CLng(s) > 24
        s = CLng(s) - 24
        v = v + 1
        If DE Then DoEvents
    Loop
End If
If v > udtMonths(MonthOfYear).DaysAMonth Then
    Do While v > udtMonths(MonthOfYear).DaysAMonth
        v = v - udtMonths(MonthOfYear).DaysAMonth
        R = R + 1
        If DE Then DoEvents
    Loop
End If
If R > 8 Then
    Do While R > 8
        R = R - 8
        w = w + 1
        If DE Then DoEvents
    Loop
End If
        
If Len(t) < 2 Then t = "0" & t
If Len(u) < 2 Then u = "0" & u
If Len(s) < 2 Then s = "0" & s
AddTimeNotReal = s & ":" & u & ":" & t & "/" & CStr(R) & ":" & CStr(v) & ":" & CStr(w)

End Function
