Attribute VB_Name = "modDesc"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modDesc
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function sDescription(Index As Long) As String
Dim ToSend As String
Dim dbIndex As Long
Dim Arr(5) As String
dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    modAppearance.GetPlayerAppearance dbIndex, Arr(1), Arr(0), Arr(2), Arr(3), Arr(4), Arr(5)
    ToSend = BRIGHTWHITE & .sPlayerName & " is a " & modEvil.GetRep(Index) & BRIGHTWHITE
    ToSend = ToSend & " of the small village Klepta. "
    ToSend = ToSend & .sPlayerName & " is from " & .sRace & " decent, " & "and seems to be trained as a" & modgetdata.GetClassPointLevel(dbIndex) & " " & BRIGHTWHITE & .sClass
    ToSend = ToSend & ". " & modgetdata.GetGenderPronoun(dbIndex, True, True) & " " & Arr(1) & ", " & Arr(2) & " hair is " & Arr(0) & ". " & .sPlayerName & "'s eyes are " & Arr(3) & ". "
    ToSend = ToSend & modgetdata.GetGenderDesc(dbIndex) & " appears to be "
    Select Case .iStr
        Case Is < 1
            ToSend = ToSend & "able-bodied and "
        Case 1 To 10
            ToSend = ToSend & "frail and "
        Case 11 To 20
            ToSend = ToSend & "feeble and "
        Case 21 To 30
            ToSend = ToSend & "weak and "
        Case 31 To 40
            ToSend = ToSend & "stout and "
        Case 41 To 50
            ToSend = ToSend & "hardy and "
        Case 51 To 60
            ToSend = ToSend & "vigorous and "
        Case 61 To 85
            ToSend = ToSend & "mighty and "
        Case 86 To 99
            ToSend = ToSend & "a goliath and "
        Case Is >= 100
            ToSend = ToSend & "godlike and "
    End Select
    Select Case .iDex
        Case Is < 1
            ToSend = ToSend & "motionless. "
        Case 1 To 10
            ToSend = ToSend & "lazy. "
        Case 11 To 20
            ToSend = ToSend & "brisk. "
        Case 21 To 30
            ToSend = ToSend & "active. "
        Case 31 To 40
            ToSend = ToSend & "energetic. "
        Case 41 To 50
            ToSend = ToSend & "intense. "
        Case 51 To 60
            ToSend = ToSend & "proficient. "
        Case 61 To 85
            ToSend = ToSend & "skilled. "
        Case 86 To 99
            ToSend = ToSend & "potent. "
        Case Is >= 100
            ToSend = ToSend & "a master. "
    End Select
    ToSend = ToSend & "While "
    Select Case .iCha
        Case Is < 1
            ToSend = ToSend & "being butt-ugly, "
        Case 1 To 10
            ToSend = ToSend & "being repulsive, "
        Case 11 To 20
            ToSend = ToSend & "being unsightly, "
        Case 21 To 30
            ToSend = ToSend & "seemily unexpecitdly being charming, "
        Case 31 To 40
            ToSend = ToSend & "looking quite average, "
        Case 41 To 50
            ToSend = ToSend & "acting very charming, "
        Case 51 To 60
            ToSend = ToSend & "being good-looking, "
        Case 61 To 85
            ToSend = ToSend & "looking gorgeous, "
        Case 86 To 99
            ToSend = ToSend & "being extremly attractive, "
        Case 100 To 175
            ToSend = ToSend & "looking like a god, "
        Case Is > 175
            ToSend = ToSend & "making you very horny, "
    End Select
    ToSend = ToSend & .sPlayerName & " "
    Select Case .iInt
        Case Is < 1
            ToSend = ToSend & "is just an oaf. "
        Case 1 To 10
            ToSend = ToSend & "is just a simpleton. "
        Case 11 To 20
            ToSend = ToSend & "is quite witty. "
        Case 21 To 30
            ToSend = ToSend & "is rather reasonable. "
        Case 31 To 40
            ToSend = ToSend & "seems to be quite cunning. "
        Case 41 To 50
            ToSend = ToSend & "looks to be rather deep. "
        Case 51 To 60
            ToSend = ToSend & "looks impartial on most things. "
        Case 61 To 85
            ToSend = ToSend & "seems very strong-minded. "
        Case 86 To 99
            ToSend = ToSend & "looks to be a great inspirationist. "
        Case Is >= 100
            ToSend = ToSend & "is a genius. "
    End Select
    ToSend = ToSend & modgetdata.GetGenderDesc(dbIndex) & " seems to be "
    Select Case .iAgil
        Case Is < 1
            ToSend = ToSend & "very non-reactive."
        Case 1 To 10
            ToSend = ToSend & "quite sluggish."
        Case 11 To 20
            ToSend = ToSend & "slighty active."
        Case 21 To 30
            ToSend = ToSend & "full of pep."
        Case 31 To 40
            ToSend = ToSend & "very active."
        Case 41 To 50
            ToSend = ToSend & "lively and excited."
        Case 51 To 60
            ToSend = ToSend & "animated and well balanced."
        Case 61 To 85
            ToSend = ToSend & "light footed and acrobatic."
        Case 86 To 99
            ToSend = ToSend & "acrobatic and extremly agile."
        Case Is >= 100
            ToSend = ToSend & "god-like in all aspects of movement."
    End Select
    
    sDescription = ToSend
End With
End Function
