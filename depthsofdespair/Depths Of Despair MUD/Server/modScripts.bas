Attribute VB_Name = "modScripts"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modScripts
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
 
Function DoIF(Index As Long, sIf As String, sCheck As String) As Boolean
Dim m As Long
Dim n As Long
Dim s As String
Dim b As String
Dim iItemID As Long
Dim dbIndex As Long
Dim Arr() As String
Dim i As Long
Dim Checking As String
Dim Comparing As String
Dim sSign As String
Dim rndNumber1 As Double
Dim rndNumber2 As Double
Dim theRndNumber As Long
If Index = 0 Then Exit Function
s = LCaseFast(TrimIt(X(Index)))
dbIndex = GetPlayerIndexNumber(Index)
b = sCheck
Select Case TrimIt(sIf)
    Case "message="
        If modSC.FastStringComp(s, b) Then
            DoIF = True
            Exit Function
        End If
    Case "in("
        If InStr(1, b, s) Then
            DoIF = True
            Exit Function
        End If
    Case "hasitems"
        SplitFast b, Arr, ","
        For i = LBound(Arr) To UBound(Arr)
            If Not IsNumeric(Arr(i)) Then
                iItemID = GetItemID(Arr(i))
            Else
                iItemID = GetItemID(, Val(Arr(i)))
            End If
            If iItemID <> 0 Then
                If InStr(1, dbMap(GetMapIndex(dbPlayers(dbIndex).lLocation)).sItems, ":" & CStr(iItemID) & "/") Then
                    DoIF = True
                Else
                    DoIF = False
                    Exit Function
                End If
            End If
            If DE Then DoEvents
        Next
    Case "isdoorlocked"
        SplitFast b, Arr, ","
        iItemID = GetMapIndex(Val(Arr(0)))
        If iItemID <> 0 Then
            With dbMap(iItemID)
                Select Case LCaseFast(Arr(1))
                    Case "direction.north"
                        If .lDN = 2 Then DoIF = True: Exit Function
                    Case "direction.south"
                        If .lDS = 2 Then DoIF = True: Exit Function
                    Case "direction.east"
                        If .lDE = 2 Then DoIF = True: Exit Function
                    Case "direction.west"
                        If .lDW = 2 Then DoIF = True: Exit Function
                    Case "direction.northwest"
                        If .lDNW = 2 Then DoIF = True: Exit Function
                    Case "direction.northeast"
                        If .lDNE = 2 Then DoIF = True: Exit Function
                    Case "direction.southwest"
                        If .lDSW = 2 Then DoIF = True: Exit Function
                    Case "direction.southeast"
                        If .lDSE = 2 Then DoIF = True: Exit Function
                    Case "direction.up"
                        If .lDU = 2 Then DoIF = True: Exit Function
                    Case "direction.down"
                        If .lDD = 2 Then DoIF = True: Exit Function
                End Select
            End With
        End If
    Case "isdoorclosed"
        SplitFast b, Arr, ","
        iItemID = GetMapIndex(Val(Arr(0)))
        If iItemID <> 0 Then
            With dbMap(iItemID)
                Select Case LCaseFast(Arr(1))
                    Case "direction.north"
                        If .lDN = 1 Then DoIF = True: Exit Function
                    Case "direction.south"
                        If .lDS = 1 Then DoIF = True: Exit Function
                    Case "direction.east"
                        If .lDE = 1 Then DoIF = True: Exit Function
                    Case "direction.west"
                        If .lDW = 1 Then DoIF = True: Exit Function
                    Case "direction.northwest"
                        If .lDNW = 1 Then DoIF = True: Exit Function
                    Case "direction.northeast"
                        If .lDNE = 1 Then DoIF = True: Exit Function
                    Case "direction.southwest"
                        If .lDSW = 1 Then DoIF = True: Exit Function
                    Case "direction.southeast"
                        If .lDSE = 1 Then DoIF = True: Exit Function
                    Case "direction.up"
                        If .lDU = 1 Then DoIF = True: Exit Function
                    Case "direction.down"
                        If .lDD = 1 Then DoIF = True: Exit Function
                End Select
            End With
        End If
    Case "isdooropen"
        SplitFast b, Arr, ","
        iItemID = GetMapIndex(Val(Arr(0)))
        If iItemID <> 0 Then
            With dbMap(iItemID)
                Select Case LCaseFast(Arr(1))
                    Case "direction.north"
                        If .lDN = 3 Then DoIF = True: Exit Function
                    Case "direction.south"
                        If .lDS = 3 Then DoIF = True: Exit Function
                    Case "direction.east"
                        If .lDE = 3 Then DoIF = True: Exit Function
                    Case "direction.west"
                        If .lDW = 3 Then DoIF = True: Exit Function
                    Case "direction.northwest"
                        If .lDNW = 3 Then DoIF = True: Exit Function
                    Case "direction.northeast"
                        If .lDNE = 3 Then DoIF = True: Exit Function
                    Case "direction.southwest"
                        If .lDSW = 3 Then DoIF = True: Exit Function
                    Case "direction.southeast"
                        If .lDSE = 3 Then DoIF = True: Exit Function
                    Case "direction.up"
                        If .lDU = 3 Then DoIF = True: Exit Function
                    Case "direction.down"
                        If .lDD = 3 Then DoIF = True: Exit Function
                End Select
            End With
        End If
    Case "iteminroom"
        If Not IsNumeric(b) Then
            iItemID = GetItemID(b)
        Else
            iItemID = GetItemID(, Val(b))
        End If
        If iItemID = 0 Then Exit Function
        iItemID = dbItems(iItemID).iID
        If InStr(1, dbMap(GetMapIndex(dbPlayers(dbIndex).lLocation)).sItems, ":" & CStr(iItemID) & "/") Then
            DoIF = True
            Exit Function
        End If
    Case "moninroom"
        If Not IsNumeric(b) Then
            iItemID = GetMonsterID(b)
        Else
            iItemID = GetMonsterID(, Val(b))
        End If
        If iItemID = 0 Then Exit Function
        iItemID = dbMonsters(iItemID).lID
        s = modGetData.GetAllMonstersInRoomMONID(dbPlayers(dbIndex).lDBLocation)
        If InStr(1, s, ":" & CStr(iItemID) & "/") Then
            DoIF = True
            Exit Function
        End If
    Case "isplayeralone"
        s = modGetData.GetAllMonstersInRoom(dbPlayers(dbIndex).lLocation, dbPlayers(dbIndex).lDBLocation)
        s = s & modGetData.GetPlayersHere(dbPlayers(dbIndex).lLocation, dbIndex)
        If s = "" Then
            DoIF = True
            Exit Function
        End If
    Case "event.find"
        m = GetEventID(sCheck, dbPlayers(dbIndex).lPlayerID)
        If m <> 0 Then
            DoIF = True
            Exit Function
        End If
    Case "event.finished"
        m = GetEventID(sCheck, dbPlayers(dbIndex).lPlayerID)
        If dbEvents(m).lIsComplete <> 0 Then
            DoIF = True
            Exit Function
        End If
    Case "event.notthere"
        If GetEventID(sCheck, dbPlayers(dbIndex).lPlayerID) = 0 Then
            DoIF = True
            Exit Function
        End If
    Case "event.hasstarted"
        m = GetEventID(sCheck, dbPlayers(dbIndex).lPlayerID)
        If m <> 0 Then
            If dbEvents(m).lIsComplete = 0 Then
                DoIF = True
                Exit Function
            End If
        End If
    Case "isadirection"
        If IsADirection(LCaseFast(X(Index))) = True Then DoIF = True
    Case "isdir"
        If "direction." & WhichDirIsIt(LCaseFast(X(Index))) = b Then DoIF = True
    Case "haveitem"
        If Not IsNumeric(b) Then
            iItemID = GetItemID(b)
        Else
            iItemID = GetItemID(, Val(b))
        End If
        If iItemID = 0 Then Exit Function
        iItemID = dbItems(iItemID).iID
        If InStr(1, dbPlayers(dbIndex).sInventory, ":" & CStr(iItemID) & "/") Then
            DoIF = True
            Exit Function
        End If
    Case "donthaveitem"
        If Not IsNumeric(b) Then
            iItemID = GetItemID(b)
        Else
            iItemID = GetItemID(, Val(b))
        End If
        If iItemID = 0 Then Exit Function
        iItemID = dbItems(iItemID).iID
        If InStr(1, dbPlayers(dbIndex).sInventory, ":" & CStr(iItemID) & "/") = 0 Then
            DoIF = True
            Exit Function
        End If
    Case "hasitemanywhere"
        If Not IsNumeric(b) Then
            iItemID = GetItemID(b)
        Else
            iItemID = GetItemID(, Val(b))
        End If
        If iItemID = 0 Then Exit Function
        iItemID = dbItems(iItemID).iID
        If InStr(1, dbPlayers(dbIndex).sInventory & modGetData.GetPlayersEq(Index), ":" & CStr(iItemID) & "/") <> 0 Then
            DoIF = True
            Exit Function
        End If
    Case "doesnthaveitemanywhere"
        If Not IsNumeric(b) Then
            iItemID = GetItemID(b)
        Else
            iItemID = GetItemID(, Val(b))
        End If
        If iItemID = 0 Then Exit Function
        iItemID = dbItems(iItemID).iID
        If InStr(1, dbPlayers(dbIndex).sInventory & modGetData.GetPlayersEq(Index), ":" & CStr(iItemID) & "/") = 0 Then
            DoIF = True
            Exit Function
        End If
    Case "hasequiped"
        If Not IsNumeric(b) Then
            iItemID = GetItemID(b)
        Else
            iItemID = GetItemID(, Val(b))
        End If
        If iItemID = 0 Then Exit Function
        iItemID = dbItems(iItemID).iID
        If InStr(1, modGetData.GetPlayersEq(Index), ":" & CStr(iItemID) & "/") <> 0 Then
            DoIF = True
            Exit Function
        End If
    Case "doesnthaveequiped"
        If Not IsNumeric(b) Then
            iItemID = GetItemID(b)
        Else
            iItemID = GetItemID(, Val(b))
        End If
        If iItemID = 0 Then Exit Function
        iItemID = dbItems(iItemID).iID
        If InStr(1, modGetData.GetPlayersEq(Index), ":" & CStr(iItemID) & "/") = 0 Then
            DoIF = True
            Exit Function
        End If
    Case "class"
        If Not IsNumeric(b) Then
            If modSC.FastStringComp(LCaseFast(dbPlayers(dbIndex).sClass), b) Then
                DoIF = True
                Exit Function
            End If
        Else
            If modSC.FastStringComp(LCaseFast(dbPlayers(dbIndex).sClass), modGetData.GetClassFromNum(CLng(b))) Then
                DoIF = True
                Exit Function
            End If
        End If
    Case "classes"
        SplitFast b, Arr, ","
        'Classes(*mage*,*warrior*,*thief*)
        For i = LBound(Arr) To UBound(Arr)
            If Not IsNumeric(Arr(i)) Then
                If modSC.FastStringComp(LCaseFast(dbPlayers(dbIndex).sClass), Arr(i)) Then
                    DoIF = True
                    Exit Function
                End If
            Else
                If modSC.FastStringComp(LCaseFast(dbPlayers(dbIndex).sClass), modGetData.GetClassFromNum(CLng(Arr(i)))) Then
                    DoIF = True
                    Exit Function
                End If
            End If
        Next
    Case "items"
        n = 0
        SplitFast b, Arr, ","
        For i = LBound(Arr) To UBound(Arr)
            If Not IsNumeric(Arr(i)) Then
                m = GetItemID(Arr(i))
                If m = 0 Then DoIF = False: Exit Function
                If InStr(1, dbPlayers(dbIndex).sInventory, ":" & dbItems(m).iID & "/") <> 0 Then n = n + 1
            Else
                m = GetItemID(, Val(Arr(i)))
                If m = 0 Then DoIF = False: Exit Function
                If InStr(1, dbPlayers(dbIndex).sInventory, ":" & dbItems(m).iID & "/") <> 0 Then n = n + 1
            End If
        Next
        If n > UBound(Arr) Then DoIF = True
    Case "race"
        If Not IsNumeric(b) Then
            If modSC.FastStringComp(LCaseFast(dbPlayers(dbIndex).sRace), b) Then
                DoIF = True
                Exit Function
            End If
        Else
            If modSC.FastStringComp(LCaseFast(dbPlayers(dbIndex).sRace), modGetData.GetRaceFromNum(CLng(b))) Then
                DoIF = True
                Exit Function
            End If
        End If
    Case "races"
        SplitFast b, Arr, ","
        For i = LBound(Arr) To UBound(Arr)
            If Not IsNumeric(Arr(i)) Then
                If modSC.FastStringComp(LCaseFast(dbPlayers(dbIndex).sRace), Arr(i)) Then
                    DoIF = True
                    Exit Function
                End If
            Else
                If modSC.FastStringComp(LCaseFast(dbPlayers(dbIndex).sRace), modGetData.GetRaceFromNum(CLng(Arr(i)))) Then
                    DoIF = True
                    Exit Function
                End If
            End If
        Next
    Case "statcheck"
        'StatCheck(stat.cha,signs.>,7)
        m = InStr(1, b, ".")
        n = InStr(m + 1, b, ",")
        Checking = Mid$(b, m + 1, n - m - 1)
        m = InStr(1, b, ",")
        m = InStr(m, b, ".")
        n = InStr(m + 1, b, ",")
        sSign = Mid$(b, m + 1, n - m - 1)
        Comparing = Mid$(b, n + 1)
        If CheckTheStat(Index, Checking, CDbl(Comparing), sSign) = True Then
            DoIF = True
            Exit Function
        End If
    Case "rnd"
        'if rnd(1,3,signs.>,4):onfail(**)
        m = InStr(1, b, "(")
        n = InStr(m + 1, b, ",")
        rndNumber1 = CDbl(Mid$(b, m + 1, n - m - 1))
        m = InStr(1, b, ",")
        n = InStr(m + 1, b, ",")
        rndNumber2 = CDbl(Mid$(b, m + 1, n - m - 1))
        m = InStr(n, b, ".")
        n = InStr(m, b, ",")
        sSign = Mid$(b, m + 1, n - m - 1)
        Comparing = Mid$(b, n + 1)
        theRndNumber = RndNumber(rndNumber1, rndNumber2)
        If CheckIt(CDbl(theRndNumber), CDbl(Comparing), sSign) = True Then
            DoIF = True
            Exit Function
        End If
    Case "itemcount"
        'itemcount(*basic sword*,signs.<=,2)
        m = InStr(1, b, ",")
        Checking = Mid$(b, 1, m - 1)
        m = InStr(1, b, ".")
        n = InStr(m + 1, b, ",")
        sSign = Mid$(b, m + 1, n - m - 1)
        Comparing = Mid$(b, n + 1)
        If Not IsNumeric(Checking) Then
            Checking = modGetData.GetItemNumFromName(Checking)
            If Checking = "(-1)" Then Exit Function
        End If
        If CheckIt(CDbl(modMain.DCount(dbPlayers(dbIndex).sInventory, ":" & Checking & "/")), CDbl(Comparing), sSign) = True Then
            DoIF = True
            Exit Function
        End If
    Case "mflag"
        m = InStr(1, b, ",")
        Checking = Mid$(b, 1, m - 1)
        n = InStr(m + 1, b, ",")
        sSign = Mid$(b, m + 1, n - m - 1)
        Comparing = Mid$(b, n + 1)
        Select Case Checking
            Case "miscflag.canattack"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Attack])
            Case "miscflag.cancastspell"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Cast Spell])
            Case "miscflag.cansneak"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Sneak])
            Case "miscflag.gibberishtalk"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Gibberish Talk])
            Case "miscflag.guildrank"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Guild Rank])
            Case "miscflag.invisible"
                n = modMiscFlag.GetMiscFlag(dbIndex, Invisible)
            Case "miscflag.caneqhead"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Head])
            Case "miscflag.caneqface"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Face])
            Case "miscflag.caneqears"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ears])
            Case "miscflag.caneqneck"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Neck])
            Case "miscflag.caneqbody"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Body])
            Case "miscflag.caneqback"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Back])
            Case "miscflag.caneqarms"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Arms])
            Case "miscflag.caneqshield"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Shield])
            Case "miscflag.caneqhands"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Hands])
            Case "miscflag.caneqlegs"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Legs])
            Case "miscflag.caneqfeet"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Feet])
            Case "miscflag.caneqwaist"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Waist])
            Case "miscflag.caneqweapon"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Weapon])
            Case "miscflag.canbedesysed"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Be De-Sysed])
            Case "miscflag.seeinvisible"
                n = modMiscFlag.GetMiscFlag(dbIndex, [See Invisible])
            Case "miscflag.seehidden"
                n = modMiscFlag.GetMiscFlag(dbIndex, [See Hidden])
            Case "miscflag.caneqring0"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 0])
            Case "miscflag.caneqring1"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 1])
            Case "miscflag.caneqring2"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 2])
            Case "miscflag.caneqring3"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 3])
            Case "miscflag.caneqring4"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 4])
            Case "miscflag.caneqring5"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 5])
            Case "miscflag.candualwield"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Dual Wield])
            Case "miscflag.cansteal"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Steal])
            Case "miscflag.canbackstab"
                n = modMiscFlag.GetMiscFlag(dbIndex, [Can Backstab])
            Case Else
                DoIF = False
                Exit Function
        End Select
        DoIF = CheckIt(CDbl(n), CDbl(Comparing), sSign)
    Case "timecheck"
        'TimeCheck(signs.>,hh:mm:ss)
        m = InStr(1, b, ".")
        n = InStr(m, b, ",")
        sSign = Mid$(b, m + 1, n - m - 1)
        Comparing = Mid$(b, n + 1)
        If modScripts.CheckTimeDif(sSign, Comparing) = True Then
            DoIF = True
            Exit Function
        Else
            Exit Function
        End If
    Case "datecheck"
        'DateCheck(signs.>,m:dd:yyyy)
        m = InStr(1, b, ".")
        n = InStr(m, b, ",")
        sSign = Mid$(b, m + 1, n - m - 1)
        Comparing = Mid$(b, n + 1)
        If modScripts.CheckDateDif(sSign, Comparing) = True Then
            DoIF = True
            Exit Function
        Else
            Exit Function
        End If
    Case "flag1="
        If dbPlayers(dbIndex).iFlag1 = CLng(b) Then
            DoIF = True
            Exit Function
        End If
    Case "flag2="
        If dbPlayers(dbIndex).iFlag2 = CLng(b) Then
            DoIF = True
            Exit Function
        End If
    Case "flag3="
        If dbPlayers(dbIndex).iFlag3 = CLng(b) Then
            DoIF = True
            Exit Function
        End If
    Case "flag4="
        If dbPlayers(dbIndex).iFlag4 = CLng(b) Then
            DoIF = True
            Exit Function
        End If
    Case "quest1="
        If modSC.FastStringComp(dbPlayers(dbIndex).sQuest1, CLng(b)) Then
            DoIF = True
            Exit Function
        End If
    Case "quest2="
        If modSC.FastStringComp(dbPlayers(dbIndex).sQuest2, CLng(b)) Then
            DoIF = True
            Exit Function
        End If
    Case "quest3="
        If modSC.FastStringComp(dbPlayers(dbIndex).sQuest3, CLng(b)) Then
            DoIF = True
            Exit Function
        End If
    Case "quest4="
        If modSC.FastStringComp(dbPlayers(dbIndex).sQuest4, CLng(b)) Then
            DoIF = True
            Exit Function
        End If
End Select
DoIF = DoIFCont(sIf, b, dbIndex)
End Function

Public Function DoIFCont(a As String, b As String, dbIndex As Long) As Boolean
Select Case a
    Case "appearance.hairlength"
        If modAppearance.GetPlayerAppearanceNumber(dbIndex, [Hair Length]) = Val(b) Then
            DoIFCont = True
            Exit Function
        End If
    Case "appearance.haircolor"
        If modAppearance.GetPlayerAppearanceNumber(dbIndex, [Hair Color]) = Val(b) Then
            DoIFCont = True
            Exit Function
        End If
    Case "appearance.hairstyle"
        If modAppearance.GetPlayerAppearanceNumber(dbIndex, [Hair Style]) = Val(b) Then
            DoIFCont = True
            Exit Function
        End If
    Case "appearance.eyecolor"
        If modAppearance.GetPlayerAppearanceNumber(dbIndex, [Eye Color]) = Val(b) Then
            DoIFCont = True
            Exit Function
        End If
    Case "appearance.beard"
        If modAppearance.GetPlayerAppearanceNumber(dbIndex, beard) = Val(b) Then
            DoIFCont = True
            Exit Function
        End If
    Case "appearance.moustache"
        If modAppearance.GetPlayerAppearanceNumber(dbIndex, moustache) = Val(b) Then
            DoIFCont = True
            Exit Function
        End If
    Case "in"
        If InStr(1, s, b) Then
            DoIFCont = True
            Exit Function
        End If
    Case "nocheck"
        DoIFCont = True
        Exit Function
    Case "hasenhanceflags"
        Select Case b
            Case "eq.weapon"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sWeapon) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.head"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sHead) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.face"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sFace) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ears"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sEars) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.neck"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sNeck) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.body"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sBody) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.back"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sBack) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.arms"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sArms) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.shield"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sShield) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.hands"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sHands) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.legs"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sLegs) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.feet"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sFeet) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.waist"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sWaist) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring0"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sRings(0)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring1"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sRings(1)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring2"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sRings(2)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring3"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sRings(3)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring4"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sRings(4)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring5"
                If modItemManip.GetItemFlagsFromUnFormattedString(dbPlayers(dbIndex).sRings(5)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
        End Select
    Case "isenchanted"
        Select Case b
            Case "eq.weapon"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sWeapon) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.head"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sHead) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.face"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sFace) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ears"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sEars) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.neck"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sNeck) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.body"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sBody) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.back"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sBack) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.arms"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sArms) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.shield"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sShield) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.hands"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sHands) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.legs"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sLegs) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.feet"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sFeet) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.waist"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sWaist) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring0"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sRings(0)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring1"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sRings(1)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring2"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sRings(2)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring3"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sRings(3)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring4"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sRings(4)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring5"
                If modItemManip.GetItemEnchantsFromUnFormattedString(dbPlayers(dbIndex).sRings(5)) <> "" Then
                    DoIFCont = True
                    Exit Function
                End If
        End Select
    Case "has"
        Select Case b
            Case "eq.weapon"
                If dbPlayers(dbIndex).sWeapon <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.head"
                If dbPlayers(dbIndex).sHead <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.face"
                If dbPlayers(dbIndex).sFace <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ears"
                If dbPlayers(dbIndex).sEars <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.neck"
                If dbPlayers(dbIndex).sNeck <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.body"
                If dbPlayers(dbIndex).sBody <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.back"
                If dbPlayers(dbIndex).sBack <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.arms"
                If dbPlayers(dbIndex).sArms <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.shield"
                If dbPlayers(dbIndex).sShield <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.hands"
                If dbPlayers(dbIndex).sHands <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.legs"
                If dbPlayers(dbIndex).sLegs <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.feet"
                If dbPlayers(dbIndex).sFeet <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.waist"
                If dbPlayers(dbIndex).sWaist <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring0"
                If dbPlayers(dbIndex).sRings(0) <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring1"
                If dbPlayers(dbIndex).sRings(1) <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring2"
                If dbPlayers(dbIndex).sRings(2) <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring3"
                If dbPlayers(dbIndex).sRings(3) <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring4"
                If dbPlayers(dbIndex).sRings(4) <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
            Case "eq.ring5"
                If dbPlayers(dbIndex).sRings(5) <> "0" Then
                    DoIFCont = True
                    Exit Function
                End If
        End Select
End Select
End Function

Public Function sScripting(Index As Long, Optional Room As Long = -1, Optional lItem As Long = -1, Optional lMonster As Long = -1, Optional sMons As String = "", Optional GetMYBASE As Boolean = False, Optional ByRef TimerInterval As Long = 0, Optional ByRef CheckFirst As Boolean = False, Optional ByRef sScript As String, Optional IgnoreTimer As Boolean = False, Optional UseThisScript As String = "-1") As Boolean
Dim Arr() As String
Dim Arr1() As String
Dim Arr2() As String
Dim lIfC As Long
Dim lOFC As Long
Dim lMBC As Long
Dim lBeC As Long
Dim lBUSC As Long
Dim i As Long
Dim j As Long
Dim o As Long
Dim m As Long
Dim n As Long
Dim b As String
Dim v As String
Dim c As String
Dim sRes As String
Dim ArrMyBase(1) As String
Dim ArrVars() As String
Dim ArrTmp() As String
Dim t As Boolean
Dim s As String
Dim R As Boolean
If Room <> -1 Then
    With dbMap(GetMapIndex(Room))
        If modSC.FastStringComp(.sScript, "0") Then sScripting = False: Exit Function
        s = .sScript
    End With
ElseIf lItem <> -1 Then
    With dbItems(GetItemID(, lItem))
        If modSC.FastStringComp(.sScript, "0") Then sScripting = False: Exit Function
        s = .sScript
    End With
ElseIf lMonster <> -1 Then
    s = sMons
ElseIf UseThisScript <> "-1" Then
    s = UseThisScript
End If
'Arr = Split(s, ";")
SplitFast s, Arr, "\\"
For i = LBound(Arr) To UBound(Arr)
    Erase Arr1
    If Arr(i) <> "" Then
        SplitFast Arr(i), Arr1, vbCrLf
        ArrTmp = Arr1
        lIfC = 0
        lOFC = 0
        lMBC = 0
        lBeC = 0
        lBUSC = 0
        For j = LBound(Arr1) To UBound(Arr1)
            Arr1(j) = LCase$(Trim$(Arr1(j)))
            If Left$(Arr1(j), 16) = "begin.usescript " Then
                lBUSC = 1
                Exit For
            End If
            If Left$(Arr1(j), 3) = "if " Then
                lIfC = lIfC + 1
                Arr1(j) = "#" & lIfC & "#" & Arr1(j)
                m = 0
                For o = j To UBound(Arr1)
                    Arr1(o) = LCase$(Trim$(Arr1(o)))
                    If Left$(Arr1(o), 3) = "if " Then m = m + 1
                    If Left$(Arr1(o), 7) = "onfail(" Then m = m - 1
                    If m = -1 And Left$(Arr1(o), 7) = "onfail(" Then
                        Arr1(o) = "#" & lIfC & "#" & Arr1(o)
                        Exit For
                    End If
                Next
            End If
            If Mid$(Arr1(j), 4, 7) = "onfail(" Then lOFC = lOFC + 1
            If Left$(Arr1(j), 7) = "mybase." Then lMBC = lMBC + 1
            If Left$(Arr1(j), 14) = "begin.declare " Then lBeC = lBeC + 1
        Next
        If lBUSC = 1 Then
            Arr1(j) = Mid$(Arr1(j), 16)
            Arr1(j) = CStr(Val(Arr1(j)))
            Arr1(j) = GetMapIndex(CLng(Val(Arr1(j))))
            If Arr1(j) <> "0" Then
                If GetMYBASE = False And Room = -1 Then
                    sScripting Index, IgnoreTimer:=True, UseThisScript:=dbMap(CLng(Val(Arr1(j)))).sScript
                ElseIf GetMYBASE = True Then
                    sScript = dbMap(CLng(Val(Arr1(j)))).sScript
                    sScripting Index, , , , , True, TimerInterval, , sScript, , sScript
                    Exit Function
                    'GoTo DoMYBASE
                End If
                Exit Function
            Else
                Exit Function
            End If
        End If
        If lIfC <> lOFC Then
            'MsgBox "If count and OnFail count do not match."
            Exit Function
        End If
        If lMBC > 2 Then
            'MsgBox "Mybase count is greater then 2."
            Exit Function
        End If
        ReDim ArrVars(0)
        If lBeC > 0 Then
            j = 0
            Do Until Left$(Arr1(j), 14) <> "begin.declare "
                If j > UBound(ArrVars) Then ReDim Preserve ArrVars(UBound(ArrVars) + 1)
                ArrVars(j) = Arr1(j)
                j = j + 1
            Loop
            For n = 0 To UBound(ArrVars)
                For o = LBound(Arr1) To UBound(Arr1) - 1
                    Arr1(o) = Arr1(o + 1)
                Next
            Next
            If ArrVars(0) <> "" Then
                c = ""
                'Var/Style/Params
                For j = LBound(ArrVars) To UBound(ArrVars)
                    n = InStr(1, ArrVars(j), " ")
                    m = InStr(n + 1, ArrVars(j), " ")
                    c = c & Mid$(ArrVars(j), n + 1, m - n - 1)
                    n = InStr(1, ArrVars(j), "style.") + 5
                    m = InStr(n + 1, ArrVars(j), " = ")
                    c = c & "/" & Mid$(ArrVars(j), n + 1, m - n - 1) & "/"
                    c = c & Mid$(ArrVars(j), m + 3)
                    ArrVars(j) = c
                    c = ""
                Next
            End If
        End If
        If lMBC > 0 Then
            ArrMyBase(0) = Mid$(Arr1(0), InStr(1, Arr1(0), ".") + 1)
            If lMBC > 1 Then
                ArrMyBase(1) = Mid$(Arr1(1), InStr(1, Arr1(1), ".") + 1)
            End If
            If GetMYBASE Then
                If Left$(ArrMyBase(0), 6) = "timer(" Then
                    m = InStr(1, ArrMyBase(0), ")")
                    TimerInterval = CLng(Val(Mid$(ArrMyBase(0), 7, m - 7)))
                    sScript = Arr(i)
                    If Not modSC.FastStringComp(ArrMyBase(1), "") Then
                        If modSC.FastStringComp(ArrMyBase(1), "checkfirst") Then CheckFirst = True
                    End If
                Else
                    If modSC.FastStringComp(ArrMyBase(0), "checkfirst") Then CheckFirst = True
                    If Not modSC.FastStringComp(ArrMyBase(1), "") Then
                        If Left$(ArrMyBase(1), 6) = "timer(" Then
                            m = InStr(1, ArrMyBase(1), ")")
                            TimerInterval = CLng(Val(Mid$(ArrMyBase(1), 7, m - 7)))
                            sScript = Arr(i)
                        End If
                    End If
                End If
                GoTo nNext
            End If
            If (Left$(ArrMyBase(0), 6) = "timer(" Or Left$(ArrMyBase(1), 6) = "timer(") And (Not IgnoreTimer) Then GoTo nNext
        End If
        If lMBC = 0 And GetMYBASE Then GoTo nNext
        lIfC = 0
        For j = LBound(Arr1) + lMBC To UBound(Arr1)
            R = False
            If Mid$(Arr1(j), 4, 3) = "if " Then
                lIfC = lIfC + 1
                b = Arr1(j)
                m = InStr(1, b, "check.") + 6
                n = InStr(m, b, "*")
                If n = 0 Then
                    n = InStr(m, b, "(")
                    If n = 0 Then
                        'MsgBox "Error parsing IF statement."
                        Exit Function
                    End If
                End If
                v = Mid$(b, m, n - m)
                m = InStr(n + 1, b, "*")
                If m = 0 Then
                    m = InStr(n + 1, b, ")")
                    If m = 0 Then
                        'MsgBox "Error parsing IF statement."
                        Exit Function
                    End If
                End If
                c = Mid$(b, n + 1, m - n - 1)
                'sReport = sReport & "IF statement (If #" & CStr(lIfC) & ") :" & vbCrLf & "   IF: " & v & " is " & c & vbCrLf
                'CHECK HERE!
                If modScripts.DoIF(Index, v, c) = True Then
                    If j + 1 < UBound(Arr1) Then
                        If Mid$(Arr1(j + 1), 4, 3) = "if " Then
                            'do nothing, continue with the loop.
                        Else
                            If Left$(Arr1(j + 1), 8) = "respond." Then
                                n = j + 1
                                t = False
                                sRes = ""
                                Do Until t = True
                                    sRes = sRes & ArrTmp(n + lBeC) & vbCrLf
                                    n = n + 1
                                    If n > UBound(Arr1) Then
                                        t = True
                                    ElseIf Left$(Arr1(n), 8) <> "respond." Then
                                        t = True
                                    End If
                                Loop
                                If sRes <> "" Then
                                    SplitFast sRes, Arr2, vbCrLf
                                    For o = LBound(Arr2) To UBound(Arr2)
                                        If Arr2(o) <> "" Then
                                            v = ""
                                            c = ""
                                            b = Arr2(o)
                                            n = InStr(1, b, ".")
                                            m = InStr(1, b, ",")
                                            If m <> 0 Then
                                                v = Mid$(b, n + 1, m - n - 1)
                                                c = Mid$(b, m + 1)
                                            Else
                                                v = Mid$(b, n + 1)
                                            End If
                                            v = LCaseFast(v)
                                            'sReport = sReport & "Respond statement:" & vbCrLf & "    RESPOND:" & v & " with " & c & vbCrLf
                                            'SEND OFF RESPOND HERE
                                            DoAbil Index, v, c, ArrVars
                                            sScripting = True
                                            If UseThisScript = "-1" And dbPlayers(GetPlayerIndexNumber(Index)).lCanClear = 0 Then X(Index) = ""
                                            
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                Else 'Check fails
                    n = j
                    m = 0
                    For o = LBound(Arr1) To UBound(Arr1)
                        If Left$(Arr1(o), 3) = Left$(Arr1(j), 3) Then
                            If Arr1(o) <> Arr1(j) Then
                                b = ArrTmp(o + lBeC)
                                m = InStr(1, b, "*")
                                n = InStr(m + 1, b, "*")
                                v = Mid$(b, m + 1, n - m - 1)
                                If v <> "" Then
                                    sSend Index, v
                                    sScripting = True
                                    If UseThisScript = "-1" Then X(Index) = ""
                                End If
                                'sReport = sReport & "OnFail statement (If #" & Mid$(Arr1(o), 2, 1) & ") :" & vbCrLf & "    ONFAIL: Send> " & v & vbCrLf
                                R = True
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
            If R Then Exit For
        Next
    End If
nNext:
Next
'If ArrMyBase(0) <> "" Then
'    sReport = "MYBASE Statement: " & vbCrLf & "    MYBASE " & ArrMyBase(0) & vbCrLf & sReport
'    If ArrMyBase(1) <> "" Then
'        sReport = "MYBASE Statement: " & vbCrLf & "    MYBASE " & ArrMyBase(1) & vbCrLf & sReport
'    End If
'End If
'sReport = sReport & vbCrLf & "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-" & vbCrLf
'sReport = sReport & "If count: " & lIfC & vbCrLf & "MyBase count: " & lMBC & vbCrLf & "OnFail count: " & lOFC
'frmMain.txtRespond = sReport
End Function



Public Function CheckTheStat(Index As Long, CheckingStat As String, ComparingTo As Double, sSign As String) As Boolean
CheckTheStat = False
Dim dbIndex As Long
dbIndex = GetPlayerIndexNumber(Index)
With dbPlayers(dbIndex)
    Select Case CheckingStat
        Case "lives"
            If CheckIt(CDbl(.iLives), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "level"
            If CheckIt(CDbl(.iLevel), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "str"
            If CheckIt(CDbl(.iStr), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "agil"
            If CheckIt(CDbl(.iAgil), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "int"
            If CheckIt(CDbl(.iInt), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "dex"
            If CheckIt(CDbl(.iDex), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "cha"
            If CheckIt(CDbl(.iCha), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "level"
            If CheckIt(CDbl(.iLevel), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "exp"
            If CheckIt(.dEXP, ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "expneeded"
            If CheckIt(.dEXPNeeded, ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "totalexp"
            If CheckIt(.dTotalEXP, ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "weapons"
            'FIX!!!?
            '********************************************************************************
            If CheckIt(CDbl(.iWeapons), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
            '********************************************************************************
        Case "armortype"
            'FIX!?
            '********************************************************************************
            If CheckIt(CDbl(.iArmorType), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
            '********************************************************************************
        Case "spelllevel"
            If CheckIt(CDbl(.iSpellLevel), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "spelltype"
            If CheckIt(CDbl(.iSpellType), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "gold"
            If CheckIt(.dGold, ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "hp"
            If CheckIt(CDbl(.lHP), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "maxhp"
            If CheckIt(CDbl(.lMaxHP), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "mana"
            If CheckIt(CDbl(.lMana), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "maxmana"
            If CheckIt(CDbl(.lMaxMana), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "ac"
            If CheckIt(CDbl(.iAC), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "acc"
            If CheckIt(CDbl(.iAcc), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "crits"
            If CheckIt(CDbl(.iCrits), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "dodge"
            If CheckIt(CDbl(.iDodge), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "maxdamage"
            If CheckIt(CDbl(.iMaxDamage), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "bank"
            If CheckIt(.dBank, ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "vision"
            If CheckIt(CDbl(.iVision), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "maxitems"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items])), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "guildleader"
            If CheckIt(CDbl(.iGuildLeader), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "evil"
            If CheckIt(CDbl(.iEvil), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.fire"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Fire)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.ice"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Ice)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.water"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Water)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.lightning"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Lightning)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.earth"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Earth)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.poison"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Poison)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.wind"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Wind)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.holy"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Holy)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "element.unholy"
            If CheckIt(CDbl(modResist.GetResistValue(dbIndex, Unholy)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "spellcasting"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlusTotal(dbIndex, [Spell Casting])), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "magicres"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlusTotal(dbIndex, [Magic Resistance])), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "perception"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlusTotal(dbIndex, Perception)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "isasysop"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlus(dbIndex, [Is A Sysop])), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "pallete"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlus(dbIndex, [Pallete Number])), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "stealth"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlusTotal(dbIndex, Steath)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "animalrelations"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlusTotal(dbIndex, [Animal Relations])), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "hunger"
            If CheckIt(dbPlayers(dbIndex).dHunger, ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "stamina"
            If CheckIt(dbPlayers(dbIndex).dStamina, ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "thieving"
            If CheckIt(CDbl(modMiscFlag.GetStatsPlusTotal(dbIndex, Thieving)), ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
        Case "classpts"
            If CheckIt(dbPlayers(dbIndex).dClassPoints, ComparingTo, sSign) = True Then CheckTheStat = True: Exit Function
    End Select
End With

End Function

Public Function CheckIt(sValue1 As Double, sValue2 As Double, sSign As String) As Boolean
CheckIt = False
Select Case sSign
    Case ">"
        If sValue1 > sValue2 Then CheckIt = True
    Case "<"
        If sValue1 < sValue2 Then CheckIt = True
    Case "="
        If sValue1 = sValue2 Then CheckIt = True
    Case ">="
        If sValue1 >= sValue2 Then CheckIt = True
    Case "<="
        If sValue1 <= sValue2 Then CheckIt = True
    Case "<>"
        If sValue1 <> sValue2 Then CheckIt = True
End Select
End Function

Public Function DoAbil(Index As Long, sDo As String, sW As String, ByRef ArrVars() As String)
Dim Arr() As String
Dim Arr2() As String
If modSC.FastStringComp(sDo, "") Then Exit Function
Dim a As String
Dim b As String
Dim s As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim dbIndex As Long
b = sDo
dbIndex = GetPlayerIndexNumber(Index)
For i = LBound(ArrVars) To UBound(ArrVars)
    If ArrVars(i) <> "" Then
    
        SplitFast ArrVars(i), Arr, "/"
        Select Case LCaseFast(Arr(1))
            Case "random"
                SplitFast Arr(2), Arr2, ","
                For j = 0 To 1
                    If Left$(Arr2(j), 5) = "stat." Then
                        Arr2(j) = Mid$(Arr2(j), 6)
                        Arr2(j) = CStr(modGetData.GetStatFromString(dbIndex, Arr2(j)))
                    End If
                Next
                a = CStr(RndNumber(CDbl(Val(Arr2(0))), CDbl(Val(Arr2(1)))))
                sW = ReplaceFast(sW, Arr(0), a, , , vbTextCompare)
                ArrVars(i) = Arr(0) & "/static/" & a
                a = ""
            Case "static"
                If Left$(Arr(2), 5) = "stat." Then
                    Arr(2) = Mid$(Arr(2), 6)
                    Arr(2) = CStr(modGetData.GetStatFromString(dbIndex, Arr(2)))
                    ArrVars(i) = Arr(0) & "/static/" & Arr(2)
                End If
                sW = ReplaceFast(sW, Arr(0), Arr(2), , , vbTextCompare)
            Case "string"
                If Left(Arr(2), 7) = "player." Then
                    Arr(2) = Mid$(Arr(2), 8)
                    Select Case Arr(2)
                        Case "name"
                            Arr(2) = dbPlayers(dbIndex).sSeenAs
                        Case "seenas"
                            Arr(2) = dbPlayers(dbIndex).sPlayerName
                        Case "namepossesive"
                            Arr(2) = dbPlayers(dbIndex).sSeenAs & "'s"
                        Case "seenaspossesive"
                            Arr(2) = dbPlayers(dbIndex).sPlayerName & "'s"
                    End Select
                End If
                sW = ReplaceFast(sW, Arr(0), Arr(2), , , vbTextCompare)
        End Select
    End If
    If DE Then DoEvents
Next
a = sW
s = sW
Select Case b
    Case "event.add"
        sAddEvent dbIndex, s
    Case "event.erase"
        sEraseEvent dbIndex, s
    Case "putiteminroom"
        If IsNumeric(s) Then
            i = GetItemID(, Val(s))
        Else
            i = GetItemID(s)
        End If
        If i <> 0 Then
            With dbMap(GetMapIndex(dbPlayers(dbIndex).lLocation))
                If .sItems = "0" Then .sItems = ""
                .sItems = .sItems & ":" & dbItems(i).iID & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(i).lDurability & "/" & dbItems(i).iUses & ";"
            End With
        End If
    Case "genmon"
        If IsNumeric(s) Then
            i = GetMonsterID(, Val(s))
        Else
            i = GetMonsterID(s)
        End If
        If i <> 0 Then modMonsters.GenAMonster dbPlayers(dbIndex).lLocation, False, dbMonsters(i).lMobGroup, dbMonsters(i).lID
    Case "takeallofitem"
        If IsNumeric(s) Then
            i = GetItemID(, Val(s))
        Else
            i = GetItemID(s)
        End If
        If i <> 0 Then
            With dbPlayers(dbIndex)
                Do Until InStr(1, .sInventory, ":" & dbItems(i).iID & "/") = 0
                    modItemManip.RemoveItemFromInv dbIndex, dbItems(i).iID
                    If DE Then DoEvents
                Loop
                If .sInventory = "" Then .sInventory = "0"
            End With
        End If
    
    Case "takeitems"
        SplitFast s, Arr, ","
        For j = LBound(Arr) To UBound(Arr)
            If IsNumeric(Arr(j)) Then
                i = GetItemID(, Val(Arr(j)))
            Else
                i = GetItemID(Arr(j))
            End If
            If i <> 0 Then
                With dbPlayers(dbIndex)
                    modItemManip.RemoveItemFromInv dbIndex, dbItems(i).iID
                    If .sInventory = "" Then .sInventory = "0"
                End With
            End If
            If DE Then DoEvents
        Next
    Case "genmons"
        SplitFast s, Arr, ","
        For j = LBound(Arr) To UBound(Arr)
            If IsNumeric(Arr(j)) Then
                i = GetMonsterID(, Val(Arr(j)))
            Else
                i = GetMonsterID(Arr(j))
            End If
            If i <> 0 Then modMonsters.GenAMonster dbPlayers(dbIndex).lLocation, False, dbMonsters(i).lMobGroup, dbMonsters(i).lID
            If DE Then DoEvents
        Next
    Case "genmonloc"
        SplitFast s, Arr, ","
        If Val(Arr(0)) > 0 Then
            If IsNumeric(Arr(1)) Then
                i = GetMonsterID(, Val(Arr(1)))
            Else
                i = GetMonsterID(Arr(1))
            End If
            If i <> 0 Then modMonsters.GenAMonster Val(Arr(0)), False, dbMonsters(i).lMobGroup, dbMonsters(i).lID
        End If
    Case "genmonsloc"
        SplitFast s, Arr, ","
        If Val(Arr(0)) > 0 Then
            For j = LBound(Arr) + 1 To UBound(Arr)
                If IsNumeric(Arr(j)) Then
                    i = GetMonsterID(, Val(Arr(j)))
                Else
                    i = GetMonsterID(Arr(j))
                End If
                If i <> 0 Then modMonsters.GenAMonster Val(Arr(0)), False, dbMonsters(i).lMobGroup, dbMonsters(i).lID
                If DE Then DoEvents
            Next
        End If
    Case "putiteminroomloc"
        SplitFast s, Arr, ","
        If Val(Arr(0)) > 0 Then
            If IsNumeric(Arr(1)) Then
                i = GetItemID(, Val(Arr(1)))
            Else
                i = GetItemID(Arr(1))
            End If
        End If
        If i <> 0 Then
            j = GetMapIndex(Val(Arr(0)))
            If j <> 0 Then
                With dbMap(j)
                    If .sItems = "0" Then .sItems = ""
                    .sItems = .sItems & ":" & dbItems(i).iID & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(i).lDurability & "/" & dbItems(i).iUses & ";"
                End With
            End If
        End If
    Case "putitemsinroomloc"
        SplitFast s, Arr, ","
        k = GetMapIndex(Val(Arr(0)))
        If Val(Arr(0)) > 0 Then
            For j = LBound(Arr) + 1 To UBound(Arr)
                If IsNumeric(Arr(j)) Then
                    i = GetItemID(, Val(Arr(j)))
                Else
                    i = GetItemID(Arr(j))
                End If
                If i <> 0 Then
                    If k <> 0 Then
                        With dbMap(k)
                            If .sItems = "0" Then .sItems = ""
                            .sItems = .sItems & ":" & dbItems(i).iID & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(i).lDurability & "/" & dbItems(i).iUses & ";"
                        End With
                    End If
                End If
                If DE Then DoEvents
            Next
        End If
    Case "putitemsinroom"
        SplitFast s, Arr, ","
        k = GetMapIndex(dbPlayers(dbIndex).lLocation)
        For j = LBound(Arr) + 1 To UBound(Arr)
            If IsNumeric(Arr(1)) Then
                i = GetItemID(, Val(Arr(j)))
            Else
                i = GetItemID(Arr(j))
            End If
            If i <> 0 Then
                With dbMap(k)
                    If .sItems = "0" Then .sItems = ""
                    .sItems = .sItems & ":" & dbItems(i).iID & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(i).lDurability & "/" & dbItems(i).iUses & ";"
                End With
            End If
            If DE Then DoEvents
        Next
    Case "takeitemsfromroomloc"
        SplitFast s, Arr, ","
        k = GetMapIndex(Val(Arr(0)))
        If Val(Arr(0)) > 0 Then
            For j = LBound(Arr) + 1 To UBound(Arr)
                If IsNumeric(Arr(j)) Then
                    i = GetItemID(, Val(Arr(j)))
                Else
                    i = GetItemID(Arr(j))
                End If
                If i <> 0 Then
                    If k <> 0 Then
                        With dbMap(k)
                            modItemManip.RemoveItemFromGround k, dbItems(i).iID
                        End With
                    End If
                End If
                If DE Then DoEvents
            Next
        End If
    Case "takeitemfromroomloc"
        SplitFast s, Arr, ","
        k = GetMapIndex(Val(Arr(0)))
        If Val(Arr(0)) > 0 Then
            If IsNumeric(Arr(1)) Then
                i = GetItemID(, Val(Arr(1)))
            Else
                i = GetItemID(Arr(1))
            End If
            If i <> 0 Then
                If k <> 0 Then
                    With dbMap(k)
                        modItemManip.RemoveItemFromGround k, dbItems(i).iID
                    End With
                End If
            End If
        End If
    Case "takeitemfromroom"
        k = GetMapIndex(dbPlayers(dbIndex).lLocation)
        If IsNumeric(s) Then
            i = GetItemID(, Val(s))
        Else
            i = GetItemID(s)
        End If
        If i <> 0 Then
            If k <> 0 Then
                With dbMap(k)
                    modItemManip.RemoveItemFromGround k, dbItems(i).iID
                End With
            End If
        End If
    Case "takeitemsfromroom"
        SplitFast s, Arr, ","
        k = GetMapIndex(dbPlayers(dbIndex).lLocation)
        For j = LBound(Arr) To UBound(Arr)
            If IsNumeric(Arr(j)) Then
                i = GetItemID(, Val(Arr(j)))
            Else
                i = GetItemID(Arr(j))
            End If
            If i <> 0 Then
                If k <> 0 Then
                    With dbMap(k)
                        modItemManip.RemoveItemFromGround k, dbItems(i).iID
                    End With
                End If
            End If
            If DE Then DoEvents
        Next
    Case "dropfrominv"
        If IsNumeric(s) Then
            i = GetItemID(, Val(s))
        Else
            i = GetItemID(s)
        End If
        If i <> 0 Then
            modItemManip.TakeItemFromInvAndPutOnGround dbIndex, dbItems(i).iID
        End If
    Case "addhunger"
        sAddHunger dbIndex, s
    Case "addstamina"
        sAddStamina dbIndex, s
    Case "appearance.hairlength", "appearance.haircolor", "appearance.hairstyle", "appearance.eyecolor", "appearance.moustache", "appearance.beard"
        sChangeAppearance dbIndex, b, s
    Case "teleport"
        sTeleport dbIndex, a
    Case "partytel"
        sPartyTel dbIndex, a
    Case "partysend"
        sPartySend dbIndex, a
    Case "gainhp"
        sGainHP dbIndex, a
    Case "gainma"
        sGainMA dbIndex, a
    Case "giveitem"
        sGiveItem dbIndex, a
    Case "addspell"
        sAddSpell dbIndex, a
    Case "addexp"
        sAddEXP dbIndex, a
    Case "addstr"
        sAddStr dbIndex, a
    Case "addcha"
        sAddCha dbIndex, a
    Case "adddex"
        sAddDex dbIndex, a
    Case "addagil"
        sAddAgil dbIndex, a
    Case "addint"
        sAddInt dbIndex, a
    Case "addevil"
        sAddEvil dbIndex, a
    Case "send"
        sSend Index, a
    Case "castsp"
        sCastSp dbIndex, a
    Case "changerace"
        sChangeRace dbIndex, a
    Case "changeclass"
        sChangeClass dbIndex, a
    Case "sendroom"
        sSendRoom dbIndex, a
    Case "givegold"
        sGiveGold dbIndex, a
    Case "takeitem"
        sTakeItem dbIndex, a
    Case "addsc"
        sAddSC dbIndex, a
    Case "changearmortype"
        sChangeArmorType dbIndex, a
    Case "changeweapontype"
        sChangeWeaponType dbIndex, a
    Case "addacc"
        sAddAcc dbIndex, a
    Case "addcrits"
        sAddCrits dbIndex, a
    Case "adddodge"
        sAddDodge dbIndex, a
    Case "givefam"
        sGiveFam dbIndex, a
    Case "takefam"
        sTakeFam dbIndex
    Case "changequest"
        sAddQuest dbIndex, s
    Case "changebank"
        sChangeBank dbIndex, s
    Case "addstattrain"
        sAddStatTrain dbIndex, s
    Case "changeflag"
        'ChangeFlag(1,3)
        sChangeFlag dbIndex, s
    Case "showtime"
        sShowTime dbIndex
    Case "changetime"
        'ChangeTime,0,0,0
        sChangeTime dbIndex, s
    Case "seenas"
        sChangeSeenAs dbIndex, s
    Case "desc"
        sChangeDesc dbIndex, s
    Case "noerase"
        dbPlayers(dbIndex).lCanClear = 1
End Select
DoAbilCont b, s, dbIndex
End Function

Public Sub DoAbilCont(b As String, s As String, dbIndex As Long)
Dim Arr() As String
Dim i As Long
Select Case b
    Case "closedoor"
        SplitFast s, Arr, ","
        i = GetMapIndex(Val(Arr(0)))
        If i <> 0 Then
            With dbMap(i)
                Select Case LCaseFast(Arr(1))
                    Case "direction.north"
                        .lDN = 1
                    Case "direction.south"
                        .lDS = 1
                    Case "direction.east"
                        .lDE = 1
                    Case "direction.west"
                        .lDW = 1
                    Case "direction.northwest"
                        .lDNW = 1
                    Case "direction.northeast"
                        .lDNE = 1
                    Case "direction.southwest"
                        .lDSW = 1
                    Case "direction.southeast"
                        .lDSE = 1
                    Case "direction.up"
                        .lDU = 1
                    Case "direction.down"
                        .lDD = 1
                End Select
            End With
        End If
    Case "lockdoor"
        SplitFast s, Arr, ","
        i = GetMapIndex(Val(Arr(0)))
        If i <> 0 Then
            With dbMap(i)
                Select Case LCaseFast(Arr(1))
                    Case "direction.north"
                        .lDN = 2
                    Case "direction.south"
                        .lDS = 2
                    Case "direction.east"
                        .lDE = 2
                    Case "direction.west"
                        .lDW = 2
                    Case "direction.northwest"
                        .lDNW = 2
                    Case "direction.northeast"
                        .lDNE = 2
                    Case "direction.southwest"
                        .lDSW = 2
                    Case "direction.southeast"
                        .lDSE = 2
                    Case "direction.up"
                        .lDU = 2
                    Case "direction.down"
                        .lDD = 2
                End Select
            End With
        End If
    Case "opendoor"
        SplitFast s, Arr, ","
        i = GetMapIndex(Val(Arr(0)))
        If i <> 0 Then
            With dbMap(i)
                Select Case LCaseFast(Arr(1))
                    Case "direction.north"
                        .lDN = 3
                    Case "direction.south"
                        .lDS = 3
                    Case "direction.east"
                        .lDE = 3
                    Case "direction.west"
                        .lDW = 3
                    Case "direction.northwest"
                        .lDNW = 3
                    Case "direction.northeast"
                        .lDNE = 3
                    Case "direction.southwest"
                        .lDSW = 3
                    Case "direction.southeast"
                        .lDSE = 3
                    Case "direction.up"
                        .lDU = 3
                    Case "direction.down"
                        .lDD = 3
                End Select
            End With
        End If
    Case "unlockdoor"
        SplitFast s, Arr, ","
        i = GetMapIndex(Val(Arr(0)))
        If i <> 0 Then
            With dbMap(i)
                Select Case LCaseFast(Arr(1))
                    Case "direction.north"
                        .lDN = 1
                    Case "direction.south"
                        .lDS = 1
                    Case "direction.east"
                        .lDE = 1
                    Case "direction.west"
                        .lDW = 1
                    Case "direction.northwest"
                        .lDNW = 1
                    Case "direction.northeast"
                        .lDNE = 1
                    Case "direction.southwest"
                        .lDSW = 1
                    Case "direction.southeast"
                        .lDSE = 1
                    Case "direction.up"
                        .lDU = 1
                    Case "direction.down"
                        .lDD = 1
                End Select
            End With
        End If
'-----------------------------------------
        
    Case "adjectives.weapon"
        dbPlayers(dbIndex).sWeapon = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sWeapon, s & "|")
    Case "adjectives.head"
        dbPlayers(dbIndex).sHead = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sHead, s & "|")
    Case "adjectives.face"
        dbPlayers(dbIndex).sFace = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sFace, s & "|")
    Case "adjectives.ears"
        dbPlayers(dbIndex).sEars = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sEars, s & "|")
    Case "adjectives.neck"
        dbPlayers(dbIndex).sNeck = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sNeck, s & "|")
    Case "adjectives.body"
        dbPlayers(dbIndex).sBody = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sBody, s & "|")
    Case "adjectives.back"
        dbPlayers(dbIndex).sBack = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sBack, s & "|")
    Case "adjectives.arms"
        dbPlayers(dbIndex).sArms = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sArms, s & "|")
    Case "adjectives.shield"
        dbPlayers(dbIndex).sShield = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sShield, s & "|")
    Case "adjectives.hands"
        dbPlayers(dbIndex).sHands = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sHands, s & "|")
    Case "adjectives.legs"
        dbPlayers(dbIndex).sLegs = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sLegs, s & "|")
    Case "adjectives.feet"
        dbPlayers(dbIndex).sFeet = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sFeet, s & "|")
    Case "adjectives.waist"
        dbPlayers(dbIndex).sWaist = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sWaist, s & "|")
    Case "adjectives.ring0"
        dbPlayers(dbIndex).sRings(0) = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sRings(0), s & "|")
    Case "adjectives.ring1"
        dbPlayers(dbIndex).sRings(1) = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sRings(1), s & "|")
    Case "adjectives.ring2"
        dbPlayers(dbIndex).sRings(2) = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sRings(2), s & "|")
    Case "adjectives.ring3"
        dbPlayers(dbIndex).sRings(3) = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sRings(3), s & "|")
    Case "adjectives.ring4"
        dbPlayers(dbIndex).sRings(4) = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sRings(4), s & "|")
    Case "adjectives.ring5"
        dbPlayers(dbIndex).sRings(5) = modItemManip.SetItemAdjectives(dbPlayers(dbIndex).sRings(5), s & "|")
        
    '-----------------------------------------
    
    Case "enchant.weapon"
        dbPlayers(dbIndex).sWeapon = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sWeapon, MakedbFlags(s))
    Case "enchant.head"
        dbPlayers(dbIndex).sHead = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sHead, MakedbFlags(s))
    Case "enchant.face"
        dbPlayers(dbIndex).sFace = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sFace, MakedbFlags(s))
    Case "enchant.ears"
        dbPlayers(dbIndex).sEars = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sEars, MakedbFlags(s))
    Case "enchant.neck"
        dbPlayers(dbIndex).sNeck = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sNeck, MakedbFlags(s))
    Case "enchant.body"
        dbPlayers(dbIndex).sBody = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sBody, MakedbFlags(s))
    Case "enchant.back"
        dbPlayers(dbIndex).sBack = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sBack, MakedbFlags(s))
    Case "enchant.arms"
        dbPlayers(dbIndex).sArms = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sArms, MakedbFlags(s))
    Case "enchant.shield"
        dbPlayers(dbIndex).sShield = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sShield, MakedbFlags(s))
    Case "enchant.hands"
        dbPlayers(dbIndex).sHands = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sHands, MakedbFlags(s))
    Case "enchant.legs"
        dbPlayers(dbIndex).sLegs = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sLegs, MakedbFlags(s))
    Case "enchant.feet"
        dbPlayers(dbIndex).sFeet = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sFeet, MakedbFlags(s))
    Case "enchant.waist"
        dbPlayers(dbIndex).sWaist = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sWaist, MakedbFlags(s))
    Case "enchant.ring0"
        dbPlayers(dbIndex).sRings(0) = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sRings(0), MakedbFlags(s))
    Case "enchant.ring1"
        dbPlayers(dbIndex).sRings(1) = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sRings(1), MakedbFlags(s))
    Case "enchant.ring2"
        dbPlayers(dbIndex).sRings(2) = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sRings(2), MakedbFlags(s))
    Case "enchant.ring3"
        dbPlayers(dbIndex).sRings(3) = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sRings(3), MakedbFlags(s))
    Case "enchant.ring4"
        dbPlayers(dbIndex).sRings(4) = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sRings(4), MakedbFlags(s))
    Case "enchant.ring5"
        dbPlayers(dbIndex).sRings(5) = modItemManip.SetItemEnchants(dbPlayers(dbIndex).sRings(5), MakedbFlags(s))
        
    '-----------------------------------------
    
    Case "flags.weapon"
        dbPlayers(dbIndex).sWeapon = modItemManip.SetItemFlags(dbPlayers(dbIndex).sWeapon, MakedbFlags(s))
    Case "flags.head"
        dbPlayers(dbIndex).sHead = modItemManip.SetItemFlags(dbPlayers(dbIndex).sHead, MakedbFlags(s))
    Case "flags.face"
        dbPlayers(dbIndex).sFace = modItemManip.SetItemFlags(dbPlayers(dbIndex).sFace, MakedbFlags(s))
    Case "flags.ears"
        dbPlayers(dbIndex).sEars = modItemManip.SetItemFlags(dbPlayers(dbIndex).sEars, MakedbFlags(s))
    Case "flags.neck"
        dbPlayers(dbIndex).sNeck = modItemManip.SetItemFlags(dbPlayers(dbIndex).sNeck, MakedbFlags(s))
    Case "flags.body"
        dbPlayers(dbIndex).sBody = modItemManip.SetItemFlags(dbPlayers(dbIndex).sBody, MakedbFlags(s))
    Case "flags.back"
        dbPlayers(dbIndex).sBack = modItemManip.SetItemFlags(dbPlayers(dbIndex).sBack, MakedbFlags(s))
    Case "flags.arms"
        dbPlayers(dbIndex).sArms = modItemManip.SetItemFlags(dbPlayers(dbIndex).sArms, MakedbFlags(s))
    Case "flags.shield"
        dbPlayers(dbIndex).sShield = modItemManip.SetItemFlags(dbPlayers(dbIndex).sShield, MakedbFlags(s))
    Case "flags.hands"
        dbPlayers(dbIndex).sHands = modItemManip.SetItemFlags(dbPlayers(dbIndex).sHands, MakedbFlags(s))
    Case "flags.legs"
        dbPlayers(dbIndex).sLegs = modItemManip.SetItemFlags(dbPlayers(dbIndex).sLegs, MakedbFlags(s))
    Case "flags.feet"
        dbPlayers(dbIndex).sFeet = modItemManip.SetItemFlags(dbPlayers(dbIndex).sFeet, MakedbFlags(s))
    Case "flags.waist"
        dbPlayers(dbIndex).sWaist = modItemManip.SetItemFlags(dbPlayers(dbIndex).sWaist, MakedbFlags(s))
    Case "flags.ring0"
        dbPlayers(dbIndex).sRings(0) = modItemManip.SetItemFlags(dbPlayers(dbIndex).sRings(0), MakedbFlags(s))
    Case "flags.ring1"
        dbPlayers(dbIndex).sRings(1) = modItemManip.SetItemFlags(dbPlayers(dbIndex).sRings(1), MakedbFlags(s))
    Case "flags.ring2"
        dbPlayers(dbIndex).sRings(2) = modItemManip.SetItemFlags(dbPlayers(dbIndex).sRings(2), MakedbFlags(s))
    Case "flags.ring3"
        dbPlayers(dbIndex).sRings(3) = modItemManip.SetItemFlags(dbPlayers(dbIndex).sRings(3), MakedbFlags(s))
    Case "flags.ring4"
        dbPlayers(dbIndex).sRings(4) = modItemManip.SetItemFlags(dbPlayers(dbIndex).sRings(4), MakedbFlags(s))
    Case "flags.ring5"
        dbPlayers(dbIndex).sRings(5) = modItemManip.SetItemFlags(dbPlayers(dbIndex).sRings(5), MakedbFlags(s))
        
    '-----------------------------------------
    
    Case "clear.flags.weapon"
        dbPlayers(dbIndex).sWeapon = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sWeapon)
    Case "clear.flags.head"
        dbPlayers(dbIndex).sHead = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sHead)
    Case "clear.flags.face"
        dbPlayers(dbIndex).sFace = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sFace)
    Case "clear.flags.ears"
        dbPlayers(dbIndex).sEars = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sEars)
    Case "clear.flags.neck"
        dbPlayers(dbIndex).sNeck = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sNeck)
    Case "clear.flags.body"
        dbPlayers(dbIndex).sBody = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sBody)
    Case "clear.flags.back"
        dbPlayers(dbIndex).sBack = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sBack)
    Case "clear.flags.arms"
        dbPlayers(dbIndex).sArms = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sArms)
    Case "clear.flags.shield"
        dbPlayers(dbIndex).sShield = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sShield)
    Case "clear.flags.hands"
        dbPlayers(dbIndex).sHands = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sHands)
    Case "clear.flags.legs"
        dbPlayers(dbIndex).sLegs = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sLegs)
    Case "clear.flags.feet"
        dbPlayers(dbIndex).sFeet = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sFeet)
    Case "clear.flags.waist"
        dbPlayers(dbIndex).sWaist = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sWaist)
    Case "clear.flags.ring0"
        dbPlayers(dbIndex).sRings(0) = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sRings(0))
    Case "clear.flags.ring1"
        dbPlayers(dbIndex).sRings(1) = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sRings(1))
    Case "clear.flags.ring2"
        dbPlayers(dbIndex).sRings(2) = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sRings(2))
    Case "clear.flags.ring3"
        dbPlayers(dbIndex).sRings(3) = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sRings(3))
    Case "clear.flags.ring4"
        dbPlayers(dbIndex).sRings(4) = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sRings(4))
    Case "clear.flags.ring5"
        dbPlayers(dbIndex).sRings(5) = modItemManip.ClearItemFlags(dbPlayers(dbIndex).sRings(5))
        
        
    '-----------------------------------------
    
    Case "clear.adjectives.weapon"
        dbPlayers(dbIndex).sWeapon = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sWeapon)
    Case "clear.adjectives.head"
        dbPlayers(dbIndex).sHead = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sHead)
    Case "clear.adjectives.face"
        dbPlayers(dbIndex).sFace = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sFace)
    Case "clear.adjectives.ears"
        dbPlayers(dbIndex).sEars = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sEars)
    Case "clear.adjectives.neck"
        dbPlayers(dbIndex).sNeck = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sNeck)
    Case "clear.adjectives.body"
        dbPlayers(dbIndex).sBody = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sBody)
    Case "clear.adjectives.back"
        dbPlayers(dbIndex).sBack = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sBack)
    Case "clear.adjectives.arms"
        dbPlayers(dbIndex).sArms = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sArms)
    Case "clear.adjectives.shield"
        dbPlayers(dbIndex).sShield = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sShield)
    Case "clear.adjectives.hands"
        dbPlayers(dbIndex).sHands = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sHands)
    Case "clear.adjectives.legs"
        dbPlayers(dbIndex).sLegs = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sLegs)
    Case "clear.adjectives.feet"
        dbPlayers(dbIndex).sFeet = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sFeet)
    Case "clear.adjectives.waist"
        dbPlayers(dbIndex).sWaist = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sWaist)
    Case "clear.adjectives.ring0"
        dbPlayers(dbIndex).sRings(0) = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sRings(0))
    Case "clear.adjectives.ring1"
        dbPlayers(dbIndex).sRings(1) = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sRings(1))
    Case "clear.adjectives.ring2"
        dbPlayers(dbIndex).sRings(2) = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sRings(2))
    Case "clear.adjectives.ring3"
        dbPlayers(dbIndex).sRings(3) = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sRings(3))
    Case "clear.adjectives.ring4"
        dbPlayers(dbIndex).sRings(4) = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sRings(4))
    Case "clear.adjectives.ring5"
        dbPlayers(dbIndex).sRings(5) = modItemManip.ClearItemAdjectives(dbPlayers(dbIndex).sRings(5))
    
    '-----------------------------------------
    
    Case "clear.enchant.weapon"
        dbPlayers(dbIndex).sWeapon = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sWeapon)
    Case "clear.enchant.head"
        dbPlayers(dbIndex).sHead = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sHead)
    Case "clear.enchant.face"
        dbPlayers(dbIndex).sFace = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sFace)
    Case "clear.enchant.ears"
        dbPlayers(dbIndex).sEars = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sEars)
    Case "clear.enchant.neck"
        dbPlayers(dbIndex).sNeck = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sNeck)
    Case "clear.enchant.body"
        dbPlayers(dbIndex).sBody = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sBody)
    Case "clear.enchant.back"
        dbPlayers(dbIndex).sBack = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sBack)
    Case "clear.enchant.arms"
        dbPlayers(dbIndex).sArms = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sArms)
    Case "clear.enchant.shield"
        dbPlayers(dbIndex).sShield = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sShield)
    Case "clear.enchant.hands"
        dbPlayers(dbIndex).sHands = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sHands)
    Case "clear.enchant.legs"
        dbPlayers(dbIndex).sLegs = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sLegs)
    Case "clear.enchant.feet"
        dbPlayers(dbIndex).sFeet = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sFeet)
    Case "clear.enchant.waist"
        dbPlayers(dbIndex).sWaist = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sWaist)
    Case "clear.enchant.ring0"
        dbPlayers(dbIndex).sRings(0) = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sRings(0))
    Case "clear.enchant.ring1"
        dbPlayers(dbIndex).sRings(1) = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sRings(1))
    Case "clear.enchant.ring2"
        dbPlayers(dbIndex).sRings(2) = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sRings(2))
    Case "clear.enchant.ring3"
        dbPlayers(dbIndex).sRings(3) = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sRings(3))
    Case "clear.enchant.ring4"
        dbPlayers(dbIndex).sRings(4) = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sRings(4))
    Case "clear.enchant.ring5"
        dbPlayers(dbIndex).sRings(5) = modItemManip.ClearItemEnchants(dbPlayers(dbIndex).sRings(5))
            
        '-----------------------------------------
End Select
End Sub

Public Function MakedbFlags(ByVal s As String) As String
Dim Arr() As String
Dim i As Long
Dim b As String
Dim g As String
SplitFast s, Arr, ","
b = ""
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" Then
        g = Mid$(Arr(i), InStr(1, Arr(i), ";") + 1)
        Select Case Left$(Arr(i), InStr(1, Arr(i), ";") - 1)
            Case "flag.teleport"
                b = b & "tel" & g & "|"
            Case "flag.stun"
                b = b & "stu" & g & "|"
            Case "flag.light"
                b = b & "lig" & g & "|"
            Case "flag.crits"
                b = b & "cri" & g & "|"
            Case "flag.accuracy"
                b = b & "acc" & g & "|"
            Case "flag.damage"
                b = b & "dam" & g & "|"
            Case "flag.strength"
                b = b & "str" & g & "|"
            Case "flag.agility"
                b = b & "agi" & g & "|"
            Case "flag.charm"
                b = b & "cha" & g & "|"
            Case "flag.dexterity"
                b = b & "dex" & g & "|"
            Case "flag.intellect"
                b = b & "int" & g & "|"
            Case "flag.currenthp"
                b = b & "chp" & g & "|"
            Case "flag.maxhp"
                b = b & "mhp" & g & "|"
            Case "flag.currentma"
                b = b & "cma" & g & "|"
            Case "flag.maxmana"
                b = b & "mma" & g & "|"
            Case "flag.hunger"
                b = b & "hun" & g & "|"
            Case "flag.stamina"
                b = b & "sta" & g & "|"
            Case "flag.ac"
                b = b & "cac" & g & "|"
            Case "flag.currentexp"
                b = b & "exp" & g & "|"
            Case "flag.totalexp"
                b = b & "txp" & g & "|"
            Case "flag.gold"
                b = b & "gol" & g & "|"
            Case "flag.dodge"
                b = b & "dod" & g & "|"
            Case "flag.bank"
                b = b & "ban" & g & "|"
            Case "flag.vision"
                b = b & "vis" & g & "|"
            Case "flag.maxitems"
                b = b & "mit" & g & "|"
            Case "flag.classpoints"
                b = b & "clp" & g & "|"
            Case "flag.evilpoints"
                b = b & "evi" & g & "|"
            Case "flag.resistfire"
                b = b & "el0" & g & "|"
            Case "flag.resistice"
                b = b & "el1" & g & "|"
            Case "flag.resistwater"
                b = b & "el2" & g & "|"
            Case "flag.resistlightning"
                b = b & "el3" & g & "|"
            Case "flag.resistearth"
                b = b & "el4" & g & "|"
            Case "flag.resistpoison"
                b = b & "el5" & g & "|"
            Case "flag.resistwind"
                b = b & "el6" & g & "|"
            Case "flag.resistholy"
                b = b & "el7" & g & "|"
            Case "flag.resistunhold"
                b = b & "el8" & g & "|"
            Case "flag.cansneak"
                b = b & "m02" & g & "|"
            Case "flag.spellcasting"
                b = b & "s01" & g & "|"
            Case "flag.magicresistance"
                b = b & "s03" & g & "|"
            Case "flag.perception"
                b = b & "s05" & g & "|"
            Case "flag.stealth"
                b = b & "s11" & g & "|"
            Case "flag.animalrelations"
                b = b & "s13" & g & "|"
            Case "flag.canattack"
                b = b & "m00" & g & "|"
            Case "flag.cancastspell"
                b = b & "m01" & g & "|"
            Case "flag.gibberishtalk"
                b = b & "m03" & g & "|"
            Case "flag.invisible"
                b = b & "m05" & g & "|"
            Case "flag.caneqhead"
                b = b & "m06" & g & "|"
            Case "flag.caneqface"
                b = b & "m07" & g & "|"
            Case "flag.caneqears"
                b = b & "m08" & g & "|"
            Case "flag.caneqneck"
                b = b & "m09" & g & "|"
            Case "flag.caneqbody"
                b = b & "m10" & g & "|"
            Case "flag.caneqback"
                b = b & "m11" & g & "|"
            Case "flag.caneqarms"
                b = b & "m12" & g & "|"
            Case "flag.caneqshield"
                b = b & "m13" & g & "|"
            Case "flag.caneqhands"
                b = b & "m14" & g & "|"
            Case "flag.caneqlegs"
                b = b & "m15" & g & "|"
            Case "flag.caneqfeet"
                b = b & "m16" & g & "|"
            Case "flag.caneqwaist"
                b = b & "m17" & g & "|"
            Case "flag.caneqweapon"
                b = b & "m18" & g & "|"
            Case "flag.canbedesysed"
                b = b & "m19" & g & "|"
            Case "flag.seeinvisible"
                b = b & "m20" & g & "|"
            Case "flag.seehidden"
                b = b & "m21" & g & "|"
            Case "flag.caneqring0"
                b = b & "m22" & g & "|"
            Case "flag.caneqring1"
                b = b & "m23" & g & "|"
            Case "flag.caneqring2"
                b = b & "m24" & g & "|"
            Case "flag.caneqring3"
                b = b & "m25" & g & "|"
            Case "flag.caneqring4"
                b = b & "m26" & g & "|"
            Case "flag.caneqring5"
                b = b & "m27" & g & "|"
            Case "flag.candualwield"
                b = b & "m28" & g & "|"
            Case "flag.cansteal"
                b = b & "m29" & g & "|"
            Case "flag.canbackstab"
                b = b & "m30" & g & "|"
            Case "enchantment.castspell"
                b = b & "csp" & g & "|"
            Case "enchantment.mindamagebonus"
                b = b & "mib" & g & "|"
            Case "enchantment.maxdamagebonus"
                b = b & "mab" & g & "|"
            Case "enchantment.swingbonus"
                b = b & "swi" & g & "|"
        End Select
    End If
    If DE Then DoEvents
Next
MakedbFlags = b
End Function

Public Sub sChangeDesc(dbIndex As Long, s As String)
With dbPlayers(dbIndex)
    s = ReplaceFast(s, "player.name", .sSeenAs)
    s = ReplaceFast(s, "player.seenas", .sPlayerName)
    If s = "" Then s = "0"
    .sOverrideDesc = s
End With
End Sub

Public Sub sChangeSeenAs(dbIndex As Long, s As String)
With dbPlayers(dbIndex)
    If Left$(s, 8) = "player." Then
        Select Case Mid$(s, 8)
            Case "name"
                s = .sPlayerName
            Case "seenas"
                s = .sSeenAs
        End Select
    End If
    If s = "" Then s = .sSeenAs
    .sPlayerName = s
End With
End Sub

Sub sAddEvent(dbIndex As Long, s As String)
Dim lID As Long
Dim i As Long
Dim CDB As Long
Dim m As Long
Dim n As Long
Dim sID As String
Dim sEnd As String
Dim sEx As String
Dim Arr() As String
Dim Arr2() As String
m = InStr(1, s, ",")
sID = Mid$(s, 1, m - 1)
n = InStr(m + 1, s, ",")
sEnd = Mid$(s, m + 1, n - m - 1)
'm = InStr(n + 1, s, ",")
sEx = Mid$(s, n + 1) ', Len(s) - m)
lID = 0
sEnd = ReplaceFast(sEnd, "/", ":")
sEx = ReplaceFast(sEx, "/", ":")
For i = LBound(dbEvents) To UBound(dbEvents)
    With dbEvents(i)
        If .lPlayerID = 0 Then
            lID = .lEventID
            CDB = i
            Exit For
        End If
    End With
    If DE Then DoEvents
Next
SplitFast sEnd, Arr, ":"
SplitFast sEx, Arr2, ":"
If UBound(Arr) < 5 Then Exit Sub
If UBound(Arr2) < 5 Then Exit Sub
With dbEvents(CDB)
    .lEventID = lID
    .lIsComplete = 0
    .sCustomID = sID
    .lPlayerID = dbPlayers(dbIndex).lPlayerID
    .sStartTime = modTime.TimeOfDay & "/" & modTime.MonthOfYear & ":" & modTime.udtMonths(modTime.MonthOfYear).CurDay & ":" & modTime.CurYear
    .sEndTime = modTime.AddTimeNotReal(CLng(Arr(0)), CLng(Arr(1)), CLng(Arr(2)), CLng(Arr(3)), CLng(Arr(4)), CLng(Arr(5)))
    .sExpire = modTime.AddTimeNotReal(CLng(Arr2(0)), CLng(Arr2(1)), CLng(Arr2(2)), CLng(Arr2(3)), CLng(Arr2(4)), CLng(Arr2(5)))
End With
End Sub

Public Sub sChangeAppearance(dbIndex As Long, a As String, s As String)
Select Case a
    Case "appearance.hairlength"
        modAppearance.SetPlayerAppearanceNumber dbIndex, [Hair Length], Val(s)
    Case "appearance.haircolor"
        modAppearance.SetPlayerAppearanceNumber dbIndex, [Hair Color], Val(s)
    Case "appearance.hairstyle"
        modAppearance.SetPlayerAppearanceNumber dbIndex, [Hair Style], Val(s)
    Case "appearance.eyecolor"
        modAppearance.SetPlayerAppearanceNumber dbIndex, [Eye Color], Val(s)
    Case "appearance.moustache"
        modAppearance.SetPlayerAppearanceNumber dbIndex, moustache, Val(s)
    Case "appearance.beard"
        modAppearance.SetPlayerAppearanceNumber dbIndex, beard, Val(s)
End Select
End Sub

Sub sEraseEvent(dbIndex As Long, s As String)
Dim dbEventId As Long
dbEventId = GetEventID(s, dbPlayers(dbIndex).lPlayerID)
If dbEventId = 0 Then Exit Sub
With dbEvents(dbEventId)
    .lIsComplete = 0
    .lPlayerID = 0
    .sCustomID = "0"
    .sEndTime = "0"
    .sExpire = "0"
    .sStartTime = "0"
End With
End Sub

Sub sChangeTime(dbIndex As Long, s As String)
Dim m As Long
Dim n As Long
Dim iHours As Long
Dim iMinutes As Long
Dim iSeconds As Long
m = InStr(1, s, ",")
iHours = CLng(Mid$(s, 1, m - 1))
n = InStr(m + 1, s, ",")
iMinutes = CLng(Mid$(s, m + 1, n - m - 1))
iSeconds = CLng(Mid$(s, n + 1))
modTime.AddTime iHours, iMinutes, iSeconds
End Sub

Sub sShowTime(Index As Long)
WrapAndSend Index, "The current time is " & modTime.TimeOfDay & modTime.GetDayNight & vbCrLf
End Sub

Sub sChangeFlag(dbIndex As Long, s As String)
Dim sFlag As String
Dim sChange As String
Dim n As Long
n = InStr(1, s, ",")
sFlag = Left$(s, n - 1)
sChange = Mid$(s, n + 1)
With dbPlayers(dbIndex)
    Select Case sFlag
        Case "1"
            .iFlag1 = CLng(sChange)
        Case "2"
            .iFlag2 = CLng(sChange)
        Case "3"
            .iFlag3 = CLng(sChange)
        Case "4"
            .iFlag4 = CLng(sChange)
    End Select
End With
End Sub

Sub sAddStamina(dbIndex As Long, s As String)
Dim dAdd As Double
   On Error GoTo sAddStatTrain_Error

dAdd = CDbl(Val(s))
With dbPlayers(dbIndex)
    .dStamina = .dStamina + dAdd
End With

   On Error GoTo 0
   Exit Sub

sAddStatTrain_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddStatTrain of Module modScripts"
End Sub

Sub sAddHunger(dbIndex As Long, s As String)
Dim dAdd As Double
   On Error GoTo sAddStatTrain_Error

dAdd = CDbl(Val(s))
With dbPlayers(dbIndex)
    .dHunger = .dHunger + dAdd
End With

   On Error GoTo 0
   Exit Sub

sAddStatTrain_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddStatTrain of Module modScripts"
End Sub

Sub sAddStatTrain(dbIndex As Long, s As String)
Dim dAdd As Double
   On Error GoTo sAddStatTrain_Error

dAdd = CDbl(Val(s))
With dbPlayers(dbIndex)
    .iIsReadyToTrain = .iIsReadyToTrain + dAdd
    If .iIsReadyToTrain < 0 Then .iIsReadyToTrain = 0
End With

   On Error GoTo 0
   Exit Sub

sAddStatTrain_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddStatTrain of Module modScripts"
End Sub

Sub sChangeBank(dbIndex As Long, s As String)
Dim sdGold As Double
   On Error GoTo sChangeBank_Error

sdGold = CDbl(Val(s))
With dbPlayers(dbIndex)
    .dBank = .dBank + sdGold
    If .dBank < 0 Then .dBank = 0
End With

   On Error GoTo 0
   Exit Sub

sChangeBank_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sChangeBank of Module modScripts"
End Sub

Sub sAddQuest(dbIndex As Long, s As String)

   On Error GoTo sAddQuest_Error
Dim sFlag As String
Dim sChange As String
Dim n As Long
n = InStr(1, s, ",")
sFlag = Left$(s, n - 1)
sChange = Mid$(s, n + 1)
With dbPlayers(dbIndex)
    Select Case sFlag
        Case "1"
            .sQuest1 = sChange
        Case "2"
            .sQuest2 = sChange
        Case "3"
            .sQuest3 = sChange
        Case "4"
            .sQuest4 = sChange
    End Select
End With

   On Error GoTo 0
   Exit Sub

sAddQuest_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddQuest of Module modScripts"
End Sub

Sub sTakeFam(Index As Long)
RemoveStats Index, True
End Sub

Sub sGiveFam(dbIndex As Long, s As String)
Dim i As Long
Dim dbFamId As Long
   On Error GoTo sGiveFam_Error
If Not IsNumeric(s) Then s = CStr(GetFamID(CLng(s))) Else s = CStr(GetFamID(, s))
dbFamId = s
i = dbFamiliars(CLng(s)).iID
s = dbFamiliars(CLng(s)).sFamName
With dbPlayers(dbIndex)
    RemoveStats .iIndex
    .sFamName = s
    .lFamID = i
    .lFamMHP = RndNumber(CDbl(dbFamiliars(dbFamId).lStartHPMin), CDbl(dbFamiliars(dbFamId).lStartHPMax))
    .dFamCEXP = 0
    .dFamEXPN = dbFamiliars(dbFamId).dEXPPerLevel
    .dFamTEXP = 0
    .lFamAcc = 0
    .sFamCustom = "0"
    .lFamCHP = .lFamMHP
    .lFamMin = dbFamiliars(dbFamId).lMinDam
    .lFamMax = dbFamiliars(dbFamId).lMaxDam
    .lFamLevel = 1
    AddStats .iIndex
End With
   On Error GoTo 0
   Exit Sub

sGiveFam_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sGiveFam of Module modScripts"
End Sub

Sub sAddDodge(dbIndex As Long, s As String)
   On Error GoTo sAddDodge_Error

If Val(s) < 1000 Then
    With dbPlayers(dbIndex)
        .iDodge = .iDodge + CLng(Val(s))
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddDodge_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddDodge of Module modScripts"
End Sub

Sub sAddCrits(dbIndex As Long, s As String)
   On Error GoTo sAdCrits_Error

If Val(s) < 500 Then
    With dbPlayers(dbIndex)
        .iCrits = .iCrits + CLng(Val(s))
    End With
End If
   On Error GoTo 0
   Exit Sub

sAdCrits_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAdCrits of Module modScripts"
End Sub

Sub sAddAcc(dbIndex As Long, s As String)
   On Error GoTo sAddAcc_Error

If Val(s) < 100 Then
    With dbPlayers(dbIndex)
        .iAcc = .iAcc + CLng(Val(s))
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddAcc_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddAcc of Module modScripts"
End Sub

Sub sChangeWeaponType(dbIndex As Long, s As String)
   On Error GoTo sChangeWeaponType_Error

If Val(s) < 18 And Val(s) > -1 Then
    With dbPlayers(dbIndex)
        .iWeapons = CLng(Val(s))
    End With
End If
   On Error GoTo 0
   Exit Sub

sChangeWeaponType_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sChangeWeaponType of Module modScripts"
End Sub

Sub sChangeArmorType(dbIndex As Long, s As String)
   On Error GoTo sChangeArmorType_Error

If Val(s) < 25 And Val(s) > -1 Then
    With dbPlayers(dbIndex)
        .iArmorType = CLng(Val(s))
    End With
End If
   On Error GoTo 0
   Exit Sub

sChangeArmorType_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sChangeArmorType of Module modScripts"
End Sub

Sub sAddSC(dbIndex As Long, s As String)
   On Error GoTo sAddSC_Error

If Val(s) < 200 Then
    modMiscFlag.SetStatsPlus dbIndex, [Spell Casting Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Spell Casting Bonus]) + Val(s)
End If
   On Error GoTo 0
   Exit Sub

sAddSC_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddSC of Module modScripts"
End Sub

Sub sTakeItem(dbIndex As Long, s As String)
'Dim dbIndex As Long
Dim iItemID As Long
   On Error GoTo sTakeItem_Error
If Not IsNumeric(s) Then
    s = modGetData.GetItemNumFromName(s)
    If s = "(-1)" Then Exit Sub
End If
iItemID = GetItemID(, CLng(s))
If iItemID = 0 Then Exit Sub
modItemManip.RemoveItemFromInv dbIndex, dbItems(iItemID).iID

   On Error GoTo 0
   Exit Sub

sTakeItem_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sTakeItem of Module modScripts"
End Sub

Sub sGiveGold(dbIndex As Long, s As String)
Dim sdGold As Double
   On Error GoTo sGiveGold_Error

sdGold = CDbl(Val(s))
If sdGold < 2000000# Then
    With dbPlayers(dbIndex)
        .dGold = .dGold + sdGold
        If .dGold < 0 Then .dGold = 0
    End With
End If
   On Error GoTo 0
   Exit Sub

sGiveGold_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sGiveGold of Module modScripts"
End Sub

'Sub SendErr(Index as long, e As String)
'Dim m as long, n as long
'm = InStr(1, e, "*")
'n = InStr(m + 1, e, "*")
'WrapAndSend Index, RED & Mid$(e, m + 1, n - m - 1) & vbCrLf
'End Sub

Sub sTeleport(dbIndex As Long, s As String)
Dim lTargetRoom As Long
   On Error GoTo sTeleport_Error

lTargetRoom = CLng(Val(s))
If GetMapIndex(lTargetRoom) = 0 Then Exit Sub Else lTargetRoom = GetMapIndex(lTargetRoom)
With dbPlayers(dbIndex)
    .lLocation = dbMap(lTargetRoom).lRoomID
End With

   On Error GoTo 0
   Exit Sub

sTeleport_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sTeleport of Module modScripts"
End Sub

Sub sPartyTel(dbIndex As Long, s As String)
Dim PartyMembers As String
Dim tArr() As String
Dim i As Long
Dim lRoom As Long
lRoom = CLng(Val(s))
lRoom = GetMapIndex(lRoom)
If lRoom = 0 Then Exit Sub
With dbPlayers(dbIndex)
    If .iPartyLeader > 0 Then
        PartyMembers = .sParty
        PartyMembers = ReplaceFast(PartyMembers, ":", "")
        SplitFast PartyMembers, tArr, ";"
        For i = 0 To UBound(tArr) 'loop the array
            If tArr(i) <> "" Then
                With dbPlayers(GetPlayerIndexNumber(CLng(tArr(i))))
                    .lLocation = dbMap(lRoom).lRoomID
                End With
            End If
            If DE Then DoEvents
        Next i
    End If
    .lLocation = dbMap(lRoom).lRoomID
End With
End Sub

Sub sPartySend(dbIndex As Long, s As String)
Dim PartyMembers As String
Dim tArr() As String
Dim i As Long
With dbPlayers(dbIndex)
    If .iPartyLeader > 0 Then
        sSend .iIndex, s
        If DE Then DoEvents
        PartyMembers = .sParty
        PartyMembers = ReplaceFast(PartyMembers, ":", "")
        SplitFast PartyMembers, tArr, ";"
        For i = 0 To UBound(tArr) 'loop the array
            If tArr(i) <> "" Then
                With dbPlayers(GetPlayerIndexNumber(CLng(tArr(i))))
                    sSend .iIndex, s
                End With
            End If
            If DE Then DoEvents
        Next i
    Else
        sSend .iIndex, s
    End If
End With
End Sub


Sub sGainHP(dbIndex As Long, s As String)
Dim lGainHP As Long
   On Error GoTo sGainHP_Error

lGainHP = CLng(Val(s))
If lGainHP < 50000# Then
    With dbPlayers(dbIndex)
        .lHP = .lHP + lGainHP
        If .lHP > .lMaxHP Then .lHP = .lMaxHP
    End With
End If
   On Error GoTo 0
   Exit Sub

sGainHP_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sGainHP of Module modScripts"
End Sub

Sub sGainMA(dbIndex As Long, s As String)
Dim lGainMA As Long
   On Error GoTo sGainMA_Error

lGainMA = CLng(Val(s))
If lGainMA < 35000# Then
    With dbPlayers(dbIndex)
        .lMana = .lMana + lGainMA
        If .lMana > .lMaxMana Then .lMana = .lMaxMana
    End With
End If
   On Error GoTo 0
   Exit Sub

sGainMA_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sGainMA of Module modScripts"
End Sub

Sub sGiveItem(dbIndex As Long, s As String)
Dim sItem As String
Dim iItemID As Long
   On Error GoTo sGiveItem_Error

sItem = s
If Not IsNumeric(sItem) Then
    sItem = GetItemID(sItem)
    If sItem = "0" Then Exit Sub
    iItemID = dbItems(CLng(sItem)).iID
Else
    sItem = GetItemID(, CLng(sItem))
    If sItem = "0" Then Exit Sub
    iItemID = dbItems(CLng(sItem)).iID
End If
With dbPlayers(dbIndex)
    If modGetData.GetPlayersTotalItems(.iIndex, dbIndex) + 1 <= modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) Then
        If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
        .sInventory = .sInventory & ":" & iItemID & "/" & dbItems(CLng(sItem)).lDurability & "/E{}F{}A{}B{0|0|0|0}/" & dbItems(CLng(sItem)).iUses & ";"
    End If
End With

   On Error GoTo 0
   Exit Sub

sGiveItem_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sGiveItem of Module modScripts"
End Sub

Sub sAddSpell(dbIndex As Long, s As String)
Dim iSpellID As Long
Dim iPlayerID As Long
   On Error GoTo sAddSpell_Error

Spell$ = s
If Not IsNumeric(Spell$) Then
    iSpellID = GetSpellID(Spell$)
Else
    iSpellID = GetSpellID(, CLng(Spell$))
End If
If iSpellID = 0 Then Exit Sub
Spell$ = dbSpells(iSpellID).lID
iPlayerID = dbIndex
If dbPlayers(iPlayerID).iSpellLevel >= dbSpells(iSpellID).iLevel Then
    If dbPlayers(iPlayerID).iSpellType = dbSpells(iSpellID).iType Then
        If InStr(1, dbPlayers(iPlayerID).sSpells, ":" & Spell$ & ";") Then Exit Sub
        If modSC.FastStringComp(dbPlayers(iPlayerID).sSpells, "0") Then dbPlayers(iPlayerID).sSpells = ""
        dbPlayers(iPlayerID).sSpells = dbPlayers(iPlayerID).sSpells & ":" & Spell$ & ";"
        If modSC.FastStringComp(dbPlayers(iPlayerID).sSpellShorts, "0") Then dbPlayers(iPlayerID).sSpellShorts = ""
        dbPlayers(iPlayerID).sSpellShorts = dbPlayers(iPlayerID).sSpellShorts & dbSpells(iSpellID).sShort & ";"
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If

   On Error GoTo 0
   Exit Sub

sAddSpell_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddSpell of Module modScripts"
End Sub

Sub sAddEXP(dbIndex As Long, s As String)
Dim ddEXP As Double
   On Error GoTo sAddEXP_Error

ddEXP = CDbl(Val(s))
If ddEXP <= 999999999# Then
    With dbPlayers(dbIndex)
        .dEXP = .dEXP + ddEXP
        .dTotalEXP = .dTotalEXP + ddEXP
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddEXP_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddEXP of Module modScripts"
End Sub

Sub sAddStr(dbIndex As Long, s As String)
Dim iAdd As Long
   On Error GoTo sAddStr_Error

If Val(s) < 1000# Then
    iAdd = CLng(Val(s))
    With dbPlayers(dbIndex)
        .iStr = .iStr + iAdd
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddStr_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddStr of Module modScripts"
End Sub

Sub sAddCha(dbIndex As Long, s As String)
Dim iAdd As Long
   On Error GoTo sAddCha_Error

If Val(s) < 1000# Then
    iAdd = CLng(Val(s))
    With dbPlayers(dbIndex)
        .iCha = .iCha + iAdd
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddCha_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddCha of Module modScripts"
End Sub

Sub sAddDex(dbIndex As Long, s As String)
Dim iAdd As Long
   On Error GoTo sAddDex_Error

If Val(s) < 1000# Then
    iAdd = CLng(Val(s))
    With dbPlayers(dbIndex)
        .iDex = .iDex + iAdd
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddDex_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddDex of Module modScripts"
End Sub

Sub sAddInt(dbIndex As Long, s As String)
Dim iAdd As Long
   On Error GoTo sAddInt_Error

If Val(s) < 1000# Then
    iAdd = CLng(Val(s))
    With dbPlayers(dbIndex)
        .iInt = .iInt + iAdd
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddInt_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddInt of Module modScripts"
End Sub

Sub sAddAgil(dbIndex As Long, s As String)
Dim iAdd As Long
   On Error GoTo sAddAgil_Error

If Val(s) < 1000# Then
    iAdd = CLng(Val(s))
    With dbPlayers(dbIndex)
        .iAgil = .iAgil + iAdd
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddAgil_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddAgil of Module modScripts"
End Sub

Sub sAddEvil(dbIndex As Long, s As String)
Dim iAdd As Long
   On Error GoTo sAddEvil_Error

If Val(s) < 1000# Then
    iAdd = CLng(Val(s))
    With dbPlayers(dbIndex)
        .iEvil = .iEvil + iAdd
    End With
End If
   On Error GoTo 0
   Exit Sub

sAddEvil_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sAddEvil of Module modScripts"
End Sub

Function sDeCode(ByVal s As String, Optional CRLF As String = vbCrLf) As String
   On Error GoTo sSend_Error
Dim t As Boolean
Dim m As Long
Dim n As Long
Dim Message As String
Dim Color As String
Dim b As String
b = s
s = LCaseFast(s)
Do Until t = True Or s = ""
    If Left$(s, 6) = "color." Then
        n = 6
        m = InStr(1, s, "&")
        If m <> 0 Then
            Color = Trim$(Mid$(s, n + 1, m - n - 1))
            Color = GetColorCode(Color)
            Message = Message & Color
            s = Mid$(s, m + 1)
            s = TrimIt(s)
            b = Mid$(b, m + 1)
            b = TrimIt(b)
        Else
            Color = Mid$(s, n + 1)
            Color = GetColorCode(Color)
            Message = Message & Color
            t = True
        End If
    ElseIf Left$(s, 7) = "newline" Then
        n = 7
        m = InStr(1, s, "&")
        If m <> 0 Then
            Message = Message & vbCrLf
            s = Mid$(s, m + 1)
            s = TrimIt(s)
            b = Mid$(b, m + 1)
            b = TrimIt(b)
        Else
            Message = Message & vbCrLf
            t = True
        End If
    ElseIf Left$(s, 1) = ";" Then
        If s = ";" Then s = "": Exit Do
        n = 1
        m = InStr(n + 1, s, ";")
        Message = Message & Mid$(b, n + 1, m - n - 1)
        s = Mid$(s, m + 1)
        b = Mid$(b, m + 1)
        Do Until s = "" Or Left$(s, 6) = "color." Or Left$(s, 7) = "newline" Or Left$(s, 1) = ";"
            s = Mid$(s, 2)
            b = Mid$(b, 2)
            If DE Then DoEvents
        Loop
        If s = "" Then t = True
    Else
        Do Until s = "" Or Left$(s, 6) = "color." Or Left$(s, 7) = "newline" Or Left$(s, 1) = ";"
            Message = Message & Mid$(b, 1, 1)
            s = Mid$(s, 2)
            b = Mid$(b, 2)
            If DE Then DoEvents
        Loop
        If s = "" Then t = True
    End If
    'Debug.Print Message
    If DE Then DoEvents
Loop
sDeCode = ReplaceFakeANSI(0, Message)

   On Error GoTo 0
   Exit Function

sSend_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sSend of Module modScripts"
End Function

Sub sSend(Index As Long, ByVal s As String, Optional CRLF As String = vbCrLf, Optional CombatMessage As Boolean = False, Optional ByRef Messages1 As String = "")
   On Error GoTo sSend_Error
Dim t As Boolean
Dim m As Long
Dim n As Long
Dim Message As String
Dim Color As String
Dim b As String
b = s
s = LCaseFast(s)
Do Until t = True Or s = ""
    If Left$(s, 6) = "color." Then
        n = 6
        m = InStr(1, s, "&")
        If m <> 0 Then
            Color = Trim$(Mid$(s, n + 1, m - n - 1))
            Color = GetColorCode(Color)
            Message = Message & Color
            s = Mid$(s, m + 1)
            s = TrimIt(s)
            b = Mid$(b, m + 1)
            b = TrimIt(b)
        Else
            Color = Mid$(s, n + 1)
            Color = GetColorCode(Color)
            Message = Message & Color
            t = True
        End If
    ElseIf Left$(s, 7) = "newline" Then
        n = 7
        m = InStr(1, s, "&")
        If m <> 0 Then
            Message = Message & vbCrLf
            s = Mid$(s, m + 1)
            s = TrimIt(s)
            b = Mid$(b, m + 1)
            b = TrimIt(b)
        Else
            Message = Message & vbCrLf
            t = True
        End If
    ElseIf Left$(s, 1) = ";" Then
        If s = ";" Then s = "": Exit Do
        n = 1
        m = InStr(n + 1, s, ";")
        Message = Message & Mid$(b, n + 1, m - n - 1)
        s = Mid$(s, m + 1)
        b = Mid$(b, m + 1)
        Do Until s = "" Or Left$(s, 6) = "color." Or Left$(s, 7) = "newline" Or Left$(s, 1) = ";"
            s = Mid$(s, 2)
            b = Mid$(b, 2)
            If DE Then DoEvents
        Loop
        If s = "" Then t = True
    Else
        Do Until s = "" Or Left$(s, 6) = "color." Or Left$(s, 7) = "newline" Or Left$(s, 1) = ";"
            Message = Message & Mid$(b, 1, 1)
            s = Mid$(s, 2)
            b = Mid$(b, 2)
            If DE Then DoEvents
        Loop
        If s = "" Then t = True
    End If
    'Debug.Print Message
    If DE Then DoEvents
Loop
If CombatMessage = False Then
    WrapAndSend Index, Message & WHITE & CRLF
Else
    Messages1 = Messages1 & Message & WHITE & vbCrLf
End If

   On Error GoTo 0
   Exit Sub

sSend_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sSend of Module modScripts"
End Sub

Public Function GetColorCode(sInput As String) As String
Select Case sInput
    Case "red"
        GetColorCode = RED
    Case "green"
        GetColorCode = GREEN
    Case "yellow"
        GetColorCode = YELLOW
    Case "blue"
        GetColorCode = BLUE
    Case "magneta"
        GetColorCode = MAGNETA
    Case "lightblue"
        GetColorCode = LIGHTBLUE
    Case "white"
        GetColorCode = WHITE
    Case "bgred"
        GetColorCode = BGRED
    Case "bggreen"
        GetColorCode = BGGREEN
    Case "bgyellow"
        GetColorCode = BGYELLOW
    Case "bgblue"
        GetColorCode = BGBLUE
    Case "bgpurple"
        GetColorCode = BGPURPLE
    Case "bglightblue"
        GetColorCode = BGLIGHTBLUE
    Case "brightyellow"
        GetColorCode = BRIGHTYELLOW
    Case "brightgreen"
        GetColorCode = BRIGHTGREEN
    Case "brightred"
        GetColorCode = BRIGHTRED
    Case "brightblue"
        GetColorCode = BRIGHTBLUE
    Case "brightmagneta"
        GetColorCode = BRIGHTMAGNETA
    Case "brightlightblue"
        GetColorCode = BRIGHTLIGHTBLUE
    Case "brightwhite"
        GetColorCode = BRIGHTWHITE
End Select
End Function

Sub sCastSp(dbIndex As Long, s As String)

   On Error GoTo sCastSp_Error
    'CastSpell(PLAYER or ROOM,23,CASTER or default)
    
    Dim sTarget As String
    Dim sSpell As String
    Dim m As Long
    Dim n As Long
    Dim sCaster As String
    Dim Index As Long
    Index = dbPlayers(dbIndex).iIndex
    m = InStr(1, s, ",")
    sTarget = Left$(s, m - 1)
    n = InStr(m + 1, s, ",")
    sSpell = Mid$(s, m + 1, n - m - 1)
    sCaster = Mid$(s, n + 1)
    If Not IsNumeric(sSpell) Then
        m = GetSpellID(sSpell)
    Else
        m = GetSpellID(, CLng(sSpell))
    End If
    If m = 0 Then Exit Sub
    Select Case sCaster
        Case "default"
            sCaster = "Unknown"
    End Select
    Select Case LCaseFast(sTarget)
        Case "player"
            modSpells.DoNonCombatSpell Index, dbIndex, m, , , True, sCaster
'        Case "player"
'            With dbSpells(m)
'                If .iuse = 1 Then
'                    modSpells.DoDamageToPlayerWithSpell Index, Clng(dbSpells(m).lID)
'                ElseIf .iuse = 0 Then
'                    modSpells.HealingSpell True, Index, Clng(dbSpells(m).lID), , False
'                ElseIf .iuse = 2 Then
'                    modSpells.TeleportSpell Index, Clng(dbSpells(m).lID), False
'                ElseIf .iuse = 3 Then
'                    modSpells.BlessMe Clng(m), Clng(dbPlayers(GetPlayerIndexNumber(Index)).lPlayerID), False
'                End If
'            End With
'        Case "room"
'
    End Select
   On Error GoTo 0
   Exit Sub

sCastSp_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sCastSp of Module modScripts"
End Sub

Sub sChangeRace(dbIndex As Long, s As String)
Dim sChange As String
   On Error GoTo sChangeRace_Error

sChange = s
With dbPlayers(dbIndex)
    .sRace = sChange
End With

   On Error GoTo 0
   Exit Sub

sChangeRace_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sChangeRace of Module modScripts"
End Sub

Sub sChangeClass(dbIndex As Long, s As String)
Dim sChange As String
   On Error GoTo sChangeClass_Error

sChange = s
With dbPlayers(dbIndex)
    .sClass = sChange
End With

   On Error GoTo 0
   Exit Sub

sChangeClass_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sChangeClass of Module modScripts"
End Sub

Sub sSendRoom(dbIndex As Long, s As String)
   On Error GoTo sSendRoom_Error
Dim t As Boolean
Dim m As Long
Dim n As Long
Dim Message As String
Dim Color As String
Dim Index As Long
Index = dbPlayers(dbIndex).iIndex
s = ReplaceFast(s, "<%p>", dbPlayers(dbIndex).sPlayerName)
Do Until t = True Or s = ""
    If Left$(s, 6) = "color." Then
        n = 6
        m = InStr(1, s, "&")
        If m <> 0 Then
            Color = Trim$(Mid$(s, n + 1, m - n - 1))
            Color = GetColorCode(Color)
            Message = Message & Color
            s = Mid$(s, m + 1)
            s = TrimIt(s)
        Else
            Color = Mid$(s, n + 1)
            Color = GetColorCode(Color)
            Message = Message & Color
            t = True
        End If
    ElseIf Left$(s, 7) = "newline" Then
        n = 7
        m = InStr(1, s, "&")
        If m <> 0 Then
            Message = Message & vbCrLf
            s = Mid$(s, m + 1)
            s = TrimIt(s)
        Else
            Message = Message & vbCrLf
            t = True
        End If
    ElseIf Left$(s, 1) = ";" Then
        If s = ";" Then s = "": Exit Do
        n = 1
        m = InStr(n + 1, s, ";")
        Message = Message & Mid$(s, n + 1, m - n - 1)
        s = Mid$(s, m + 1)
        Do Until s = "" Or Left$(s, 6) = "color." Or Left$(s, 7) = "newline" Or Left$(s, 1) = ";"
            s = Mid$(s, 2)
            If DE Then DoEvents
        Loop
        If s = "" Then t = True
    Else
        Do Until s = "" Or Left$(s, 6) = "color." Or Left$(s, 7) = "newline" Or Left$(s, 1) = ";"
            Message = Message & Mid$(s, 1, 1)
            s = Mid$(s, 2)
            If DE Then DoEvents
        Loop
        If s = "" Then t = True
    End If
    'Debug.Print Message
    If DE Then DoEvents
Loop
'Message = StrConv(Message, vbProperCase)
SendToAllInRoom Index, Message & WHITE & vbCrLf, CStr(dbPlayers(dbIndex).lLocation)

   On Error GoTo 0
   Exit Sub

sSendRoom_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sSendRoom of Module modScripts"
End Sub

Public Function CheckTimeDif(sSign As String, OtherTime As String) As Boolean
Dim l As Long, m As Long, n As Long
Dim LL As Long, mm As Long, nn As Long
Dim i As Long, j As Long

i = InStr(1, OtherTime, ":")
l = CLng(Mid$(OtherTime, 1, i - 1))
j = InStr(i + 1, OtherTime, ":")
m = CLng(Mid$(OtherTime, i + 1, j - i - 1))
j = j + 1
n = CLng(Mid$(OtherTime, j))

i = InStr(1, modTime.TimeOfDay, ":")
LL = CLng(Mid$(modTime.TimeOfDay, 1, i - 1))
j = InStr(i + 1, modTime.TimeOfDay, ":")
mm = CLng(Mid$(modTime.TimeOfDay, i + 1, j - i - 1))
j = j + 1
nn = CLng(Mid$(modTime.TimeOfDay, j))

Select Case sSign
    Case ">"
        If LL > l Then
            CheckTimeDif = True
        ElseIf LL = l Then
            If mm > m Then
                CheckTimeDif = True
            ElseIf mm = m Then
                If nn > n Then
                    CheckTimeDif = True
                Else
                    CheckTimeDif = False
                End If
            Else
                CheckTimeDif = False
            End If
        Else
            CheckTimeDif = False
        End If
    Case "<"
        If LL < l Then
            CheckTimeDif = True
        ElseIf LL = l Then
            If mm < m Then
                CheckTimeDif = True
            ElseIf mm = m Then
                If nn < n Then
                    CheckTimeDif = True
                Else
                    CheckTimeDif = False
                End If
            Else
                CheckTimeDif = False
            End If
        Else
            CheckTimeDif = False
        End If
    Case "="
        If (LL = l) And (mm = m) And (nn = n) Then
            CheckTimeDif = True
        Else
            CheckTimeDif = False
        End If
    Case ">="
        If LL > l Then
            CheckTimeDif = True
        ElseIf LL = l Then
            If mm > m Then
                CheckTimeDif = True
            ElseIf mm = m Then
                If nn >= n Then
                    CheckTimeDif = True
                Else
                    CheckTimeDif = False
                End If
            Else
                CheckTimeDif = False
            End If
        Else
            CheckTimeDif = False
        End If
    Case "<="
        If LL < l Then
            CheckTimeDif = True
        ElseIf LL = l Then
            If mm < m Then
                CheckTimeDif = True
            ElseIf mm = m Then
                If nn <= n Then
                    CheckTimeDif = True
                Else
                    CheckTimeDif = False
                End If
            Else
                CheckTimeDif = False
            End If
        Else
            CheckTimeDif = False
        End If
    Case "<>"
        If LL <> l Then
            CheckTimeDif = True
        ElseIf LL = l Then
            If mm <> m Then
                CheckTimeDif = True
            ElseIf mm = m Then
                If nn <> n Then
                    CheckTimeDif = True
                Else
                    CheckTimeDif = False
                End If
            Else
                CheckTimeDif = False
            End If
        Else
            CheckTimeDif = False
        End If
End Select
End Function

Public Function CheckDateDif(sSign As String, OtherTime As String) As Boolean
Dim l As Long, m As Long, n As Long
Dim LL As Long, mm As Long, nn As Long
Dim i As Long, j As Long

i = InStr(1, OtherTime, ":")
l = CLng(Mid$(OtherTime, 1, i - 1))
j = InStr(i + 1, OtherTime, ":")
m = CLng(Mid$(OtherTime, i + 1, j - i - 1))
j = j + 1
n = CLng(Mid$(OtherTime, j))

LL = CLng(modTime.MonthOfYear)
mm = CLng(modTime.udtMonths(modTime.MonthOfYear).CurDay)
nn = CLng(modTime.CurYear)

Select Case sSign
    Case ">"
        If n > nn Then
            CheckDateDif = True
        ElseIf n = nn Then
            If l > LL Then
                CheckDateDif = True
            ElseIf l = LL Then
                If m > mm Then
                    CheckDateDif = True
                Else
                    CheckDateDif = False
                End If
            Else
                CheckDateDif = False
            End If
        Else
            CheckDateDif = False
        End If
    Case "<"
        If n < nn Then
            CheckDateDif = True
        ElseIf n = nn Then
            If l < LL Then
                CheckDateDif = True
            ElseIf l = LL Then
                If m < mm Then
                    CheckDateDif = True
                Else
                    CheckDateDif = False
                End If
            Else
                CheckDateDif = False
            End If
        Else
            CheckDateDif = False
        End If
    Case "="
        If n = nn Then
            If l = LL Then
                If m = mm Then
                    CheckDateDif = True
                Else
                    CheckDateDif = False
                End If
            Else
                CheckDateDif = False
            End If
        Else
            CheckDateDif = False
        End If
    Case ">="
        If n > nn Then
            CheckDateDif = True
        ElseIf n = nn Then
            If l > LL Then
                CheckDateDif = True
            ElseIf l = LL Then
                If m >= mm Then
                    CheckDateDif = True
                Else
                    CheckDateDif = False
                End If
            Else
                CheckDateDif = False
            End If
        Else
            CheckDateDif = False
        End If
    Case "<="
        If n < nn Then
            CheckDateDif = True
        ElseIf n = nn Then
            If l < LL Then
                CheckDateDif = True
            ElseIf l = LL Then
                If m <= mm Then
                    CheckDateDif = True
                Else
                    CheckDateDif = False
                End If
            Else
                CheckDateDif = False
            End If
        Else
            CheckDateDif = False
        End If
    Case "<>"
        If n <> nn Then
            CheckDateDif = True
        ElseIf n = nn Then
            If l <> LL Then
                CheckDateDif = True
            ElseIf l = LL Then
                If m <> mm Then
                    CheckDateDif = True
                Else
                    CheckDateDif = False
                End If
            Else
                CheckDateDif = False
            End If
        Else
            CheckDateDif = False
        End If
End Select
End Function




'Public Function OLDSCRIPT(Index As Long, Optional Room As Long = -1, Optional lItem As Long = -1) As Boolean
''Dim s As String
'Dim tArr1() As String, tArr2() As String
'Dim i As Long, j As Long
'Dim m As Long, n As Long
'Dim s As String, l As Long
'Dim iItemID As Long
'Dim sItemID As String
'Dim aryClasses() As String
'Dim aryRaces() As String
'Dim rndNumber1 As Double, rndNumber2 As Double
'Dim intCompare As Long, theSign As String
'Dim TempVal As String, theRndNumber As Long
'Dim Checking As String
'Dim Comparing As String
'Dim sSign As String
''If x(Index) = "" Then sScripting = False: Exit Function
'If Room <> -1 Then
'    With dbMap(GetMapIndex(Room))
'        If modSC.FastStringComp(.sScript, "0") Then sScripting = False: Exit Function
'        s = .sScript
'    End With
'ElseIf lItem <> -1 Then
'    With dbItems(GetItemID(, lItem))
'        If modSC.FastStringComp(.sScript, "0") Then sScripting = False: Exit Function
'        s = .sScript
'    End With
'End If
's = ReplaceFast(s, vbCrLf, "")
'SplitFast LCaseFast(s), tArr1, ";"
'For i = LBound(tArr1) To UBound(tArr1)
'    Erase tArr2
'    SplitFast tArr1(i), tArr2, ":"
'    For j = LBound(tArr2) To UBound(tArr2)
'        s = LCaseFast(tArr2(j))
'        If modSC.FastStringComp(LCaseFast(Left$(s, 3)), "if ") Then
'            If InStr(1, s, "message=") Then
'                m = InStr(1, s, "*")
'                n = InStr(m + 1, s, "*")
'                If modSC.FastStringComp(TrimIt(LCaseFast(x(Index))), LCaseFast(Mid$(s, m + 1, n - m - 1))) Then
'                    If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                        DoAbil Index, tArr2(j + 1)
'                        sScripting = True
'                        Exit For
'                    Else
'                        'do nothing
'                    End If
'                    'message=*test*
'                Else
'                    Exit For
'                End If
'            ElseIf InStr(1, s, "in(") Then
'                m = InStr(1, s, "*")
'                n = InStr(m + 1, s, "*")
'                If InStr(1, LCaseFast(x(Index)), LCaseFast(Mid$(s, m + 1, n - m - 1))) Then
'                    'in(*test*)
'                    If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                        DoAbil Index, tArr2(j + 1)
'                        sScripting = True
'                        Exit For
'                    End If
'                    'message=*test*
'                Else
'                    Exit For
'                End If
'            ElseIf InStr(1, s, "haveitem(") Then
'                m = InStr(1, s, "*")
'                n = InStr(m + 1, s, "*")
'                sItemID = LCaseFast(Mid$(s, m + 1, n - m - 1))
'                If Not IsNumeric(sItemID) Then
'                    iItemID = GetItemID(sItemID)
'                Else
'                    iItemID = GetItemID(, CLng(sItemID))
'                End If
'                If iItemID = 0 Then
'                    Exit For
'                End If
'                iItemID = dbItems(iItemID).iID
'                If InStr(1, dbPlayers(GetPlayerIndexNumber(Index)).sInventory, ":" & iItemID & "/") Then
'                    'HaveItem(*test*)
'                    If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                        DoAbil Index, tArr2(j + 1)
'                        sScripting = True
'                        Exit For
'                    End If
'                    'message=*test*
'                Else
'                    Exit For
'                End If
'            ElseIf InStr(1, s, "class(") Then
'                m = InStr(1, s, "*")
'                n = InStr(m + 1, s, "*")
'                If Not IsNumeric(LCaseFast(Mid$(s, m + 1, n - m - 1))) Then
'                    If modSC.FastStringComp(LCaseFast(dbPlayers(GetPlayerIndexNumber(Index)).sClass), LCaseFast(Mid$(s, m + 1, n - m - 1))) Then
'                        If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                            DoAbil Index, tArr2(j + 1)
'                            sScripting = True
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                Else
'                    If modSC.FastStringComp(LCaseFast(dbPlayers(GetPlayerIndexNumber(Index)).sClass), modgetdata.GetClassFromNum(CLng(Mid$(s, m + 1, n - m - 1)))) Then
'                        If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                            DoAbil Index, tArr2(j + 1)
'                            sScripting = True
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                End If
'                'Class(*mage*)
'            ElseIf InStr(1, s, "classes(") Then
'                m = InStr(1, s, "(")
'                n = InStr(m + 1, s, ")")
'                SplitFast ReplaceFast(LCaseFast(Mid$(s, m + 1, n - m - 1)), "*", ""), aryClasses, ","
'                'Class(*mage*,*warrior*,*thief*)
'                For l = LBound(aryClasses) To UBound(aryClasses)
'                    If Not IsNumeric(aryClasses(l)) Then
'                        If modSC.FastStringComp(LCaseFast(dbPlayers(GetPlayerIndexNumber(Index)).sClass), aryClasses(l)) Then
'                            If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                                DoAbil Index, tArr2(j + 1)
'                                sScripting = True
'                                Exit For
'                            End If
'                        Else
'                            Exit For
'                        End If
'                    Else
'                        If modSC.FastStringComp(LCaseFast(dbPlayers(GetPlayerIndexNumber(Index)).sClass), modgetdata.GetClassFromNum(CLng(aryClasses(l)))) Then
'                            If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                                DoAbil Index, tArr2(j + 1)
'                                sScripting = True
'                                Exit For
'                            End If
'                        Else
'                            Exit For
'                        End If
'                    End If
'                Next l
'            ElseIf InStr(1, s, "race(") Then
'                m = InStr(1, s, "*")
'                n = InStr(m + 1, s, "*")
'                If Not IsNumeric(LCaseFast(Mid$(s, m + 1, n - m - 1))) Then
'                    If modSC.FastStringComp(LCaseFast(dbPlayers(GetPlayerIndexNumber(Index)).sRace), LCaseFast(Mid$(s, m + 1, n - m - 1))) Then
'                        If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                            DoAbil Index, tArr2(j + 1)
'                            sScripting = True
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                Else
'                    If modSC.FastStringComp(LCaseFast(dbPlayers(GetPlayerIndexNumber(Index)).sRace), modgetdata.GetRaceFromNum(CLng(Mid$(s, m + 1, n - m - 1)))) Then
'                        If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                            DoAbil Index, tArr2(j + 1)
'                            sScripting = True
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                End If
'                'Race(*elf*)
'            ElseIf InStr(1, s, "races(") Then
'                m = InStr(1, s, "(")
'                n = InStr(m + 1, s, ")")
'                SplitFast ReplaceFast(LCaseFast(Mid$(s, m + 1, n - m - 1)), "*", ""), aryClasses, ","
'                'Class(*mage*,*warrior*,*thief*)
'                For l = LBound(aryRaces) To UBound(aryRaces)
'                    If Not IsNumeric(aryClasses(l)) Then
'                        If modSC.FastStringComp(LCaseFast(dbPlayers(GetPlayerIndexNumber(Index)).sRace), aryRaces(l)) Then
'                            If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                                DoAbil Index, tArr2(j + 1)
'                                sScripting = True
'                                Exit For
'                            End If
'                        Else
'                            Exit For
'                        End If
'                    Else
'                        If modSC.FastStringComp(LCaseFast(dbPlayers(GetPlayerIndexNumber(Index)).sRace), modgetdata.GetRaceFromNum(CLng(aryRaces(l)))) Then
'                            If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                                DoAbil Index, tArr2(j + 1)
'                                sScripting = True
'                                Exit For
'                            End If
'                        Else
'                            Exit For
'                        End If
'                    End If
'                Next l
'                'Race(*human*,*elf*,*dwarf*)
'            ElseIf InStr(1, s, "check(") Then
'                m = InStr(1, s, "*")
'                n = InStr(m + 1, s, "*")
'                Checking = LCaseFast(Mid$(s, m + 1, n - m - 1))
'                m = InStr(1, s, ",")
'                n = InStr(m + 1, s, ")")
'                'StatCheck(*Cha*,>7)
'                Comparing = LCaseFast(Mid$(s, m + 1, n - m - 1))
'                sSign = Left$(Comparing, IIf(Mid$(Comparing, 2, 1) = "=", 2, 1))
'                Comparing = Right$(Comparing, Len(Comparing) - Len(sSign))
'                If CheckTheStat(Index, Checking, CDbl(Comparing), sSign) = True Then
'                    If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                        DoAbil Index, tArr2(j + 1)
'                        sScripting = True
'                        Exit For
'                    End If
'                Else
'                    Exit For
'                End If
'            ElseIf InStr(1, s, "rnd(") Then
'                m = InStr(1, s, "(")
'                n = InStr(m + 1, s, ",")
'                rndNumber1 = CDbl(Mid$(s, m + 1, n - m - 1))
'                m = InStr(1, s, ",")
'                n = InStr(m + 1, s, ",")
'                rndNumber2 = CDbl(Mid$(s, m + 1, n - m - 1))
'                m = InStr(n, s, ")")
'                TempVal = Mid$(s, n + 1, m - n - 1)
'                theSign = Left$(TempVal, IIf(Mid$(Comparing, 2, 1) = "=", 2, 1))
'                intCompare = Right$(TempVal, Len(TempVal) - Len(theSign))
'                theRndNumber = RndNumber(rndNumber1, rndNumber2)
'                If CheckIt(CDbl(theRndNumber), CDbl(intCompare), theSign) = True Then
'                    If tArr2(j + 1) <> "cont" Then
'                        DoAbil Index, tArr2(j + 1)
'                        sScripting = True
'                        Exit For
'                    End If
'                Else
'                    Exit For
'                End If
'            ElseIf InStr(1, s, "itemcount(") Then
'                m = InStr(1, s, "*")
'                n = InStr(m + 1, s, "*")
'                Checking = LCaseFast(Mid$(s, m + 1, n - m - 1))
'                m = InStr(1, s, ",")
'                n = InStr(m + 1, s, ")")
'                'ItemCount(*basic sword*,>3)
'                Comparing = LCaseFast(Mid$(s, m + 1, n - m - 1))
'                sSign = Left$(Comparing, IIf(Mid$(Comparing, 2, 1) = "=", 2, 1))
'                Comparing = Right$(Comparing, Len(Comparing) - Len(sSign))
'                If Not IsNumeric(Comparing) Then
'                    Checking = modgetdata.GetItemNumFromName(Checking)
'                    If Checking = "(-1)" Then Exit For
'                End If
'                If CheckIt(CDbl(modMain.DCount(dbPlayers(GetPlayerIndexNumber(Index)).sInventory, ":" & Checking & ";")), CDbl(Comparing), sSign) = True Then
'                    If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                        DoAbil Index, tArr2(j + 1)
'                        sScripting = True
'                        Exit For
'                    End If
'                Else
'                    Exit For
'                End If
'            ElseIf InStr(1, s, "timecheck(") Then
'                'TimeCheck(>,hh:mm:ss)
'                m = InStr(1, s, "(")
'                n = InStr(m, s, ",")
'                m = m + 1
'                n = n - m - 1
'                sSign = Mid$(s, m, n)
'                m = InStr(1, s, ",")
'                n = InStr(m, s, ")")
'                m = m + 1
'                n = n - m - 1
'                Comparing = Mid$(s, m, n)
'                If modScripts.CheckTimeDif(sSign, Comparing) = True Then
'                    If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                        DoAbil Index, tArr2(j + 1)
'                        sScripting = True
'                        Exit For
'                    End If
'                Else
'                    Exit For
'                End If
'            ElseIf InStr(1, s, "datecheck(") Then
'                'DateCheck(>,m:dd:yyyy)
'                m = InStr(1, s, "(")
'                n = InStr(m, s, ",")
'                m = m + 1
'                n = n - m - 1
'                sSign = Mid$(s, m, n)
'                m = InStr(1, s, ",")
'                n = InStr(m, s, ")")
'                m = m + 1
'                n = n - m - 1
'                Comparing = Mid$(s, m, n)
'                If modScripts.CheckDateDif(sSign, Comparing) = True Then
'                    If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                        DoAbil Index, tArr2(j + 1)
'                        sScripting = True
'                        Exit For
'                    End If
'                Else
'                    Exit For
'                End If
'            ElseIf InStr(1, s, "flag1=") Then
'                m = InStr(1, s, "=")
'                Comparing = Mid$(s, m + 1)
'                With dbPlayers(GetPlayerIndexNumber(Index))
'                    If .iFlag1 = CLng(Comparing) Then
'                        If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                            DoAbil Index, tArr2(j + 1)
'                            sScripting = True
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                End With
'            ElseIf InStr(1, s, "flag2=") Then
'                m = InStr(1, s, "=")
'                Comparing = Mid$(s, m)
'                With dbPlayers(GetPlayerIndexNumber(Index))
'                    If .iFlag2 = CLng(Comparing) Then
'                        If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                            DoAbil Index, tArr2(j + 1)
'                            sScripting = True
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                End With
'            ElseIf InStr(1, s, "flag3=") Then
'                m = InStr(1, s, "=")
'                Comparing = Mid$(s, m)
'                With dbPlayers(GetPlayerIndexNumber(Index))
'                    If .iFlag3 = CLng(Comparing) Then
'                        If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                            DoAbil Index, tArr2(j + 1)
'                            sScripting = True
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                End With
'            ElseIf InStr(1, s, "flag4=") Then
'                m = InStr(1, s, "=")
'                Comparing = Mid$(s, m)
'                With dbPlayers(GetPlayerIndexNumber(Index))
'                    If .iFlag4 = CLng(Comparing) Then
'                        If Not modSC.FastStringComp(tArr2(j + 1), "cont") Then
'                            DoAbil Index, tArr2(j + 1)
'                            sScripting = True
'                            Exit For
'                        End If
'                    Else
'                        Exit For
'                    End If
'                End With
'            End If
'        ElseIf Not modSC.FastStringComp(tArr2(j), "cont") Then
'            'no ifs, it just does it
'            DoAbil Index, tArr2(j)
'        End If
'    Next j
'Next i
'End Function
