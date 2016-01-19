Attribute VB_Name = "modMain"
'*************************************************************************************
Rem*************************************************************************************
Rem***************       Code create by Chris Van Hooser          **********************
Rem***************                  (c)2002                       **********************
Rem*************** You may use this code and freely distribute it **********************
Rem***************   If you have any questions, please email me   **********************
Rem***************          at theendorbunker@attbi.com.          **********************
Rem***************       Thanks for downloading my project        **********************
Rem***************        and i hope you can use it well.         **********************
Rem***************                modMain                         **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Public Type MapArea
    Xl As Long
    Yl As Long
    sExits As String
    sIsRoom As Boolean
    lIndoor As Long
    sTitle As String
    sDesc As String
    lAuto As Long
    lRealID As Long
    lMob As Long
    lMaxRegen As Long
    lLight As Long
    lDeath As Long
    lJoinRoom As Long
    sJoinExit As String
    lAlreadyExist As Long
    lChecked As Long
    dN As Long
    dS As Long
    ddE As Long
    dW As Long
    dNW As Long
    dNE As Long
    dSW As Long
    dSE As Long
    cN As Long
    cS As Long
    cE As Long
    cW As Long
    cNW As Long
    cNE As Long
    cSW As Long
    cSE As Long
End Type
Public udtMapArea(899) As MapArea
Public Enum SetCBOWhich
    [Armor Type] = 0
    [Weapon Type] = 1
    [Magic Type] = 2
    [Magic Level] = 3
    [Vision Level] = 4
    [Element Type] = 5
    [Spell Use] = 6
End Enum
Public DB As Database
Public MRS          As Recordset
Public MRSMAP       As Recordset
Public MRSCLASS     As Recordset
Public MRSRACE      As Recordset
Public MRSITEM      As Recordset
Public MRSMONSTER   As Recordset
Public MRSEMOTIONS  As Recordset
Public MRSSPELLS    As Recordset
Public MRSFAMILIARS As Recordset
Public MRSSHOPS     As Recordset
Public MRSEVENTS    As Recordset
Public dbClass() As UDTClasses
Public dbEmotions() As UDTEmotions
Public dbFamiliars() As UDTFamiliars
Public dbItems() As UDTItems
Public dbMap() As UDTMap
Public dbMonsters() As UDTMonsters
Public dbPlayers() As UDTPlayers
Public dbRaces() As UDTRaces
Public dbSpells() As UDTSpells
Public dbShops() As UDTShops
Public dbEvents() As UDTEvents
Public dbArenas() As UDTMap
Public dbLetters() As UDTLetter
Public dbDoor() As UDTMap
Const sValue As String = "db1nr2me18f5m7c39k0md3f8mw31o9b56dsacry3gde1acve5hmjw"
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKey As Any, ByVal lpString As String, ByVal lpFileName As String)
Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)

Function DSVal(sdValue As String) As String
Dim Temp As String
Temp = Mid$(sdValue, 2, 1)
Temp = Temp & Mid$(sdValue, 5, 1)
Temp = Temp & Mid$(sdValue, 9, 1)
Temp = Temp & Mid$(sdValue, 14, 1)
Temp = Temp & Mid$(sdValue, 20, 1)
Temp = Temp & Mid$(sdValue, 27, 1)
Temp = Temp & Mid$(sdValue, 35, 1)
Temp = Temp & Mid$(sdValue, 44, 1)
DSVal = Temp
End Function

Sub Main()
'On Error GoTo Main_Error
If GetINI("PATH", "Path") <> "Error" Then
    On Error GoTo Main_Error
    modSec.dB_set_and_load "spike technolog", App.Path & "\data.mud", modSec.uJunkIt("S3t2]aJJMnWH³‰da“ f)aÄZw€(vrqr@‹€¿d’K¢¿‘•\• i«+{nb¡vKJJ°v‡@14z•ºem½FZ”4^`b²—teN‰Œ`RÃqe·D« ‡‚p_HvŸz§’ªÅD4•=Ba‡fhMmÀseBDek,Ç·M¥a=¥‰£¯")
    modSec.dB_set_and_load "MUD", App.Path & "\data.mud", "4.0"
    Set DB = OpenDatabase(GetINI("PATH", "Path"), False, False, ";pwd=" & DSVal(sValue))
    modUpdateDatabase.UpdateMRSSets
    Load frmSplash
    frmSplash.Show
    frmSplash.LoadServer
Else
    MsgBox "MUD Server Editor has determined you have not yet set the path to the database." & vbCrLf & "Please direct to the correct path on the next dialog screen.", vbOKOnly + vbQuestion, "First run"
    Load mdiMain
    mdiMain.CDLMain.DialogTitle = "Locate data.mud"
    mdiMain.CDLMain.InitDir = App.Path
    mdiMain.CDLMain.Filter = "MUD Data Files|*.mud"
    mdiMain.CDLMain.ShowOpen
    If Err.Number = cdlCancel Then
        MsgBox "You have not provided the location of the database, therefore you cannot continue.", vbCritical, "Not database provided"
        Unload mdiMain
        Exit Sub
    End If
    If Right$(mdiMain.CDLMain.FileName, 8) = "data.mud" Then
        WriteINI "PATH", "Path", mdiMain.CDLMain.FileName
        modSec.dB_set_and_load "spike technolog", mdiMain.CDLMain.FileName, modSec.uJunkIt("S3t2]aJJMnWH³‰da“ f)aÄZw€(vrqr@‹€¿d’K¢¿‘•\• i«+{nb¡vKJJ°v‡@14z•ºem½FZ”4^`b²—teN‰Œ`RÃqe·D« ‡‚p_HvŸz§’ªÅD4•=Ba‡fhMmÀseBDek,Ç·M¥a=¥‰£¯")
        modSec.dB_set_and_load "MUD", mdiMain.CDLMain.FileName, "4.0"
        Set DB = OpenDatabase(mdiMain.CDLMain.FileName, False, False, ";pwd=" & DSVal(sValue))
        modUpdateDatabase.UpdateMRSSets
        Load frmSplash
        frmSplash.Show
        frmSplash.LoadServer
    Else
        MsgBox "You have not provided the location of the correct database, therefore you cannot continue.", vbCritical, "Not database provided"
        Unload mdiMain
        Exit Sub
    End If
End If
On Error GoTo 0
Exit Sub
Main_Error:
MsgBox "An error occured on startup." & vbCrLf & "Cause may be a corrupted database, or wrong database path." & vbCrLf & "Restart the editor to correct the problem.", vbCritical, "Error"
On Error Resume Next
Kill App.Path & "\Data.dat"
Unload mdiMain
Unload frmSplash
End Sub

Public Function GetINI(sHeader As String, sKey As String) As String
Dim strSpace As String, theLength As Long
strSpace = Space(255)
theLength = GetPrivateProfileString(sHeader, sKey, "Error", strSpace, 255, App.Path & "\data.dat")
strSpace = Replace(strSpace, Chr(0), "")
strSpace = Trim(strSpace)
GetINI = strSpace
End Function

Public Sub WriteINI(sHeader As String, sKey As String, sString As String)
WritePrivateProfileString sHeader, sKey, sString, App.Path & "\data.dat"
End Sub

Public Sub SetLstSelected(LST As UltraBox, FindWhat As String)
Dim i As Long
For i = 1 To LST.ListCount
    If LST.list(i) = FindWhat Then
        LST.SetSelected i, True
        Exit For
    End If
Next
End Sub

Public Sub SetListIndex(CBO As ComboBox, SwitchTo As String)
Dim i As Long
For i = 0 To CBO.ListCount - 1
    If CBO.list(i) = SwitchTo Then
        CBO.ListIndex = i
        Exit For
    End If
Next
End Sub

Public Sub SetCBOSelectByID(CBO As ComboBox, Num As String)
Dim i As Long
For i = 0 To CBO.ListCount - 1
    If Left$(CBO.list(i), 2 + Len(Num)) = "(" & Num & ")" Then
        CBO.ListIndex = i
        Exit For
    End If
Next
End Sub

Public Function DCount(sInString As String, sCountString As String) As Long
'finds out how many a certain string appears in another string
Dim iNextOccur As Long
iNextOccur = 0
'If InStr(iNextOccur, sInString, sCountString) = 0 Then Exit Function Else DCount = DCount + 1
Do
    iNextOccur = InStr(iNextOccur + 1, sInString, sCountString)
    If iNextOccur > 0 Then DCount = DCount + 1 Else Exit Do
    DoEvents
Loop
End Function

Public Sub FeedAList(f As FlagOptions, WhichLoad As String)
Dim i As Long
Select Case WhichLoad
    Case "rooms"
        f.ClearFeed
        For i = LBound(dbMap) To UBound(dbMap)
            With dbMap(i)
                f.FeedMe CStr("(" & .lRoomID & ") " & .sRoomTitle), .lRoomID
            End With
        Next
        f.FillNow
    Case "items"
        f.ClearFeed
        For i = LBound(dbItems) To UBound(dbItems)
            With dbItems(i)
                f.FeedMe CStr("(" & .iID & ") " & .sItemName), .iID
            End With
        Next
        f.FillNow
    Case "spells"
        f.ClearFeed
        For i = LBound(dbSpells) To UBound(dbSpells)
            With dbSpells(i)
                f.FeedMe CStr("(" & .lID & ") " & .sSpellName), .lID
            End With
        Next
        f.FillNow
    Case "familiars"
        f.ClearFeed
        For i = LBound(dbFamiliars) To UBound(dbFamiliars)
            With dbFamiliars(i)
                f.FeedMe CStr("(" & .iID & ") " & .sFamName), .iID
            End With
        Next
        f.FillNow
    Case "guilds"
        With f
            .ClearFeed
            .FeedMe "(5) Leader", 5
            .FeedMe "(4) General", 4
            .FeedMe "(3) Lieutenant", 3
            .FeedMe "(2) Soldier", 2
            .FeedMe "(1) Normal", 1
            .FeedMe "(0) Scrub", 0
        End With
        f.FillNow
End Select
End Sub

Public Function DeterStyle(sShortFlg As String, ByRef WhichLoad As String) As Long
Select Case Left$(sShortFlg, 3)
    Case "mhp" 'Max Hitpoints
        DeterStyle = 0
    Case "str" 'Strength
        DeterStyle = 0
    Case "agi" 'Agility
        DeterStyle = 0
    Case "int" 'Intellect
        DeterStyle = 0
    Case "cha" 'Charm
        DeterStyle = 0
    Case "dex" 'Dexterity
        DeterStyle = 0
    Case "pac" 'Armor Class
        DeterStyle = 0
    Case "acc" 'Accurracy
        DeterStyle = 0
    Case "cri" 'Crits
        DeterStyle = 0
    Case "mma" 'Max Mana
        DeterStyle = 0
    Case "dam" 'damage bonus
        DeterStyle = 0
    Case "dod" 'dodge
        DeterStyle = 0
    Case "h/l"
        DeterStyle = 0
    Case "m/l"
        DeterStyle = 0
    Case "mit"
        DeterStyle = 0
    Case "acl"
        DeterStyle = 0
    Case "vis"
        DeterStyle = 0
    Case "pts"
        DeterStyle = 0
    Case "sne"
        DeterStyle = 2
    Case "cbs"
        DeterStyle = 2
    Case "cdw"
        DeterStyle = 2
    Case "tel"
        DeterStyle = 3
        WhichLoad = "rooms"
    Case "stu"
        DeterStyle = 0
    Case "lig"
        DeterStyle = 0
    Case "chp"
        DeterStyle = 0
    Case "cma"
        DeterStyle = 0
    Case "hun"
        DeterStyle = 0
    Case "sta"
        DeterStyle = 0
    Case "cac"
        DeterStyle = 0
    Case "evi"
        DeterStyle = 0
    Case "pap"
        DeterStyle = 0
    Case "mat"
        DeterStyle = 3
        WhichLoad = "items"
    Case "snd"
        DeterStyle = 1
    Case "sro"
        DeterStyle = 1
    Case "gsp"
        DeterStyle = 3
        WhichLoad = "spells"
    Case "gfa"
        DeterStyle = 3
        WhichLoad = "familiars"
    Case "el0"
        DeterStyle = 0
    Case "el1"
        DeterStyle = 0
    Case "el2"
        DeterStyle = 0
    Case "el3"
        DeterStyle = 0
    Case "el4"
        DeterStyle = 0
    Case "el5"
        DeterStyle = 0
    Case "el6"
        DeterStyle = 0
    Case "el7"
        DeterStyle = 0
    Case "el8"
        DeterStyle = 0
    Case "m00"
        DeterStyle = 2
    Case "m01"
        DeterStyle = 2
    Case "m02"
        DeterStyle = 2
    Case "m03"
        DeterStyle = 2
    Case "m04"
        DeterStyle = 3
        WhichLoad = "guilds"
    Case "m05"
        DeterStyle = 2
    Case "m06"
        DeterStyle = 2
    Case "m07"
        DeterStyle = 2
    Case "m08"
        DeterStyle = 2
    Case "m09"
        DeterStyle = 2
    Case "m10"
        DeterStyle = 2
    Case "m11"
        DeterStyle = 2
    Case "m12"
        DeterStyle = 2
    Case "m13"
        DeterStyle = 2
    Case "m14"
        DeterStyle = 2
    Case "m15"
        DeterStyle = 2
    Case "m16"
        DeterStyle = 2
    Case "m17"
        DeterStyle = 2
    Case "m18"
        DeterStyle = 2
    Case "m19"
        DeterStyle = 2
    Case "m20"
        DeterStyle = 2
    Case "m21"
        DeterStyle = 2
    Case "m22"
        DeterStyle = 2
    Case "m23"
        DeterStyle = 2
    Case "m24"
        DeterStyle = 2
    Case "m25"
        DeterStyle = 2
    Case "m26"
        DeterStyle = 2
    Case "m27"
        DeterStyle = 2
    Case "s01"
        DeterStyle = 0
    Case "s03"
        DeterStyle = 0
    Case "s05"
        DeterStyle = 0
    Case "s09"
        DeterStyle = 0
    Case "s11"
        DeterStyle = 0
    Case "s13"
        DeterStyle = 0
    Case "the"
        DeterStyle = 2
    Case "thi"
        DeterStyle = 0
End Select
End Function

Public Function TranslateFlag(sInput As String) As String
'Accuracy
'Armor Class Bonus
'Can Backstab
'Can Sneak
'Character Points Bonus Per Level
'Crital Chance Bonus
'Dodge Bonus
'Hitpoint Bonus Per Level
'Mana Bonus Per Level
'Max Damage Bonus

Select Case Left$(sInput, 3)
    Case "the"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Can Steal                = TRUE"
            Case Else
                TranslateFlag = "Can Steal                = FALSE"
        End Select
    Case "thi"
        TranslateFlag = "Thieving Bonus           + " & Mid$(sInput, 4)
    Case "csp"
        TranslateFlag = "Cast Spell               : " & Mid$(sInput, 4)
    Case "mhp" 'Max Hitpoints
        TranslateFlag = "Max Hitpoints            + " & Mid$(sInput, 4)
    Case "str" 'Strength
        TranslateFlag = "Strength                 + " & Mid$(sInput, 4)
    Case "agi" 'Agility
        TranslateFlag = "Agility                  + " & Mid$(sInput, 4)
    Case "int" 'Intellect
        TranslateFlag = "Intellect                + " & Mid$(sInput, 4)
    Case "cha" 'Charm
        TranslateFlag = "Charm                    + " & Mid$(sInput, 4)
    Case "dex" 'Dexterity
        TranslateFlag = "Dexterity                + " & Mid$(sInput, 4)
    Case "pac" 'Armor Class
        TranslateFlag = "Armor Class              + " & Mid$(sInput, 4)
    Case "acc" 'Accurracy
        TranslateFlag = "Accuracy Bonus           + " & Mid$(sInput, 4)
    Case "cri" 'Crits
        TranslateFlag = "Critical Hit Bonus       + " & Mid$(sInput, 4)
    Case "mma" 'Max Mana
        TranslateFlag = "Max Mana                 + " & Mid$(sInput, 4)
    Case "dam" 'damage bonus
        TranslateFlag = "Max Damage Bonus         + " & Mid$(sInput, 4)
    Case "dod" 'dodge
        TranslateFlag = "Dodge Bonus              + " & Mid$(sInput, 4)
    Case "h/l"
        TranslateFlag = "Hitpoint Bonus Per Level + " & Mid$(sInput, 4)
    Case "m/l"
        TranslateFlag = "Mana Bonus Per Level     + " & Mid$(sInput, 4)
    Case "mit"
        TranslateFlag = "Max Items Bonus          + " & Mid$(sInput, 4)
    Case "acl"
        TranslateFlag = "Armor Class Bonus        + " & Mid$(sInput, 4)
    Case "vis"
        TranslateFlag = "Vision Bonus             + " & Mid$(sInput, 4)
    Case "pts"
        TranslateFlag = "Char. Points Bonus P/LVL + " & Mid$(sInput, 4)
    Case "sne"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Can Sneak                = TRUE"
            Case Else
                TranslateFlag = "Can Sneak                = FALSE"
        End Select
    Case "cbs"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Can Backstab             = TRUE"
            Case Else
                TranslateFlag = "Can Backstab             = FALSE"
        End Select
    Case "tel"
        TranslateFlag = "Teleport To Room         : " & Mid$(sInput, 4)
    Case "stu"
        TranslateFlag = "Stun                     : " & Mid$(sInput, 4)
    Case "lig"
        TranslateFlag = "Player's Vision          + " & Mid$(sInput, 4)
    Case "chp"
        TranslateFlag = "Current Hit Points       + " & Mid$(sInput, 4)
    Case "cma"
        TranslateFlag = "Current Mana Points      + " & Mid$(sInput, 4)
    Case "hun"
        TranslateFlag = "Current Hunger Status    + " & Mid$(sInput, 4)
    Case "sta"
        TranslateFlag = "Current Stamina Status   + " & Mid$(sInput, 4)
    Case "cac"
        TranslateFlag = "Current Armor Class      + " & Mid$(sInput, 4)
    Case "evi"
        TranslateFlag = "Current Evil Points      + " & Mid$(sInput, 4)
    Case "pap"
        TranslateFlag = "Current Amount Of Paper  + " & Mid$(sInput, 4)
    Case "mat"
        TranslateFlag = "Give Item To Player      : " & Mid$(sInput, 4)
    Case "snd"
        TranslateFlag = "Send To Player           : " & Mid$(sInput, 4)
    Case "sro"
        TranslateFlag = "Send Message To Room     : " & Mid$(sInput, 4)
    Case "gsp"
        TranslateFlag = "Give Spell To Player     : " & Mid$(sInput, 4)
    Case "gfa"
        TranslateFlag = "Give Familiar To Player  : " & Mid$(sInput, 4)
    Case "cdw"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Can Dual Wield   = TRUE"
            Case Else
                TranslateFlag = "Can Dual Wield   = FALSE"
        End Select
        'fire/ice/water/lightning/earth/poison/wind/holy/unholy
    Case "el0"
        TranslateFlag = "Fire Resistance          + " & Mid$(sInput, 4)
    Case "el1"
        TranslateFlag = "Ice Resistance           + " & Mid$(sInput, 4)
    Case "el2"
        TranslateFlag = "Water Resistance         + " & Mid$(sInput, 4)
    Case "el3"
        TranslateFlag = "Lightning Resistance     + " & Mid$(sInput, 4)
    Case "el4"
        TranslateFlag = "Earth Resistance         + " & Mid$(sInput, 4)
    Case "el5"
        TranslateFlag = "Poison Resistance        + " & Mid$(sInput, 4)
    Case "el6"
        TranslateFlag = "Wind Resistance          + " & Mid$(sInput, 4)
    Case "el7"
        TranslateFlag = "Holy Resistance          + " & Mid$(sInput, 4)
    Case "el8"
        TranslateFlag = "Unholy Resistance        + " & Mid$(sInput, 4)
    Case "m00"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Attack        = TRUE"
            Case Else
                TranslateFlag = "Player Can Attack        = FALSE"
        End Select
    Case "m01"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Cast Spell    = TRUE"
            Case Else
                TranslateFlag = "Player Can Cast Spell    = FALSE"
        End Select
    Case "m02"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Sneak         = TRUE"
            Case Else
                TranslateFlag = "Player Can Sneak         = FALSE"
        End Select
    Case "m03"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Talks Gibberish   = TRUE"
            Case Else
                TranslateFlag = "Player Talks Gibberish   = FALSE"
        End Select
    Case "m04"
        TranslateFlag = "Players Guild Rank       = " & Mid$(sInput, 4)
    Case "m05"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Is Invisible      = TRUE"
            Case Else
                TranslateFlag = "Player Is Invisible      = FALSE"
        End Select
    Case "m06"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Head       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Head       = FALSE"
        End Select
    Case "m07"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Face       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Face       = FALSE"
        End Select
    Case "m08"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Ears       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Ears       = FALSE"
        End Select
    Case "m09"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Neck       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Neck       = FALSE"
        End Select
    Case "m10"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Body       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Body       = FALSE"
        End Select
    Case "m11"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Back       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Back       = FALSE"
        End Select
    Case "m12"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Arms       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Arms       = FALSE"
        End Select
    Case "m13"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Shield     = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Shield     = FALSE"
        End Select
    Case "m14"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Hands      = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Hands      = FALSE"
        End Select
    Case "m15"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Legs       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Legs       = FALSE"
        End Select
    Case "m16"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Feet       = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Feet       = FALSE"
        End Select
    Case "m17"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Waist      = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Waist      = FALSE"
        End Select
    Case "m18"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Weapon     = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Weapon     = FALSE"
        End Select
    Case "m19"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Be De-Sysed   = TRUE"
            Case Else
                TranslateFlag = "Player Can Be De-Sysed   = FALSE"
        End Select
    Case "m20"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can See Invisible = TRUE"
            Case Else
                TranslateFlag = "Player Can See Invisible = FALSE"
        End Select
    Case "m21"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can See Hidden    = TRUE"
            Case Else
                TranslateFlag = "Player Can See Hidden    = FALSE"
        End Select
    Case "m22"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Ring 0     = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Ring 0     = FALSE"
        End Select
    Case "m23"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Ring 1     = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Ring 1     = FALSE"
        End Select
    Case "m24"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Ring 2     = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Ring 2     = FALSE"
        End Select
    Case "m25"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Ring 3     = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Ring 3     = FALSE"
        End Select
    Case "m26"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Ring 4     = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Ring 4     = FALSE"
        End Select
    Case "m27"
        Select Case Mid$(sInput, 4)
            Case "1"
                TranslateFlag = "Player Can Eq Ring 5     = TRUE"
            Case Else
                TranslateFlag = "Player Can Eq Ring 5     = FALSE"
        End Select
    Case "s01"
        TranslateFlag = "Players Spell Casting    + " & Mid$(sInput, 4)
    Case "s03"
        TranslateFlag = "Players Magic Resistance + " & Mid$(sInput, 4)
    Case "s05"
        TranslateFlag = "Players Perception       + " & Mid$(sInput, 4)
    Case "s09"
        TranslateFlag = "Players Max Items        + " & Mid$(sInput, 4)
    Case "s11"
        TranslateFlag = "Players Stealth          + " & Mid$(sInput, 4)
    Case "s13"
        TranslateFlag = "Players Animal Relations + " & Mid$(sInput, 4)
End Select
End Function

Public Function MakeDBFlag(sInput As String) As String

Dim s As String
Dim t As String
s = Left$(sInput, 25)
s = Trim$(s)
s = ShortFlag(s)
t = Mid$(sInput, 28)
t = Trim$(t)
t = Replace$(t, "TRUE", "1")
t = Replace$(t, "FALSE", "0")
MakeDBFlag = s & t
End Function

Public Sub PopulateCBOFlag(CBO As ComboBox)
'Select Case sInput
CBO.Clear
    CBO.AddItem "Can Steal"
        ' "the"
    CBO.AddItem "Thieving Bonus"
        ' "thi"
    CBO.AddItem "Cast Spell"
        ' "csp"
    CBO.AddItem "Max Hitpoints" 'Max Hitpoints
        ' "mhp"
    CBO.AddItem "Strength" 'Strength
        ' "str"
    CBO.AddItem "Agility" 'Agility
        ' "agi"
    CBO.AddItem "Intellect" 'Intellect
        ' "int"
    CBO.AddItem "Charm" 'Charm
        ' "cha"
    CBO.AddItem "Dexterity" 'Dexterity
        ' "dex"
    CBO.AddItem "Armor Class" 'Armor Class
        ' "pac"
    CBO.AddItem "Accuracy"  'Accurracy
        ' "acc"
    CBO.AddItem "Critical Hit Bonus" 'Crits
        ' "cri"
    CBO.AddItem "Max Mana" 'Max Mana
        ' "mma"
    CBO.AddItem "Max Damage Bonus" 'damage bonus
        ' "dam"
    CBO.AddItem "Dodge Bonus" 'dodge
        ' "dod"
    CBO.AddItem "Hitpoint Bonus Per Level"
        ' "h/l"
    CBO.AddItem "Mana Bonus Per Level"
        ' "m/l"
    CBO.AddItem "Max Items Bonus"
        ' "mit"
    CBO.AddItem "Armor Class Bonus"
        ' "acl"
    CBO.AddItem "Vision Bonus"
        ' "vis"
    CBO.AddItem "Character Points Bonus Per Level"
        ' "pts"
    CBO.AddItem "Can Sneak"
        ' "sne"
    CBO.AddItem "Can Backstab"
        ' "cbs"
    CBO.AddItem "Teleport To Room"
        ' "tel"
    CBO.AddItem "Stun"
        ' "stu"
    CBO.AddItem "Player's Vision"
        ' "lig"
    CBO.AddItem "Current Hit Points"
        ' "chp"
    CBO.AddItem "Current Mana Points"
        ' "cma"
    CBO.AddItem "Current Hunger Status"
        ' "hun"
    CBO.AddItem "Current Stamina Status"
        ' "sta"
    CBO.AddItem "Current Armor Class"
        ' "cac"
    CBO.AddItem "Current Evil Points"
        ' "evi"
    CBO.AddItem "Current Amount Of Paper"
        ' "pap"
    CBO.AddItem "Give Item To Player"
        ' "mat"
    CBO.AddItem "Send To Player"
        ' "snd"
    CBO.AddItem "Send Message To Room"
        ' "sro"
    CBO.AddItem "Give Spell To Player"
        ' "gsp"
    CBO.AddItem "Give Familiar To Player"
        ' "gfa"
    CBO.AddItem "Fire Resistance"
        ' "el0"
    CBO.AddItem "Ice Resistance"
        ' "el1"
    CBO.AddItem "Water Resistance"
        ' "el2"
    CBO.AddItem "Lightning Resistance"
        ' "el3"
    CBO.AddItem "Earth Resistance"
        ' "el4"
    CBO.AddItem "Poison Resistance"
        ' "el5"
    CBO.AddItem "Wind Resistance"
        ' "el6"
    CBO.AddItem "Holy Resistance"
        ' "el7"
    CBO.AddItem "Unholy Resistance"
        ' "el8"
    CBO.AddItem "Player Can Attack"
        ' "m00"
    CBO.AddItem "Player Can Cast Spell"
        ' "m01"
    CBO.AddItem "Player Can Sneak"
        ' "m02"
    CBO.AddItem "Player Talks Gibberish"
        ' "m03"
    CBO.AddItem "Players Guild Rank"
        ' "m04"
    CBO.AddItem "Player Is Invisible"
        ' "m05"
    CBO.AddItem "Player Can Eq Head"
        ' "m06"
    CBO.AddItem "Player Can Eq Face"
        ' "m07"
    CBO.AddItem "Player Can Eq Ears"
        ' "m08"
    CBO.AddItem "Player Can Eq Neck"
        ' "m09"
    CBO.AddItem "Player Can Eq Body"
        ' "m10"
    CBO.AddItem "Player Can Eq Back"
        ' "m11"
    CBO.AddItem "Player Can Eq Arms"
        ' "m12"
    CBO.AddItem "Player Can Eq Shield"
        ' "m13"
    CBO.AddItem "Player Can Eq Hands"
        ' "m14"
    CBO.AddItem "Player Can Eq Legs"
        ' "m15"
    CBO.AddItem "Player Can Eq Feet"
        ' "m16"
    CBO.AddItem "Player Can Eq Waist"
        ' "m17"
    CBO.AddItem "Player Can Eq Weapon"
        ' "m18"
    CBO.AddItem "Player Can Be De-Sysed"
        ' "m19"
    CBO.AddItem "Player Can See Invisible"
        ' "m20"
    CBO.AddItem "Player Can See Hidden"
        ' "m21"
    CBO.AddItem "Player Can Eq Ring 0"
        ' "m22"
    CBO.AddItem "Player Can Eq Ring 1"
        ' "m23"
    CBO.AddItem "Player Can Eq Ring 2"
        ' "m24"
    CBO.AddItem "Player Can Eq Ring 3"
        ' "m25"
    CBO.AddItem "Player Can Eq Ring 4"
        ' "m26"
    CBO.AddItem "Player Can Eq Ring 5"
        ' "m27"
    CBO.AddItem "Players Spell Casting"
        ' "s01"
    CBO.AddItem "Players Magic Resistance"
        ' "s03"
    CBO.AddItem "Players Perception"
        ' "s05"
    CBO.AddItem "Players Max Items"
        ' "s09"
    CBO.AddItem "Players Stealth"
        ' "s11"
    CBO.AddItem "Players Animal Relations"
        ' "s13"
'End Select
End Sub


Public Function ShortFlag(sInput As String) As String
Select Case sInput
    Case "Can Steal"
        ShortFlag = "the"
    Case "Thieving Bonus"
        ShortFlag = "thi"
    Case "Cast Spell"
        ShortFlag = "csp"
    Case "Max Hitpoints" 'Max Hitpoints
        ShortFlag = "mhp"
    Case "Strength" 'Strength
        ShortFlag = "str"
    Case "Agility" 'Agility
        ShortFlag = "agi"
    Case "Intellect" 'Intellect
        ShortFlag = "int"
    Case "Charm" 'Charm
        ShortFlag = "cha"
    Case "Dexterity" 'Dexterity
        ShortFlag = "dex"
    Case "Armor Class" 'Armor Class
        ShortFlag = "pac"
    Case "Accuracy Bonus", "Accuracy" 'Accurracy
        ShortFlag = "acc"
    Case "Critical Hit Bonus" 'Crits
        ShortFlag = "cri"
    Case "Max Mana" 'Max Mana
        ShortFlag = "mma"
    Case "Max Damage Bonus" 'damage bonus
        ShortFlag = "dam"
    Case "Dodge Bonus" 'dodge
        ShortFlag = "dod"
    Case "Hitpoint Bonus Per Level"
        ShortFlag = "h/l"
    Case "Mana Bonus Per Level"
        ShortFlag = "m/l"
    Case "Max Items Bonus"
        ShortFlag = "mit"
    Case "Armor Class Bonus"
        ShortFlag = "acl"
    Case "Vision Bonus"
        ShortFlag = "vis"
    Case "Char. Points Bonus P/LVL", "Character Points Bonus Per Level"
        ShortFlag = "pts"
    Case "Can Sneak"
        ShortFlag = "sne"
    Case "Can Backstab"
        ShortFlag = "cbs"
    Case "Teleport To Room"
        ShortFlag = "tel"
    Case "Stun"
        ShortFlag = "stu"
    Case "Player's Vision"
        ShortFlag = "lig"
    Case "Current Hit Points"
        ShortFlag = "chp"
    Case "Current Mana Points"
        ShortFlag = "cma"
    Case "Current Hunger Status"
        ShortFlag = "hun"
    Case "Current Stamina Status"
        ShortFlag = "sta"
    Case "Current Armor Class"
        ShortFlag = "cac"
    Case "Current Evil Points"
        ShortFlag = "evi"
    Case "Current Amount Of Paper"
        ShortFlag = "pap"
    Case "Give Item To Player"
        ShortFlag = "mat"
    Case "Send To Player"
        ShortFlag = "snd"
    Case "Send Message To Room"
        ShortFlag = "sro"
    Case "Give Spell To Player"
        ShortFlag = "gsp"
    Case "Give Familiar To Player"
        ShortFlag = "gfa"
    Case "Fire Resistance"
        ShortFlag = "el0"
    Case "Ice Resistance"
        ShortFlag = "el1"
    Case "Water Resistance"
        ShortFlag = "el2"
    Case "Lightning Resistance"
        ShortFlag = "el3"
    Case "Earth Resistance"
        ShortFlag = "el4"
    Case "Poison Resistance"
        ShortFlag = "el5"
    Case "Wind Resistance"
        ShortFlag = "el6"
    Case "Holy Resistance"
        ShortFlag = "el7"
    Case "Unholy Resistance"
        ShortFlag = "el8"
    Case "Player Can Attack"
        ShortFlag = "m00"
    Case "Player Can Cast Spell"
        ShortFlag = "m01"
    Case "Player Can Sneak"
        ShortFlag = "m02"
    Case "Player Talks Gibberish"
        ShortFlag = "m03"
    Case "Players Guild Rank"
        ShortFlag = "m04"
    Case "Player Is Invisible"
        ShortFlag = "m05"
    Case "Player Can Eq Head"
        ShortFlag = "m06"
    Case "Player Can Eq Face"
        ShortFlag = "m07"
    Case "Player Can Eq Ears"
        ShortFlag = "m08"
    Case "Player Can Eq Neck"
        ShortFlag = "m09"
    Case "Player Can Eq Body"
        ShortFlag = "m10"
    Case "Player Can Eq Back"
        ShortFlag = "m11"
    Case "Player Can Eq Arms"
        ShortFlag = "m12"
    Case "Player Can Eq Shield"
        ShortFlag = "m13"
    Case "Player Can Eq Hands"
        ShortFlag = "m14"
    Case "Player Can Eq Legs"
        ShortFlag = "m15"
    Case "Player Can Eq Feet"
        ShortFlag = "m16"
    Case "Player Can Eq Waist"
        ShortFlag = "m17"
    Case "Player Can Eq Weapon"
        ShortFlag = "m18"
    Case "Player Can Be De-Sysed"
        ShortFlag = "m19"
    Case "Player Can See Invisible"
        ShortFlag = "m20"
    Case "Player Can See Hidden"
        ShortFlag = "m21"
    Case "Player Can Eq Ring 0"
        ShortFlag = "m22"
    Case "Player Can Eq Ring 1"
        ShortFlag = "m23"
    Case "Player Can Eq Ring 2"
        ShortFlag = "m24"
    Case "Player Can Eq Ring 3"
        ShortFlag = "m25"
    Case "Player Can Eq Ring 4"
        ShortFlag = "m26"
    Case "Player Can Eq Ring 5"
        ShortFlag = "m27"
    Case "Players Spell Casting"
        ShortFlag = "s01"
    Case "Players Magic Resistance"
        ShortFlag = "s03"
    Case "Players Perception"
        ShortFlag = "s05"
    Case "Players Max Items"
        ShortFlag = "s09"
    Case "Players Stealth"
        ShortFlag = "s11"
    Case "Players Animal Relations"
        ShortFlag = "s13"
    Case "Can Dual Wield"
        ShortFlag = "cdw"
End Select
End Function

Public Function GetHelp(LongFlag As String) As String
Select Case LongFlag
    Case "Max Hitpoints" 'Max Hitpoints
        GetHelp = "This flag will increase the players Max Hit Points by the value set. The value may be a negative number."
    Case "Strength" 'Strength
        GetHelp = "This flag will increase the players Strength by the value set. The value may be a negative number."
    Case "Agility" 'Agility
        GetHelp = "This flag will increase the players Agility by the value set. The value may be a negative number."
    Case "Intellect" 'Intellect
        GetHelp = "This flag will increase the players Intellect by the value set. The value may be a negative number."
    Case "Charm" 'Charm
        GetHelp = "This flag will increase the players Charm by the value set. The value may be a negative number."
    Case "Dexterity" 'Dexterity
        GetHelp = "This flag will increase the players Dexterity by the value set. The value may be a negative number."
    Case "Armor Class" 'Armor Class
        GetHelp = "This flag will increase the players Armor Class by the value set. The value may be a negative number."
    Case "Accuracy Bonus", "Accuracy"
        GetHelp = "This flag will increase the players Accuracy Bonus by the value set. The value may be a negative number."
    Case "Critical Hit Bonus" 'Crits
        GetHelp = "This flag will increase the players Critical Hit Bonus by the value set. The value may be a negative number."
    Case "Max Mana" 'Max Mana
        GetHelp = "This flag will increase the players Max Mana by the value set. The value may be a negative number."
    Case "Max Damage Bonus" 'damage bonus
        GetHelp = "This flag will increase the players Max Damage Bonus by the value set. The value may be a negative number."
    Case "Dodge Bonus" 'dodge
        GetHelp = "This flag will increase the players Dodge Bonus by the value set. The value may be a negative number."
    Case "Hitpoint Bonus Per Level"
        GetHelp = "This flag will increase the players Max Hit Points by the value set each time they train for a new level. The value may be a negative number."
    Case "Mana Bonus Per Level"
        GetHelp = "This flag will increase the players Max Mana by the value set each time they train for a new level. The value may be a negative number."
    Case "Max Items Bonus"
        GetHelp = "This flag will increase the players Max Items they are able to carry by the value set. The value may be a negative number."
    Case "Armor Class Bonus"
        GetHelp = "This flag will increase the players Armor Class by the value set. The value may be a negative number."
    Case "Vision Bonus"
        GetHelp = "This flag will increase the players Vision by the value set. The value may be a negative number."
    Case "Char. Points Bonus P/LVL", "Character Points Bonus Per Level"
        GetHelp = "This flag will increase the players Character Points by the value set each time they train. Training automatically gives 1 point. The value may be a negative number."
    Case "Can Sneak"
        GetHelp = "This flag toggles whether the character is proficient at sneaking. Value can be 1; Can Sneak, or 0; Can NOT Sneak"
    Case "Can Backstab"
        GetHelp = "This flag toggles whether the character is proficient at backstabbing. Value can be 1; Can Backstab, or 0; Can NOT Backstab. NOTE: Without the sneaking flag, this will be useless to the character."
    Case "Can Dual Wield"
        GetHelp = "This flag toggles whether the character can dual wield weapons."
End Select
End Function


Public Sub SetCBOlstIndex(CBO As ComboBox, Val As Long, WhichOne As SetCBOWhich)
Select Case WhichOne
    Case 0
        Select Case Val
            Case 0
                SetListIndex CBO, "0,1= Nothing"
            Case 1
                SetListIndex CBO, "0,1= Nothing"
            Case 2
                SetListIndex CBO, "2  = Silk"
            Case 3
                SetListIndex CBO, "3  = Padded"
            Case 4
                SetListIndex CBO, "4  = Robes"
            Case 5
                SetListIndex CBO, "5  = Soft Leather"
            Case 6
                SetListIndex CBO, "6  = Hard Leather"
            Case 7
                SetListIndex CBO, "7  = Studded Leather"
            Case 8
                SetListIndex CBO, "8  = Scale"
            Case 9
                SetListIndex CBO, "9  = Studded Scale"
            Case 10
                SetListIndex CBO, "10 = Chain"
            Case 11
                SetListIndex CBO, "11 = Plate"
            Case 12
                SetListIndex CBO, "12 = Silk And Padded Only"
            Case 13
                SetListIndex CBO, "13 = Leather Only"
            Case 14
                SetListIndex CBO, "14 = Scale Only"
            Case 15
                SetListIndex CBO, "15 = Scale And Padded Only"
            Case 16
                SetListIndex CBO, "16 = Chain And Padded Only"
            Case 17
                SetListIndex CBO, "17 = Plate And Chain And Padded Only"
            Case 18
                SetListIndex CBO, "18 = Leather And Scale Only"
            Case 19
                SetListIndex CBO, "19 = Plate Only"
            Case 20
                SetListIndex CBO, "20 = Plate And Padded Only"
            Case 21
                SetListIndex CBO, "21 = Silk And Leather Only"
            Case 22
                SetListIndex CBO, "22 = Robes Only"
            Case 23
                SetListIndex CBO, "23 = Robes And Silk Only"
            Case 24
                SetListIndex CBO, "24 = Robes And Padded Only"
        End Select
    Case 1
        Select Case Val
            Case 0
                SetListIndex CBO, "0  = None"
            Case 1
                SetListIndex CBO, "1  = 1h Sharp Short"
            Case 2
                SetListIndex CBO, "2  = 1h Sharp Long"
            Case 3
                SetListIndex CBO, "3  = 2h Bows"
            Case 4
                SetListIndex CBO, "4  = 1h Blunt Short"
            Case 5
                SetListIndex CBO, "5  = 1h Blunt Long"
            Case 6
                SetListIndex CBO, "6  = 2h Sharp Short"
            Case 7
                SetListIndex CBO, "7  = 2h Sharp Long"
            Case 8
                SetListIndex CBO, "8  = 2h Blunt Short"
            Case 9
                SetListIndex CBO, "9  = 2h Blunt Long"
            Case 10
                SetListIndex CBO, "10 = 2h Staves"
            Case 11
                SetListIndex CBO, "11 = 1h Only"
            Case 12
                SetListIndex CBO, "12 = 2h Bows only And Sharp Short"
            Case 13
                SetListIndex CBO, "13 = 2h Staves only And Sharp Short"
            Case 14
                SetListIndex CBO, "14 = 2h only"
            Case 15
                SetListIndex CBO, "15 = Sharp only"
            Case 16
                SetListIndex CBO, "16 = Blunt only"
            Case 17
                SetListIndex CBO, "17 = Any"
        End Select
    Case 2
        Select Case Val
            Case 0
                SetListIndex CBO, "0 - None"
            Case 1
                SetListIndex CBO, "1 - Magery"
            Case 2
                SetListIndex CBO, "2 - Druish"
            Case 3
                SetListIndex CBO, "3 - Priestly"
            Case 4
                SetListIndex CBO, "4 - Kai"
            Case 5
                SetListIndex CBO, "5 - General"
            Case 6
                SetListIndex CBO, "6 - Unholy"
            Case 7
                SetListIndex CBO, "7 - Psychic"
            Case 8
                SetListIndex CBO, "8 - Bardic"
            Case 9
                SetListIndex CBO, "9 - Witch"
            Case 10
                SetListIndex CBO, "10 - Teleporter"
        End Select
    Case 3
        Select Case Val
            Case 0
                SetListIndex CBO, "0 - None"
            Case 1
                SetListIndex CBO, "1 - Basic"
            Case 2
                SetListIndex CBO, "2 - Intermediate"
            Case 3
                SetListIndex CBO, "3 - Advanced"
            Case 4
                SetListIndex CBO, "4 - Expert"
            Case 5
                SetListIndex CBO, "5 - Master"
        End Select
    Case 4
        Select Case Val
            Case -4
                SetListIndex CBO, "-4 Horrible Vision"
            Case -3
                SetListIndex CBO, "-3 Terrible Vision"
            Case -2
                SetListIndex CBO, "-2 Bad Vision"
            Case -1
                SetListIndex CBO, "-1 Below Average Vision"
            Case 0
                SetListIndex CBO, "0 Average Vision"
            Case 1
                SetListIndex CBO, "1 Above Average Vision"
            Case 2
                SetListIndex CBO, "2 Good Vision"
            Case 3
                SetListIndex CBO, "3 Excellent Vision"
            Case 4
                SetListIndex CBO, "4 Near Perfect Vision"
            Case 5
                SetListIndex CBO, "5 Perfect Vision"
        End Select
    Case 5
        Select Case Val
            Case -1
                SetListIndex CBO, "(-1) Normal"
            Case 0
                SetListIndex CBO, "(0) - Fire"
            Case 1
                SetListIndex CBO, "(1) - Ice"
            Case 2
                SetListIndex CBO, "(2) - Water"
            Case 3
                SetListIndex CBO, "(3) - Lightning"
            Case 4
                SetListIndex CBO, "(4) - Earth"
            Case 5
                SetListIndex CBO, "(5) - Poison"
            Case 6
                SetListIndex CBO, "(6) - Wind"
            Case 7
                SetListIndex CBO, "(7) - Holy"
            Case 8
                SetListIndex CBO, "(8) - Unholy"
        End Select
    Case 6
        Select Case Val
            Case 0
                SetListIndex CBO, "0 - Healing"
            Case 1
                SetListIndex CBO, "1 - Combat"
            Case 2
                SetListIndex CBO, "2 - Teleport"
            Case 3
                SetListIndex CBO, "3 - Bless"
            Case 4
                SetListIndex CBO, "4 - Room Spell"
            Case 5
                SetListIndex CBO, "5 - Party Spell"
        End Select
End Select
End Sub

Public Sub DoItemFlags(dbIndex As Long, dbItemID As Long, Roll As Long, Optional ByRef WasUsed As Long, Optional Inverse As Boolean = False, Optional Flags2 As Boolean = False, Optional AllowSpell As Boolean = True, Optional ThisISNotAnItem As Boolean = False, Optional FlagsAsString As String = "")
'tel# 'teleport, -1,-2, room num
'stu# 'stun
'lig# 'light
'cri# 'crits
'Acc# 'acc
'dam# 'damage
'Str# 'strengh
'agi# 'agility
'cha# 'charm
'dex# 'dexterity
'int# 'intelect
'chp# 'current hp
'mHP# 'max hp
'cma# 'current mana
'mma# 'max mana
'hun# 'hunger
'sta# 'stamina
'cac# 'current AC
'dod# 'dodge
'vis# 'vision
'mit# 'max items
'evi# 'evil points
'pap# 'paper
'mat# 'make item
'SndABC 'Send a message
'sRoABC 'Send a message to the room
'gsp# 'Give spell
'gfa# 'give familiar
'el0
'el1
'el2
'el3
'el4
'el5
'el6
'el7
'el8
'm01-m19
Dim i As Long
Dim sVal As String
Dim dVal As Double
Dim aFlgs() As String
Dim iSpellID As Long
Dim FamID As Long
Dim dbFamId As Long
If Not Flags2 And Not ThisISNotAnItem Then
'    SplitFast dbItems(dbItemID).sFlags, aFlgs, ";"
'    If InStr(1, dbItems(dbItemID).sFlags, "gsp") <> 0 And Not Inverse And AllowSpell And clsSC.FastStringComp(dbItems(dbItemID).sWorn, "scroll") Then
'        For i = LBound(aFlgs) To UBound(aFlgs)
'            dVal = CDbl(Val(Mid$(aFlgs(i), 4)))
'            Select Case Left$(aFlgs(i), 3)
'                Case "gsp"
'                    iSpellID = GetSpellID(, CLng(dVal))
'                    If iSpellID = 0 Then GoTo tNext
'                    If dbPlayers(dbIndex).iSpellLevel >= dbSpells(iSpellID).iLevel Then
'                        If dbPlayers(dbIndex).iSpellType = dbSpells(iSpellID).iType Then
'                            If InStr(1, dbPlayers(dbIndex).sSpells, ":" & dbSpells(iSpellID).lID & ";") Then
'
'                                WasUsed = 1
'                            End If
'
'                        Else
'
'                            WasUsed = 1
'                        End If
'                    Else
'
'                        WasUsed = 1
'                    End If
'                Case Else
'
'            End Select
'tNext:
'            If DE Then DoEvents
'        Next
'    End If
'    If WasUsed = 1 Then Exit Sub
    SplitFast dbItems(dbItemID).sFlags, aFlgs, ";"
ElseIf Not ThisISNotAnItem Then
    SplitFast dbItems(dbItemID).sFlags2, aFlgs, ";"
Else
    SplitFast FlagsAsString, aFlgs, ";"
End If
For i = LBound(aFlgs) To UBound(aFlgs)
    If aFlgs(i) <> "" Then
        sVal = Mid$(aFlgs(i), 4)
        If IsNumeric(sVal) Then
            dVal = CDbl(Val(sVal))
            If dVal = -3 Then dVal = lRoll
            If Inverse Then dVal = -dVal
        End If
        Select Case Left$(aFlgs(i), 3)
            Case "lig"
                With dbPlayers(dbIndex)
                    .iVision = .iVision + dVal
                    
                End With
            Case "cri"
                With dbPlayers(dbIndex)
                    .iCrits = .iCrits + 1
                    
                End With
            Case "acc"
                With dbPlayers(dbIndex)
                    .iAcc = .iAcc + dVal
                    
                End With
            Case "dam"
                With dbPlayers(dbIndex)
                    .iMaxDamage = .iMaxDamage + dVal
                    
                End With
            Case "str"
                With dbPlayers(dbIndex)
                    .iStr = .iStr + dVal
                    
                End With
            Case "agi"
                With dbPlayers(dbIndex)
                    .iAgil = .iAgil + dVal
                    
                End With
            Case "cha"
                With dbPlayers(dbIndex)
                    .iCha = .iCha + dVal
                    
                End With
            Case "dex"
                With dbPlayers(dbIndex)
                    .iDex = .iDex + dVal
                    
                End With
            Case "int"
                With dbPlayers(dbIndex)
                    .iInt = .iInt + dVal
                    
                End With
            Case "chp"
                With dbPlayers(dbIndex)
                    .lHP = .lHP + dVal
                    If .lHP > .lMaxHP Then .lHP = .lMaxHP
                    
                End With
            Case "mhp"
                With dbPlayers(dbIndex)
                    .lMaxHP = .lMaxHP + dVal
                    
                End With
            Case "cma"
                With dbPlayers(dbIndex)
                    .lMana = .lMana + dVal
                    If .lMana > .lMaxMana Then .lMana = .lMaxMana
                    
                End With
            Case "mma"
                With dbPlayers(dbIndex)
                    .lMaxMana = .lMaxMana + dVal
                    
                End With
            Case "hun"
                With dbPlayers(dbIndex)
                    .dHunger = .dHunger + dVal
                    
                End With
            Case "sta"
                With dbPlayers(dbIndex)
                    .dStamina = .dStamina + dVal
                    
                End With
            Case "cac"
                With dbPlayers(dbIndex)
                    .iAC = .iAC + dVal
                    
                End With
            Case "dod"
                With dbPlayers(dbIndex)
                    .iDodge = .iDodge + dVal
                    
                End With
            Case "vis"
                With dbPlayers(dbIndex)
                    .iVision = .iVision + dVal
                    
                End With
            Case "mit"
                modMiscFlag.SetStatsPlus dbIndex, [Max Items Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Max Items Bonus]) + CLng(dVal)
            Case "evi"
                With dbPlayers(dbIndex)
                    .iEvil = .iEvil + dVal
                    
                End With
            Case "pap"
                With dbPlayers(dbIndex)
                    
                    .lPaper = .lPaper + dVal
                    If .lPaper < 0 Then .lPaper = 0
                    
                End With
            Case "mat"
                dbFamId = GetItemID(, CLng(dVal))
                If dbFamId = 0 Then GoTo nNext
                With dbPlayers(dbIndex)
                    If modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) + 1 < modMiscFlag.GetStatsPlusTotal(dbIndex, [Max Items]) Then
                        If .sInventory = "0" Then .sInventory = ""
                        .sInventory = .sInventory & ":" & dbItems(dbFamId).iID & "/" & dbItems(dbFamId).iUses & "/" & dbItems(dbFamId).lDurability & ";"
                    Else
                        With dbMap(GetMapIndex(.lLocation))
                            If .sItems = "0" Then .sItems = ""
                            .sItems = .sItems & ":" & dbItems(dbFamId).iID & "/" & dbItems(dbFamId).iUses & "/" & dbItems(dbFamId).lDurability & ";"
                        End With
                    End If
                End With
'            Case "gsp"
'                If Not Inverse And Not ThisISNotAnItem And Not Flags2 And AllowSpell And clsSC.FastStringComp(dbItems(dbItemID).sWorn, "scroll") Then
'                    iSpellID = GetSpellID(, CLng(dVal))
'                    If iSpellID = 0 Then GoTo nNext
'                    If clsSC.FastStringComp(dbPlayers(dbIndex).sSpells, "0") Then dbPlayers(dbIndex).sSpells = ""
'                    dbPlayers(dbIndex).sSpells = dbPlayers(dbIndex).sSpells & ":" & dbSpells(iSpellID).lID & ";"
'                    If clsSC.FastStringComp(dbPlayers(dbIndex).sSpellShorts, "0") Then dbPlayers(dbIndex).sSpellShorts = ""
'                    dbPlayers(dbIndex).sSpellShorts = dbPlayers(dbIndex).sSpellShorts & dbSpells(iSpellID).sShort & ";"
'                End If
'            Case "gfa"
'                FamID = CLng(dVal)
'                dbFamID = GetFamID(FamID)
'                If dbFamID = 0 Then GoTo nNext
'                RemoveStats dbPlayers(dbIndex).iIndex
'                With dbPlayers(dbIndex)
'                    .sFamName = dbFamiliars(dbFamID).sFamName
'                    .iFamID = FamID
'                    .lFamMaxHP = RndNumber(CDbl(dbFamiliars(dbFamID).lStartHPMin), CDbl(dbFamiliars(dbFamID).lStartHPMax))
'                    .dFamEXP = 0
'                    .lFamCurrentHP = .lFamMaxHP
'                End With
'                AddStats dbPlayers(dbIndex).iIndex
            Case "sas"
                If Inverse Then
                    With dbPlayers(dbIndex)
                        .sPlayerName = .sSeenAs
                    End With
                Else
                    With dbPlayers(dbIndex)
                        .sPlayerName = sVal
                    End With
                End If
            Case "des"
                If Inverse Then
                    With dbPlayers(dbIndex)
                        .sOverrideDesc = "0"
                    End With
                Else
                    With dbPlayers(dbIndex)
                        .sOverrideDesc = sVal
                    End With
                End If
            Case "thi"
                modMiscFlag.SetStatsPlus dbIndex, [Thieving Bonus], modMiscFlag.GetStatsPlus(dbIndex, [Thieving Bonus]) + CLng(dVal)
        End Select
        If Left$(aFlgs(i), 2) Like "el#" Then
            If Inverse Then
                Select Case dVal
                    Case -1
                        dVal = 0
                    Case 0
                        dVal = 1
                    Case Else
                        dVal = 0
                End Select
            End If
            modResist.UpdateResistValue dbIndex, CLng(Val(Mid$(aFlgs(i), 3, 1))), CLng(dVal)
        End If
        If Left$(aFlgs(i), 3) Like "m##" Then
            If Inverse Then
                Select Case dVal
                    Case -1
                        dVal = 0
                    Case 0
                        dVal = 1
                    Case Else
                        dVal = 0
                End Select
            End If
            modMiscFlag.SetMiscFlag dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2))), CLng(dVal)
        End If
        If Left(aFlgs(i), 3) Like "s##" Then
            Select Case Val(Mid$(aFlgs(i), 2, 2))
                Case 1, 3, 5, 9, 11, 13
                    modMiscFlag.SetStatsPlus dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2))), modMiscFlag.GetStatsPlus(dbIndex, CLng(Val(Mid$(aFlgs(i), 2, 2)))) + CLng(dVal)
            End Select
        End If
    End If
nNext:
    If DE Then DoEvents
Next
If Not ThisISNotAnItem Then
    If Inverse Then
        dbPlayers(dbIndex).iAC = dbPlayers(dbIndex).iAC - dbItems(dbItemID).iAC
    Else
        dbPlayers(dbIndex).iAC = dbPlayers(dbIndex).iAC + dbItems(dbItemID).iAC
    End If
End If
End Sub
