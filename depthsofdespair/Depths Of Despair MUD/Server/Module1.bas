Attribute VB_Name = "modMain"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modMain
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Rem color.brightgreen & ;Running version 1.00 ; & color.brightyellow & ;ALPHA; & color.brightgreen & ; of Depths Of Despair MUD;

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public ColorLst() As String
Public HairLen() As String
Public HairStyle() As String


'=====================================================================================
'Database stored in memory

Public lAgeD As Long
Public lDeath As Long
Public lIsPvP As Long
Public lPvPLevel As Long

'Public modgetdata As clsGet
'Public clsSC As clsCompare
'Public clsReg As CReadWriteEasyReg
Public bUpdate As Boolean
Public sGraphic As String

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

Public dbMBTimer() As UDTMyBaseTimer


Rem Public Type Arrays
Rem used to hold a bunch of the types, for each user/monster
Public pWeapon()    As Weapon
Public aMons()      As Monster
Rem///////////////////////////////



Rem/////////////////////////////
Rem Public Arrays for users online,
Public SpellCombat()        As Boolean
Public MaxUsers             As Long
Public X()                  As String
Public PNAME()              As String
Public pLogOn()             As Boolean
Public pLogOnPW()           As Boolean
Public pPoint()             As Long
Rem///////////////////////////////

Rem/////////////////////////////
Rem Array to hold all the possible emotions
Rem for the players to use
Public Emotions()   As String
Rem//////////////////////////////

Rem////////////////////////////
Rem Variables for the max amount of
Rem monsters, and the current amount
Rem that are in the game
Public MaxMonsters  As Long
Public AmountMons   As Long
Rem/////////////////////////////


Rem The database
Public db           As Database
Rem All the recordsets that everyone will use
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
Rem/////////////////////////////////

Rem/////////////////////////////
Public sBuild As String

Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As Currency) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Dim cStartTime As Currency
Dim cPerfFreq As Currency

Public Function StartTimer() As Long
  If QueryPerformanceFrequency(cPerfFreq) = False Then
    Debug.Print "High-perf counter not supported"
  End If
  QueryPerformanceCounter cStartTime
End Function

Public Function FinishTimer() As Currency
  Dim cCurrentTime As Currency
  
  QueryPerformanceCounter cCurrentTime
  FinishTimer = (cCurrentTime - cStartTime) / cPerfFreq
End Function

Sub Main()
On Error GoTo Main_Error
Dim f As Long
Dim s As String
Dim Arr() As String
Dim i As Long
f = FreeFile
Open App.Path & "\colors.aimg" For Binary As #f
    s = Input$(LOF(f), f)
Close #f
SplitFast s, Arr, vbCrLf
f = 0
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" Then f = f + 1
Next
ReDim ColorLst(f - 1)
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" Then
        ColorLst(i) = Arr(i)
    End If
Next

f = FreeFile
Open App.Path & "\hairlength.aimg" For Binary As #f
    s = Input$(LOF(f), f)
Close #f
SplitFast s, Arr, vbCrLf
f = 0
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" Then f = f + 1
Next
ReDim HairLen(f - 1)
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" Then
        HairLen(i) = Arr(i)
    End If
Next

f = FreeFile
Open App.Path & "\hairstyle.aimg" For Binary As #f
    s = Input$(LOF(f), f)
Close #f
SplitFast s, Arr, vbCrLf
f = 0
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" Then f = f + 1
Next
ReDim HairStyle(f - 1)
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" Then
        HairStyle(i) = Arr(i)
    End If
Next

ReDim dbLetters(1 To 1)
ReDim dbMBTimer(0)
Load frmSplash
frmSplash.Show
Screen.MousePointer = vbHourglass
modDatabase.OpenDatabaseConnection
'Set modgetdata = New clsGet
'Set clsReg = New CReadWriteEasyReg
'Set clsSC = New clsCompare
UpdateList "Loading Depths of Despair MUD Server... }b(}n}i" & Time & "}n}b)"
sBuild = ByValGetBuildVal

frmSplash.LoadServer
On Error GoTo 0
Exit Sub
Main_Error:
Screen.MousePointer = vbDefault
WriteINI "MaxPlayers", "20"
WriteINI "DeathLevel", "-5"
WriteINI "DeathRoom", "1"
WriteINI "MaxMonsters", "600"
WriteINI "ShowTime", "True"
WriteINI "LogFile", "True"
WriteINI "Logons", "1"
WriteINI "PvPE", "1"
WriteINI "PvPL", "10"
WriteINI "Age", "0"
MsgBox "An error has occured in startup." & vbCrLf & "Cause may be a corrupted database, or corrupted user defined information." & vbCrLf & "User defined information will now be set to default." & vbCrLf & "If you wish to restore the database, on the next prompt, select YES." & vbCrLf & "Please restart the server after the message dissapears.", vbCritical, "Error"
If MsgBox("Do you wish to restore the database?" & vbCrLf & "It will resort to what it was since you last shut down the server.", vbQuestion + vbYesNo, "Restore Database") = vbYes Then
    FileCopy App.Path & "\BACKUP.BAK", App.Path & "\data.mud"
End If
Unload frmMain
End Sub

Public Function DCount(sInString As String, sCountString As String) As Long
Dim iNextOccur As Long
iNextOccur = 0
Do
    iNextOccur = InStr(iNextOccur + 1, sInString, sCountString)
    If iNextOccur > 0 Then DCount = DCount + 1 Else Exit Do
    If DE Then DoEvents
Loop
End Function

Public Function RndNumber(Min As Double, Max As Double) As Long
Randomize Timer
RndNumber = CLng(RoundFast(CDbl((Rnd * (Max - Min)) + Min), 0))
End Function

Public Sub UpdateList(AddWhat$, Optional IsError As Boolean = False)
If frmMain.lstEvents.ListCount >= 1000 Then frmMain.lstEvents.Clear
If modSC.FastStringComp(GetINI("LogFile"), "True") Then WriteToLog AddWhat$
If IsError Then
    frmMain.lstEvents.AddItem AddWhat$, FCOLOR:=vbBlack, BCOLOR:=vbRed
Else
    frmMain.lstEvents.AddItem AddWhat$
End If
frmMain.lstEvents.SetSelected frmMain.lstEvents.ListCount, True
End Sub

Public Function WaitFor(MS As Long)
Dim Start As Long
Start = GetTickCount
While Start + MS > GetTickCount
    If DE Then DoEvents
Wend
End Function

Sub WriteToLog(WriteWhat$)
    Dim FileNumber&
    FileNumber = FreeFile
    Open App.Path & "\mud.log" For Append Shared As #FileNumber
        Print #FileNumber, "[" & Date & "|" & Time & "] - " & WriteWhat$
    Close #FileNumber
End Sub

Sub CheckLogFileSize()
Dim tString As String
On Error GoTo CheckLogFileSize_Error
tString = FileLen(App.Path & "\mud.log")
If Val(tString) > 1048576 Then
    AlwaysOnTop frmSplash, False
    If MsgBox("The log file is getting quite large." & vbCrLf & "Would you like to clear it?", vbQuestion + vbOKCancel, "Large log file") = vbOK Then
        Open App.Path & "\mud.log" For Output As #1
            Print #1, ""
        Close #1
    End If
    AlwaysOnTop frmSplash, True
End If
On Error GoTo 0
Exit Sub
CheckLogFileSize_Error:
End Sub

Public Function CheckBIP(ip As String) As Boolean
Dim FileNumber&
Dim s As String
Dim tArr() As String
Dim i As Long
FileNumber = FreeFile
Open App.Path & "\ipb.list" For Binary As #FileNumber
    If DE Then DoEvents
    s = Input$(LOF(1), 1)
Close #FileNumber
SplitFast s, tArr, vbCrLf
For i = LBound(tArr) To UBound(tArr)
    If tArr(i) <> "" Then
        If modSC.FastStringComp(ip, tArr(i)) Then
            CheckBIP = True
            Exit Function
        End If
    End If
    If DE Then DoEvents
Next
End Function

