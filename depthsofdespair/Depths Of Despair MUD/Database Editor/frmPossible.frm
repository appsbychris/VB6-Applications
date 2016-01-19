VERSION 5.00
Begin VB.Form frmPossible 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Possible Race/Class Combonations And EXP chart for them"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   6135
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin ServerEditor.FlagOptions flgOpts 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   4470
      _extentx        =   7885
      _extenty        =   556
      style           =   3
   End
   Begin VB.ListBox lstEXP 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      IntegralHeight  =   0   'False
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ListBox lstCombos 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   45
   End
End
Attribute VB_Name = "frmPossible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
modMain.FeedAList flgOpts, "rooms"
flgOpts.FillNow
End Sub

Private Sub Form_Load()
'GetPossibilities
End Sub

Sub GetPossibilities()
Dim aRaces() As String
Dim aClasses() As String
Dim sRaces As String
Dim sClasses As String
Dim sCombos As String
Dim aCombos() As String
With RSRace
    .MoveFirst
    Do
        sRaces = sRaces & !Name & ";"
        .MoveNext
    Loop Until .EOF
End With
With RSClass
    .MoveFirst
    Do
        sClasses = sClasses & !Name & ";"
        .MoveNext
    Loop Until .EOF
End With
aRaces() = Split(Left$(sRaces, Len(sRaces) - 1), ";")
aClasses() = Split(Left$(sClasses, Len(sClasses) - 1), ";")
For i = 0 To UBound(aRaces)
    For a = 0 To UBound(aClasses)
        sCombos = sCombos & aRaces(i) & "<>" & aClasses(a) & ";"
    
    Next
    
Next
aCombos() = Split(Left$(sCombos, Len(sCombos) - 1), ";")
For i = 0 To UBound(aCombos)
    lstCombos.AddItem aCombos(i)
Next
lblLabel.Caption = "There are a total of " & lstCombos.ListCount & " possible race/class mixes."
End Sub

Sub GetEXPChart()
Dim sRace As String
Dim sClass As String
Dim dRaceEXP As Double
Dim dClassEXP As Double
Dim dTotalEXP As Double
Dim aEXP() As String
Dim sEXP As String
lstEXP.Clear
sRace = Left$(lstCombos.Text, InStr(1, lstCombos.Text, "<>") - 1)
sClass = Mid$(lstCombos.Text, InStr(1, lstCombos.Text, "<>") + 2, Len(lstCombos.Text) - InStr(1, lstCombos.Text, "<>") + 2)
With RSRace
    .MoveFirst
    Do
        If sRace = !Name Then
            dRaceEXP = CDbl(!Exp)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
With RSClass
    .MoveFirst
    Do
        If sClass = !Name Then
            dClassEXP = CDbl(!Exp)
            Exit Do
        ElseIf Not .EOF Then
            .MoveNext
        End If
    Loop Until .EOF
End With
dTotalEXP = dRaceEXP + dClassEXP
sEXP = sEXP & "[1] " & dTotalEXP & ";"
For i = 2 To 200
    Select Case i
        Case 2 To 10:
            dTotalEXP = dTotalEXP + (((dTotalEXP * 5.4) / i) + (dTotalEXP / i - 1.9))
        Case 11 To 30:
            dTotalEXP = dTotalEXP + (((dTotalEXP * 2.9) / (i - 1.5)) + dTotalEXP / (i - 1.8))
        Case 31 To 40:
            dTotalEXP = dTotalEXP + (((dTotalEXP * 12.3) / (i - 0.2)) + dTotalEXP / (i - 1.9))
        Case 41 To 50:
            dTotalEXP = dTotalEXP + (((dTotalEXP * 1.2) / (i - 0.4)) + dTotalEXP / (i - 1.6))
        Case 51 To 55:
            dTotalEXP = dTotalEXP + (((dTotalEXP * 1.1) / (i - 0.7)) + dTotalEXP / (i - 1.2))
        Case 56 To 75:
            dTotalEXP = dTotalEXP + (((dTotalEXP * 2.8) / (i - 0.1)) + dTotalEXP / (i - 1.4))
        Case 76 To 200:
            dTotalEXP = dTotalEXP + (((dTotalEXP * 1.7) / (i - 0.9)) + dTotalEXP / (i - 1.9))
    End Select
    sEXP = sEXP & "[" & i & "] " & FormatNumber(dTotalEXP, 0, vbUseDefault, vbFalse, vbTrue) & ";"
Next

aEXP = Split(Left$(sEXP, Len(sEXP) - 1), ";")
For i = 0 To UBound(aEXP)
    lstEXP.AddItem aEXP(i)
Next
End Sub
Private Sub lstCombos_Click()
GetEXPChart
End Sub
