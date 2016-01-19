VERSION 5.00
Object = "*\AReg.vbp"
Begin VB.Form frmTest 
   Caption         =   "Demo of Reg By Chris Van Hooser"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin QuickReg.QReg Reg1 
      Left            =   3960
      Top             =   3720
      _ExtentX        =   1058
      _ExtentY        =   1376
      Strength        =   0
      ScrambleStrength=   0
      SerialNumber    =   "1234"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Strength of Scramble"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   4455
      Begin VB.OptionButton optScramStr 
         Caption         =   "Toughest"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optScramStr 
         Caption         =   "Tougher"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optScramStr 
         Caption         =   "Tough"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optScramStr 
         Caption         =   "Minor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraStrKeyGen 
      Caption         =   "Strength of Key Gening"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   4455
      Begin VB.OptionButton optStrChoices 
         Caption         =   "XtraMax"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optStrChoices 
         Caption         =   "Maximum"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optStrChoices 
         Caption         =   "Medium"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optStrChoices 
         Caption         =   "Minimum"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame fraCheckType 
      Caption         =   "Check Type-"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   4455
      Begin VB.OptionButton optChoices 
         Caption         =   "Combo of Both"
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Serial Number"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optChoices 
         Caption         =   "Computer Name"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSetSerialNum 
      Caption         =   "Set Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "1234"
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Key"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "Gen A Key"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtTF 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCheck_Click()
txtTF.Text = Reg1.CheckAKey(txtKey.Text)
End Sub

Private Sub cmdGen_Click()
txtKey.Text = Reg1.GenAKey
End Sub

Private Sub cmdSetSerialNum_Click()
If txtSerialNumber.Text <> "" Then Reg1.SerialNumber = txtSerialNumber.Text
End Sub

Private Sub optChoices_Click(Index As Integer)
Select Case Index
    Case 0:
        Reg1.CheckType = ComputerName
    Case 1:
        Reg1.CheckType = cSerialNumber
    Case 2:
        Reg1.CheckType = ComboOfBoth
End Select
End Sub

Private Sub optScramStr_Click(Index As Integer)
Select Case Index
    Case 0:
        Reg1.ScrambleStrength = Minor
    Case 1:
        Reg1.ScrambleStrength = Tough
    Case 2:
        Reg1.ScrambleStrength = Tougher
    Case 3:
        Reg1.ScrambleStrength = Toughest
End Select
End Sub

Private Sub optStrChoices_Click(Index As Integer)
Select Case Index
    Case 0:
        Reg1.Strength = Minimum
    Case 1:
        Reg1.Strength = Medium
    Case 2:
        Reg1.Strength = Maximum
    Case 3:
        Reg1.Strength = XtraMax
 End Select
End Sub
