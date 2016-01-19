VERSION 5.00
Begin VB.Form frmMapMonsters 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Edit Monsters"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdd 
      Caption         =   "^add"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdDrop 
      Caption         =   "d r op    v"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1680
      Width           =   375
   End
   Begin VB.ListBox lstMonsters 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2535
   End
   Begin VB.ListBox lstMonstersThere 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMapMonsters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem*************************************************************************************
Rem*************************************************************************************
Rem***************       Code create by Chris Van Hooser          **********************
Rem***************                  (c)2002                       **********************
Rem*************** You may use this code and freely distribute it **********************
Rem***************   If you have any questions, please email me   **********************
Rem***************          at theendorbunker@attbi.com.          **********************
Rem***************       Thanks for downloading my project        **********************
Rem***************        and i hope you can use it well.         **********************
Rem***************                frmMapMonsters                  **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Private Sub cmdAdd_Click()
If lstMonsters.ListIndex <> -1 Then lstMonstersThere.AddItem lstMonsters.List(lstMonsters.ListIndex)
End Sub

Private Sub cmdDrop_Click()
If lstMonstersThere.ListIndex > -1 Then lstMonstersThere.RemoveItem (lstMonstersThere.ListIndex)
End Sub

Private Sub cmdOK_Click()
Dim tVar As String
If lstMonstersThere.ListCount = 0 Then
    frmMap.txtMonsters.Text = "0;"
    frmMap.Enabled = True
    Unload Me
    Exit Sub
End If
For i = 0 To lstMonstersThere.ListCount - 1
    tVar = tVar & lstMonstersThere.List(i) & ";"
Next
frmMap.txtMonsters.Text = tVar
frmMap.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
FillMonstersList
FillMonstersThere frmMap.txtMonsters.Text
End Sub

Sub FillMonstersThere(sString As String)
lstMonstersThere.Clear
If sString = "0;" Then Exit Sub
Dim tArr() As String
tArr = Split(Left$(sString, Len(sString) - 1), ";")
For i = 0 To UBound(tArr)
    lstMonstersThere.AddItem tArr(i)
Next
End Sub

Sub FillMonstersList()
With RSMonster
    .MoveFirst
    Do
        lstMonsters.AddItem !MonsterName
        .MoveNext
    Loop Until .EOF
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMap.Enabled = True
End Sub
