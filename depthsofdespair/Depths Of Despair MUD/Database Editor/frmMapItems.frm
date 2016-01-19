VERSION 5.00
Begin VB.Form frmMapItems 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Room Floor Item Editor"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
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
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
      Width           =   375
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
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
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.ListBox lstItems 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ListBox lstItemsThere 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMapItems"
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
Rem***************                frmMapItems                     **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Private Sub cmdAdd_Click()
If lstItems.ListIndex > -1 Then lstItemsThere.AddItem lstItems.List(lstItems.ListIndex)
End Sub

Private Sub cmdDrop_Click()
If lstItemsThere.ListIndex > -1 Then lstItemsThere.RemoveItem (lstItemsThere.ListIndex)
End Sub

Private Sub cmdOK_Click()
Dim tVar As String
If lstItemsThere.ListCount = 0 Then
    frmMap.txtItems.Text = "nothing;"
    frmMap.Enabled = True
    Unload Me
    Exit Sub
End If
For i = 0 To lstItemsThere.ListCount - 1
    tVar = tVar & lstItemsThere.List(i) & ";"
Next
frmMap.txtItems.Text = tVar
frmMap.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
LoadItemsToList
FillItemsThere frmMap.txtItems
End Sub

Sub LoadItemsToList()
lstItems.Clear
With RSItem
    .MoveFirst
    Do
        lstItems.AddItem !ItemName
        .MoveNext
    Loop Until .EOF
End With
End Sub

Sub FillItemsThere(sString As String)
lstItemsThere.Clear
If sString = "nothing;" Then Exit Sub
Dim tArr() As String
tArr = Split(Left$(sString, Len(sString) - 1), ";")
For i = 0 To UBound(tArr)
    lstItemsThere.AddItem tArr(i)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMap.Enabled = True
End Sub
