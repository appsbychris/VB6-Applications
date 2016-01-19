VERSION 5.00
Begin VB.Form frmMapShop 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Edit a Shop"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   3420
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
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
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
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
      Left            =   720
      TabIndex        =   2
      Top             =   1560
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
      Height          =   1260
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
   End
   Begin VB.ListBox lstShop 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmMapShop"
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
Rem***************                frmMapShop                      **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Private Sub cmdAdd_Click()
If lstItems.ListIndex > -1 Then lstShop.AddItem lstItems.List(lstItems.ListIndex)
End Sub

Private Sub cmdDrop_Click()
If lstShop.ListIndex > -1 Then lstShop.RemoveItem (lstShop.ListIndex)
End Sub

Private Sub cmdOK_Click()
If lstShop.ListCount = 0 Then
    frmMap.txtShopItems.Text = "0"
    frmMap.Enabled = True
    Unload Me
    Exit Sub
End If
Dim tVar As String
With RSItem
    For i = 0 To lstShop.ListCount - 1
        .MoveFirst
        Do
            If !ItemName = lstShop.List(i) Then
                tVar = tVar & !ID & ";"
                Exit Do
            ElseIf Not .EOF Then
                .MoveNext
            End If
        Loop Until .EOF
    Next
End With
frmMap.txtShopItems.Text = tVar
frmMap.Enabled = True
Unload Me
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

Sub LoadShop(sString As String)
lstShop.Clear
If sString = "0" Then Exit Sub
Dim tArr() As String
tArr = Split(Left$(sString, Len(sString) - 1), ";")
With RSItem
    For i = 0 To UBound(tArr)
        .MoveFirst
        Do
            If CInt(!ID) = CInt(tArr(i)) Then
                lstShop.AddItem !ItemName
                Exit Do
            ElseIf Not .EOF Then
                .MoveNext
            End If
        Loop Until .EOF
    Next
End With
End Sub

Private Sub Form_Load()
LoadItemsToList
LoadShop frmMap.txtShopItems.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMap.Enabled = True
End Sub
