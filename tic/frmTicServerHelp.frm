VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Help Index"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10950
   Icon            =   "frmTicServerHelp.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   6120
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser webHelp 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      ExtentX         =   18865
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
'*************************************************************************************
'***************       Code create by Chris Van Hooser          **********************
'***************                  (c)2001                       **********************
'*************** You may use this code and freely distribute it **********************
'***************   If you have any questions, please email me   **********************
'***************          at theendorbunker@attbi.com.          **********************
'***************       Thanks for downloading my project        **********************
'***************        and i hope you can use it well.         **********************
'***************                TicServer                       **********************
'***************                tic.vbp                         **********************
'*************************************************************************************
'*************************************************************************************

Private Sub Form_Load()
Call Form_Resize 'resize the webbrowser to the correct width
webHelp.Navigate App.Path & "\help\tic_server.htm" 'load the help file
End Sub

Private Sub Form_Resize()
webHelp.Top = 0 'make top and left 0
webHelp.Left = 0
webHelp.Height = Me.Height - 425 'and make it perfectly fit the form
webHelp.Width = Me.Width - 100
End Sub

