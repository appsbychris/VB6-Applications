VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2745
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1560
      Top             =   1080
   End
   Begin SHDocVwCtl.WebBrowser webGif 
      Height          =   3615
      Left            =   -360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -360
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
      Location        =   ""
   End
End
Attribute VB_Name = "Form3"
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
'***************                TicBoard                        **********************
'***************                TicBoard.vbp                    **********************
'*************************************************************************************
'*************************************************************************************
Private Sub Form_Load()
On Error GoTo Form_Load_Error
'load the splash screen
'Get the gif file from the resource file
Dim b() As Byte, strText$ 'byte array, and a string
b() = LoadResData(101, "TICGIF") 'get the gif file
strText$ = StrConv(b, vbUnicode) 'convert it to unicode
Open App.Path & "\tic.gif" For Output As #1 'save it tempararily
    Print #1, strText$
Close #1
'get the gif, and show it
webGif.Navigate App.Path & "\tic.gif"
On Error GoTo 0
Exit Sub
Form_Load_Error:
MsgBox "An error has occured" & vbCrLf & "Error #" & Err.Number & ": " & Err.Description & vbCrLf & "Procedure that caused error: Form_Load in Form, Form3"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'delete the gif file
Kill App.Path & "\tic.gif"
'load the first form
Load Form1
Form1.Show 'show it
Timer1.Enabled = False 'disable this timer
Unload Me 'unload the form
End Sub
