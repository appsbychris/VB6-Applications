VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plug-ins"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":08CA
   ScaleHeight     =   2865
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   360
      Top             =   1080
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   3120
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   1560
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   40
      TabIndex        =   1
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TickCount As Integer

Private Sub Form_Load()
On Error Resume Next
'Set the path for the file list box
File1.Path = App.Path
'set tickout to begining of 1
TickCount = 1
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'makes the cheap "progress"/"busy" thing work
With Label1
    If TickCount = 1 Then
        .Caption = "|"
        TickCount = 2
        Exit Sub
    ElseIf TickCount = 2 Then
        .Caption = "/"
        TickCount = 3
        Exit Sub
    ElseIf TickCount = 3 Then
        .Caption = "-"
        TickCount = 4
        Exit Sub
    ElseIf TickCount = 4 Then
        .Caption = "\"
        TickCount = 5
        Exit Sub
    ElseIf TickCount = 5 Then
        .Caption = "-"
        TickCount = 1
        Exit Sub
    End If
End With

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'Call the Sub to scan for plugins
Call ScanPlug
'disable timer2
Timer2.Enabled = False

End Sub

Sub ScanPlug()
On Error Resume Next
'load form1 into memory
Load Form1
Dim X As Integer
'search for the plugins, and show progress
For X = 0 To File1.ListCount - 1
    If InStr(1, LCase(File1.List(X)), "plug") Then
        List1.AddItem "...Plug-in found...Adding..."
        Form1.Plugs.AddItem File1.Path & "\" & File1.List(X)
        List1.AddItem "...Plug-in added...searching for more..."
        List1.Selected(List1.ListCount - 1) = True
    End If
Next
'Once they are all found, enable timer 3
List1.AddItem "...All plug-ins found...loading..."
List1.Selected(List1.ListCount - 1) = True
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
'show form 1
Form1.Show
Unload Me
End Sub
