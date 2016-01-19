VERSION 5.00
Begin VB.Form frmOverview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Overview"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13350
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   591
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   890
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmOverview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aMap() As String
Private Type typForm
    lCenterY As Long
    lCenterX As Long
    lHeight As Long
    lWidth As Long
End Type
Private Type typExits
    North() As Long
    South() As Long
    East()  As Long
    West()  As Long
    NW()    As Long
    SW()    As Long
    SE()    As Long
    NE()    As Long
End Type
Dim udtForm As typForm

Private Sub Form_Load()
With udtForm
    .lHeight = Me.ScaleHeight
    .lWidth = Me.ScaleWidth
    .lCenterY = (.lHeight \ 2)
    .lCenterX = (.lWidth \ 2)
End With
End Sub

Private Sub LoadMapArray(Optional lStart As Long = 1)
Dim x As Long

x = GetMapIndex(lStart)

End Sub
