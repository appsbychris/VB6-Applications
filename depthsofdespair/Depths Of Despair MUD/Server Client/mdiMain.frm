VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11700
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar mTB 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1138
      ButtonWidth     =   2593
      ButtonHeight    =   979
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Connection"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Load frmConsole
frmConsole.Show
End Sub

Private Sub mTB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Dim newconn As New frmConsole
        newconn.Show
End Select
End Sub
