VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000A&
   Caption         =   "Editor"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   11025
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CDLMain 
      Left            =   2640
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnufads 
         Caption         =   "Click"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCharacters 
      Caption         =   "&Characters"
      Begin VB.Menu mnuPlayers 
         Caption         =   "&Players"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRaces 
         Caption         =   "&Races"
      End
      Begin VB.Menu mnuClasses 
         Caption         =   "&Classes"
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnupos 
         Caption         =   "Possibilities"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMainItems 
      Caption         =   "I&tems"
      Begin VB.Menu mnuItems 
         Caption         =   "&Items"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpells 
         Caption         =   "&Spells"
      End
      Begin VB.Menu mnuDash00006 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShops 
         Caption         =   "S&hops"
      End
   End
   Begin VB.Menu mnuMisx 
      Caption         =   "&Misc"
      Begin VB.Menu mnuEmotions 
         Caption         =   "&Emotions"
      End
      Begin VB.Menu mnuDashsdklaf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFamiliars 
         Caption         =   "&Familiars"
      End
   End
   Begin VB.Menu mnuLocations 
      Caption         =   "&Locations"
      Begin VB.Menu mnuMap 
         Caption         =   "&Map"
      End
      Begin VB.Menu mnuQuickMap 
         Caption         =   "&Quick"
      End
   End
   Begin VB.Menu mnuMainMonsters 
      Caption         =   "M&onsters"
      Begin VB.Menu mnuMonsters 
         Caption         =   "M&onsters"
      End
   End
End
Attribute VB_Name = "mdiMain"
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
Rem***************                mdiMain                         **********************
Rem***************                ServerEditor                    **********************
Rem***************                Editor.vbp                      **********************
Rem*************************************************************************************
Rem*************************************************************************************

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
Set MRS = Nothing
Set MRSITEM = Nothing
Set MRSMAP = Nothing
Set MRSMONSTER = Nothing
Set MRSCLASS = Nothing
Set MRSRACE = Nothing
Set MRSEMOTIONS = Nothing
Set MRSSPELLS = Nothing
Set MRSEVENTS = Nothing
DB.Close
Set DB = Nothing
modSec.dB_set_and_load "Standard Jet DB", App.Path & "\data.mud", "spike technolog"
modSec.dB_set_and_load "4.0", App.Path & "\data.mud", "MUD"
End Sub

Private Sub mnuClasses_Click()
Load frmClasses
frmClasses.Show
End Sub

Private Sub mnuEmotions_Click()
Load frmEmotions
frmEmotions.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnufads_Click()
Load frmClickAndAdd
frmClickAndAdd.Show
End Sub

Private Sub mnuFamiliars_Click()
Load frmFamiliars
frmFamiliars.Show
End Sub

Private Sub mnuItems_Click()
Load frmItems
frmItems.Show
End Sub

Private Sub mnuMap_Click()
Load frmMapp
frmMapp.Show
End Sub

Private Sub mnuMonsters_Click()
Load frmMonsters
frmMonsters.Show
End Sub

Private Sub mnuPlayers_Click()
Load frmPlayers
frmPlayers.Show
End Sub

Private Sub mnupos_Click()
Load frmPossible
frmPossible.Show
End Sub

Private Sub mnuQuickMap_Click()
Load frmMapEdit
frmMapEdit.Show
End Sub

Private Sub mnuRaces_Click()
Load frmRace
frmRace.Show
End Sub

Private Sub mnuShops_Click()
Load frmShops
frmShops.Show
End Sub

Private Sub mnuSpells_Click()
Load frmSpells
frmSpells.Show
End Sub
