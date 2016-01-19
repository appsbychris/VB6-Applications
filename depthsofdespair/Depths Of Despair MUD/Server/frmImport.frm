VERSION 5.00
Begin VB.Form frmImport 
   Caption         =   "Import Database"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3075
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox fl 
      Height          =   2430
      Left            =   3000
      Pattern         =   "*.mud"
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.DirListBox dr 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox drv 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin DoDMudServer.eButton cmdImport 
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&Import"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      hCol            =   12632256
      bCol            =   12632256
      CA              =   2
   End
   Begin DoDMudServer.eButton cmdCancel 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      hCol            =   12632256
      bCol            =   12632256
      CA              =   2
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : frmImport
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private db2           As Database
Private MRSMAP2       As Recordset
Private MRSCLASS2     As Recordset
Private MRSRACE2      As Recordset
Private MRSITEM2      As Recordset
Private MRSMONSTER2   As Recordset
Private MRSEMOTIONS2  As Recordset
Private MRSSPELLS2    As Recordset
Private MRSFAMILIARS2 As Recordset
Private MRSSHOPS2     As Recordset

Private Sub UpdateMRSSets2()
modSec.dB_set_and_load "spike technolog", fl.Path & "/" & fl.FileName, modSec.uJunkIt("S3t2]aJJMnWH≥âdaì†f)aƒZwÄ(vrqr@ãÄøçdíK¢øëï\ï i´+{nb°vKJJ∞vá@14zï∫emΩFZî4^`b≤óteNâå`R√qe∑D´ áÇp_Hvüzßêí™≈D4ï=BaáfêhMm¿seBDek,«∑M•a=•âù£Ø")
modSec.dB_set_and_load "MUD", fl.Path & "/" & fl.FileName, "4.0"
Set db2 = OpenDatabase(fl.Path & "/" & fl.FileName, False, False, modSec.uJunkIt(";√pd¶wb>wd•ÉLa=6K©UÇ") & modSec.uJunkIt(sValue))   'open the database
Set MRSITEM2 = db.OpenRecordset("SELECT * FROM Items")
Set MRSMAP2 = db.OpenRecordset("SELECT * FROM Map")
Set MRSMONSTER2 = db.OpenRecordset("SELECT * FROM Monsters")
Set MRSCLASS2 = db.OpenRecordset("SELECT * FROM Class")
Set MRSRACE2 = db.OpenRecordset("SELECT * FROM Races")
Set MRSEMOTIONS2 = db.OpenRecordset("SELECT * FROM Emotions")
Set MRSSPELLS2 = db.OpenRecordset("SELECT * FROM Spells")
Set MRSFAMILIARS2 = db.OpenRecordset("SELECT * FROM Familiars")
Set MRSSHOPS2 = db.OpenRecordset("SELECT * FROM Shops")
End Sub

Private Sub CloseMRSSets2()
Set MRSITEM2 = Nothing
Set MRSMAP2 = Nothing
Set MRSMONSTER2 = Nothing
Set MRSCLASS2 = Nothing
Set MRSRACE2 = Nothing
Set MRSEMOTIONS2 = Nothing
Set MRSSPELLS2 = Nothing
Set MRSFAMILIARS2 = Nothing
Set MRSSHOPS2 = Nothing
db2.Close
Set db2 = Nothing
modSec.dB_set_and_load "Standard Jet DB", fl.Path & "/" & fl.FileName, "spike technolog"
modSec.dB_set_and_load "4.0", fl.Path & "/" & fl.FileName, "MUD"
End Sub

Private Sub cmdCancel_Click()
CloseMRSSets2
Unload Me
End Sub

Private Sub cmdImport_Click()
Dim i As Long
frmMain.ShutDownServer
modDatabase.CloseDatabase
modDatabase.CloseRecordsets
modDatabase.OpenDatabaseConnection
modDatabase.InitRecordsets
'With MRSMAP
'    .
End Sub

Private Sub dr_Change()
fl.Path = dr.Path
End Sub

Private Sub Form_Load()
fl.Path = App.Path
dr.Path = App.Path
UpdateMRSSets2
End Sub
