VERSION 5.00
Begin VB.Form frmClickAndAdd 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin VB.PictureBox picBuffer 
      Height          =   1185
      Left            =   5205
      ScaleHeight     =   1125
      ScaleWidth      =   1125
      TabIndex        =   7
      Top             =   450
      Width           =   1185
   End
   Begin VB.PictureBox picNSEW 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   1320
      Picture         =   "frmClickAndAdd.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   1215
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox picNWSWNESE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   2445
      Picture         =   "frmClickAndAdd.frx":3042
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   1215
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox picUD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3570
      Picture         =   "frmClickAndAdd.frx":6084
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   1215
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox picKUD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3540
      Picture         =   "frmClickAndAdd.frx":90C6
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox picKNWSWNESE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   2415
      Picture         =   "frmClickAndAdd.frx":C108
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   75
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox picKNSEW 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   1230
      Picture         =   "frmClickAndAdd.frx":F14A
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox picIn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   105
      Picture         =   "frmClickAndAdd.frx":1218C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "frmClickAndAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DX As New DirectX7
Dim DDRAW As DirectDraw7
Dim Primary As DirectDrawSurface7
Dim SurfDesc As DDSURFACEDESC2
Dim dds1 As DDSURFACEDESC2
Dim picBMP As DirectDrawSurface7
Dim Clipper As DirectDrawClipper


Private Type Tiletype
    sTitle As String
    lID As Long
    R As RECT
End Type
Dim m() As Tiletype

Private Sub init()
Set DDRAW = DX.DirectDrawCreate("")

DDRAW.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL

SurfDesc.lFlags = DDSD_CAPS
SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE

Set Primary = DDRAW.CreateSurface(SurfDesc)

dds1.lFlags = DDSD_CAPS
dds1.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN

Set picBMP = DDRAW.CreateSurfaceFromFile(App.Path & "\nsew.bmp", dds1)

Set Clipper = DDRAW.CreateClipper(0)

Clipper.SetHWnd Me.hWnd
Dim key As DDCOLORKEY
key.high = 255
key.low = 255
Primary.SetClipper Clipper
picBMP.SetColorKey DDCKEY_SRCBLT, key
End Sub

Private Sub Command1_Click()
Me.DrawDoorNonLock "n"
End Sub

Private Sub Form_Load()
init
Dim i As Long
Dim j As Long
ReDim m(20)
For i = 0 To Me.Width \ 80 Step 80
    For j = 0 To Me.Height \ 80 Step 80
        'With m
    Next
Next
End Sub

Public Sub DrawDoorNonLock(Dir As String)
Dim lSprite As Long
Dim lMask As Long
Dim R As RECT
Dim r1 As RECT
'lSprite = CreateCompatibleDC(GetDC(0))
'lMask = CreateCompatibleDC(GetDC(0))
Select Case LCase$(Left$(Dir, 1))
    Case "n"
        'SelectObject lSprite, CreateCompatibleBitmap(picNSEW.hDC, 32, 10)
        'SelectObject lMask, CreateCompatibleBitmap(picNSEW.hDC, 32, 10)
'        With r
'            .Top = 0
'            .Left = 0
'            .Bottom = 10
'            .Right = 32
'        End With
        DX.GetWindowRect Me.hWnd, R
        'r1.Right = dds1.lWidth
        'r1.Bottom = dds1.lHeight
        With r1
            .Top = 0
            .Left = 64 \ 4
            .Right = 64 - (64 \ 4)
            .Bottom = 10
        End With
        With R
            '.Top = 0
            '.Left = 64 \ 4
            .Right = .Top + (r1.Right - r1.Left)
            .Bottom = 10 + (r1.Bottom - r1.Top)
        End With
        Primary.DrawBox R.Left, R.Top, R.Right, R.Bottom
End Select
'DeleteDC lSprite
'DeleteDC lMask
End Sub

Public Sub DrawMap(StartRoom As Long)

End Sub

Private Sub picBuffer_Click()
DrawDoorNonLock "n"
End Sub
