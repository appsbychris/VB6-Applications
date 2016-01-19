VERSION 5.00
Begin VB.Form frmGraphics 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon Graphics"
   ClientHeight    =   9450
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGraphics.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   630
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   718
   StartUpPosition =   3  'Windows Default
   Begin DoDMudServer.ucColors clrsMain 
      Height          =   1080
      Left            =   720
      TabIndex        =   2
      Top             =   2760
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   1905
   End
   Begin DoDMudServer.ucANSI ansOpts 
      Height          =   2730
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   4815
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   5415
      Left            =   120
      MouseIcon       =   "frmGraphics.frx":08CA
      MousePointer    =   99  'Custom
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   0
      Top             =   3960
      Width           =   10575
   End
   Begin DoDMudServer.eButton cmdOK 
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Style           =   2
      Cap             =   "&Done"
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
         Shortcut        =   ^Y
      End
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : frmGraphics
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type CharSlot
    lx As Long
    lY As Long
    c As String
    co As Long
    b As Boolean
End Type
Private lW As Long
Private lH As Long
Private Const DT_LEFT As Long = &H0
Private charMAP(1399) As CharSlot
Private OldChar As Long
Private b As Boolean
Private prevCHAR As Long

Private Sub init()
Dim i As Long
Dim k As Long
lW = picMain.TextWidth("X")
lH = picMain.TextHeight("W")
'lH = lH - 3
'lW = lW + 3
'If lH * 20 < picMain.ScaleHeight Then
picMain.Height = lH * 20
'If lW * 70 < picMain.ScaleWidth Then
picMain.Width = lW * 70
With charMAP(0)
    .lx = 0
    .lY = 0
End With
For i = LBound(charMAP) + 1 To UBound(charMAP)
    With charMAP(i)
        .lx = charMAP(i - 1).lx + lW
        .lY = charMAP(i - 1).lY
        If .lx + lW > picMain.ScaleWidth Then
            k = k + 1
            .lx = 0
            .lY = lH * k
        End If
    End With
Next
prevCHAR = -1
End Sub

Private Sub ansOpts_GotFocus()
picMain.SetFocus
End Sub

Private Sub clrsMain_GotFocus()
picMain.SetFocus
End Sub

Private Sub cmdOK_Click()
Dim i As Long
Dim s As String
Dim l As String
Dim j As String
For i = LBound(charMAP) To UBound(charMAP)
    With charMAP(i)
        If .c <> "" And .c <> Chr$(0) Then
            If .lx = 0 And i <> 0 Then s = s & vbCrLf
            j = GetColorCode(.co, .b)
            If j = l Then
                s = s & .c
            Else
                s = s & j & .c
            End If
            l = j
        Else
            If .lx = 0 And i <> 0 Then s = s & vbCrLf
            s = s & rWHITE & " "
        End If
    End With
Next
Open App.Path & "\intgrp.ansi" For Output As #1
    Print #1, s
Close #1
Unload Me
End Sub

Private Sub LoadPrev()
Dim i As Long
Dim j As Long
Dim s As String
Dim k As String
Dim m As Long
Dim f As Long
Dim R As RECT
Dim ff As Long
Dim pb As Boolean
Dim Arr() As String
Open App.Path & "\intgrp.ansi" For Binary As #1
    s = Input$(LOF(1), 1)
Close #1
f = Get_Color("")
SplitFast s, Arr, vbCrLf
For i = LBound(Arr) To UBound(Arr)
    If Arr(i) <> "" Then
        k = Arr(i)
        Do Until LenB(k) = 0
            With charMAP(j)
                If Left$(k, 9) Like "[[]#m[[]##m" Then
                    m = Set_Color(Left$(k, 9))
                    If m = 1 Then .b = True Else .b = False
                    .co = Get_Color(Left$(k, 9))
                    f = .co
                    k = Mid$(k, 10)
                    .c = Left$(k, 1)
                    k = Mid$(k, 2)
                Else
                    .co = f
                    .b = pb
                    .c = Left$(k, 1)
                    k = Mid$(k, 2)
                End If
                pb = .b
                R.Top = .lY
                R.Left = .lx
                R.Bottom = .lY + lH
                R.Right = .lx + lW
                If .c <> "" Then
                    If .b Then ff = CreateSolidBrush(.co) Else ff = CreateSolidBrush(vbBlack)
                    FillRect picMain.hdc, R, ff
                    DeleteObject ff
                    If .b Then
                        SetTextColor picMain.hdc, &HC0C0C0
                        DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
                    Else
                        DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
                    End If
                End If
                j = j + 1
                DoEvents
            End With
        Loop
    End If
    DoEvents
Next
EraseCursor
DrawCursor
picMain.Refresh
End Sub

Private Function Get_Color(ESC_Sequence As String) As Long
Select Case ESC_Sequence
    Case rRED
        Get_Color = &H80&
    Case rbRED
        Get_Color = &HFF&
    Case rGREEN
        Get_Color = &H8000&
    Case rbGREEN
        Get_Color = &HFF00&
    Case rYELLOW
        Get_Color = &HC0C0&
    Case rbYELLOW
        Get_Color = &HFFFF&
    Case rBLUE
        Get_Color = &H800000
    Case rbBLUE
        Get_Color = &HFF0000
    Case rMAGNETA
        Get_Color = &H800080
    Case rbMAGNETA
        Get_Color = &HFF00FF
    Case rLIGHTBLUE
        Get_Color = &H808000
    Case rbLIGHTBLUE
        Get_Color = &HFFFF00
    Case rWHITE
        Get_Color = &HC0C0C0
    Case rbWHITE
        Get_Color = &HFFFFFF
    Case rBGRED
        Get_Color = &H80&
    Case rBGGREEN
        Get_Color = &H8000&
    Case rBGYELLOW
        Get_Color = &HC0C0&
    Case rBGBLUE
        Get_Color = &H800000
    Case rBGPURPLE
        Get_Color = &H400040
    Case rBGLIGHTBLUE
        Get_Color = &HFFFF00
    Case Else
        Get_Color = &HC0C0C0
End Select
End Function

Private Function Set_Color(ESC_Sequence As String) As Integer
Set_Color = 0
With picMain
    Select Case ESC_Sequence
        Case rRED
            SetTextColor .hdc, &H80&
        Case rbRED
            SetTextColor .hdc, &HFF&
        Case rGREEN
            SetTextColor .hdc, &H8000&
        Case rbGREEN
            SetTextColor .hdc, &HFF00&
        Case rYELLOW
            SetTextColor .hdc, &HC0C0&
        Case rbYELLOW
            SetTextColor .hdc, &HFFFF&
        Case rBLUE
            SetTextColor .hdc, &H800000
        Case rbBLUE
            SetTextColor .hdc, &HFF0000
        Case rMAGNETA
            SetTextColor .hdc, &H800080
        Case rbMAGNETA
            SetTextColor .hdc, &HFF00FF
        Case rLIGHTBLUE
            SetTextColor .hdc, &H808000
        Case rbLIGHTBLUE
            SetTextColor .hdc, &HFFFF00
        Case rWHITE
            SetTextColor .hdc, &HC0C0C0
        Case rbWHITE
            SetTextColor .hdc, &HFFFFFF
        Case rBGRED
            SetTextColor .hdc, &H80&
            Set_Color = 1
        Case rBGGREEN
            SetTextColor .hdc, &H8000&
            Set_Color = 1
        Case rBGYELLOW
            SetTextColor .hdc, &HC0C0&
            Set_Color = 1
        Case rBGBLUE
            SetTextColor .hdc, &H800000
            Set_Color = 1
        Case rBGPURPLE
            SetTextColor .hdc, &H400040
            Set_Color = 1
        Case rBGLIGHTBLUE
            SetTextColor .hdc, &HFFFF00
            Set_Color = 1
        Case Else
            SetTextColor .hdc, &HC0C0C0
    End Select
End With
End Function

Private Function GetColorCode(l As Long, b As Boolean) As String
Select Case l
    Case &H80& 'red
        If Not b Then GetColorCode = rRED Else GetColorCode = rBGRED
    Case &H8000& 'green
        If Not b Then GetColorCode = rGREEN Else GetColorCode = rBGGREEN
    Case &HC0C0& 'yellow
        If Not b Then GetColorCode = rYELLOW Else GetColorCode = rBGYELLOW
    Case &H800000    'blue
        If Not b Then GetColorCode = rBLUE Else GetColorCode = rBGBLUE
    Case &H800080    'magneta
        If Not b Then GetColorCode = rMAGNETA Else GetColorCode = rBGPURPLE
    Case &H808000    'Lightblue
        If Not b Then GetColorCode = rLIGHTBLUE Else GetColorCode = rBGLIGHTBLUE
    Case &HC0C0C0    'White
        GetColorCode = rWHITE
    Case &HFF&       'bred
        GetColorCode = rbRED
    Case &HFF00&     'bgreen
        GetColorCode = rbGREEN
    Case &HFFFF&     'byellow
        GetColorCode = rbYELLOW
    Case &HFF0000    'bBlue
        GetColorCode = rbBLUE
    Case &HFF00FF    'bMagneta
        GetColorCode = rbMAGNETA
    Case &HFFFF00    'blightblue
        GetColorCode = rbLIGHTBLUE
    Case &HFFFFFF    'bwhite
        GetColorCode = rbWHITE
End Select
End Function


Private Sub Form_Load()
init
LoadPrev
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'ShowCursor 1
End Sub

Private Sub mnuClear_Click()
Dim i As Long
For i = LBound(charMAP) To UBound(charMAP)
    With charMAP(i)
        .b = False
        .c = ""
        .co = 0
    End With
Next
prevCHAR = -1
picMain.Cls
DrawCursor
picMain.Refresh
End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
Dim R As RECT
Dim l As Long
Select Case KeyCode
    Case vbKeyLeft
        EraseCursor
        prevCHAR = prevCHAR - 1
        DrawCursor
        picMain.Refresh
    Case vbKeyRight
        EraseCursor
        prevCHAR = prevCHAR + 1
        DrawCursor
        picMain.Refresh
    Case vbKeyDown
        EraseCursor
        If prevCHAR + 70 <= UBound(charMAP) Then prevCHAR = prevCHAR + 70
        DrawCursor
        picMain.Refresh
    Case vbKeyUp
        EraseCursor
        If prevCHAR - 70 <= UBound(charMAP) Then prevCHAR = prevCHAR - 70
        DrawCursor
        picMain.Refresh
    Case vbKeyDelete
        EraseCursor
        With charMAP(prevCHAR + 1)
            .c = ""
            .co = 0
            .b = False
            R.Top = .lY
            R.Bottom = .lY + lH
            R.Left = .lx
            R.Right = .lx + lW
            l = CreateSolidBrush(vbBlack)
            FillRect picMain.hdc, R, l
            DeleteObject l
        End With
        'prevCHAR = prevCHAR - 1
        DrawCursor
        picMain.Refresh
End Select
End Sub

Private Sub picMain_KeyPress(KeyAscii As Integer)
Dim R As RECT
Dim j As Long
Select Case KeyAscii
    Case 13, 8
        If KeyAscii = 8 Then
            EraseCursor
            With charMAP(prevCHAR)
                .c = ""
                .co = 0
                .b = False
                R.Top = .lY
                R.Bottom = .lY + lH
                R.Left = .lx
                R.Right = .lx + lW
                l = CreateSolidBrush(vbBlack)
                FillRect picMain.hdc, R, l
                DeleteObject l
            End With
            prevCHAR = prevCHAR - 1
            DrawCursor
            picMain.Refresh
        Else
            EraseCursor
            With charMAP(prevCHAR + 1)
                .c = ansOpts.StringChar
                .b = clrsMain.IsBC
                .co = clrsMain.lColor
                R.Top = .lY
                R.Left = .lx
                R.Bottom = .lY + lH
                R.Right = .lx + lW
                If .c <> "" Then
                    If .b Then j = CreateSolidBrush(.co) Else j = CreateSolidBrush(vbBlack)
                    FillRect picMain.hdc, R, j
                    DeleteObject j
                    If .b Then
                        SetTextColor picMain.hdc, &HC0C0C0
                        DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
                    Else
                        SetTextColor picMain.hdc, .co
                        DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
                    End If
                End If
            End With
            prevCHAR = prevCHAR + 1
            DrawCursor
            picMain.Refresh
        End If
    Case Else
        If prevCHAR + 1 <= UBound(charMAP) Then
            EraseCursor
            With charMAP(prevCHAR + 1)
                .c = Chr$(KeyAscii)
                .b = clrsMain.IsBC
                .co = clrsMain.lColor
                R.Top = .lY
                R.Left = .lx
                R.Bottom = .lY + lH
                R.Right = .lx + lW
                If .c <> "" And .c <> Chr$(0) Then
                    If .b Then j = CreateSolidBrush(.co) Else j = CreateSolidBrush(vbBlack)
                    FillRect picMain.hdc, R, j
                    DeleteObject j
                    If .b Then
                        SetTextColor picMain.hdc, &HC0C0C0
                        DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
                    Else
                        SetTextColor picMain.hdc, .co
                        DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
                    End If
                End If
            End With
            prevCHAR = prevCHAR + 1
            DrawCursor
            picMain.Refresh
        End If
End Select
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'20 height
'70 width
Dim R As RECT
Dim i As Long
Dim j As Long
Dim xx As Long
For i = LBound(charMAP) To UBound(charMAP)
    With charMAP(i)
        If X >= .lx And X <= .lx + lW Then
            If Y >= .lY And Y <= .lY + lH Then
                'DING
                R.Top = .lY
                R.Left = .lx
                R.Bottom = .lY + lH
                R.Right = .lx + lW
                .c = ansOpts.StringChar
                .co = clrsMain.lColor
                .b = clrsMain.IsBC
                xx = i
                Exit For
            End If
        End If
    End With
Next
With charMAP(xx)
    R.Top = .lY
    R.Left = .lx
    R.Bottom = .lY + lH
    R.Right = .lx + lW
    If .c <> "" And .c <> Chr$(0) Then
        If .b Then j = CreateSolidBrush(.co) Else j = CreateSolidBrush(vbBlack)
        FillRect picMain.hdc, R, j
        DeleteObject j
        If .b Then
            SetTextColor picMain.hdc, &HC0C0C0
            DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
        Else
            SetTextColor picMain.hdc, .co
            DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
        End If
    Else
        j = CreateSolidBrush(vbBlack)
        FillRect picMain.hdc, R, j
        DeleteObject j
    End If
End With
EraseCursor
prevCHAR = xx
DrawCursor
'i = CreateSolidBrush(vbBlack)
'FillRect picMain.hdc, R, i
'DeleteObject i
'DrawText picMain.hdc, ansOpts.StringChar, 1, R, DT_LEFT
picMain.Refresh
End Sub

Private Sub EraseCursor()
Dim R As RECT
Dim l As Long
If prevCHAR + 1 <= UBound(charMAP) Then
With charMAP(prevCHAR + 1)
    R.Top = .lY
    R.Bottom = .lY + lH
    R.Left = .lx
    R.Right = .lx + lW
    l = CreateSolidBrush(vbBlack)
    FillRect picMain.hdc, R, l
    DeleteObject l
    If .b Then j = CreateSolidBrush(.co) Else j = CreateSolidBrush(vbBlack)
    FillRect picMain.hdc, R, j
    DeleteObject j
    If .b Then
        SetTextColor picMain.hdc, &HC0C0C0
        DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
    Else
        SetTextColor picMain.hdc, .co
        DrawText picMain.hdc, .c, Len(.c), R, DT_LEFT
    End If
End With
End If
End Sub

Private Sub DrawCursor()
Dim R As RECT
Dim l As Long
If prevCHAR < 1 Then prevCHAR = 1
If prevCHAR + 1 <= UBound(charMAP) Then
    With charMAP(prevCHAR + 1)
        R.Top = .lY + lH - 3
        R.Bottom = .lY + lH
        R.Left = .lx
        R.Right = .lx + lW
        l = CreateSolidBrush(vbWhite)
        FillRect picMain.hdc, R, l
        DeleteObject l
    End With
End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'With picMain
'    If GetCapture <> .hwnd Then SetCapture .hwnd
'    If X < 0 Or Y < 0 Or X > .ScaleWidth Or Y > .ScaleHeight Then
'        If b Then
'            ReleaseCapture
'            ShowCursor 1
'            b = False
'        End If
'        Exit Sub
'    ElseIf Not b Then
'        ShowCursor 0
'        b = True
'    End If
'End With
Dim R As RECT
Dim i As Long
Dim j As Long
For i = LBound(charMAP) To UBound(charMAP)
    With charMAP(i)
        If X >= .lx And X <= .lx + lW Then
            If Y >= .lY And Y <= .lY + lH Then
                'DING
                If i <> OldChar Then
                    With charMAP(OldChar)
                        R.Top = .lY
                        R.Left = .lx
                        R.Bottom = .lY + lH
                        R.Right = .lx + lW
                        If .c <> "" And .c <> Chr$(0) Then
                            
                            If .b Then j = CreateSolidBrush(.co) Else j = CreateSolidBrush(vbBlack)
                            FillRect picMain.hdc, R, j
                            DeleteObject j
                            If .b Then
                                SetTextColor picMain.hdc, &HC0C0C0
                                DrawText picMain.hdc, charMAP(OldChar).c, Len(charMAP(OldChar).c), R, DT_LEFT
                            Else
                                SetTextColor picMain.hdc, .co
                                DrawText picMain.hdc, charMAP(OldChar).c, Len(charMAP(OldChar).c), R, DT_LEFT
                            End If
                        Else
                            j = CreateSolidBrush(vbBlack)
                            FillRect picMain.hdc, R, j
                            DeleteObject j
                        End If
                    End With
                End If
                R.Top = .lY
                R.Left = .lx
                R.Bottom = .lY + lH
                R.Right = .lx + lW
                OldChar = i
                Exit For
            End If
        End If
    End With
Next

i = CreateSolidBrush(vbGreen)
FillRect picMain.hdc, R, i
DeleteObject i
If charMAP(OldChar).c = "" And charMAP(OldChar).c <> Chr$(0) Then
    SetTextColor picMain.hdc, vbBlack
    DrawText picMain.hdc, ansOpts.StringChar, 1, R, DT_LEFT
Else
    SetTextColor picMain.hdc, &HC0C0C0
    DrawText picMain.hdc, charMAP(OldChar).c, Len(charMAP(OldChar).c), R, DT_LEFT
End If
DrawCursor
picMain.Refresh
End Sub
