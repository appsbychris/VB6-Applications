VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParseForANSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub ParseAndDrawLine(frm As Form, X As Long, Y As Long, sText As String, ByRef EndX As Long, ByRef EndY As Long)
Dim iEnters As Integer
Dim tArr() As String
Dim tArr2() As String
Dim i As Long
Dim a As Long
Dim b As Long
Dim ToPrint As String
Dim sESC As String
Dim sStore As String
Dim lFinalWidth As Long
Dim lFinalHeight As Long
Dim bMult As Boolean
Dim uDims As udtSIZE
Dim m As Long
Dim n As Long
Dim s As String
Dim a As String
m = InStr(1, sText, "[")
If m <> 0 Then
    n = InStr(m + 2, sText, "[")
    If n = 0 Then
        n = InStr(m + 2, sText, " ")
        s = Mid$(sText, m, n - m)
        tArr(i) = cReplace.Replace11(sText, s, "", 1, 0)
        a = Me.DeterANSICommand(s)
        Select Case a
            Case "color"
                Set_Color frm, s
            Case "move"
                
            Case "dmove"
                n = InStr(1, s, "[")
                m = InStr(n, s, ";")
                a = CLng(Mid$(s, n + 1, m - n - 1))
                n = InStr(m, s, "H")
                b = CLng(Mid$(s, m + 1, n - m - 1))
                With cpCursor
                    .lx = a
                    .ly = b
                End With
            Case "clear"
                frm.Cls
                sText = ""
        End Select
    End If
    
End If
    
'ToPrint = tArr(iEnters)
'ToPrint = ReplaceANSIColors(ToPrint)
'GetTextExtentPoint32 frm.hdc, ToPrint, Len(ToPrint), uDims
'lFinalWidth = uDims.lWidth
'lFinalHeight = iEnters

End Sub

Private Function DeterANSICommand(s As String) As String
Select Case Right$(s, 1)
    Case "m"
        DeterANSICommand = "color"
    Case "D"
        DeterANSICommand = "move"
    Case "J"
        DeterANSICommand = "clear"
    Case "H"
        DeterANSICommand = "dmove"
End Select
End Function

Private Function Set_Color(frm As Form, ESC_Sequence As String) As Integer
Set_Color = 0
With frm
    Select Case ESC_Sequence
        Case BLACK
            SetTextColor .hdc, vbBlack
        Case RED
            SetTextColor .hdc, &H80&
        Case bRED
            SetTextColor .hdc, &HFF&
        Case GREEN
            SetTextColor .hdc, &H8000&
        Case bGREEN
            SetTextColor .hdc, &HFF00&
        Case YELLOW
            SetTextColor .hdc, &HC0C0&
        Case bYELLOW
            SetTextColor .hdc, &HFFFF&
        Case BLUE
            SetTextColor .hdc, &H800000
        Case bBLUE
            SetTextColor .hdc, &HFF0000
        Case MAGNETA
            SetTextColor .hdc, &H800080
        Case bMAGNETA
            SetTextColor .hdc, &HFF00FF
        Case LIGHTBLUE
            SetTextColor .hdc, &H808000
        Case bLIGHTBLUE
            SetTextColor .hdc, &HFFFF00
        Case WHITE
            SetTextColor .hdc, &HC0C0C0
        Case bWHITE
            SetTextColor .hdc, &HFFFFFF
        Case BGRED
            SetTextColor .hdc, &H80&
            Set_Color = 1
        Case BGGREEN
            SetTextColor .hdc, &H8000&
            Set_Color = 1
        Case BGYELLOW
            SetTextColor .hdc, &HC0C0&
            Set_Color = 1
        Case BGBLUE
            SetTextColor .hdc, &H800000
            Set_Color = 1
        Case BGPURPLE
            SetTextColor .hdc, &H400040
            Set_Color = 1
        Case BGLIGHTBLUE
            SetTextColor .hdc, &HFFFF00
            Set_Color = 1
        Case Else
            SetTextColor .hdc, &HC0C0C0
    End Select
End With
End Function

Private Sub DrawBackLine(frm As Form, s As String)
Dim orgX As Long
Dim orgY As Long
Dim ly2 As Long
Dim lx2 As Long
Dim uRECT As RECT
With uRECT
    .lTop = cpCursor.ly
    .lLeft = cpCursor.lx
    .lBottom = cpCursor.ly + cpCursor.lHeight
    .lRight = cpCursor.lx + (cpCursor.lWidth * Len(s))
End With
With cpCursor
    orgY = .ly
    orgX = .lx
End With
With frm
    lSolidBrush = CreateSolidBrush(GetTextColor(.hdc))
    FillRect .hdc, uRECT, lSolidBrush
    SetTextColor .hdc, vbWhite
    TextOut .hdc, orgX, orgY, UseText, Len(s)
    DeleteObject lSolidBrush
End With
With cpCursor
    .ly = orgY
    .lx = uRECT.lRight
End With
End Sub
