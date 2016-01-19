Attribute VB_Name = "Module1"
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
Public Declare Function GetTickCount Lib "kernel32" () As Long 'api for milisecond precision pauses
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lbRect As RECT) As Long 'trim off borders API
Public Type RECT 'rect type
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
End Type
'border triming
Public Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'get the pictures behind the listbox using bitblt
Public Declare Function BitBlt Lib "gdi32" ( _
           ByVal hDestDC As Long, _
           ByVal X As Long, _
           ByVal Y As Long, _
           ByVal nWidth As Long, _
           ByVal nHeight As Long, _
           ByVal hSrcDC As Long, _
           ByVal xSrc As Long, _
           ByVal ySrc As Long, _
           ByVal dwRop As Long _
) As Long
'///////Subclassing things/////////
Public Const GWL_WNDPROC = (-4)
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal cColor As Long) As Long
Public Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_VSCROLL = &H115
Private Declare Function InvalidateRect Lib "USER32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Public oldLbx1Proc As Long
Public oldWindowProc As Long
Public gBGBrush As Long
'//////////////////////////////

Public Function NewWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If uMsg = WM_CTLCOLORLISTBOX And gBGBrush <> 0 Then 'if it needs to paint
    SetBkMode wParam, 1 'Make the words print transparently
    SetTextColor wParam, &H8000&
    NewWindowProc = gBGBrush 'use our brush instead
Else
    'continue normaly
    NewWindowProc = CallWindowProc(oldWindowProc, hWnd, uMsg, wParam, lParam)
End If
End Function

Public Function NewLbxProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Force the control to repaint itself every time the scroll message is received.
If uMsg = WM_VSCROLL Then
    InvalidateRect hWnd, 0, 0
    NewLbxProc = CallWindowProc(oldLbx1Proc, hWnd, uMsg, wParam, lParam)
ElseIf uMsg = &H14 Then
    NewLbxProc = 0
Else
    NewLbxProc = CallWindowProc(oldLbx1Proc, hWnd, uMsg, wParam, lParam)
End If
End Function

Public Function WaitFor(MS As Long)
If MS = 1 Or MS = 2 Then MS = 300 'change 1 to 300 miliseconds
Dim start
start = GetTickCount
'pause the the specified amount of time
While start + MS > GetTickCount
    DoEvents
Wend
End Function


