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
'***************                TicBoard                        **********************
'***************                TicBoard.vbp                    **********************
'*************************************************************************************
'*************************************************************************************

Public YName$ 'players name
Public CardsInHand(15) As Integer 'Your current hand
'*************************************************************************************
'For the timer
Public Declare Function GetTickCount Lib "kernel32" () As Long
'*************************************************************************************

'*************************************************************************************
'move the form
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'*************************************************************************************

'*************************************************************************************
'Reshape the form/Trim off listbox borders
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lbRect As RECT) As Long
Public Type RECT
    lLeft As Long
    lTop As Long
    lRight As Long
    lBottom As Long
End Type
'*************************************************************************************

'*************************************************************************************
'get pictures for backgrounds
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
'*************************************************************************************

'*************************************************************************************
'Put window ontop
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
'*************************************************************************************

'*************************************************************************************
'Subclassing stuff
Public Const GWL_WNDPROC = (-4)
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal cColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_COMMAND = &H111
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_VSCROLL = &H115
Private Const RDW_INVALIDATE = &H1
Private Const EM_GETRECT = &HB2
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public oldLbx1Proc As Long 'listbox1
Public oldLbx2Proc As Long 'listbox2
Public oldSortListProc As Long 'lstsort on form7
Public oldWindowProc As Long 'windows
Public oldWindowProc2 As Long
Public oldWindowProc3 As Long
Public oldWindowSortProc As Long
'*************************************************************************************
'brushes
Public gBGBrush As Long
Public gBGBrush2 As Long
Public txtBoxBrush1 As Long
Public txtBoxBrush2 As Long
Public txtConBox1 As Long
Public txtConBox2 As Long
Public SortBrush As Long
'*************************************************************************************
'*************************************************************************************

Public Function NewWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo NewWindowProc_Error
'if we get the message to paint the listbox
If uMsg = WM_CTLCOLORLISTBOX And gBGBrush <> 0 And gBGBrush2 <> 0 Then
    SetBkMode wParam, 1 'make text print transparently
    If lParam = Form1.lstCurrent.hwnd Then 'if lstcurrent
        NewWindowProc = gBGBrush 'make a new background
    ElseIf lParam = Form1.lstDiscard.hwnd Then 'if lstdiscard
        NewWindowProc = gBGBrush2 'make new background
    End If
Else
    'if not, continue with that process
    NewWindowProc = CallWindowProc(oldWindowProc, hwnd, uMsg, wParam, lParam)
End If
On Error GoTo 0
Exit Function
NewWindowProc_Error:
End Function

Public Function NewWindowSortProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo NewWindowSortProc_Error
'if we get the message to paint the listbox
If uMsg = WM_CTLCOLORLISTBOX And SortBrush <> 0 Then
    SetBkMode wParam, 1 'make text print transparently
    SetTextColor wParam, vbGreen 'make text green
    NewWindowSortProc = SortBrush 'make a new background
Else
    'if not, continue with that process
    NewWindowSortProc = CallWindowProc(oldWindowSortProc, hwnd, uMsg, wParam, lParam)
End If
On Error GoTo 0
Exit Function
NewWindowSortProc_Error:
End Function

Public Function NewWindowProc2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo NewWindowProc2_Error
Dim aRect As RECT 'rect
If (uMsg = WM_CTLCOLORLISTBOX Or uMsg = WM_CTLCOLOREDIT) Then 'if the message
                'if any of these
    SetBkMode wParam, 1 'Make the words print transparently
    If lParam = Form1.txtChat.hwnd Then 'if its txtchat
        SetTextColor wParam, vbGreen 'make text green
        NewWindowProc2 = txtBoxBrush1 'paint out background
    ElseIf lParam = Form1.txtTalk.hwnd Then 'if its txttalk then
        SetTextColor wParam, vbGreen 'make the text green
        NewWindowProc2 = txtBoxBrush2 'use our background
    End If
ElseIf uMsg = WM_COMMAND And GetProp(lParam, "DoRedraw") = -1 Then 'if we get
            'this message, and our custom property is -1 then
    SendMessage lParam, EM_GETRECT, 0, aRect
    InvalidateRect lParam, aRect, 1
    UpdateWindow lParam
Else
    'continue with original process
    NewWindowProc2 = CallWindowProc(oldWindowProc2, hwnd, uMsg, wParam, lParam)
End If
On Error GoTo 0
Exit Function
NewWindowProc2_Error:
End Function

Public Function NewWindowProc3(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo NewWindowProc3_Error
If uMsg = &H133 Or uMsg = &H111 And txtBoxBrush1 <> 0 And txtBoxBrush2 <> 0 Then 'if we
            'get the paint textbox message
    SetBkMode wParam, 1 'make words print transparently
    If lParam = Form2.Text1.hwnd Then 'if text1
        NewWindowProc3 = txtConBox1 'skin it
    ElseIf lParam = Form2.Text2.hwnd Then 'if text2
        NewWindowProc3 = txtConBox2 'skin it
    End If
Else
    'return to normal process
    NewWindowProc3 = CallWindowProc(oldWindowProc3, hwnd, uMsg, wParam, lParam)
End If
On Error GoTo 0
Exit Function
NewWindowProc3_Error:
End Function

Public Function NewLbxProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Dim aRect As RECT
If uMsg = WM_VSCROLL Or uMsg = &H111 Then 'if we get the scroll message
    InvalidateRect lParam, aRect, 1 'if we get the scroll message
    NewLbxProc = CallWindowProc(oldLbx1Proc, hwnd, uMsg, wParam, lParam) 'don't paint the bg
ElseIf uMsg = &H14 Then
    NewLbxProc = 0 'clear it
Else
    'continue with old process
    NewLbxProc = CallWindowProc(oldLbx1Proc, hwnd, uMsg, wParam, lParam)
End If
End Function

Public Function NewLbxProc2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Dim aRect As RECT
If uMsg = WM_VSCROLL Or uMsg = &H111 Then 'if we get the scroll message
    InvalidateRect lParam, aRect, 1 'if we get the scroll message
    NewLbxProc2 = CallWindowProc(oldLbx2Proc, hwnd, uMsg, wParam, lParam)
ElseIf uMsg = &H14 Then
    NewLbxProc2 = 0 'clear it
Else
    NewLbxProc2 = CallWindowProc(oldLbx2Proc, hwnd, uMsg, wParam, lParam) 'continue with old process
End If
End Function

Public Function NewSortListProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Dim aRect As RECT
If uMsg = WM_VSCROLL Or uMsg = &H111 Then 'if we get the scroll message
    InvalidateRect lParam, aRect, 1 'if we get the scroll message
    NewSortListProc = CallWindowProc(oldSortListProc, hwnd, uMsg, wParam, lParam)
ElseIf uMsg = &H14 Then
    NewSortListProc = 0 'clear it
Else
    NewSortListProc = CallWindowProc(oldSortListProc, hwnd, uMsg, wParam, lParam) 'continue with old process
End If
End Function

