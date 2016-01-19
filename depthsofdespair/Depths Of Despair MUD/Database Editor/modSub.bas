Attribute VB_Name = "modSub"
Private Const WM_VSCROLL As Long = &H115
Private Const GWL_WNDPROC = (-4)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Type HILOWord
  loword As Integer
  hiword As Integer
End Type
Private Const WM_CREATE As Long = &H1
Private Const WM_DESTROY As Long = &H2
Private Const WM_MOUSEWHEEL As Long = &H20A


Public Sub SubclassTextbox(txtBox As TextBox)
Dim oldWindowProc As Long
SubclassCallback 0, 0, 0, 0
With txtBox
    SetProp .hWnd, "txtBoxPtr", ObjPtr(.Parent)
    If GetProp(.hWnd, "OriginalCallback") = 0 Then
        oldWindowProc = SetWindowLong(.hWnd, GWL_WNDPROC, AddressOf SubclassCallback)
        SetProp .hWnd, "OriginalCallback", oldWindowProc
    End If
End With
End Sub

Private Function SubclassCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lOriginalP As Long
Dim lCtrlPtr As Long
Dim oIntel As IntelliSense
Dim udtHiLo As HILOWord
Dim udtXY As HILOWord
If hWnd = 0 And uMsg = 0 And wParam = 0 And lParam = 0 Then Exit Function
lOriginalP = GetProp(hWnd, "OriginalCallback")
If uMsg = WM_DESTROY And lOriginalP <> 0 Then
    SetWindowLong hWnd, GWL_WNDPROC, lorginalp
    RemoveProp hWnd, "OriginalCallback"
    RemoveProp hWnd, "txtBoxPtr"
    SubclassCallback = CallWindowProc(lOriginalP, hWnd, uMsg, wParam, lParam)
ElseIf lOriginalP <> 0 Then
    SubclassCallback = CallWindowProc(lOriginalP, hWnd, uMsg, wParam, lParam)
    lCtrlPtr = GetProp(hWnd, "txtBoxPtr")
    If lCtrlPtr <> 0 Then
        If uMsg = WM_VSCROLL Or uMsg = WM_MOUSEWHEEL Then
            CopyMemory oIntel, lCtrlPtr, 4&
            oIntel.PaintNumbers
            CopyMemory oIntel, 0&, 4&
        End If
    End If
Else
    SubclassCallback = DefWindowProc(hWnd, uMsg, wParam, lParam)
End If
End Function
