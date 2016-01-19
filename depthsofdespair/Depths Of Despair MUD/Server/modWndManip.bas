Attribute VB_Name = "modWndManip"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modWndManip
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Rem////////////////////
Rem API's and const for removing the X on the form
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Const MF_BYPOSITION = &H400&
Rem////////////////////////////////////

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long 'For the timer
Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
        myfrm.Left / Screen.TwipsPerPixelX, _
        myfrm.Top / Screen.TwipsPerPixelY, _
        myfrm.Width / Screen.TwipsPerPixelX, _
        myfrm.Height / Screen.TwipsPerPixelY, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Sub RemoveX(Frm As Form)
Dim lSYSTEM_MENU As Long
'Get the handle to this windows system menu
lSYSTEM_MENU = GetSystemMenu(Frm.hwnd, 0)
'This will disable the close button
RemoveMenu lSYSTEM_MENU, 6, MF_BYPOSITION
'remove the seperator bar
RemoveMenu lSYSTEM_MENU, 5, MF_BYPOSITION
End Sub
