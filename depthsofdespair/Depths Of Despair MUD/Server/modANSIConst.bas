Attribute VB_Name = "modANSIConst"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modANSIConst
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Rem////////////////////////////
Rem Public color ANSI constants
Rem Fake ones so the word wrap will work right
Public Const RED            As String = "®"
Public Const GREEN          As String = "à"
Public Const YELLOW         As String = "á"
Public Const BLUE           As String = "½"
Public Const MAGNETA        As String = "¾"
Public Const LIGHTBLUE      As String = "Þ"
Public Const WHITE          As String = "æ"

Public Const BGRED          As String = "Æ"
Public Const BGGREEN        As String = "Ý"
Public Const BGYELLOW       As String = "«"
Public Const BGBLUE         As String = "§"
Public Const BGPURPLE       As String = "ê" 'ê
Public Const BGLIGHTBLUE    As String = "Ü"
'àáâ

Public Const BRIGHTYELLOW     As String = "¢"
Public Const BRIGHTGREEN      As String = "£"
Public Const BRIGHTRED        As String = "ª"
Public Const BRIGHTBLUE       As String = "¬"
Public Const BRIGHTMAGNETA    As String = "ë" '¶ë
Public Const BRIGHTLIGHTBLUE  As String = "¡"
Public Const BRIGHTWHITE      As String = "±"

Rem Ones to replace the fake ones with the real ones
Public Const BLACK           As String = "[0m[30m"

Public Const rRED            As String = "[0m[31m"
Public Const rbRED           As String = "[1m[31m"

Public Const rGREEN          As String = "[0m[32m"
Public Const rbGREEN         As String = "[1m[32m"

Public Const rYELLOW         As String = "[0m[33m"
Public Const rbYELLOW        As String = "[1m[33m"

Public Const rBLUE           As String = "[0m[34m"
Public Const rbBLUE          As String = "[1m[34m"


Public Const rMAGNETA        As String = "[0m[35m"
Public Const rbMAGNETA       As String = "[1m[35m"

Public Const rLIGHTBLUE      As String = "[0m[36m"
Public Const rbLIGHTBLUE     As String = "[1m[36m"

Public Const rWHITE          As String = "[0m[37m"
Public Const rbWHITE         As String = "[1m[37m"

Public Const rBGRED          As String = "[0m[41m"
Public Const rBGGREEN        As String = "[0m[42m"
Public Const rBGYELLOW       As String = "[0m[43m"
Public Const rBGBLUE         As String = "[0m[44m"
Public Const rBGPURPLE       As String = "[0m[45m"
Public Const rBGLIGHTBLUE    As String = "[0m[46m"
Public Const rBGBRIGHTYELLOW As String = "[1m[43m"

Public Const MOVELEFTONE     As String = "[1D"
Public Const ERASETOLEFT     As String = "[0K"
Public Const MOVELEFT23      As String = "[23D"
Public Const MOVERIGHTNUM    As String = "[#D"
Public Const ANSICLS         As String = "[2J"
Public Const MOVECURSOR      As String = "[R;CH"
Public Const UP_ARROW        As String = "[A"
Public Const DOWN_ARROW      As String = "[B"
Public Const RIGHT_ARROW     As String = "[C"
Public Const LEFT_ARROW      As String = "[D"
Public Const MOVE_LEFT_X     As String = "[XD"

Public Const EraseRow        As String = "[1K"


'23
Rem/////////////////////////////

Public Function SetMoveCursor(Row&, Col&) As String
SetMoveCursor = ReplaceFast(MOVECURSOR, "R", CStr(Row&))
SetMoveCursor = ReplaceFast(SetMoveCursor, "C", CStr(Col&))
End Function
