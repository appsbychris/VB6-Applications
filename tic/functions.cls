VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'Enum for GetSystemMetrics
Public Enum SysMet
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYHTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_CXMIN = 24
    SM_CYMIN = 25
    SM_CXSIZE = 26
    SM_CYSIZE = 27
    SM_CXMINTRACK = 28
    SM_CYMINTRACK = 29
    SM_CXDOUBLECLK = 30
    SM_CYDOUBLECLK = 31
    SM_CXICONSPACING = 32
    SM_CYICONSPACING = 33
    SM_MENUDROPALIGNMENT = 34
    SM_PENWINDOWS = 35
    SM_DBCSENABLED = 36
    SM_CMOUSEBUTTONS = 37
    SM_CMETRICS = 38
    SM_CLEANBOOT = 39
    SM_CXMAXIMIZED = 40
    SM_CXMAXTRACK = 41
    SM_CXMENUCHECK = 42
    SM_CXMENUSIZE = 43
    SM_CXMINIMIZED = 44
    SM_CYMAXIMIZED = 45
    SM_CYMAXTRACK = 46
    SM_CYMENUCHECK = 47
    SM_CYMENUSIZE = 48
    SM_CYMINIMIZED = 49
    SM_CYSMCAPTION = 50
    SM_MIDEASTENABLED = 51
    SM_NETWORK = 52
    SM_SLOWMACHINE = 53
End Enum
'decalre for getsystemmetrics
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Sub BubbleSort(List() As Integer)
'//Basic bubble sort routine//
Dim iFirst As Integer
Dim iLast As Integer
Dim iStart As Integer
Dim iEnd As Integer
Dim iTemp As Integer
iFirst = LBound(List)
iLast = UBound(List)
'puts the array in order from smallest to greatest
For iStart = iFirst To iLast - 1
    For iEnd = iStart + 1 To iLast
        If List(iStart) > List(iEnd) Then
            iTemp = List(iEnd)
            List(iEnd) = List(iStart)
            List(iStart) = iTemp
        End If
    Next iEnd
Next iStart
End Sub

Private Function Check3s(tCards() As Integer, Wilds%, Round%) As Boolean
'//Routine to place wilds correctly and check it its correct//
If Wilds% = 0 Then 'if there are no wilds, then just check it
    If ThreeOfAKind(tCards()) = True Then 'checkit
        Check3s = True 'set to true
    Else 'if check fails, set this function to false
        Check3s = False
    End If
Else 'if there are wilds
    Dim ChangeTo$ 'hold
    ChangeTo$ = "" 'holds the new wild value
    If Round% <> 14 Or Round% <> 15 Then 'if not round Ace or Two
        For i = 0 To UBound(tCards) 'find the NON-Wild card in the hand
            If Val(tCards(i)) <> Round% - 1 Then '(-1 because the cards value is 1 less then the round)
                ChangeTo$ = tCards(i) 'and save that value
                Exit For
            End If
        Next
    Else
        For i = 0 To UBound(tCards)
            If Val(tCards(i)) <> Round% - 14 Then '(-14 because Ace is 0, and Two is 1, so to get a wild from the round, we
                    'have to minus 14 from the round)
                ChangeTo$ = tCards(i) 'save the value
                Exit For
            End If
        Next
    End If
    If ChangeTo$ <> "" Then 'make sure there is something to change to
        If Round% <> 14 Or Round% <> 15 Then 'if not round Ace or Two
            For i = 0 To UBound(tCards)
                If Val(tCards(i)) = Round% - 1 Then
                    tCards(i) = ChangeTo$ 'change all wilds to the correct card
                End If
            Next
        Else
            'if round ace or two
            For i = 0 To UBound(tCards)
                If Val(tCards(i)) = Round% - 14 Then
                    tCards(i) = ChangeTo$ 'change to non-wild card
                End If
            Next
        End If
        If ThreeOfAKind(tCards()) = True Then 'check it
            Check3s = True 'if true, then set this true
        Else
            Check3s = False 'if not, set this to false
        End If
    Else
        If ThreeOfAKind(tCards()) = True Then 'check it
            Check3s = True 'if true, then make this true
        Else
            Check3s = False 'if not, make this false
        End If
    End If
End If
End Function

Private Function CheckAce(Cards() As Integer, Wilds%) As Boolean
'sub to check to see if the ace is high or low
For i = 0 To UBound(Cards()) 'reduce each card
    Select Case Cards(i)
        Case 0 To 12:
            Cards(i) = Cards(i) '-0
        Case 13 To 25:
            Cards(i) = (Cards(i) - 13) '-13
        Case 26 To 38:
            Cards(i) = (Cards(i) - (13 * 2)) '-26
        Case 39 To 51:
            Cards(i) = (Cards(i) - (13 * 3)) '-39
    End Select
Next
Call BubbleSort(Cards()) 'sort the cards in numrical order
For i = 0 To UBound(Cards)
    If Cards(i) = 0 Then 'if there is an ace
        If Cards(UBound(Cards)) = 12 Then 'if theres a king
            CheckAce = True 'make this function true
            Exit Function
        ElseIf Wilds% = 1 Then 'if there is 1 wild
            For a = 0 To UBound(Cards)
                If Cards(a) = 11 Or Cards(a) = 12 Then 'if there is a queen or king
                    CheckAce = True 'make this true
                    Exit Function
                End If
            Next
        ElseIf Wilds% > 1 Then 'if there is more then 1 wild
            For a = 0 To UBound(Cards)
                If Cards(a) = 12 Or Cards(a) = 11 _
                    Or Cards(a) = 10 Or Cards(a) = 9 Then 'if ten,jack,queen,king
                    CheckAce = True 'make this true
                    Exit Function
                End If
            Next
        End If
    End If
Next
End Function

Private Function CheckStr(Cards() As Integer, Wilds%, Round%) As Boolean
'function to check straights..and place wilds appropiately
CheckStr = False 'default to false
If Round% <> 14 Then 'if its not round Ace
    If CheckAce(Cards(), Wilds%) = True Then 'if we need to make the Ace a high value
        For i = 0 To UBound(Cards)
            Select Case Cards(i)
                Case 0:
                    Cards(i) = 13 'if its spades
                    Exit For
                Case 13:
                    Cards(i) = 26 'if its diamonds
                    Exit For
                Case 26:
                    Cards(i) = 39 'if its clubs
                    Exit For
                Case 39:
                    Cards(i) = 52 'if its hearts
                    Exit For
            End Select
        Next
    End If
End If
'if there are no wilds
If Wilds% = 0 Then
    If Straight(Cards()) = True Then 'see if it works
        CheckStr = True 'if yes, make this true
    Else
        CheckStr = False 'if not make this false
    End If
Else
    If Round% <> 14 Or Round% <> 15 Then 'if not round Ace or Two
        For i = 0 To UBound(Cards)
            If Cards(i) = Round% - 1 Then Cards(i) = 999 'if wild card, change to 999
            If Cards(i) - 13 = Round% - 1 Then Cards(i) = 999
            If Cards(i) - 26 = Round% - 1 Then Cards(i) = 999
            If Cards(i) - 39 = Round% - 1 Then Cards(i) = 999
        Next
    Else
        'if round Ace or Two
        For i = 0 To UBound(Cards)
            If Cards(i) = Round% - 14 Then Cards(i) = 999 'if wild card, change to 999
            If Cards(i) - 13 = Round% - 14 Then Cards(i) = 999
            If Cards(i) - 26 = Round% - 14 Then Cards(i) = 999
            If Cards(i) - 39 = Round% - 14 Then Cards(i) = 999
        Next
    End If
    Call BubbleSort(Cards()) 'sort the cards
    On Error Resume Next 'error handling
    For i = 0 To UBound(Cards)
        If Wilds% = 1 Then 'if there is only 1 wild
            If Cards(i) + 1 <> Cards(i + 1) And Cards(i) <> 999 Then
                For a = 0 To UBound(Cards)
                    If Cards(a) = 999 Then
                        Cards(a) = Cards(i) + 1 'set it to the correct value
                        Exit For
                    End If
                Next
            End If
        Else 'if there are more then 1 wild
            If Cards(i) <> -1 Then 'if not a blank card
                If Cards(i) + 1 <> Cards(i + 1) Then 'if the cards don't go in a row
                    For a = 0 To UBound(Cards)
                        If Cards(a) = 999 Then 'find a wild,
                            Cards(a) = Cards(i) + 1 'and make it the correct value
                            BubbleSort Cards() 'sort the cards
                            i = 0 'make i 0 again
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    Next
    Call BubbleSort(Cards()) 'sort the cards
    If Straight(Cards()) = True Then 'check it
        CheckStr = True 'if its good, make this true
    Else
        CheckStr = False 'if not, make this false
    End If
End If
End Function

Public Function GetSysMetrics(GetWhat As SysMet) As Long
'function to get the system metrics
GetSysMetrics = GetSystemMetrics(SetSysMetIndex(GetWhat))
End Function

Public Function IsCorrect(Cards() As Integer, Round%, CheckWhat$) As Boolean
Dim Wilds% 'wild counter
Dim i As Integer 'loop counter
Dim tCards() As Integer 'temp array
ReDim tCards(UBound(Cards())) As Integer 'get the right dimensions
For i = 0 To UBound(Cards()) 'reduce the cards
    Select Case Cards(i)
        Case 0 To 12:
            tCards(i) = Cards(i) '-0
        Case 13 To 25:
            tCards(i) = (Cards(i) - 13) '-13
        Case 26 To 38:
            tCards(i) = (Cards(i) - (13 * 2)) '-26
        Case 39 To 51:
            tCards(i) = (Cards(i) - (13 * 3)) '-39
    End Select
Next
Wilds% = 0 'default to 0
If Round% <> 14 Or Round <> 15 Then 'if not Ace or Two round
    For i = 0 To UBound(tCards)
        If tCards(i) = Round% - 1 Then 'if theres a wild
            Wilds% = Wilds% + 1 'add 1 to the total
        End If
    Next
Else 'if round ace or two
    For i = 0 To UBound(tCards)
        If tCards(i) = Round% - 14 Then 'if a wild
            Wilds% = Wilds% + 1 'add 1 to the total
        End If
    Next
End If
If CheckWhat$ = "PAIRS" Then 'if we need to check 3ofakind
    If Wilds% >= 2 And UBound(Cards) >= 4 Then 'if there are more then
                '2 or more wilds, and 4 or more cards
        If CheckStr(Cards(), Wilds%, Round%) = False Then 'first check if
                    'its a straight
                'if not, check threeofakind
            If Check3s(tCards(), Wilds%, Round%) = True Then
                IsCorrect = True 'if yes, make this true
            Else
                IsCorrect = False 'if no, make this false
            End If
        Else
            IsCorrect = True 'if yes to straight, make this trUe
        End If
    Else
        'check threeofakind
        If Check3s(tCards(), Wilds%, Round%) = True Then
            IsCorrect = True 'if yes, make this true
        Else
            If CheckStr(Cards(), Wilds%, Round%) = False Then
                IsCorrect = False 'check a straight, if not, make this false
            Else
                IsCorrect = True 'if it works, make this true
            End If
        End If
    End If
Else
    'check a straight
    If CheckStr(Cards(), Wilds%, Round%) = True Then 'if yes
        IsCorrect = True 'then make this true
    Else
        If Check3s(tCards(), Wilds%, Round%) = False Then 'check threeofakind
            IsCorrect = False 'if no, make this false
        Else
            IsCorrect = True 'if yes, make this true
        End If
    End If
End If
End Function

Private Function SetSysMetIndex(SysMetVal As SysMet) As Long
'get the index number from the enum
Select Case SysMetVal
    Case 0:
        SetSysMetIndex = 0
    Case 1:
        SetSysMetIndex = 1
    Case 2:
        SetSysMetIndex = 2
    Case 3:
        SetSysMetIndex = 3
    Case 4:
        SetSysMetIndex = 4
    Case 5:
        SetSysMetIndex = 5
    Case 6:
        SetSysMetIndex = 6
    Case 7:
        SetSysMetIndex = 7
    Case 8:
        SetSysMetIndex = 8
    Case 9:
        SetSysMetIndex = 9
    Case 10:
        SetSysMetIndex = 10
    Case 11:
        SetSysMetIndex = 11
    Case 12:
        SetSysMetIndex = 12
    Case 13:
        SetSysMetIndex = 13
    Case 14:
        SetSysMetIndex = 14
    Case 15:
        SetSysMetIndex = 15
    Case 16:
        SetSysMetIndex = 16
    Case 17:
        SetSysMetIndex = 17
    Case 18:
        SetSysMetIndex = 18
    Case 19:
        SetSysMetIndex = 19
    Case 20:
        SetSysMetIndex = 20
    Case 21:
        SetSysMetIndex = 21
    Case 22:
        SetSysMetIndex = 22
    Case 23:
        SetSysMetIndex = 23
    Case 24:
        SetSysMetIndex = 28
    Case 25:
        SetSysMetIndex = 29
    Case 26:
        SetSysMetIndex = 30
    Case 27:
        SetSysMetIndex = 31
    Case 28:
        SetSysMetIndex = 34
    Case 29:
        SetSysMetIndex = 35
    Case 30:
        SetSysMetIndex = 36
    Case 31:
        SetSysMetIndex = 37
    Case 32:
        SetSysMetIndex = 38
    Case 33:
        SetSysMetIndex = 39
    Case 34:
        SetSysMetIndex = 40
    Case 35:
        SetSysMetIndex = 41
    Case 36:
        SetSysMetIndex = 42
    Case 37:
        SetSysMetIndex = 43
    Case 38:
        SetSysMetIndex = 44
    Case 39:
        SetSysMetIndex = 67
    Case 40:
        SetSysMetIndex = 61
    Case 41:
        SetSysMetIndex = 59
    Case 42:
        SetSysMetIndex = 71
    Case 43:
        SetSysMetIndex = 54
    Case 44:
        SetSysMetIndex = 57
    Case 45:
        SetSysMetIndex = 62
    Case 46:
        SetSysMetIndex = 60
    Case 47:
        SetSysMetIndex = 72
    Case 48:
        SetSysMetIndex = 55
    Case 49:
        SetSysMetIndex = 58
    Case 50:
        SetSysMetIndex = 51
    Case 51:
        SetSysMetIndex = 74
    Case 52:
        SetSysMetIndex = 63
    Case 53:
        SetSysMetIndex = 73
End Select
End Function

Private Function Straight(Cards() As Integer) As Boolean
Dim X%, c%, a%, RemStr As Boolean, HowLong% 'counters/temp values/ etc.
BubbleSort Cards() 'sort the cards
X% = Cards(0) 'get the first value
c% = 1 'set to default
RemStr = False 'set to default
For a% = 0 To UBound(Cards)
    If a% + 1 > UBound(Cards) Then Exit For 'exit if value gets too high
    If Cards(a%) <> -1 Then 'if not a blank card
        If (Cards(a%) + 1) = Cards(a% + 1) Then 'if the val of cards(a) plus 1,
                    'is equal to the next Cards(a+1) then continue
            HowLong% = 2 'there are 2 cards of the straight so far
            For i = a% To UBound(Cards)
                If Cards(a%) <> -1 Then 'if not a blank
                    If a% + HowLong% > UBound(Cards) Then Exit For 'if num to high
                    If Cards(a% + (HowLong% - 1)) + 1 = Cards(a% + HowLong%) Then
                            'if the cards go in order
                        HowLong% = HowLong% + 1 'add 1 to howlong
                        RemStr = True 'set this to true so we know we
                                'how a straight
                    Else
                        Exit For 'exit the loop
                    End If
                End If
            Next
            If RemStr = True Then 'if theres a straight
                If HowLong% = UBound(Cards) + 1 Then 'if the straight is the
                        'whole hand
                    Straight = True 'set this function to true
                    Exit For 'exit the loop
                Else
                    'if not, make it false
                    Straight = False
                    Exit For 'exit the loop
                End If
            End If
        End If
    End If
    RemStr = False 'default to false
Next
End Function

Private Function ThreeOfAKind(tCards() As Integer) As Boolean
Dim xx As Integer 'counters/temp values
Dim X%, p$
X% = 0
p$ = tCards(0) 'get the card they are trying to match
For xx = 0 To UBound(tCards) 'if each value  matches what it should be
    If Val(tCards(xx)) = Val(p$) Then X% = X% + 1 'add 1 to x
Next
If X% - 1 = UBound(tCards) Then 'if all the cards are the same value
    ThreeOfAKind = True 'make this true
Else
    ThreeOfAKind = False 'if not, make this false
End If
End Function

Public Function WaitFor(MS As Long)
If MS = 1 Or MS = 2 Then MS = 400 '(i changed my waitfor statment...
            'most of the way through for more precision, and i didnt
            'want to change all waitfor 1 to waitfor 400, so this was
            'easyier, and takes up less code
Dim start
start = GetTickCount
'pause for the specified amount of time
While start + MS > GetTickCount
    DoEvents
Wend
End Function

Public Sub PutOnTop(Frm As Form, TopIt As Boolean)
    Dim tVal As Long
    If TopIt Then 'if we want it ontop then
        tVal = HWND_TOPMOST 'set to the new value
    Else
        tVal = HWND_NOTOPMOST 'if not on top, set to this vale
    End If
    'put it ontop/take it off
    SetWindowPos Frm.hwnd, tVal, Frm.Left / Screen.TwipsPerPixelX, _
        Frm.Top / Screen.TwipsPerPixelY, Frm.Width / Screen.TwipsPerPixelX, _
        Frm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


