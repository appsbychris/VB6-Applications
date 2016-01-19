Attribute VB_Name = "modBank"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modBank
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function BankOptions(Index As Long, Optional dbIndex As Long) As Boolean
If Deposit(Index, dbIndex) = True Then BankOptions = True: Exit Function
If WithDrawl(Index, dbIndex) = True Then BankOptions = True: Exit Function
If ShowAccount(Index, dbIndex) = True Then BankOptions = True: Exit Function
End Function

Public Function Deposit(Index As Long, dbIndex As Long) As Boolean
Dim DepAmount As Double

If modSC.FastStringComp(LCaseFast(Left$(X(Index), 3)), "dep") Then  'keyword
    Deposit = True
    X(Index) = Mid$(X(Index), InStr(1, X(Index), " ") + 1)
    DepAmount = Val(TrimIt(X(Index)))
    If DepAmount < 1 Then
        WrapAndSend Index, RED & "You must specify an amount!" & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    If dbMap(dbPlayers(dbIndex).lDBLocation).iType <> 5 Then
        WrapAndSend Index, RED & "This is not a bank!" & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(dbIndex)
        If DepAmount <= .dGold Then
            .dGold = .dGold - DepAmount
            .dBank = .dBank + DepAmount
            WrapAndSend Index, LIGHTBLUE & "You deposit " & DepAmount & " gold into your account." & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " deposits some money." & WHITE & vbCrLf, .lLocation
            X(Index) = ""
        Else
            WrapAndSend Index, RED & "You don't have the much gold!" & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
    End With
End If
End Function

Public Function WithDrawl(Index As Long, Optional dbIndex As Long) As Boolean
Dim WithAmount As Double
Dim dTemp As Double
If modSC.FastStringComp(LCaseFast(Left$(X(Index), 4)), "with") Then   'keyword
    WithDrawl = True
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    If dbMap(dbPlayers(dIndex).lDBLocation).iType <> 5 Then
        WithDrawl = False
        Exit Function
    End If
    X(Index) = Mid$(X(Index), InStr(1, X(Index), " ") + 1)
    WithAmount = Val(TrimIt(X(Index)))
    If WithAmount < 1 Then
        WrapAndSend Index, RED & "You must specify an amount!" & vbCrLf & WHITE
        X(Index) = ""
        Exit Function
    End If
    With dbPlayers(dbIndex)
        If WithAmount <= .dBank Then
            If .dGold + WithAmount > modGetData.GetPlayersMaxGold(Index, dbIndex) Then
                WithAmount = modGetData.GetPlayersMaxGold(Index, dbIndex) - .dGold
            End If
            .dGold = .dGold + WithAmount
            .dBank = .dBank - WithAmount
            WrapAndSend Index, LIGHTBLUE & "You withdrawl " & WithAmount & " gold." & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " withdrawls some money." & WHITE & vbCrLf, .lLocation
            X(Index) = ""
        Else
            WrapAndSend Index, RED & "You don't have the much gold in the bank!" & vbCrLf & WHITE
            X(Index) = ""
            Exit Function
        End If
    End With
End If
End Function

Public Function ShowAccount(Index As Long, dbIndex As Long) As Boolean
Dim ToSend$
If modSC.FastStringComp(LCaseFast(X(Index)), "bank") Then
    ShowAccount = True
    If dbIndex = 0 Then dbIndex = GetPlayerIndexNumber(Index)
    ToSend$ = BRIGHTBLUE & "On deposit: " & vbCrLf
    ToSend$ = ToSend$ & GREEN & dbPlayers(dbIndex).dBank & " gold."
    WrapAndSend Index, ToSend$ & WHITE & vbCrLf
    X(Index) = ""
End If
End Function
