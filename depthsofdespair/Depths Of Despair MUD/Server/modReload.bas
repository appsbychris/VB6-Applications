Attribute VB_Name = "modReload"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modReload
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function Reload(Index As Long) As Boolean
Dim s As String
Dim lAmount As Long
Dim dbItemID As Long
Dim sItemID As String
Dim sArrow As String
Dim m As Long
Dim n As Long
Dim i As Long
Dim j As Long
Dim l As Long
Dim f As Long
Dim Arr() As String
Dim dbIndex As Long
s = LCaseFast(X(Index))
If s Like "rel* #* *" Or s Like "rel* *" Then
    If s Like "rel* #* *" Then
        m = InStr(1, s, " ")
        s = Mid$(s, m + 1)
        m = InStr(1, s, " ")
        lAmount = Val(Left$(s, m))
        sItemID = Mid$(s, m + 1)
    ElseIf s Like "rel* *" Then
        m = InStr(1, s, " ")
        s = Mid$(s, m + 1)
        sItemID = s
        lAmount = 1
    End If
    Reload = True
    sItemID = SmartFind(Index, sItemID, Inventory_Item, True, sArrow)
    If InStr(1, sItemID, Chr$(0)) > 0 Then sItemID = Mid$(sItemID, InStr(1, sItemID, Chr$(0)) + 1)
    dbItemID = GetItemID(sItemID)
    If dbItemID = 0 Then
        WrapAndSend Index, RED & "You don't have that!" & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    s = ""
    j = 0
    dbIndex = GetPlayerIndexNumber(Index)
    For i = 1 To lAmount
        If InStr(1, dbPlayers(dbIndex).sInventory, ":" & dbItems(dbItemID).iID & "/") <> 0 Then
            Call SmartFind(Index, sItemID, Inventory_Item, True, sArrow)
            s = s & sArrow & ";"
            modItemManip.RemoveItemFromInv dbIndex, dbItems(dbItemID).iID
            j = j + 1
        End If
        If DE Then DoEvents
    Next
    SplitFast s, Arr, ";"
    l = 0
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) <> "" Then
            l = l + modItemManip.GetItemUsesFromUnFormattedString(Arr(i))
        End If
        If DE Then DoEvents
    Next
    With dbPlayers(dbIndex)
        If (modItemManip.GetItemBulletsID(.sWeapon) = dbItems(dbItemID).iID) Or modItemManip.GetItemBulletsID(.sWeapon) = 0 Then
            If modItemManip.GetItemBulletsLeft(.sWeapon) + l _
                > modItemManip.GetItemUsesFromUnFormattedString(.sWeapon) _
                    Then
                        '--------------------
                        WrapAndSend Index, RED & "You don't have enough space!" & WHITE & vbCrLf
                        X(Index) = ""
                        If .sInventory = "0" Then .sInventory = ""
                        .sInventory = .sInventory & s
                        Exit Function
                        '--------------------
            End If
        Else
            If l > modItemManip.GetItemUsesFromUnFormattedString(.sWeapon) Then
                '--------------------
                WrapAndSend Index, RED & "You don't have enough space!" & WHITE & vbCrLf
                X(Index) = ""
                If .sInventory = "0" Then .sInventory = ""
                .sInventory = .sInventory & s
                Exit Function
                '--------------------
            End If
            If .sInventory = "0" Then .sInventory = ""
            l = modItemManip.GetItemBulletsID(.sWeapon)
            m = modItemManip.GetItemBulletsLeft(.sWeapon)
            .sWeapon = modItemManip.SetItemBullets(.sWeapon, "0|0|0|0")
            l = GetItemID(, l)
            With dbItems(l)
                If .iUses <> 1 Then
                    Do Until m = 0
                        If m >= .iUses Then
                            dbPlayers(dbIndex).sInventory = dbPlayers(dbIndex).sInventory & ":" & .iID & "/" & .lDurability & "/E{}F{}A{}B{0|0|0|0}/" & .iUses & ";"
                            m = m - .iUses
                        Else
                            dbPlayers(dbIndex).sInventory = dbPlayers(dbIndex).sInventory & ":" & .iID & "/" & .lDurability & "/E{}F{}A{}B{0|0|0|0}/" & m & ";"
                            m = 0
                        End If
                        If DE Then DoEvents
                    Loop
                Else
                    For j = 1 To m
                        dbPlayers(dbIndex).sInventory = dbPlayers(dbIndex).sInventory & ":" & .iID & "/" & .lDurability & "/E{}F{}A{}B{0|0|0|0}/1;"
                        If DE Then DoEvents
                    Next
                End If
            End With
        End If
        For f = 1 To lAmount
            m = modItemManip.GetItemBulletsLeft(.sWeapon)
            n = modItemManip.GetItemIDFromUnFormattedString(sArrow)
            With dbItems(dbItemID)
                i = .iMagical
                j = .iAC
            End With
            l = modItemManip.GetItemUsesFromUnFormattedString(sArrow)
            If l = 0 Then l = 1
            .sWeapon = modItemManip.SetItemBullets(.sWeapon, modItemManip.SetupItemBullets(.sWeapon, l + m, n, i, j))
            If DE Then DoEvents
        Next
        If lAmount > 1 Then
            WrapAndSend Index, LIGHTBLUE & "You ready " & lAmount & " " & modgetdata.GetItemsNameAddS(dbItemID) & " for your " & dbItems(GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sWeapon))).sItemName & "." & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " readies " & lAmount & " " & modgetdata.GetItemsNameAddS(dbItemID) & "." & WHITE & vbCrLf, .lLocation
        Else
            WrapAndSend Index, LIGHTBLUE & "You ready " & lAmount & " " & dbItems(dbItemID).sItemName & " for your " & dbItems(GetItemID(, modItemManip.GetItemIDFromUnFormattedString(.sWeapon))).sItemName & "." & WHITE & vbCrLf
            SendToAllInRoom Index, LIGHTBLUE & .sPlayerName & " readies " & lAmount & " " & dbItems(dbItemID).sItemName & "." & WHITE & vbCrLf, .lLocation
        End If
        X(Index) = ""
    End With
End If
End Function

Public Function UnloadWeapon(Index As Long) As Boolean
Dim l As Long
Dim j As Long
Dim m As Long
Dim dbIndex As Long
If LCaseFast(X(Index)) Like "unlo*" Then
    UnloadWeapon = True
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        l = modItemManip.GetItemBulletsID(.sWeapon)
        m = modItemManip.GetItemBulletsLeft(.sWeapon)
        l = GetItemID(, l)
        If l = 0 Then
            WrapAndSend Index, RED & "You don't have anything loaded!" & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
        If .sInventory = "0" Then .sInventory = ""
        With dbItems(l)
            If .iUses <> 1 Then
                Do Until m = 0
                    If m >= .iUses Then
                        dbPlayers(dbIndex).sInventory = dbPlayers(dbIndex).sInventory & ":" & .iID & "/" & .lDurability & "/E{}F{}A{}B{0|0|0|0}/" & .iUses & ";"
                        m = m - .iUses
                    Else
                        dbPlayers(dbIndex).sInventory = dbPlayers(dbIndex).sInventory & ":" & .iID & "/" & .lDurability & "/E{}F{}A{}B{0|0|0|0}/" & m & ";"
                        m = 0
                    End If
                    If DE Then DoEvents
                Loop
            Else
                For j = 1 To m
                    dbPlayers(dbIndex).sInventory = dbPlayers(dbIndex).sInventory & ":" & .iID & "/" & .lDurability & "/E{}F{}A{}B{0|0|0|0}/1;"
                    If DE Then DoEvents
                Next
            End If
        End With
        .sWeapon = modItemManip.SetItemBullets(.sWeapon, "0|0|0|0")
    End With
    
    X(Index) = ""
    WrapAndSend Index, LIGHTBLUE & "You unload your weapon." & WHITE & vbCrLf
End If
End Function
