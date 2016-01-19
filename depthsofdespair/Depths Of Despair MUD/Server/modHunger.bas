Attribute VB_Name = "modHunger"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modHunger
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function Eat(Index As Long) As Boolean
Dim s As String
Dim sUn As String
Dim dbItemID As Long
Dim tArr() As String
Dim dbIndex As Long
Dim i As Long
Dim SendOthers As String
If Left$(LCaseFast(X(Index)), 4) = "eat " Then
    Eat = True
    s = ReplaceFast(X(Index), "eat ", "")
    s = SmartFind(Index, s, Inventory_Item, True, sUn)
    If InStr(1, s, Chr$(0)) > 0 Then s = Mid$(s, InStr(1, s, Chr$(0)) + 1)
    dbItemID = GetItemID(s)
    If dbItemID = 0 Then
        WrapAndSend Index, RED & "You don't seem to have that." & WHITE & vbCrLf
        X(Index) = ""
        Exit Function
    End If
    dbIndex = GetPlayerIndexNumber(Index)
    With dbPlayers(dbIndex)
        If InStr(1, .sInventory, ":" & CStr(dbItems(dbItemID).iID) & "/") = 0 Then
            WrapAndSend Index, RED & "You don't seem to have that." & WHITE & vbCrLf
            X(Index) = ""
            Exit Function
        End If
    End With
    With dbItems(dbItemID)
        
        If modSC.FastStringComp(.sWorn, "corpse") Or modSC.FastStringComp(.sWorn, "food") Or modSC.FastStringComp(.sWorn, "ofood") Then
            If modSC.FastStringComp(.sWorn, "corpse") Then
                If RndNumber(0, 1) = 0 Then
                    WrapAndSend Index, BRIGHTRED & "After eating the corpse, you begin to feel sick." & WHITE & vbCrLf, False
                    SplitFast .sDamage, tArr, ":"
                    With dbPlayers(dbIndex)
                        i = RndNumber(CDbl(tArr(0)), CDbl(tArr(1)))
                        .lHP = .lHP - i
                        .dHunger = .dHunger - (i \ 2)
                        .dStamina = .dStamina - i
                    End With
                    SendOthers = GREEN & dbPlayers(dbIndex).sPlayerName & " eats a corpse, and looks sick!" & WHITE & vbCrLf
                    WrapAndSend Index, BRIGHTRED & "You take " & CStr(i) & " damage!" & WHITE & vbCrLf
                    modItemManip.RemoveItemFromInv dbIndex, modItemManip.GetItemIDFromUnFormattedString(sUn)
                    SendToAllInRoom Index, SendOthers, dbPlayers(dbIndex).lLocation
                    X(Index) = ""
                    Exit Function
                End If
            End If
            If modSC.FastStringComp(.sWorn, "ofood") Then
                If RndNumber(0, 1) = 0 Then
                    WrapAndSend Index, BRIGHTRED & "After eating the " & s & ", you begin to feel sick." & WHITE & vbCrLf, False
                    SplitFast .sDamage, tArr, ":"
                    With dbPlayers(dbIndex)
                        i = RndNumber(CDbl(tArr(0)), CDbl(tArr(1)))
                        .lHP = .lHP - i
                        .dHunger = .dHunger - (i \ 2)
                        .dStamina = .dStamina - i
                    End With
                    SendOthers = GREEN & dbPlayers(dbIndex).sPlayerName & " eats a " & s & " and looks sick!" & WHITE & vbCrLf
                    WrapAndSend Index, BRIGHTRED & "You take " & CStr(i) & " damage!" & WHITE & vbCrLf
                    modItemManip.RemoveItemFromInv dbIndex, modItemManip.GetItemIDFromUnFormattedString(sUn)
                    SendToAllInRoom Index, SendOthers, dbPlayers(dbIndex).lLocation
                    X(Index) = ""
                    Exit Function
                End If
            End If
            SplitFast .sDamage, tArr, ":"
            With dbPlayers(dbIndex)
                i = RndNumber(CDbl(tArr(0)), CDbl(tArr(1)))
                .lHP = .lHP + i
                If .lHP > .lMaxHP Then .lHP = .lMaxHP
                .dHunger = .dHunger + i
                .dStamina = .dStamina + (i \ 3)
            End With
            modItemManip.RemoveItemFromInv dbIndex, modItemManip.GetItemIDFromUnFormattedString(sUn)
            SendOthers = LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & " eats a " & s & "." & WHITE & vbCrLf
            WrapAndSend Index, LIGHTBLUE & "After eating the " & s & ", you begin to feel revitalized." & WHITE & vbCrLf
            SendToAllInRoom Index, SendOthers, dbPlayers(dbIndex).lLocation
        Else
            WrapAndSend Index, RED & "You can't eat that." & WHITE & vbCrLf
        End If
    End With
    X(Index) = ""
End If
End Function

Public Sub DropOutDoorFood(dbMapIndex As Long)
Dim s As String
Dim i As Long
Dim tArr() As String
Dim dd As Long
SplitFast dbMap(dbMapIndex).sOutDoorFood, tArr, ";"
If UBound(tArr) < 1 And tArr(0) = "0" Then Exit Sub
i = RndNumber(0, UBound(tArr) - 1)
If tArr(i) = "0" Then Exit Sub
If DCount(dbMap(dbMapIndex).sItems, ":" & tArr(i) & "/") > 3 Then Exit Sub
If modSC.FastStringComp(dbMap(dbMapIndex).sItems, "0") Then dbMap(dbMapIndex).sItems = ""
dd = GetItemID(, Val(tArr(i)))
tArr(i) = ":" & tArr(i) & "/1/E{}F{}A{}B{0|0|0|0}/" & dbItems(dd).iUses & ";"
dbMap(dbMapIndex).sItems = dbMap(dbMapIndex).sItems & tArr(i) & ";"
SendToAllInRoom 0, LIGHTBLUE & "You notice a " & dbItems(dd).sItemName & " fall to the ground here." & WHITE & vbCrLf, dbMap(dbMapIndex).lRoomID
End Sub
