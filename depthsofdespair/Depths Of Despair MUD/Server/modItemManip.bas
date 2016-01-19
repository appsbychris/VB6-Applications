Attribute VB_Name = "modItemManip"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modItemManip
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Sub TakeLetterFromHiddenAndPutItInInv(dbIndex As Long, ByVal lLID As Long)
With dbMap(dbPlayers(dbIndex).lDBLocation)
    .sHLetters = ReplaceFast(.sHLetters, ":" & CStr(lLID) & ";", "")
    If modSC.FastStringComp(.sHLetters, "") Then .sHLetters = "0"
    With dbPlayers(dbIndex)
        If modSC.FastStringComp(.sLetters, "0") Then .sLetters = ""
        .sLetters = .sLetters & ":" & CStr(lLID) & ";"
    End With
End With
End Sub

Sub TakeLetterFromGroundAndPutItInInv(dbIndex As Long, ByVal lLID As Long)
With dbMap(dbPlayers(dbIndex).lDBLocation)
    .sLetters = ReplaceFast(.sLetters, ":" & CStr(lLID) & ";", "")
    If modSC.FastStringComp(.sLetters, "") Then .sLetters = "0"
    With dbPlayers(dbIndex)
        If modSC.FastStringComp(.sLetters, "0") Then .sLetters = ""
        .sLetters = .sLetters & ":" & CStr(lLID) & ";"
    End With
End With
End Sub

Sub TakeLetterFromInvAndDropIt(dbIndex As Long, ByVal lLID As Long)
With dbPlayers(dbIndex)
    .sLetters = ReplaceFast(.sLetters, ":" & CStr(lLID) & ";", "")
    If modSC.FastStringComp(.sLetters, "") Then .sLetters = "0"
    With dbMap(.lDBLocation)
        If modSC.FastStringComp(.sLetters, "0") Then .sLetters = ""
        .sLetters = .sLetters & ":" & CStr(lLID) & ";"
    End With
End With
End Sub

Sub TakeLetterFromInvAndPutInAnotherInv(dbIndex As Long, dbIndex2 As Long, ByVal lLID As Long)
With dbPlayers(dbIndex)
    .sLetters = ReplaceFast(.sLetters, ":" & CStr(lLID) & ";", "")
    If modSC.FastStringComp(.sLetters, "") Then .sLetters = "0"
End With
With dbPlayers(dbIndex2)
    If .sLetters = "0" Then .sLetters = ""
    .sLetters = .sLetters & ":" & CStr(lLID) & ";"
End With
End Sub

Sub TakeLetterFromInvAndHideIt(dbIndex As Long, ByVal lLID As Long)
With dbPlayers(dbIndex)
    .sLetters = ReplaceFast(.sLetters, ":" & CStr(lLID) & ";", "")
    If modSC.FastStringComp(.sLetters, "") Then .sLetters = "0"
    With dbMap(.lDBLocation)
        If modSC.FastStringComp(.sHLetters, "0") Then .sHLetters = ""
        .sHLetters = .sHLetters & ":" & CStr(lLID) & ";"
    End With
End With
End Sub

Public Function GetListOfLettersFromInv(dbIndex As Long) As String
Dim s As String
Dim tArr() As String
Dim i As Long
With dbPlayers(dbIndex)
    s = .sLetters
    If s = "0" Then
        GetListOfLettersFromInv = ""
        Exit Function
    End If
    s = ReplaceFast(s, ":", "")
    SplitFast s, tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        If Not modSC.FastStringComp(tArr(i), "") Then GetListOfLettersFromInv = GetListOfLettersFromInv & "note: " & dbLetters(GetLetterID(, CLng(tArr(i)))).sTitle & ","
        If DE Then DoEvents
    Next
End With
End Function

Public Function GetListOfLettersFromGround(ByVal dbLoc As Long) As String
Dim s As String
Dim tArr() As String
Dim i As Long
With dbMap(dbLoc)
    s = .sLetters
    If s = "0" Then
        GetListOfLettersFromGround = ""
        Exit Function
    End If
    s = ReplaceFast(s, ":", "")
    SplitFast s, tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        If Not modSC.FastStringComp(tArr(i), "") Then GetListOfLettersFromGround = GetListOfLettersFromGround & "note: " & dbLetters(GetLetterID(, CLng(tArr(i)))).sTitle & ","
        If DE Then DoEvents
    Next
End With
End Function

Public Function GetListOfLettersFromHidden(ByVal dbLoc As Long) As String
Dim s As String
Dim tArr() As String
Dim i As Long
With dbMap(dbLoc)
    s = .sHLetters
    If s = "0" Then
        GetListOfLettersFromHidden = ""
        Exit Function
    End If
    s = ReplaceFast(s, ":", "")
    SplitFast s, tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        If Not modSC.FastStringComp(tArr(i), "") Then GetListOfLettersFromHidden = GetListOfLettersFromHidden & "note: " & dbLetters(GetLetterID(, CLng(tArr(i)))).sTitle & ","
        If DE Then DoEvents
    Next
End With
End Function

Public Function FindItemInInvFromIndexNumber(dbIndex As Long, ByVal iItemID As Long) As String
Dim tArr() As String
Dim i As Long
Dim m As Long
With dbPlayers(dbIndex)
    SplitFast Left$(.sInventory, Len(.sInventory) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            FindItemInInvFromIndexNumber = tArr(i)
            Exit Function
        End If
        If DE Then DoEvents
    Next
End With
End Function

Sub JoinInventory(dbIndex As Long, tArr() As String)
Dim i As Long
Dim s As String
For i = LBound(tArr) To UBound(tArr)
    If tArr(i) <> "" Then
        If Left$(tArr(i), 1) <> ":" Then tArr(i) = ":" & tArr(i)
        s = s & tArr(i) & IIf(Right$(tArr(i), 1) = ";", "", ";")
    End If
    If DE Then DoEvents
Next
With dbPlayers(dbIndex)
    .sInventory = s
    If modSC.FastStringComp(.sInventory, "") Then .sInventory = "0"
End With
End Sub

Sub SubtractOneFromItemUseINV(dbIndex As Long, ByVal iItemID As Long, ByVal iCurrentUses As Long, ByVal iCurrentDur As Long)
Dim tArr() As String
Dim i As Long
Dim m As Long
Dim n As Long
Dim j As Long
Dim s As String
With dbPlayers(dbIndex)
    s = .sInventory
    SplitFast Left$(s, Len(s) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = modItemManip.GetItemIDFromUnFormattedString(tArr(i))
        n = modItemManip.GetItemDurFromUnFormattedString(tArr(i))
        j = modItemManip.GetItemUsesFromUnFormattedString(tArr(i))
        If (iItemID = m) And (iCurrentUses = j) And (iCurrentDur = n) Then
            s = tArr(i)
            If j - 1 > 0 Then
                modItemManip.SetItemUses dbIndex, s, (j - 1)
            Else
                modItemManip.LastUseFlags2 dbIndex, iItemID
                modItemManip.RemoveItemFromInv dbIndex, iItemID
            End If
            Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
End Sub

Sub SubtractOneFromItemDUR(dbIndex As Long, ByVal iItemID As Long, ByVal iCurrentUses As Long, ByVal iCurrentDur As Long)
Dim tArr() As String
Dim i As Long
Dim m As Long
Dim n As Long
Dim j As Long
Dim s As String
With dbPlayers(dbIndex)
    s = modGetData.GetPlayersEq(.iIndex)
    SplitFast Left$(s, Len(s) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = modItemManip.GetItemIDFromUnFormattedString(tArr(i))
        n = modItemManip.GetItemDurFromUnFormattedString(tArr(i))
        j = modItemManip.GetItemUsesFromUnFormattedString(tArr(i))
        If (iItemID = m) And (iCurrentUses = j) And (iCurrentDur = n) Then
            s = tArr(i)
            If n - 1 > 0 Then
                modItemManip.SetItemDur dbIndex, s, (n - 1)
            Else
                modItemManip.AdjustStats dbIndex, CLng(m), 0
                modItemManip.RemoveItemFromEQ dbIndex, s
                SetWeaponStats .iIndex
            End If
            Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
End Sub

Sub SubtractNFromItemDUR(dbIndex As Long, ByVal lN As Long, ByVal iItemID As Long, ByVal iCurrentUses As Long, ByVal iCurrentDur As Long)
Dim tArr() As String
Dim i As Long
Dim m As Long
Dim n As Long
Dim j As Long
Dim s As String
With dbPlayers(dbIndex)
    s = modGetData.GetPlayersEq(.iIndex)
    SplitFast Left$(s, Len(s) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = modItemManip.GetItemIDFromUnFormattedString(tArr(i))
        n = modItemManip.GetItemDurFromUnFormattedString(tArr(i))
        j = modItemManip.GetItemUsesFromUnFormattedString(tArr(i))
        If (iItemID = m) And (iCurrentUses = j) And (iCurrentDur = n) Then
            s = tArr(i)
            If n - lN > 0 Then
                modItemManip.SetItemDur dbIndex, s, (n - lN)
            Else
                modItemManip.AdjustStats dbIndex, CLng(m), 0
                modItemManip.RemoveItemFromEQ dbIndex, s
                SetWeaponStats .iIndex
            End If
            Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
End Sub

Sub LastUseFlags2(dbIndex As Long, ByVal iItemID As Long)
Dim dbItemID As Long
dbItemID = GetItemID(, CLng(iItemID))
With dbItems(dbItemID)
    If .lOnLastUseDoFlags2 = 1 Then modUseItems.DoFlags dbIndex, .sFlags2
End With
End Sub

Sub SetItemUses(dbIndex As Long, ByVal UnFormatedItemString As String, ByVal lNewUses As Long)
Dim tArr() As String
Dim i As Long
With dbPlayers(dbIndex)
    SplitFast Left$(.sInventory, Len(.sInventory) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        If modSC.FastStringComp(tArr(i), UnFormatedItemString) Then
            tArr(i) = ":" & modItemManip.GetItemIDFromUnFormattedString(tArr(i)) & "/" & _
                        CStr(lNewDur) & "/E{" & _
                        modItemManip.GetItemEnchantsFromUnFormattedString(tArr(i)) & "}F{" & _
                        modItemManip.GetItemFlagsFromUnFormattedString(tArr(i)) & "}A{" & _
                        modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)) & "}B{" & _
                        modItemManip.GetItemBulletsFromUnFormattedString(tArr(i)) & "}/" & _
                        modItemManip.GetItemUsesFromUnFormattedString(tArr(i))
            JoinInventory dbIndex, tArr
            Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
End Sub

Sub SetItemDur(dbIndex As Long, ByVal UnFormatedItemString As String, ByVal lNewDur As Long)
Dim tArr() As String
Dim i As Long
Dim s As String
With dbPlayers(dbIndex)
    s = modGetData.GetPlayersEq(.iIndex)
    SplitFast Left$(s, Len(s) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        If modSC.FastStringComp(tArr(i), UnFormatedItemString) Then
            tArr(i) = ":" & modItemManip.GetItemIDFromUnFormattedString(tArr(i)) & "/" & _
                        CStr(lNewDur) & "/E{" & _
                        modItemManip.GetItemEnchantsFromUnFormattedString(tArr(i)) & "}F{" & _
                        modItemManip.GetItemFlagsFromUnFormattedString(tArr(i)) & "}A{" & _
                        modItemManip.GetItemAdjectivesFromUnFormattedString(tArr(i)) & "}B{" & _
                        modItemManip.GetItemBulletsFromUnFormattedString(tArr(i)) & "}/" & _
                        modItemManip.GetItemUsesFromUnFormattedString(tArr(i))
            Select Case i
                Case 0
                    .sArms = tArr(i)
                Case 1
                    .sBack = tArr(i)
                Case 2
                    .sBody = tArr(i)
                Case 3
                    .sEars = tArr(i)
                Case 4
                    .sFace = tArr(i)
                Case 5
                    .sFeet = tArr(i)
                Case 6
                    .sHands = tArr(i)
                Case 7
                    .sHead = tArr(i)
                Case 8
                    .sLegs = tArr(i)
                Case 9
                    .sNeck = tArr(i)
                Case 10
                    .sShield = tArr(i)
                Case 11
                    .sWaist = tArr(i)
                Case 12
                    .sWeapon = tArr(i)
            End Select
            Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
End Sub

Public Function GetItemDurFromUnFormattedString(ByVal s As String) As Long
Dim m As Long
Dim n As Long
   On Error GoTo GetItemDurFromUnFormattedString_Error

m = InStr(1, s, ":")
n = InStr(m, s, "/")
m = InStr(n + 1, s, "/")
GetItemDurFromUnFormattedString = CLng(Mid$(s, n + 1, m - n - 1))

   On Error GoTo 0
   Exit Function

GetItemDurFromUnFormattedString_Error:
    GetItemDurFromUnFormattedString = -1
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetItemDurFromUnFormattedString of Module modItemManip"
End Function

Public Function ClearItemEnchants(ByVal s As String) As String
Dim m As Long
Dim u As String
Dim n As Long

   On Error GoTo ClearItemEnchants_Error

m = InStr(1, s, "E{") + 1
u = Left$(s, m)
s = Mid$(s, m + 1)
n = InStr(1, s, "}")
s = Mid$(s, n)
ClearItemEnchants = u & s

   On Error GoTo 0
   Exit Function

ClearItemEnchants_Error:
End Function

Public Function ClearItemAdjectives(ByVal s As String) As String
Dim m As Long
Dim u As String
Dim n As Long
   On Error GoTo ClearItemAdjectives_Error

m = InStr(1, s, "A{") + 1
u = Left$(s, m)
s = Mid$(s, m + 1)
n = InStr(1, s, "}")
s = Mid$(s, n)
ClearItemAdjectives = u & s

   On Error GoTo 0
   Exit Function

ClearItemAdjectives_Error:
End Function

Public Function ClearItemFlags(ByVal s As String) As String
Dim m As Long
Dim u As String
Dim n As Long
   On Error GoTo ClearItemFlags_Error

m = InStr(1, s, "F{") + 1
u = Left$(s, m)
s = Mid$(s, m + 1)
n = InStr(1, s, "}")
s = Mid$(s, n)
ClearItemFlags = u & s

   On Error GoTo 0
   Exit Function

ClearItemFlags_Error:
End Function

Public Function SetItemEnchants(ByVal s As String, ByVal sWith As String) As String
Dim m As Long
Dim u As String

   On Error GoTo SetItemEnchants_Error

m = InStr(1, s, "E{") + 1
u = Left$(s, m)
s = Mid$(s, m + 1)

SetItemEnchants = u & sWith & s

   On Error GoTo 0
   Exit Function

SetItemEnchants_Error:
End Function

Public Function SetItemAdjectives(ByVal s As String, ByVal sWith As String) As String
Dim m As Long
Dim u As String

   On Error GoTo SetItemAdjectives_Error

m = InStr(1, s, "A{") + 1
u = Left$(s, m)
s = Mid$(s, m + 1)

SetItemAdjectives = u & sWith & s

   On Error GoTo 0
   Exit Function

SetItemAdjectives_Error:
End Function

Public Function SetItemFlags(ByVal s As String, ByVal sWith As String) As String
Dim m As Long
Dim u As String

   On Error GoTo SetItemFlags_Error

m = InStr(1, s, "F{") + 1
u = Left$(s, m)
s = Mid$(s, m + 1)

SetItemFlags = u & sWith & s

   On Error GoTo 0
   Exit Function

SetItemFlags_Error:
End Function

Public Function SetupItemBullets(ByVal s As String, lB As Long, lID As Long, lMag As Long, lMa As Long) As String
Dim Arr() As String
s = modItemManip.GetItemBulletsFromUnFormattedString(s)
SplitFast s, Arr, "|"
Arr(0) = lB
Arr(1) = lID
Arr(2) = lMag
Arr(3) = lMa
SetupItemBullets = Arr(0) & "|" & Arr(1) & "|" & Arr(2) & "|" & Arr(3)
End Function

Public Function SetItemBullets(ByVal s As String, ByVal sWith As String) As String
Dim m As Long
Dim u As String
Dim n As Long
   On Error GoTo SetItemBullets_Error

m = InStr(1, s, "B{") + 1
u = Left$(s, m)
n = InStr(m, s, "}")
s = Mid$(s, n)

SetItemBullets = u & sWith & s

   On Error GoTo 0
   Exit Function

SetItemBullets_Error:
End Function

Public Function GetItemEnchantsFromUnFormattedString(ByVal s As String) As String
Dim m As Long
Dim n As Long
   On Error GoTo GetItemEnchantsFromUnFormattedString_Error

m = InStr(1, s, "E{") + 1
n = InStr(m, s, "}")
GetItemEnchantsFromUnFormattedString = Mid$(s, m + 1, n - m - 1)

   On Error GoTo 0
   Exit Function

GetItemEnchantsFromUnFormattedString_Error:
End Function

Public Function GetItemBulletsMana(ByVal s As String) As Long
Dim Arr() As String
   On Error GoTo GetItemBulletsMana_Error

s = modItemManip.GetItemBulletsFromUnFormattedString(s)
SplitFast s, Arr, "|"
GetItemBulletsMana = Arr(3)

   On Error GoTo 0
   Exit Function

GetItemBulletsMana_Error:
End Function

Public Function GetItemBulletsMagical(ByVal s As String) As Long
Dim Arr() As String
   On Error GoTo GetItemBulletsMagical_Error

s = modItemManip.GetItemBulletsFromUnFormattedString(s)
SplitFast s, Arr, "|"
GetItemBulletsMagical = Arr(2)

   On Error GoTo 0
   Exit Function

GetItemBulletsMagical_Error:
End Function

Public Function GetItemBulletsID(ByVal s As String) As Long
Dim Arr() As String
   On Error GoTo GetItemBulletsID_Error

s = modItemManip.GetItemBulletsFromUnFormattedString(s)
SplitFast s, Arr, "|"
GetItemBulletsID = Arr(1)

   On Error GoTo 0
   Exit Function

GetItemBulletsID_Error:
End Function

Public Function GetItemBulletsLeft(ByVal s As String) As Long
Dim Arr() As String
   On Error GoTo GetItemBulletsLeft_Error

s = modItemManip.GetItemBulletsFromUnFormattedString(s)
SplitFast s, Arr, "|"
GetItemBulletsLeft = Arr(0)

   On Error GoTo 0
   Exit Function

GetItemBulletsLeft_Error:
End Function

Public Function GetItemBulletsFromUnFormattedString(ByVal s As String) As String
Dim m As Long
Dim n As Long
   On Error GoTo GetItemBulletsFromUnFormattedString_Error

m = InStr(1, s, "B{") + 1
n = InStr(m, s, "}")
GetItemBulletsFromUnFormattedString = Mid$(s, m + 1, n - m - 1)

   On Error GoTo 0
   Exit Function

GetItemBulletsFromUnFormattedString_Error:
End Function

Public Function GetItemAdjectivesFromUnFormattedString(ByVal s As String) As String
Dim m As Long
Dim n As Long
   On Error GoTo GetItemAdjectivesFromUnFormattedString_Error

m = InStr(1, s, "A{") + 1
n = InStr(m, s, "}")
GetItemAdjectivesFromUnFormattedString = Mid$(s, m + 1, n - m - 1)

   On Error GoTo 0
   Exit Function

GetItemAdjectivesFromUnFormattedString_Error:
End Function

Public Function GetItemFlagsFromUnFormattedString(ByVal s As String) As String
Dim m As Long
Dim n As Long
   On Error GoTo GetItemFlagsFromUnFormattedString_Error

m = InStr(1, s, "F{") + 1
n = InStr(m, s, "}")
GetItemFlagsFromUnFormattedString = Mid$(s, m + 1, n - m - 1)

   On Error GoTo 0
   Exit Function

GetItemFlagsFromUnFormattedString_Error:
End Function

Public Function GetItemUsesFromUnFormattedString(ByVal s As String) As Long
Dim m As Long
   On Error GoTo GetItemUsesFromUnFormattedString_Error

m = InStrRev(s, "/")
GetItemUsesFromUnFormattedString = CLng(Mid$(s, m + 1))

   On Error GoTo 0
   Exit Function

GetItemUsesFromUnFormattedString_Error:
    GetItemUsesFromUnFormattedString = -1
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetItemUsesFromUnFormattedString of Module modItemManip"
End Function

Public Function GetItemIDFromUnFormattedString(ByVal s As String) As Long
Dim m As Long
Dim n As Long

   On Error GoTo GetItemIDFromUnFormattedString_Error

m = InStr(1, s, ":")
n = InStr(m, s, "/")
GetItemIDFromUnFormattedString = CLng(Mid$(s, m + 1, n - m - 1))

   On Error GoTo 0
   Exit Function

GetItemIDFromUnFormattedString_Error:

    GetItemIDFromUnFormattedString = -1
End Function

Sub TakeFromYourInvAndPutInAnothersInv(dbIndex1 As Long, dbIndex2 As Long, ByVal iItemID As Long)
Dim tArr() As String
Dim i As Long
Dim m As Long
Dim s As String
With dbPlayers(dbIndex1)
    SplitFast Left$(.sInventory, Len(.sInventory) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            s = tArr(i) & ";"
            tArr(i) = ""
            JoinInventory dbIndex1, tArr
            Exit For
        End If
        If DE Then DoEvents
    Next
End With
With dbPlayers(dbIndex2)
    If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
    .sInventory = .sInventory & s
End With
End Sub

Sub RemoveItemFromInv(dbIndex As Long, ByVal iItemID As Long)
Dim tArr() As String
Dim i As Long
Dim m As Long
With dbPlayers(dbIndex)
    SplitFast Left$(.sInventory, Len(.sInventory) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            tArr(i) = ""
            JoinInventory dbIndex, tArr
            Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
End Sub

Sub RemoveItemFromGround(dbMapIndex As Long, ByVal iItemID As Long)
Dim tArr() As String
Dim i As Long
Dim m As Long
With dbMap(dbMapIndex)
    SplitFast Left$(.sItems, Len(.sItems) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            tArr(i) = ""
            modItemManip.JoinMapItems dbMapIndex, tArr
            Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
End Sub

Sub RemoveItemFromEQ(dbIndex As Long, ByVal sUnformated As String)
Dim tArr() As String
Dim i As Long
Dim s As String
With dbPlayers(dbIndex)
    s = modGetData.GetPlayersEq(.iIndex)
    SplitFast Left$(s, Len(s) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        If modSC.FastStringComp(tArr(i), sUnformated) Then
            Select Case i
                Case 0
                    .sArms = "0"
                Case 1
                    .sBack = "0"
                Case 2
                    .sBody = "0"
                Case 3
                    .sEars = "0"
                Case 4
                    .sFace = "0"
                Case 5
                    .sFeet = "0"
                Case 6
                    .sHands = "0"
                Case 7
                    .sHead = "0"
                Case 8
                    .sLegs = "0"
                Case 9
                    .sNeck = "0"
                Case 10
                    .sShield = "0"
                Case 11
                    .sWaist = "0"
                Case 12
                    .sWeapon = "0"
                Case 13 To 18
                    .sRings(i - 13) = "0"
            End Select
            modItemManip.SendBrokenItemMessage dbIndex, CLng(i), modItemManip.GetItemIDFromUnFormattedString(sUnformated)
            Exit Sub
        End If
        If DE Then DoEvents
    Next
End With
End Sub

Sub TakeItemFromInvAndPutOnGround(dbIndex As Long, ByVal iItemID As Long)
Dim tArr() As String
Dim dbMapIndex As Long
Dim i As Long
Dim s As String
Dim m As Long
With dbPlayers(dbIndex)
    dbMapIndex = .lDBLocation
    SplitFast Left$(.sInventory, Len(.sInventory) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            s = tArr(i)
            tArr(i) = ""
            JoinInventory dbIndex, tArr
            Exit For
        End If
        If DE Then DoEvents
    Next
    If dbMap(dbMapIndex).sItems = "0" Then dbMap(dbMapIndex).sItems = ""
    dbMap(dbMapIndex).sItems = dbMap(dbMapIndex).sItems & s & ";"
End With
End Sub

Sub TakeItemFromInvAndHideIt(dbIndex As Long, ByVal iItemID As Long)
Dim tArr() As String
Dim dbMapIndex As Long
Dim i As Long
Dim s As String
Dim m As Long
With dbPlayers(dbIndex)
    dbMapIndex = .lDBLocation
    SplitFast Left$(.sInventory, Len(.sInventory) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            s = tArr(i)
            tArr(i) = ""
            JoinInventory dbIndex, tArr
            Exit For
        End If
        If DE Then DoEvents
    Next
    If dbMap(dbMapIndex).sHidden = "0" Then dbMap(dbMapIndex).sHidden = ""
    dbMap(dbMapIndex).sHidden = dbMap(dbMapIndex).sHidden & s & ";"
End With
End Sub

Sub JoinMapItems(dbMapIndex, tArr() As String, Optional bHidden As Boolean = False)
Dim i As Long
Dim s As String
For i = LBound(tArr) To UBound(tArr)
    If tArr(i) <> "" Then
        s = s & tArr(i) & ";"
    End If
    If DE Then DoEvents
Next
If Not bHidden Then
    With dbMap(dbMapIndex)
        .sItems = s
        If modSC.FastStringComp(.sItems, "") Or modSC.FastStringComp(.sItems, "0;") Then .sItems = "0"
    End With
Else
    With dbMap(dbMapIndex)
        .sHidden = s
        If modSC.FastStringComp(.sHidden, "") Or modSC.FastStringComp(.sHidden, "0;") Then .sHidden = "0"
    End With
End If
End Sub

Sub TakeItemFromGroundAndPutInInv(dbIndex As Long, ByVal iItemID As Long)
Dim tArr() As String
Dim dbMapIndex As Long
Dim i As Long
Dim m As Long
Dim s As String
With dbPlayers(dbIndex)
    dbMapIndex = .lDBLocation
End With
With dbMap(dbMapIndex)
    SplitFast Left$(.sItems, Len(.sItems) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            s = tArr(i)
            tArr(i) = ""
            JoinMapItems dbMapIndex, tArr
            Exit For
        End If
        If DE Then DoEvents
    Next
End With
With dbPlayers(dbIndex)
    If .sInventory = "0" Then .sInventory = ""
    .sInventory = .sInventory & s & ";"
End With
End Sub

Sub TakeHiddenItemAndPutInInv(dbIndex As Long, ByVal iItemID As Long)
Dim tArr() As String
Dim dbMapIndex As Long
Dim i As Long
Dim m As Long
Dim s As String
With dbPlayers(dbIndex)
    dbMapIndex = .lDBLocation
End With
With dbMap(dbMapIndex)
    SplitFast Left$(.sHidden, Len(.sHidden) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            s = tArr(i)
            tArr(i) = ""
            JoinMapItems dbMapIndex, tArr, True
            Exit For
        End If
        If DE Then DoEvents
    Next
End With
With dbPlayers(dbIndex)
    If .sInventory = "0" Then .sInventory = ""
    .sInventory = .sInventory & s & ";"
End With
End Sub

Sub TakeItemFromInvAndEqIt(dbIndex As Long, ByVal iItemID As Long, Optional DualWield As Boolean = False)
Dim tArr() As String
Dim i As Long
Dim m As Long
Dim s As String
Dim t As String
Dim bF As Boolean
Dim sWorn As String
Dim iOldItemID As Long
Dim iNewItemID As Long
With dbPlayers(dbIndex)
    SplitFast Left$(.sInventory, Len(.sInventory) - 1), tArr, ";"
    For i = LBound(tArr) To UBound(tArr)
        m = CLng(modItemManip.GetItemIDFromUnFormattedString(tArr(i)))
        If iItemID = m Then
            s = tArr(i)
            tArr(i) = ""
            JoinInventory dbIndex, tArr
            Exit For
        End If
        If DE Then DoEvents
    Next
    iItemID = GetItemID(, CLng(iItemID))
End With
With dbItems(iItemID)
    sWorn = .sWorn
End With
With dbPlayers(dbIndex)
    Select Case sWorn
        Case "arms"
            If modSC.FastStringComp(.sArms, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sArms = s
            Else
                
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sArms))
                t = .sArms
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sArms = s
            End If
        Case "back"
            If modSC.FastStringComp(.sBack, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sBack = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sBack))
                t = .sBack
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sBack = s
            End If
        Case "body"
            If modSC.FastStringComp(.sBody, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sBody = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sBody))
                t = .sBody
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sBody = s
            End If
        Case "ears"
            If modSC.FastStringComp(.sEars, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sEars = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sEars))
                t = .sEars
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sEars = s
            End If
        Case "face"
            If modSC.FastStringComp(.sFace, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sFace = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sFace))
                t = .sFace
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sFace = s
            End If
        Case "hands"
            If modSC.FastStringComp(.sHands, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sHands = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sHands))
                t = .sHands
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sHands = s
            End If
        Case "head"
            If modSC.FastStringComp(.sHead, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sHead = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sHead))
                t = .sHead
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sHead = s
            End If
            
        Case "legs"
            If modSC.FastStringComp(.sLegs, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sLegs = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sLegs))
                t = .sLegs
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sLegs = s
            End If
        Case "neck"
            If modSC.FastStringComp(.sNeck, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sNeck = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sNeck))
                t = .sNeck
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sNeck = s
            End If
        Case "shield"
            If modSC.FastStringComp(.sShield, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sShield = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sShield))
                t = .sShield
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sShield = s
            End If
        Case "waist"
            If modSC.FastStringComp(.sWaist, "0") Then
                modItemManip.AdjustStats dbIndex, 0, iItemID
                .sWaist = s
            Else
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sWaist))
                t = .sWaist
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iItemID, True
                .sWaist = s
            End If
        Case "weapon"
            If Not DualWield Then
                If modSC.FastStringComp(.sWeapon, "0") Then
                    modItemManip.AdjustStats dbIndex, 0, iItemID
                    .sWeapon = s
                Else
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sWeapon))
                    t = .sWeapon
                    modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, iOldItemID, True, DualWield
                    .sWeapon = s
                End If
            Else
                If modSC.FastStringComp(.sShield, "0") Then
                    modItemManip.AdjustStats dbIndex, 0, iItemID
                    .sShield = s
                Else
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sShield))
                    t = .sShield
                    modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                    modItemManip.TakeEqItemAndPlaceInInv dbIndex, iOldItemID, True
                    .sShield = s
                End If
            End If
        Case "ring"
            bF = False
            For i = 0 To 5
                If modSC.FastStringComp(.sRings(i), "0") Then
                    Select Case i
                        Case 0
                            If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 0]) = 1 Then GoTo nNextI
                        Case 1
                            If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 1]) = 1 Then GoTo nNextI
                        Case 2
                            If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 2]) = 1 Then GoTo nNextI
                        Case 3
                            If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 3]) = 1 Then GoTo nNextI
                        Case 4
                            If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 4]) = 1 Then GoTo nNextI
                        Case 5
                            If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ring 5]) = 1 Then GoTo nNextI
                    End Select
                    modItemManip.AdjustStats dbIndex, 0, iItemID
                    .sRings(i) = s
                    bF = True
                    Exit For
                End If
nNextI:
                If DE Then DoEvents
            Next
            If Not bF Then
                i = RndNumber(0, 5)
                iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(.sRings(i)))
                t = .sRings(i)
                modItemManip.AdjustStats dbIndex, iOldItemID, iItemID
                modItemManip.TakeEqItemAndPlaceInInv dbIndex, iOldItemID, True
                .sRings(i) = s
            End If
    End Select
    If iItemID <> 0 Then
        iNewItemID = iItemID
        If modItemManip.GetItemFlagsFromUnFormattedString(s) <> "" Then modUseItems.DoFlags dbIndex, modItemManip.GetItemFlagsFromUnFormattedString(s), "|"
        If dbItems(iNewItemID).iOnEquipKillDur <> 0 Then
            If modSC.FastStringComp(.sKillDurItems, "0") Then .sKillDurItems = ""
            .sKillDurItems = .sKillDurItems & dbItems(iNewItemID).sWorn & "/" & dbItems(iNewItemID).iOnEquipKillDur & ";"
        End If
    End If
    If iOldItemID <> 0 Then
        iOldItemID = GetItemID(, iOldItemID)
        If iOldItemID = 0 Then Exit Sub
        If modItemManip.GetItemFlagsFromUnFormattedString(t) <> "" Then modUseItems.DoFlags dbIndex, modItemManip.GetItemFlagsFromUnFormattedString(t), "|", Inverse:=True
        If dbItems(iOldItemID).iOnEquipKillDur <> 0 Then
            .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(iOldItemID).sWorn & "/" & dbItems(iOldItemID).iOnEquipKillDur & ";", "", 1, 1)
            If modSC.FastStringComp(.sKillDurItems, "") Then .sKillDurItems = "0"
        End If
    End If
End With
End Sub

Sub TakeEqItemAndPlaceInInv(dbIndex As Long, ByVal iItemID As Long, Optional AlreadyAdjusted As Boolean = False, Optional CheckDual As Boolean = True)
Dim t As String
Dim sWorn As String
Dim i As Long
Dim iOldItemID As Long
iItemID = GetItemID(, CLng(iItemID))
If iItemID = 0 Then Exit Sub
With dbItems(iItemID)
    sWorn = .sWorn
End With
With dbPlayers(dbIndex)
    Select Case sWorn
        Case "arms"
            If Not modSC.FastStringComp(.sArms, "0") Then
                t = .sArms
                .sArms = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "back"
            If Not modSC.FastStringComp(.sBack, "0") Then
                t = .sBack
                .sBack = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "body"
            If Not modSC.FastStringComp(.sBody, "0") Then
                t = .sBody
                .sBody = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "ears"
            If Not modSC.FastStringComp(.sEars, "0") Then
                t = .sEars
                .sEars = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "face"
            If Not modSC.FastStringComp(.sFace, "0") Then
                t = .sFace
                .sFace = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "hands"
            If Not modSC.FastStringComp(.sHands, "0") Then
                t = .sHands
                .sHands = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "head"
            If Not modSC.FastStringComp(.sHead, "0") Then
                t = .sHead
                .sHead = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "legs"
            If Not modSC.FastStringComp(.sLegs, "0") Then
                If Not AlreadyAdjusted Then
                    t = .sLegs
                    .sLegs = "0"
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
               If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "neck"
            If Not modSC.FastStringComp(.sNeck, "0") Then
                t = .sNeck
                .sNeck = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "shield"
            If Not modSC.FastStringComp(.sShield, "0") Then
                t = .sShield
                .sShield = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "waist"
            If Not modSC.FastStringComp(.sWaist, "0") Then
                t = .sWaist
                .sWaist = "0"
                If Not AlreadyAdjusted Then
                    iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                    modItemManip.AdjustStats dbIndex, iOldItemID, 0
                End If
                If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                .sInventory = .sInventory & t & ";"
            End If
        Case "weapon"
            If Not modSC.FastStringComp(.sWeapon, "0") Then
                If modItemManip.GetItemIDFromUnFormattedString(.sWeapon) = dbItems(iItemID).iID Then
                    t = .sWeapon
                    .sWeapon = "0"
                    If Not AlreadyAdjusted Then
                        iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                        modItemManip.AdjustStats dbIndex, iOldItemID, 0
                    End If
                    If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                    .sInventory = .sInventory & t & ";"
                    If .iDualWield = 1 And CheckDual Then
                        .sWeapon = .sShield
                        .sShield = "0"
                        .iDualWield = 0
                    End If
                End If
            End If
            If .iDualWield = 1 And CheckDual Then
                If Not modSC.FastStringComp(.sShield, "0") Then
                    t = .sShield
                    .sShield = "0"
                    .iDualWield = 0
                    If Not AlreadyAdjusted Then
                        iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                        modItemManip.AdjustStats dbIndex, iOldItemID, 0
                    End If
                    If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                    .sInventory = .sInventory & t & ";"
                End If
            End If
        Case "ring"
            For i = 0 To 5
                If modItemManip.GetItemIDFromUnFormattedString(.sRings(i)) = dbItems(iItemID).iID Then
                    t = .sRings(i)
                    .sRings(i) = "0"
                    If Not AlreadyAdjusted Then
                        iOldItemID = CLng(modItemManip.GetItemIDFromUnFormattedString(t))
                        modItemManip.AdjustStats dbIndex, iOldItemID, 0
                    End If
                    If modSC.FastStringComp(.sInventory, "0") Then .sInventory = ""
                    .sInventory = .sInventory & t & ";"
                    Exit For
                End If
                If DE Then DoEvents
            Next
    End Select
    If Not AlreadyAdjusted Then
        If iItemID <> 0 Then
            If modItemManip.GetItemFlagsFromUnFormattedString(t) <> "" Then modUseItems.DoFlags dbIndex, modItemManip.GetItemFlagsFromUnFormattedString(t), "|", Inverse:=True
            If dbItems(iItemID).iOnEquipKillDur <> 0 Then
                .sKillDurItems = ReplaceFast(.sKillDurItems, dbItems(iItemID).sWorn & "/" & dbItems(iItemID).iOnEquipKillDur & ";", "", 1, 1)
                If modSC.FastStringComp(.sKillDurItems, "") Then .sKillDurItems = "0"
            End If
        End If
    End If
End With
End Sub

Sub AdjustStats(dbIndex As Long, ByVal oldItemID As Long, ByVal newItemID As Long)
If oldItemID <> 0 Then modUseItems.DoFlags dbIndex, dbItems(GetItemID(, oldItemID)).sFlags, Inverse:=True
If newItemID <> 0 Then modUseItems.DoFlags dbIndex, dbItems(newItemID).sFlags
End Sub

Sub SendBrokenItemMessage(dbIndex As Long, ByVal iPos As Long, ByVal iItemID As Long)
Dim sItemName As String
Dim pIndex As Long
Dim sWorn As String
Dim s As String
Dim dbItemID As Long
dbItemID = GetItemID(, iItemID)
sItemName = dbItems(dbItemID).sItemName
pIndex = dbPlayers(dbIndex).iIndex
If dbItems(dbItemID).iInGame > 0 Then dbItems(dbItemID).iInGame = dbItems(dbItemID).iInGame - 1
sWorn = TrimIt(ReplaceFast(modGetData.GetWornLocation(iPos), "(", ""))
sWorn = ReplaceFast(sWorn, ")", "")
s = BGRED & "Your " & sItemName & " breaks apart and falls off your " & sWorn & "." & WHITE & vbCrLf
WrapAndSend pIndex, s
SendToAllInRoom pIndex, LIGHTBLUE & dbPlayers(dbIndex).sPlayerName & "'s " & sItemName & " crumbles apart!" & WHITE & vbCrLf, dbPlayers(dbIndex).lLocation
End Sub
