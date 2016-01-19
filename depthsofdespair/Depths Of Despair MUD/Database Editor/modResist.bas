Attribute VB_Name = "modResist"
'fire/ice/water/lightning/earth/poison/wind
Public Enum GetResist
    Fire = 0
    Ice = 1
    Water = 2
    Lightning = 3
    Earth = 4
    Poison = 5
    Wind = 6
    Holy = 7
    Unholy = 8
End Enum

Public Function GetResistValue(dbIndex As Long, WhichOne As GetResist) As Long
Dim Arr() As String
With dbPlayers(dbIndex)
    SplitFast .sElements, Arr, "/"
End With
GetResistValue = CLng(Val(Arr(WhichOne)))
End Function

Public Sub UpdateResistValue(dbIndex As Long, WhichOne As GetResist, lValue As Long)
Dim Arr() As String
Dim i As Long
With dbPlayers(dbIndex)
    SplitFast .sElements, Arr, "/"
    Arr(WhichOne) = CLng(Val(Arr(WhichOne))) + lValue
    .sElements = ""
    For i = LBound(Arr) To UBound(Arr)
        .sElements = .sElements & Arr(i) & "/"
        If DE Then DoEvents
    Next
    .sElements = Left$(.sElements, Len(.sElements) - 1)
End With
End Sub

