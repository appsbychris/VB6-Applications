Attribute VB_Name = "Module1"



Public Sub SplitFast(Expression$, ResultSplit$(), Optional Delimiter$ = " ")
' By Chris Lucas, cdl1051@earthlink.net, 20011208
    Dim c&, SLen&, DelLen&, tmp&, Results&()

    SLen = LenB(Expression) \ 2
    DelLen = LenB(Delimiter) \ 2

    ' Bail if we were passed an empty delimiter or an empty expression
    If SLen = 0 Or DelLen = 0 Then
        ReDim Preserve ResultSplit(0 To 0)
        ResultSplit(0) = Expression
        Exit Sub
    End If

    ' Count delimiters and remember their positions
    ReDim Preserve Results(0 To SLen)
    tmp = InStr(Expression, Delimiter)

    Do While tmp
        Results(c) = tmp
        c = c + 1
        tmp = InStr(Results(c - 1) + 1, Expression, Delimiter)
    Loop

    ' Size our return array
    ReDim Preserve ResultSplit(0 To c)

    ' Populate the array
    If c = 0 Then
        ' lazy man's call
        ResultSplit(0) = Expression
    Else
        ' typical call
        ResultSplit(0) = VBA.Left$(Expression, Results(0) - 1)
        For c = 0 To c - 2
            ResultSplit(c + 1) = VBA.Mid$(Expression, _
                Results(c) + DelLen, _
                Results(c + 1) - Results(c) - DelLen)
        Next c
        ResultSplit(c + 1) = VBA.Right$(Expression, SLen - Results(c) - DelLen + 1)
    End If

End Sub


