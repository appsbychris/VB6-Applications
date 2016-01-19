Attribute VB_Name = "modFast"
Public Type SafeArray1D
  cDims       As Integer
  fFeatures   As Integer
  cbElements  As Long
  cLocks      As Long
  pvData      As Long
  cElements   As Long
  lLBound     As Long
End Type

Public Const FADF_AUTO As Long = &H1        '// Array is allocated on the stack.
Public Const FADF_FIXEDSIZE As Long = &H10  '// Array may not be resized or reallocated.


' ==============================================================================
' "RtlMoveMemory" and popular synonyms
Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
    dst As Any, _
    src As Any, _
    ByVal nBytes As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    dst As Any, _
    src As Any, _
    ByVal nBytes As Long)
' (re-)typed aliases
Private Declare Sub CopyMemLng Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    ByVal dst As Long, _
    ByVal src As Long, _
    ByVal nBytes As Long)
Private Declare Sub PokeInt Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Addr As Long, _
    Value As Long, _
    Optional ByVal nBytes As Long = 2)
Private Declare Sub PokeLng Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Addr As Long, _
    Value As Long, _
    Optional ByVal nBytes As Long = 4)

' ==============================================================================
' "RtlZeroMemory"
Private Declare Sub RtlZeroMemory Lib "kernel32" ( _
    dst As Any, _
    ByVal nBytes As Long)

' ==============================================================================
' "RtlFillMemory"
Private Declare Sub RtlFillMemory Lib "kernel32" ( _
    dst As Any, _
    ByVal nBytes As Long, _
    ByVal bFill As Byte)

' ==============================================================================
' "VarPtr" and popular synonyms

Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ptr() As Any) As Long     '<-- VB6
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ptr() As Any) As Long     '<-- VB6

' ==============================================================================
' "SysAllocStringByteLen"
Private Declare Function SysAllocStringByteLen Lib "oleaut32" ( _
    ByVal olestr As Long, _
    ByVal BLen As Long) As Long


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
        ResultSplit(0) = Left$(Expression, Results(0) - 1)
        For c = 0 To c - 2
            ResultSplit(c + 1) = Mid$(Expression, _
                Results(c) + DelLen, _
                Results(c + 1) - Results(c) - DelLen)
        Next c
        ResultSplit(c + 1) = Right$(Expression, SLen - Results(c) - DelLen + 1)
    End If

End Sub

Public Function ReplaceFast(ByRef Text As String, _
    ByRef sOld As String, ByRef sNew As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Count As Long = 2147483647, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
  ) As String
' by Jost Schwider, jost@schwider.de, 20001218

  If LenB(sOld) Then

    If Compare = vbBinaryCompare Then
      Replace09Bin ReplaceFast, Text, Text, _
          sOld, sNew, Start, Count
    Else
      Replace09Bin ReplaceFast, Text, LCaseFast(Text), _
          LCaseFast(sOld), sNew, Start, Count
    End If

  Else 'Suchstring ist leer:
    ReplaceFast = Text
  End If
End Function

Private Static Sub Replace09Bin(ByRef result As String, _
    ByRef Text As String, ByRef Search As String, _
    ByRef sOld As String, ByRef sNew As String, _
    ByVal Start As Long, ByVal Count As Long _
  )
' by Jost Schwider, jost@schwider.de, 20001218
  Dim TextLen As Long
  Dim OldLen As Long
  Dim NewLen As Long
  Dim ReadPos As Long
  Dim WritePos As Long
  Dim CopyLen As Long
  Dim Buffer As String
  Dim BufferLen As Long
  Dim BufferPosNew As Long
  Dim BufferPosNext As Long
  
  'Ersten Treffer bestimmen:
  If Start < 2 Then
    Start = InStrB(Search, sOld)
  Else
    Start = InStrB(Start + Start - 1, Search, sOld)
  End If
  If Start Then
  
    OldLen = LenB(sOld)
    NewLen = LenB(sNew)
    Select Case NewLen
    Case OldLen 'einfaches Überschreiben:
    
      result = Text
      For Count = 1 To Count
        MidB$(result, Start) = sNew
        Start = InStrB(Start + OldLen, Search, sOld)
        If Start = 0 Then Exit Sub
      Next Count
      Exit Sub
    
    Case Is < OldLen 'Ergebnis wird kürzer:
    
      'Buffer initialisieren:
      TextLen = LenB(Text)
      If TextLen > BufferLen Then
        Buffer = Text
        BufferLen = TextLen
      End If
      
      'Ersetzen:
      ReadPos = 1
      WritePos = 1
      If NewLen Then
      
        'Einzufügenden Text beachten:
        For Count = 1 To Count
          CopyLen = Start - ReadPos
          If CopyLen Then
            BufferPosNew = WritePos + CopyLen
            MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
            MidB$(Buffer, BufferPosNew) = sNew
            WritePos = BufferPosNew + NewLen
          Else
            MidB$(Buffer, WritePos) = sNew
            WritePos = WritePos + NewLen
          End If
          ReadPos = Start + OldLen
          Start = InStrB(ReadPos, Search, sOld)
          If Start = 0 Then Exit For
        Next Count
      
      Else
      
        'Einzufügenden Text ignorieren (weil leer):
        For Count = 1 To Count
          CopyLen = Start - ReadPos
          If CopyLen Then
            MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
            WritePos = WritePos + CopyLen
          End If
          ReadPos = Start + OldLen
          Start = InStrB(ReadPos, Search, sOld)
          If Start = 0 Then Exit For
        Next Count
      
      End If
      
      'Ergebnis zusammenbauen:
      If ReadPos > TextLen Then
        result = LeftB$(Buffer, WritePos - 1)
      Else
        MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
        result = LeftB$(Buffer, WritePos + LenB(Text) - ReadPos)
      End If
      Exit Sub
    
    Case Else 'Ergebnis wird länger:
    
      'Buffer initialisieren:
      TextLen = LenB(Text)
      BufferPosNew = TextLen + NewLen
      If BufferPosNew > BufferLen Then
        Buffer = Space$(BufferPosNew)
        BufferLen = LenB(Buffer)
      End If
      
      'Ersetzung:
      ReadPos = 1
      WritePos = 1
      For Count = 1 To Count
        CopyLen = Start - ReadPos
        If CopyLen Then
          'Positionen berechnen:
          BufferPosNew = WritePos + CopyLen
          BufferPosNext = BufferPosNew + NewLen
          
          'Ggf. Buffer vergrößern:
          If BufferPosNext > BufferLen Then
            Buffer = Buffer & Space$(BufferPosNext)
            BufferLen = LenB(Buffer)
          End If
          
          'String "patchen":
          MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
          MidB$(Buffer, BufferPosNew) = sNew
        Else
          'Position bestimmen:
          BufferPosNext = WritePos + NewLen
          
          'Ggf. Buffer vergrößern:
          If BufferPosNext > BufferLen Then
            Buffer = Buffer & Space$(BufferPosNext)
            BufferLen = LenB(Buffer)
          End If
          
          'String "patchen":
          MidB$(Buffer, WritePos) = sNew
        End If
        WritePos = BufferPosNext
        ReadPos = Start + OldLen
        Start = InStrB(ReadPos, Search, sOld)
        If Start = 0 Then Exit For
      Next Count
      
      'Ergebnis zusammenbauen:
      If ReadPos > TextLen Then
        result = LeftB$(Buffer, WritePos - 1)
      Else
        BufferPosNext = WritePos + TextLen - ReadPos
        If BufferPosNext < BufferLen Then
          MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
          result = LeftB$(Buffer, BufferPosNext)
        Else
          result = LeftB$(Buffer, WritePos - 1) & MidB$(Text, ReadPos)
        End If
      End If
      Exit Sub
    
    End Select
  
  Else 'Kein Treffer:
    result = Text
  End If
End Sub



Public Static Function RoundFast(dblNumber As Double, Optional ByVal numDecimalPlaces As Long) As Double
' by Donald, donald@xbeat.net, 20001018
  
  Dim fInit As Boolean
  Dim numDecimalPlacesPrev As Long
  Dim dFac As Double
  Dim dFacInv As Double
  Dim dTmp As Double
  
  ' calc factor once for this depth of rounding
  If Not fInit Or numDecimalPlacesPrev <> numDecimalPlaces Then
    dFac = 10 ^ numDecimalPlaces
    dFacInv = 10 ^ -numDecimalPlaces
    numDecimalPlacesPrev = numDecimalPlaces
    fInit = True
  End If
  
  If dblNumber >= 0 Then
    dTmp = dblNumber * dFac + 0.5
    RoundFast = Int(dTmp) * dFacInv
  Else
    dTmp = -dblNumber * dFac + 0.5
    RoundFast = -Int(dTmp) * dFacInv
  End If
  
End Function


Public Function LCaseFast(ByRef sString As String) As String
    Static saDst As SafeArray1D
    Static aDst%()
    Static pDst&, psaDst&
    Static init As Long
    Dim c As Long
    Dim lLen As Long
    Static iLUT(0 To 400) As Integer
    
    If init Then
    Else
        saDst.cDims = 1
        saDst.cbElements = 2
        saDst.cElements = &H7FFFFFFF
        
        pDst = VarPtr(saDst)
        psaDst = ArrPtr(aDst)
        
        ' init LUT
        For c = 0 To 255: iLUT(c) = AscW(LCase$(Chr$(c))): Next
        For c = 256 To 400: iLUT(c) = c: Next
        iLUT(352) = 353
        iLUT(338) = 339
        iLUT(381) = 382
        iLUT(376) = 255
        
        init = 1
    End If
    
    lLen = Len(sString)
    RtlMoveMemory ByVal VarPtr(LCaseFast), _
        SysAllocStringByteLen(StrPtr(sString), lLen + lLen), 4
    saDst.pvData = StrPtr(LCaseFast)
    RtlMoveMemory ByVal psaDst, pDst, 4
    
    For c = 0 To lLen - 1
      Select Case aDst(c)
      Case 65 To 381
        aDst(c) = iLUT(aDst(c))
      End Select
    Next
    
    RtlMoveMemory ByVal psaDst, 0&, 4
    
End Function

Public Function TrimIt(s As String) As String
s = RTrim$(s)
s = LTrim$(s)
TrimIt = s
End Function




