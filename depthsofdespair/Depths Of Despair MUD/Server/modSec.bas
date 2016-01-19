Attribute VB_Name = "modSec"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modSec
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Private Const v1574621 = "504B�-3��57Z�1qs�l*�8�H����-�,quk��54���6@f�"
Private Const v3847282 = "C�hOfrN��iR�~�s���:� �c����Vy[�u�SVa�iTH~6�Jnwz4�c�)(� _�sm[|�`��Hj\(��bf}�V�oe�qU����Z�}Ro������bveb;��s=<�XN6q����+`oe���2��(���4-|?nr�-�8{�Lf��`�?Xry"
Private Const v7379273 = "3h3H.4R1E4w�J-E��L]7-�0|�Zj<~CapiRtH����,rR5N:Q�RA2(�-3Dd���+*�OgtA�W�a,J}�M4*<2b\V~^v�A03�nb|k+2[�=>>Wv�7�Zs06�g��a*�-c��I�.R��]F�)JEkH\�k������<�T�(>4g~�snd�9F�/�r��Qx�T)��@GB�(Pt����;�6aR�|WL��>cs��J��5��"
Private Const v2872744 = "6S48�9�t�-y6ƨ3{���D4�D��|�7P{yb�E�-@��y9���6-µ�sA-�69�Z}-�����8��AWƩ5�\@I-?��c�2oWq��]6Q��v{;k�-lL[R8���8�/Q��P�-J�6N��Q?W\�>m1��"
Private Const WndK As String = "�W�?�Ͱ�����e��7�9P��Ǒ�0?��Z�0ƅz��7m�-��G˿X\ES�0=X�Q�ypFOoA���W@fEu��y����O�F7e������1�.�����C����:cɱ.PWYU+��u�iw�ǥ�WC�[y���i�}���`@��WEfd.��X�����?�����/Y�R�f��sRoʺlZ��2���v�_�v�����C�OwDY0ZO�9h{����~�̉@e�LC�V9��+��ER�S�h͂�I���VL�k��XC.�������iwbY)����ǚBGc���8�n9��{kA0�dt���ZX`2��B*�*g���Q�*RZ?l�VPQ�j`�41�xzg�1"
Rem Encrptyed Value
Public Const sValue As String = "b�r�81��E75`�9meu_qm3P|���Ds�����D�1u-�?McK�"
Rem////////////////////////////
Public Function DSVal(sdValue As String) As String
Rem sub to decrypt value
Dim Temp As String
Temp = Mid$(sdValue, 2, 1)
Temp = Temp & Mid$(sdValue, 5, 1)
Temp = Temp & Mid$(sdValue, 9, 1)
Temp = Temp & Mid$(sdValue, 14, 1)
Temp = Temp & Mid$(sdValue, 20, 1)
Temp = Temp & Mid$(sdValue, 27, 1)
Temp = Temp & Mid$(sdValue, 35, 1)
Temp = Temp & Mid$(sdValue, 44, 1)
DSVal = Temp
End Function

Public Function uJunkIt(sArg As String) As String
Dim i As Long
Dim j As Long
Dim sBuild As String
For i = 1 To Len(sArg)
    j = j + 1
    sBuild = sBuild & Mid$(sArg, j, 1)
    j = j + i
Next
uJunkIt = sBuild
End Function

Private Function dWndProcVal1(Arg As String) As String
dWndProcVal1 = uJunkIt(Arg)
End Function

Private Function dWndProcVal2(Arg As String) As String
dWndProcVal2 = uJunkIt(Arg)
End Function

Private Function dWndProcVal3(Arg As String) As String
dWndProcVal3 = uJunkIt(Arg)
End Function

Private Function dWndProcVal4(Arg As String) As String
dWndProcVal4 = uJunkIt(Arg)
End Function

Public Function WndProcDeter(Arg As String) As Boolean
Dim sPN As String
Dim sUn As String
Dim sSN As String
Dim AscNums As String
Dim AllThree As String
Dim sCID As String
Dim i As Long
Dim rNum As Long
Dim sOutPut As String
Dim lLen As Long
Dim Num1 As Long
Dim Num2 As Long
Dim b As Boolean
'v1574621 = "0"
'Private Const v3847282 = "0"
'Private Const v7379273 = "0"
'Private Const v2872744 = "0"
sPN = dWndProcVal1(v1574621)
sUn = dWndProcVal2(v3847282)
sSN = dWndProcVal3(v7379273)
sCID = dWndProcVal4(v2872744)
sPN = Replace$(sPN, "-", "")
sUn = Replace$(sUn, " ", "")
sSN = Replace$(sSN, "-", "")
sCID = Replace$(sCID, "-", "")
AllThree = sPN & sUn & sSN
For i = 1 To Len(AllThree)
    AscNums = AscNums & Asc(Mid$(AllThree, i, 1))
Next
lLen = Len(AscNums) - 1
b = False
For i = 1 To 12
    rNum = Mid$(AscNums, lLen \ i, 1) + Mid$(sCID, Len(sCID) \ i, 1)
    If rNum > lLen Then rNum = Mid$(sCID, Len(sCID) \ i, 1)
    If rNum <= 0 Then rNum = Mid$(sCID, Len(sCID) \ i, 1)
    rNum = CLng(Mid$(AscNums, rNum, 1))
    If rNum <= 0 Then rNum = Mid$(AscNums, lLen \ i, 1) And Mid$(sCID, Len(sCID) \ i, 1)
    sOutPut = sOutPut & CLng(CLng(Mid$(AscNums, rNum + 1, 1) + CLng(IIf(b, 1, 2))))
    If rNum > 4 Then b = True Else b = False
Next
sOutPut = Mid$(sOutPut, Len(sPN), 2)
Num1 = 0
For i = 1 To Len(sPN)
    Num1 = Num1 + Asc(Mid$(sPN, i, 1))
Next
sOutPut = sOutPut & Num1 \ CLng(Mid$(sCID, IIf(b, 1, 2), 1))
Num1 = 0
b = Not b
For i = 1 To Len(sUn)
    Num1 = Num1 + Asc(Mid$(sUn, i, 1))
Next
sOutPut = sOutPut & Num1 \ CLng(Mid$(sCID, IIf(b, 3, 4), 1))
b = Not b
Num1 = 0
For i = 1 To Len(sSN)
    Num1 = Num1 + Asc(Mid$(sSN, i, 1))
Next
sOutPut = sOutPut & Num1 \ CLng(Mid$(sCID, IIf(b, 5, 7), 1))
Num1 = 0
b = Not b
For i = 1 To Len(sCID)
    Num1 = Num1 + Asc(Mid$(sCID, i, 1))
Next
sOutPut = sOutPut & Num1 \ CLng(Mid$(sCID, IIf(b, 9, 11), 1))
For i = 1 To Len(sOutPut)
    If (CLng(Mid$(sOutPut, i, 1)) < 9) And (CLng(Mid$(sOutPut, i, 1)) > 0) Then
        Mid$(sOutPut, i, 1) = (CLng(Mid$(sOutPut, i, 1)) + CLng(IIf(b, 1, -1)))
        b = Not b
    End If
Next
sOutPut = Mid$(sOutPut, 1, 3) & "-" & Mid$(sOutPut, 4, 3) & "-" & Mid$(sOutPut, 7, 3) & "-" & Mid$(sOutPut, 10)  'Format$(sOutput, "#### #### ### ##")
sOutPut = wndEnc(sOutPut, CDbl(sCID))
End Function

Private Function wndEnc(sArg As String, dKey As Double) As String
Dim i As Long
Dim rNums As String
Dim Temp As Long
Dim sOutPut As String
Dim v2 As Long
Temp = Len(sArg)
dKey = dKey * (1 / 3)
rNums = CStr(dKey)
For i = 1 To Temp
    v2 = Asc(Mid$(rNums, CLng(i Mod Len(rNums)) + 1, 1))
    If v2 < 71 Then v2 = v2 + (130 - (v2 + CLng(IIf(i And 1 = 1, 4, 7))))
    sOutPut = sOutPut & Chr$(Asc(Mid$(sArg, i, 1)) Xor v2)
Next
For j = 70 To 150 Step 10
    sOutPut = GetHexString(sOutPut)
    sArg = sOutPut
    sOutPut = ""
    For i = 1 To Temp
        v2 = Asc(Mid$(rNums, CLng(i Mod Len(rNums)) + 1, 1))
        If v2 < 71 Then v2 = v2 + (j - (v2 + CLng(IIf(i And 1 = 1, 4, 7))))
        sOutPut = sOutPut & Chr$(Asc(Mid$(sArg, i, 1)) Xor v2)
    Next
Next
sArg = sOutPut
sOutPut = ""
For i = 1 To 5
    sOutPut = sOutPut & Asc(Mid$(sArg, i, 1)) + CLng(IIf(i And 1 = 1, i, i + i))
Next
sArg = sOutPut
sOutPut = ""
For i = 1 To Temp
    v2 = Asc(Mid$(rNums, CLng(i Mod Len(rNums)) + 1, 1))
    If v2 < 71 Then v2 = v2 + (130 - (v2 + CLng(IIf(i And 1 = 1, 4, 7))))
    sOutPut = sOutPut & Chr$(Asc(Mid$(sArg, i, 1)) Xor v2)
Next
wndEnc = sOutPut
End Function

Public Function GetHexString(strText As String) As String
Dim i As Long
For i = 1 To Len(strText)
    GetHexString = GetHexString & Hex(Asc(Mid$(strText, i, 1)))
Next
End Function

Public Sub dB_set_and_load(ByVal sSearch As String, ByVal fn As String, ByVal sReplace As String)
Dim ff As Long
Dim s As String
Dim ptr As Long
Dim n As Long
Dim iLen As Long
Dim iFound As Long
Dim bDone As Boolean
Const BLOCKSIZE As Long = 500   'number of bytes to read at one time
iLen = Len(sSearch)  'length of string to find
If iLen <> Len(sReplace) Then Exit Sub
If iLen Then
    ff = FreeFile
    ptr = 0  'starting location
    Open fn For Binary As #ff
        n = LOF(ff)   'number of bytes in file
        If n Then
            Do
                If (n - ptr) > BLOCKSIZE Then
                    s = Space$(BLOCKSIZE)
                Else
                    s = Space$(n - ptr)
                    bDone = True    'will exit because end of file
                End If
                Get #ff, ptr + 1, s
                iFound = InStr(s, sSearch)
                If iFound Then
                    'FindInFile = ptr + iFound
                    Put #ff, iFound, sReplace
                    bDone = True   'will exit because a match has been found
                Else
                    ptr = ptr + BLOCKSIZE - iLen
                End If
            Loop While Not bDone
        End If
    Close #ff
End If
End Sub

Public Function GetASCII(Arg As String) As String
Dim i As Long
For i = 1 To Len(Arg)
    GetASCII = GetASCII & Asc(Mid$(Arg, i, 1))
    If DE Then DoEvents
Next
End Function

Public Function WndPWHash(Arg As String) As Boolean
Dim sPN As String
Dim sUn As String
Dim sSN As String
Dim AscNums As String
Dim AllThree As String
Dim sCID As String
Dim i As Long
Dim rNum As Long
Dim sOutPut As String
Dim lLen As Long
Dim Num1 As Long
Dim Num2 As Long
Dim b As Boolean
sPN = Arg 'dWndProcVal1(v1574621)
sUn = StrReverse(Arg) 'dWndProcVal2(v3847282)
sSN = StrReverse(Arg) & StrReverse(GetASCII(Arg)) & Arg
sCID = GetASCII(Arg)

AllThree = sPN & sUn & sSN
For i = 1 To Len(AllThree)
    AscNums = AscNums & Asc(Mid$(AllThree, i, 1))
    If DE Then DoEvents
Next
lLen = Len(AscNums) - 1
b = False
For i = 1 To 12
    rNum = Mid$(AscNums, lLen \ i, 1) + Mid$(sCID, Len(sCID) \ i, 1)
    If rNum > lLen Then rNum = Mid$(sCID, Len(sCID) \ i, 1)
    If rNum <= 0 Then rNum = Mid$(sCID, Len(sCID) \ i, 1)
    rNum = CLng(Mid$(AscNums, rNum, 1))
    If rNum <= 0 Then rNum = Mid$(AscNums, lLen \ i, 1) And Mid$(sCID, Len(sCID) \ i, 1)
    sOutPut = sOutPut & CLng(CLng(Mid$(AscNums, rNum + 1, 1) + CLng(IIf(b, 1, 2))))
    If rNum > 4 Then b = True Else b = False
    If DE Then DoEvents
Next
sOutPut = Mid$(sOutPut, Len(sPN), 2)
Num1 = 0
For i = 1 To Len(sPN)
    Num1 = Num1 + Asc(Mid$(sPN, i, 1))
    If DE Then DoEvents
Next
sOutPut = sOutPut & Num1 \ CLng(Mid$(sCID, IIf(b, 1, 2), 1))
Num1 = 0
b = Not b
For i = 1 To Len(sUn)
    Num1 = Num1 + Asc(Mid$(sUn, i, 1))
    If DE Then DoEvents
Next
sOutPut = sOutPut & Num1 \ CLng(Mid$(sCID, IIf(b, 3, 4), 1))
b = Not b
Num1 = 0
For i = 1 To Len(sSN)
    Num1 = Num1 + Asc(Mid$(sSN, i, 1))
    If DE Then DoEvents
Next
sOutPut = sOutPut & Num1 \ CLng(Mid$(sCID, IIf(b, 5, 7), 1))
Num1 = 0
b = Not b
For i = 1 To Len(sCID)
    Num1 = Num1 + Asc(Mid$(sCID, i, 1))
    If DE Then DoEvents
Next
sOutPut = sOutPut & Num1 \ CLng(Mid$(sCID, IIf(b, 9, 11), 1))
For i = 1 To Len(sOutPut)
    If (CLng(Mid$(sOutPut, i, 1)) < 9) And (CLng(Mid$(sOutPut, i, 1)) > 0) Then
        Mid$(sOutPut, i, 1) = (CLng(Mid$(sOutPut, i, 1)) + CLng(IIf(b, 1, -1)))
        b = Not b
    End If
    If DE Then DoEvents
Next
sOutPut = Mid$(sOutPut, 1, 3) & "-" & Mid$(sOutPut, 4, 3) & "-" & Mid$(sOutPut, 7, 3) & "-" & Mid$(sOutPut, 10)  'Format$(sOutput, "#### #### ### ##")
'sOutPut = wndEnc(sOutPut, CDbl(sCID))

End Function

Private Function GetWndVal(ByVal InText As String, Optional ByVal Password As String = "") As String

Dim i As Long, j As Long
Dim s As String
j = 0
If Len(Password) Then
    For i = 1 To Len(InText)
        s = s & Chr$(Asc(Mid$(InText, i, 1)) Xor Asc(Mid$(Password, j + 1, 1)))
        j = (j + 1) Mod Len(Password)
    Next i
Else
    For i = 1 To Len(InText)
        s = s & Chr$(&HFF Xor Asc(Mid$(InText, i, 1)))
    Next i
End If
GetWndVal = s
End Function


Public Function ByValGetBuildVal() As String
Dim i As Long
Dim j As Long
Dim ssBuild As String
Dim sArg As String
Open App.Path & "\build.bld" For Binary As #1
    sArg = Input$(LOF(1), 1)
Close #1
For i = 1 To Len(sArg)
    j = j + 1
    ssBuild = ssBuild & Mid$(sArg, j, 1)
    j = j + i
Next
ByValGetBuildVal = GetWndVal(ssBuild, GetWndVal(uJunkIt(WndK)))  ', "542354324535468735432774")
End Function
