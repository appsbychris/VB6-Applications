VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script IMAGE editor"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "&Add New"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox txtKeyWord 
      Height          =   285
      Left            =   6240
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   6240
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox txtMethod 
      Height          =   285
      Left            =   6240
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.ListBox lstMethods 
      Height          =   7275
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox lstKeyWords 
      Height          =   7275
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "KeyWord"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Desc"
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Method"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tpWords
    sKeyWord As String
    sMethod As String
    sDesc As String
End Type
Private udtWrds() As tpWords
Private SelMeth As Long

Public Sub SaveWrdsFile(sFile As String)
Dim tmpArr() As tpWords
Dim i As Long
Dim s As String
Dim f As String
Redo:
ReDim tmpArr(0)
For i = LBound(udtWrds) To UBound(udtWrds)
    With udtWrds(i)
        If .sKeyWord <> "" Then
            If s = "" Then s = .sKeyWord
            If s = .sKeyWord Then
                ReDim Preserve tmpArr(UBound(tmpArr) + 1)
                tmpArr(UBound(tmpArr)).sKeyWord = s
                tmpArr(UBound(tmpArr)).sDesc = .sDesc
                tmpArr(UBound(tmpArr)).sMethod = .sMethod
                .sKeyWord = ""
            End If
        End If
    End With
Next
If UBound(tmpArr) > 0 Then
    f = f & "[Start Method List=" & tmpArr(1).sKeyWord & "]" & vbCrLf
    For i = LBound(tmpArr) + 1 To UBound(tmpArr)
        f = f & "[" & tmpArr(i).sMethod & "," & tmpArr(i).sDesc & "]" & vbCrLf
    Next
    f = f & "[End Method List=" & tmpArr(1).sKeyWord & "]" & vbCrLf
    s = ""
    GoTo Redo
Else
    f = Left$(f, Len(f) - 2)
    Open sFile For Output As #1
        Print #1, f
    Close #1
End If
LoadWrdsFile sFile
End Sub

Public Sub LoadWrdsFile(sFile As String)
Dim s As String
Dim i As Long
Dim m As Long
Dim n As Long
Dim Arr() As String
Dim Arr2() As String
Dim sKeyWord As String
Dim sMethods As String
Dim sTrigger As String
Open sFile For Binary As #1
    s = Input$(LOF(1), 1)
Close #1
Do While Len(s) > 0
    m = InStr(1, s, "[Start Method List=")
    If m = 0 Then Exit Do
    n = InStr(m, s, "=")
    m = InStr(n, s, "]")
    sTrigger = Mid$(s, n + 1, m - n - 1)
    n = InStr(m, s, "[End Method List=" & sTrigger & "]")
    Arr = Split(Mid$(s, m + 1, n - m - 1), vbCrLf)
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) <> "" Then
            Arr(i) = Replace$(Arr(i), "[", "")
            Arr(i) = Replace$(Arr(i), "]", "")
            Arr(i) = Replace$(Arr(i), vbCrLf, "")
            Arr2 = Split(Arr(i), ",", 2)
            ReDim Preserve udtWrds(UBound(udtWrds) + 1)
            With udtWrds(UBound(udtWrds))
                .sDesc = Arr2(1)
                .sKeyWord = sTrigger
                .sMethod = Arr2(0)
            End With
        End If
        DoEvents
    Next
    s = Mid$(s, n + Len("[End Method List=" & sTrigger & "]"))
    DoEvents
Loop
End Sub

Private Sub cmdAddNew_Click()
ReDim Preserve udtWrds(UBound(udtWrds) + 1)
With udtWrds(UBound(udtWrds))
    .sKeyWord = txtKeyWord.Text
    .sMethod = txtMethod.Text
    .sDesc = txtDesc.Text
End With
lstKeyWords.Clear
lstMethods.Clear
For i = LBound(udtWrds) + 1 To UBound(udtWrds)
    b = False
    For j = 0 To lstKeyWords.ListCount - 1
        If lstKeyWords.List(j) = udtWrds(i).sKeyWord Then
            b = True
            Exit For
        End If
    Next
    If Not b Then
        lstKeyWords.AddItem udtWrds(i).sKeyWord
    End If
Next
txtKeyWord.Text = ""
txtMethod.Text = ""
txtDesc.Text = ""
txtKeyWord.SetFocus
End Sub

Private Sub cmdEdit_Click()
With udtWrds(SelMeth)
    .sKeyWord = txtKeyWord.Text
    .sDesc = txtDesc.Text
    .sMethod = txtMethod.Text
End With
lstKeyWords.Clear
lstMethods.Clear
For i = LBound(udtWrds) + 1 To UBound(udtWrds)
    b = False
    For j = 0 To lstKeyWords.ListCount - 1
        If lstKeyWords.List(j) = udtWrds(i).sKeyWord Then
            b = True
            Exit For
        End If
    Next
    If Not b Then
        lstKeyWords.AddItem udtWrds(i).sKeyWord
    End If
Next
txtKeyWord.Text = ""
txtMethod.Text = ""
txtDesc.Text = ""
End Sub

Private Sub cmdSave_Click()
SaveWrdsFile App.Path & "\scriptdef.aimg"
End Sub

Private Sub Form_Load()
ReDim udtWrds(0)
Dim i As Long
Dim j As Long
Dim b As Boolean
LoadWrdsFile App.Path & "\scriptdef.aimg"
For i = LBound(udtWrds) + 1 To UBound(udtWrds)
    b = False
    For j = 0 To lstKeyWords.ListCount - 1
        If lstKeyWords.List(j) = udtWrds(i).sKeyWord Then
            b = True
            Exit For
        End If
    Next
    If Not b Then
        lstKeyWords.AddItem udtWrds(i).sKeyWord
    End If
Next
End Sub

Private Sub lstKeyWords_Click()
Dim i As Long
lstMethods.Clear
For i = LBound(udtWrds) + 1 To UBound(udtWrds)
    If udtWrds(i).sKeyWord = lstKeyWords.List(lstKeyWords.ListIndex) Then
        lstMethods.AddItem udtWrds(i).sMethod
    End If
Next
txtKeyWord.Text = ""
txtMethod.Text = ""
txtDesc.Text = ""
End Sub

Private Sub lstMethods_Click()
Dim i As Long
For i = LBound(udtWrds) + 1 To UBound(udtWrds)
    If udtWrds(i).sKeyWord = lstKeyWords.List(lstKeyWords.ListIndex) Then
        If udtWrds(i).sMethod = lstMethods.List(lstMethods.ListIndex) Then
            With udtWrds(i)
                txtKeyWord.Text = .sKeyWord
                txtMethod.Text = .sMethod
                txtDesc.Text = .sDesc
            End With
            SelMeth = i
            Exit For
        End If
    End If
Next
End Sub
