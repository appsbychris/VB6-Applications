VERSION 5.00
Begin VB.UserControl FlagOptions 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox cboCombo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox txtString 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
   End
   Begin ServerEditor.NumOnlyText txtValue 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      AllowNeg        =   -1  'True
      Align           =   0
      MaxLength       =   5
      Enabled         =   -1  'True
      Backcolor       =   -2147483643
   End
End
Attribute VB_Name = "FlagOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum ViewStyles
    NumericInput = 0
    StringInput = 1
    ComboInputBoolean = 2
    ComboInputFeed = 3
End Enum

Private lCurStyle As ViewStyles
Private sFed As String
Private sVals As String

Public Property Get ViewStyle() As ViewStyles
ViewStyle = lCurStyle
End Property

Public Property Let ViewStyle(ByVal v As ViewStyles)
lCurStyle = v
DoChange
PropertyChanged "Style"
End Property

Private Sub cboCombo_Change()
If lCurStyle = ComboInputFeed Then txtValue.Text = cboCombo.ItemData(cboCombo.ListIndex)
End Sub

Private Sub cboCombo_Click()
cboCombo_Change
End Sub

Private Sub txtValue_Change()
Dim i As Long
If lCurStyle = ComboInputFeed Then
    For i = 0 To cboCombo.ListCount - 1
        If cboCombo.ItemData(i) = Val(txtValue.Text) Then
            cboCombo.ListIndex = i
            Exit For
        End If
    Next
End If
End Sub

Private Sub UserControl_Initialize()
DoChange
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    lCurStyle = .ReadProperty("Style")
End With
DoChange
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Style", lCurStyle
End With
End Sub

Private Sub DoChange()
Dim i As Long
Dim Arr() As String
Dim Arr2() As String
Select Case lCurStyle
    Case 0
        With txtValue
            .Top = 0
            .Left = 0
            .Visible = True
        End With
        cboCombo.Visible = False
        txtString.Visible = False
        With UserControl
            .Width = txtValue.Width
            .Height = txtValue.Height
        End With
    Case 1
        With txtString
            .Top = 0
            .Left = 0
            .Visible = True
        End With
        txtValue.Visible = False
        cboCombo.Visible = False
        With UserControl
            .Width = txtString.Width
            .Height = txtString.Height
        End With
    Case 2
        With cboCombo
            .Clear
            .AddItem "TRUE"
            .ItemData(0) = 1
            .AddItem "FALSE"
            .ItemData(1) = 0
            .Top = 0
            .Left = 0
            .ListIndex = 0
            .Visible = True
        End With
        txtValue.Visible = False
        txtString.Visible = False
        With UserControl
            .Width = cboCombo.Width
            .Height = cboCombo.Height
        End With
    Case 3
        Arr = Split(sFed, vbCrLf)
        Arr2 = Split(sVals, vbCrLf)
        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> "" Then
                cboCombo.AddItem Arr(i)
                cboCombo.ItemData(cboCombo.NewIndex) = CLng(Val(Arr2(i)))
            End If
        Next
        With txtValue
            .Top = 0
            .Left = 0
            .Visible = True
        End With
        With cboCombo
            .Top = 0
            .Left = txtValue.Width + 20
            If .ListCount > 0 Then .ListIndex = 0
            .Visible = True
        End With
        txtString.Visible = False
        With UserControl
            .Width = cboCombo.Width + txtValue.Width
            .Height = cboCombo.Height
        End With
End Select
End Sub

Public Sub FeedMe(AString As String, AndAValue As Long)
sFed = sFed & AString & vbCrLf
sVals = sVals & CStr(AndAValue) & vbCrLf
End Sub

Public Sub ClearFeed()
sFed = ""
sVals = ""
cboCombo.Clear
End Sub

Public Sub FillNow()
DoChange
End Sub

Public Sub SetVal(lVal As Long)
On Error Resume Next
Select Case lCurStyle
    Case 0, 3
        txtValue.Text = CStr(lVal)
    Case 2
        cboCombo.ListIndex = lVal
End Select
End Sub

Public Sub SetStr(sStr As String)
If lCurStyle = StringInput Then txtString.Text = sStr
End Sub

Public Function GetCurVal() As Long
Select Case lCurStyle
    Case 0, 3
        GetCurVal = Val(txtValue.Text)
    Case 2
        GetCurVal = cboCombo.ItemData(cboCombo.ListIndex)
    Case Else
        GetCurVal = -1
End Select
End Function

Public Function GetCurStr() As String
If lCurStyle = StringInput Then
    GetCurStr = txtString.Text
Else
    GetCurStr = "-1"
End If
End Function
