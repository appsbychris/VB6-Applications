VERSION 5.00
Begin VB.UserControl NumOnlyText 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtMain 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "NumOnlyText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : NumOnlyText
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Dim ff As StdFont
Dim bB As Boolean
Public Event Change()
'Dim bBold As Boolean

'Public Enum FontSizes
'    vb8 = 8
'    vb10 = 10
'    vb12 = 12
'    vb14 = 14
'    vb16 = 16
'    vb18 = 18
'    vb20 = 20
'    vb28 = 28
'    vb36 = 36
'    vb48 = 48
'    vb72 = 72
'End Enum

Private Sub txtMain_Change()
Dim i As Long
Dim m As Long
Dim a As String
m = txtMain.SelStart
a = txtMain.Text
For i = 1 To Len(a)
    Select Case Asc(Mid$(a, i, 1))
        Case 48 To 57, 45, 46
        
        Case Else
            
            Mid$(a, i, 1) = " "
    End Select
    DoEvents
Next
If bB Then
    If DCount(a, "-") > 1 Then
        For i = 2 To Len(a)
            Select Case Asc(Mid$(a, i, 1))
                Case 45
                    Mid$(a, i, 1) = " "
            End Select
            DoEvents
        Next
    End If
End If
txtMain.Text = Replace$(a, " ", "")
If txtMain.Text = "" Then txtMain.Text = "0"
If Len(txtMain.Text) < m Then txtMain.SelStart = Len(txtMain.Text) Else txtMain.SelStart = m
RaiseEvent Change
End Sub

Private Sub txtMain_GotFocus()
txtMain.SelStart = 0
txtMain.SelLength = Len(txtMain.Text)
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)
CheckKeyAsciiForNumber KeyAscii, bB
End Sub

Private Sub UserControl_Initialize()
'With UserControl
'    .Width = txtMain.Width
'    .Height = txtMain.Height
'End With

If ff Is Nothing Then Set ff = txtMain.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
With PropBag
    Set ff = .ReadProperty("Font")
    Set txtMain.Font = ff
    txtMain.Text = .ReadProperty("Text")
    bB = .ReadProperty("AllowNeg")
    txtMain.Alignment = .ReadProperty("Align")
    txtMain.MaxLength = .ReadProperty("MaxLength")
    txtMain.Enabled = .ReadProperty("Enabled")
    txtMain.BackColor = .ReadProperty("Backcolor")
End With
End Sub

Private Sub UserControl_Resize()
With txtMain
    .Width = UserControl.Width
    .Height = UserControl.Height - 20
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
With PropBag
    .WriteProperty "Font", ff
    .WriteProperty "Text", txtMain.Text
    .WriteProperty "AllowNeg", bB
    .WriteProperty "Align", txtMain.Alignment
    .WriteProperty "MaxLength", txtMain.MaxLength
    .WriteProperty "Enabled", txtMain.Enabled
    .WriteProperty "Backcolor", txtMain.BackColor
End With
End Sub

Public Property Get Font() As StdFont
Set Font = ff

End Property

Public Property Set Font(ByVal f As StdFont)
Set ff = f
Set txtMain.Font = ff
PropertyChanged "Font"
End Property

Public Property Get Text() As String
Text = txtMain.Text
End Property

Public Property Let Text(ByVal s As String)
txtMain.Text = s
PropertyChanged "Text"
End Property

Public Property Get AllowNegative() As Boolean
AllowNegative = bB
End Property

Public Property Let AllowNegative(ByVal b As Boolean)
bB = b
PropertyChanged "AllowNeg"
End Property

'Public Property Get Bold() As Boolean
'Bold = bBold
'End Property
'
'Public Property Let Bold(ByVal b As Boolean)
'bBold = b
'txtMain.FontBold = bBold
'PropertyChanged "Bold"
'End Property

Public Property Get Alignment() As AlignmentConstants
Alignment = txtMain.Alignment
End Property

Public Property Let Alignment(ByVal a As AlignmentConstants)
txtMain.Alignment = a
PropertyChanged "Align"
End Property

'Public Property Get FontSize() As FontSizes
'FontSize = txtMain.FontSize
'End Property
'
'Public Property Let FontSize(ByVal Fs As FontSizes)
'txtMain.FontSize = Fs
'PropertyChanged "FontSize"
'End Property

Private Sub CheckKeyAsciiForNumber(ByRef KeyAscii As Integer, Optional AllowMinus As Boolean = False)
If AllowMinus Then If KeyAscii = 45 Then Exit Sub
Select Case KeyAscii
    Case 48 To 57, vbKeyBack, vbKeyDelete, vbKeyLeft, vbKeyRight
        
    Case Else
        KeyAscii = 0
End Select
End Sub

Public Property Get MaxLength() As Long
MaxLength = txtMain.MaxLength
End Property

Public Property Let MaxLength(ByVal l As Long)
txtMain.MaxLength = l
PropertyChanged "MaxLength"
End Property


Public Property Get Enabled() As Boolean
Enabled = txtMain.Enabled
End Property

Public Property Let Enabled(ByVal b As Boolean)
txtMain.Enabled = b
End Property


Public Property Get BackColor() As OLE_COLOR
BackColor = txtMain.BackColor
End Property

Public Property Let BackColor(ByVal v As OLE_COLOR)
txtMain.BackColor = v
PropertyChanged "Backcolor"
End Property
