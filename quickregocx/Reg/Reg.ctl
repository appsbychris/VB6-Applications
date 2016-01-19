VERSION 5.00
Begin VB.UserControl QReg 
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Reg.ctx":0000
   ScaleHeight     =   780
   ScaleWidth      =   600
   ToolboxBitmap   =   "Reg.ctx":17B2
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "QReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************************
'*************************************************************************************
'***************       Code create by Chris Van Hooser          **********************
'***************                  (c)2001                       **********************
'*************** You may use this code and freely distribute it **********************
'***************   If you have any questions, please email me   **********************
'***************          at theendorbunker@home.com.           **********************
'***************       Thanks for downloading my project        **********************
'***************        and i hope you can use it well.         **********************
'*************************************************************************************
'*************************************************************************************

'*************************************************************************************
'*******************    API calls, Enums, and Variable decloration    ****************
'*************************************************************************************

'API call that gets the computer name
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
'My enum for the types of Keys it makes/Checks
Public Enum TypeOfCheck
    ComputerName = 0
    cSerialNumber = 1
    ComboOfBoth = 2
End Enum
'My enum for the strength of the key (higher number, longer key)
Public Enum StrOf
    Minimum = 0
    Medium = 1
    Maximum = 2
    XtraMax = 3
End Enum
'Power of the scrambler, (highernumber, the longer the key gets)
Public Enum ScramStr
    Minor = 0
    Tough = 1
    Tougher = 2
    Toughest = 3
End Enum
'Variables
Dim TCompNa$ 'Computer Name
Dim ScrStr As ScramStr 'Enum of Scramble Strength
Dim SoF As StrOf 'Enum for Strength of key
Dim ToC As TypeOfCheck 'Enum for the type of check
Dim WtC As String 'What to Check (For serial number)
Dim ScrambleStr As Integer 'The strength for the scrambler
Dim NoPC As String 'NameOfPc
Dim Str As Integer 'Strength of key gen

'**************************************************************************************
'*****************    Public Properties of the control    *****************************
'**************************************************************************************

'Set the property of checktype to my enum
Public Property Get CheckType() As TypeOfCheck
Attribute CheckType.VB_Description = "Setting for the type of data you want to use to make the key, either being the computer name, or a serial number, or both combined."
    'Set the property to whatever my variabl of TypeOFCheck is
    CheckType = ToC
End Property

Public Property Let CheckType(ByVal NewVal As TypeOfCheck)
    'Change my variable of TypeOfCheck to the new value
    ToC = NewVal
    PropertyChanged
End Property

'Set the property to my enum
Public Property Get ScrambleStrength() As ScramStr
    'set ScrambleStrength to my variable
    ScrambleStrength = ScrStr
End Property

Public Property Let ScrambleStrength(ByVal NewVal As ScramStr)
    'set my varibale to the newval
    ScrStr = NewVal
    PropertyChanged
End Property

'Set the property of Strengh to my enum
Public Property Get Strength() As StrOf
Attribute Strength.VB_Description = "Sets the strength of the ken generator."
    'set the propert to whatever my variable of StrOf is
    Strength = SoF
End Property

Public Property Let Strength(ByVal NewVal As StrOf)
    'Change my variable to whatever the newval is
    SoF = NewVal
    PropertyChanged
End Property

'make a property for the serialnumber, and set it as a string
Public Property Get SerialNumber() As String
Attribute SerialNumber.VB_Description = "This value can be any string value.  Use this if you do not with to have the key based on the user's computers name."
    'Set the serialnumber property to whatever wtc(What To Check) is
    SerialNumber = WtC
End Property

Public Property Let SerialNumber(ByVal NewVal As String)
    'change wtc to the newval
    WtC = NewVal
    Label1.Caption = WtC
    PropertyChanged
End Property

'**************************************************************************************
'*********************    Public Functions of the control    **************************
'**************************************************************************************

'The function the user uses to gen a key,
'used as x = Reg1.GenAKey
Public Function GenAKey() As String
Attribute GenAKey.VB_Description = "Generates a key based on the settings you specify."
    'if the toc (Type Of Check) is computername then
    If ToC = ComputerName Then
        'Get the Computername
        Call SetCompName
        'Get the strength number
        Call GetStrNumber
        'Gen the key
        GenAKey = GenKey(TCompNa$, Str)
    'But if it is a SerialNUmber check then
    ElseIf ToC = cSerialNumber Then
        'make sure wtc isnt nothing
        If Trim(WtC) <> "" Then
            'Get the strength number
            Call GetStrNumber
            'gen the key
            GenAKey = GenKey(WtC, Str)
        End If
    'If the user selected combo then
    ElseIf ToC = ComboOfBoth Then
        'if wtc isn't nothing
        If Trim(WtC) <> "" Then
            'Get the comp Name
            Call SetCompName
            'Get the strength number
            Call GetStrNumber
            'Dim some random variables
            Dim x1%, x2%
            'Set them to random numbers
            x1% = Int(Rnd * (8 - 1) + 1)
            x2% = Int((Rnd * (4 - 1) + 1))
            'Make a small array
            Dim tArray(3) As String
            tArray(0) = "a"
            tArray(1) = "g"
            tArray(2) = "n"
            tArray(3) = "y"
            'Gen a key for a computer name, place whatever x1 is, then
            'whatever tArray(of x2 minus 1) is, and then x1 minus 1, and then
            'Gen a key for the serial number and then add whatever x1 plus
            '1 is, and then pu the letter of tArray(of x2 minus 1 is)
            'And set this function to whatever it gets
            GenAKey = GenKey(TCompNa$, Str) & x1% & tArray(x2% - 1) & x1% - 1 & GenKey(WtC, Str) & x1% + 1 & tArray(x2% - 1)
        End If
    End If
    'for the prog to have wtc equal whatever is in label1's caption
    WtC = Label1.Caption
End Function

Public Function CheckAKey(CheckWhat$) As Boolean
Attribute CheckAKey.VB_Description = "Checks a certian key with the set values and gives a boolean response wether it is correct or not."
'This is the function the user uses to check a key
'Used as if Reg1.CheckAKey("thekeyhere") = True then MsgBox "Worked"
    'If it is a computer name checking
    If ToC = ComputerName Then
        'set the pc name
        Call SetCompName
        'get the strength number
        Call GetStrNumber
        'Check the key
        If CheckKey(TCompNa$, Str, CheckWhat$, False) = True Then
            'if its true, then set the function
            'to be true
            CheckAKey = True
        Else
            'if not, set hte function to false
            CheckAKey = False
        End If
    'If they are checking a serialnumber then
    ElseIf ToC = cSerialNumber Then
        'If wtc isnt nothing
        If Trim(WtC) <> "" Then
            'get the strength number
            Call GetStrNumber
            'Check the key
            If CheckKey(WtC, Str, CheckWhat$, True) = True Then
                'If its good, set the function true
                CheckAKey = True
            Else
                'if not, set it false
                CheckAKey = False
            End If
        End If
    'If they are checking a combo then
    ElseIf ToC = ComboOfBoth Then
        'If wtc isnt nothing
        If Trim(WtC) <> "" Then
            'set the pc name
            Call SetCompName
            'Get the strength number
            Call GetStrNumber
            'Dom some variables
            Dim tTemp$, x1%
            'Set ttemp equal to the last character of what
            'to check
            tTemp$ = Right(CheckWhat$, 1)
            'trim off that last character
            CheckWhat$ = Left(CheckWhat$, Len(CheckWhat$) - 1)
            'Set x1 equal to the lastcharacter of checkwhat
            x1% = Val(Right(CheckWhat$, 1))
            'Since we added 1 when making it, subtract 1
            x1% = x1% - 1
            'Trim off the last character
            CheckWhat$ = Left(CheckWhat$, Len(CheckWhat$) - 1)
            'Dim an integer
            Dim r%
            'Set it to find where we put the mid section of out 2 keys
            r% = InStr(1, CheckWhat$, x1% & tTemp$ & x1% - 1)
            'If it found it, continue on
            If r% <> 0 Then
                'check the key for the computer name, its the first half of
                'the CheckWhat variable
                If CheckKey(TCompNa$, Str, Left(CheckWhat$, r% - 1), False) = True Then
                    'if that key is correct, then
                    'trim it off
                    For i = 1 To r% + 2
                        CheckWhat$ = Mid(CheckWhat$, 2)
                    Next
                    'Now check the serial number key
                    If CheckKey(WtC, Str, CheckWhat$, True) = True Then
                        'if it works, set the function to true
                        CheckAKey = True
                    Else
                        'if not, set it to false
                        CheckAKey = False
                    End If
                Else
                    'If not, set function to false
                    CheckAKey = False
                End If
            Else
                'if not, set it to be false
                CheckAKey = False
            End If
        End If
    End If
    'set Wtc to label1's caption
    WtC = Label1.Caption
End Function

'**************************************************************************************
'***********************    Public Subs for the control    ****************************
'**************************************************************************************

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
    MsgBox "QuickReg was created by Chris Van Hooser." & vbCrLf _
        & "This control is free to use and distibute, as long" & vbCrLf & _
        " as it is not altered in any way, shape or form." & _
        vbCrLf & "(c)2001 Chris Van Hooser" & vbCrLf & _
        "theendorbunker@home.com", vbInformation + vbOKOnly, "About"
End Sub

'**************************************************************************************
'***********************    UserControl Events    *************************************
'**************************************************************************************

Private Sub UserControl_Initialize()
'Initialize the random number generator
Randomize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'Setup the propbag to read things
CheckType = PropBag.ReadProperty("CheckType", TypeOfCheck.ComputerName)
Strength = PropBag.ReadProperty("Strength", StrOf.Medium)
SerialNumber = PropBag.ReadProperty("SerialNumber", "")
ScrambleStrength = PropBag.ReadProperty("ScrambleStrength", ScramStr.Tougher)
End Sub

Private Sub UserControl_Resize()
'Make it so it cant get bigger or smaller then 600x780
UserControl.Width = 600
UserControl.Height = 780
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Write all of the properties to the propbag
PropBag.WriteProperty "CheckType", ToC, TypeOfCheck.ComputerName
PropBag.WriteProperty "Strength", SoF, StrOf.Medium
PropBag.WriteProperty "ScrambleStrength", ScrStr, ScramStr.Tougher
PropBag.WriteProperty "SerialNumber", Label1.Caption, ""
End Sub

'**************************************************************************************
'************************    Private Subs for the Control    **************************
'**************************************************************************************

Private Sub SetScramStr()
'Get the scramble strengths
Select Case ScrStr
    Case 0:
        'lowest setting is 5
        ScrambleStr = 5
    Case 1:
        'next is 9
        ScrambleStr = 9
    Case 2:
        'then is 12
        ScrambleStr = 12
    Case 3:
        'and max is 15
        ScrambleStr = 15
End Select
End Sub

Private Sub SetCompName()
'Get the PC name
Dim p As Long
p = PCName(NoPC)
TCompNa$ = NoPC
'Get rid of Chr(0) in it
TCompNa$ = Replace(TCompNa$, Chr(0), "")
'Get rid of extra spaces
TCompNa$ = Trim(TCompNa$)
'set wtc to what label1's caption is
WtC = Label1.Caption
End Sub

Private Sub GetStrNumber()
'Function to set the strength
    Select Case SoF
        'If lowest, set to 9
        Case 0:
            Str = 9
        'Next level is 15
        Case 1:
            Str = 15
        'Then 24
        Case 2:
            Str = 24
        'And max is 34
        Case 3:
            Str = 34
    End Select
    'set wtc to label1's caption
    WtC = Label1.Caption
End Sub

'**************************************************************************************
'********************    Private Functions for the control    *************************
'**************************************************************************************

Private Function PCName(sName As String) As Long
'Function to get the computername
Dim NameSize As Long
Dim x As Long
sName = Space$(16)
NameSize = Len(sName)
x = GetComputerName(sName, NameSize)
'Force the comp to save whats in WtC into label1
WtC = Label1.Caption
End Function

Private Function GenKey(strCompName$, intLength%) As String
'Basic key Gening
'Variables
Dim Temp$, TempCompName$, AscFirstChar%, AscThirdChar%
'The min length for a key can be 3, so check for it
If Len(strCompName$) < 3 Then Exit Function
'Make the case for the name alternate (ex: tHiSiSaTeSt)
AlternateCase strCompName$
'Save the original name for later use
TempCompName$ = strCompName$
'Get the ascii of the first and third character
AscFirstChar% = Asc(Left(TempCompName$, 1))
AscThirdChar% = Asc(Mid(TempCompName$, 3, 1))
'Change each character in the name to the ascii of the name
For i = 1 To Len(strCompName)
    Temp = Temp & Asc(Left(strCompName$, 1))
    strCompName$ = Mid(strCompName$, 2)
Next
'Random number variables
Dim intRand1%, intRand2%
'Make the first variable an integer from 1 to 9
intRand1% = Int(Rnd * (9 - 1) + 1)
'Make the second number a variable from 113 to 97
intRand2% = Int(Rnd * (113 - 97) + 97)
'If whatever the ascii of the name plus the first random
'number is greater then 255, make it 0
If Asc(Left(TempCompName$, 1)) + intRand1% > 255 Then intRand1% = 0
'Reverse the string to make it more confusing
Temp$ = StrReverse(Temp$)
'Put a lot of things together---
'Whatever the character for intRand2 is, and then whatever the ascii of
'the first letter of the Name is plus what ever intrand1 is, then whatever
'temp was in the firstplace, then the character of intrand2 and intrand1
'added together is, then whatever intrand2 is, and then whatever the
'character of intrand2 minus intrand1, and finaly intrand1 at the end
Temp$ = Chr(intRand2%) & Asc(Left(TempCompName$, 1)) _
    + intRand1% & Temp$ & Chr(intRand2% + intRand1%) _
    & intRand2% & Chr(intRand2% - intRand1%) & intRand1%
'And again of putting more things together
'Howeverlong the name is, and then the first ::intLength:: characters (intLength
'equals the value of the strength of the check) and then the character of
'intRand2 and intRand2 added together, then intRand2, and then the length of
'the ascii of the first character in the name, and then the ascii of the
'length of the 3rd character in the name plus intrand2, and then the character
'of intrand2 minus intrand1, and at the end, intrand1
Temp$ = Len(TempCompName$) & Left(Temp$, intLength%) & Chr(intRand2% + _
    intRand1%) & intRand2% & Len(AscFirstChar%) & _
    Asc(Len(AscThirdChar%) + intRand2%) & Chr(intRand2% - intRand1%) _
    & intRand1%
'Ok, now we scramble to key to make it even harder to crack.
Temp$ = ScrambleKey(Temp$, intRand1%)
'Now we put in information so that we can make sure the code works when
'we go and check it
'Take the length of the name, then the character of intrand2, and then the first half
'of whatever temp happends to be, and then the character of intrand2 minus intrand1
'and then the character of the ascii of the thirdcharacter in the name, and then the
'last half of whatever temp is, and then the character of intrand2 and intrand1 added
'together, and then the character of the ascii of the first character in the name
'and last we put what intrand1 minus 1 is.
Temp$ = Len(TempCompName$) & Chr(intRand2%) & Left(Temp$, Int(Len(Temp$) / 2)) _
    & Chr(intRand2% - intRand1%) & Chr(AscThirdChar) & Right(Temp$, Int(Len(Temp$) / 2)) & _
    Chr(intRand2% + intRand1%) & Chr(AscFirstChar) & intRand1% - 1
'Now we alternate the case of the key to make it look different
AlternateCase Temp$
'and set GenKey to equal the key
GenKey = Temp$
'Force the prog to save whatever WtC is in label1
WtC = Label1.Caption
End Function

Private Function CheckKey(strCompName$, intLength%, strKey$, cSerNum As Boolean) As Boolean
Dim Temp$, intRand1%, intRand2%, TempCompName$, iLen%, AscFirstChar%, AscThirdChar%
'Now to check if the key works.
'If either the name or the key is less then 3 characters,
'exit the function
If Len(strCompName$) < 3 Or Len(strKey$) < 3 Then Exit Function
'We need to save the length of the lenghth of the key for later use
iLen% = Len(strCompName$)
iLen% = Len(iLen%)
'Check to find if cSerNum is true or false, and adjust the
'iLen integer to the correct value in order to find the
'position of intRand2
If cSerNum = False And iLen% > 2 Then iLen% = iLen% + 1
If cSerNum = True And iLen% > 1 Then iLen% = iLen% + 1
'Alternate the case of the name
AlternateCase strCompName$
'Put the name in a temp storagespot for later use
TempCompName$ = strCompName$
'Get the ascii of the first and third character
AscFirstChar% = Asc(Left(TempCompName$, 1))
AscThirdChar% = Asc(Mid(TempCompName$, 3, 1))
'Get the first random number that was used to make the key
'It is stored as the lastcharacter.  Since when we made the key, we
'subracted 1 from its total, we add 1 to it
intRand1% = Right(strKey$, 1) + 1
'intRand2 is stored in either the 2nd or 3rd posistion usually,
'it depends on the length of the length of the key
intRand2% = Asc(LCase(Mid(strKey$, iLen%, 1)))
'Now converto the name to ascii
For i = 1 To Len(strCompName$)
    Temp$ = Temp$ & Asc(Left(strCompName$, 1))
    strCompName$ = Mid$(strCompName$, 2)
Next
'Now we make a new key, depending on the random numbers used to first make it
'Reverse the string to make it more confusing
Temp$ = StrReverse(Temp$)
'Put a lot of things together---
'Whatever the character for intRand2 is, and then whatever the ascii of
'the first letter of the Name is plus what ever intrand1 is, then whatever
'temp was in the firstplace, then the character of intrand2 and intrand1
'added together is, then whatever intrand2 is, and then whatever the
'character of intrand2 minus intrand1, and finaly intrand1 at the end
Temp$ = Chr(intRand2%) & Asc(Left(TempCompName$, 1)) + intRand1% _
    & Temp$ & Chr(intRand2% + intRand1%) & intRand2% & Chr(intRand2% _
    - intRand1%) & intRand1%
'And again of putting more things together
'Howeverlong the name is, and then the first ::intLength:: characters (intLength
'equals the value of the strength of the check) and then the character of
'intRand2 and intRand2 added together, then intRand2, and then the length of
'the ascii of the first character in the name, and then the ascii of the
'length of the 3rd character in the name plus intrand2, and then the character
'of intrand2 minus intrand1, and at the end, intrand1
Temp$ = Len(TempCompName$) & Left(Temp$, intLength%) & Chr(intRand2% _
    + intRand1%) & intRand2% & Len(AscFirstChar%) & _
    Asc(Len(AscThirdChar%) + intRand2%) & Chr(intRand2% - intRand1%) & intRand1%
'Ok, now we scramble to key to make it even harder to crack.
Temp$ = ScrambleKey(Temp$, intRand1%)
'Take the length of the name, then the character of intrand2, and then the first half
'of whatever temp happends to be, and then the character of intrand2 minus intrand1
'and then the character of the ascii of the thirdcharacter in the name, and then the
'last half of whatever temp is, and then the character of intrand2 and intrand1 added
'together, and then the character of the ascii of the first character in the name
'and last we put what intrand1 minus 1 is.
Temp$ = Len(TempCompName$) & Chr(intRand2%) & Left(Temp$, Int(Len(Temp$) / 2)) _
    & Chr(intRand2% - intRand1%) & Chr(AscThirdChar) & Right(Temp$, Int(Len(Temp$) / 2)) & _
    Chr(intRand2% + intRand1%) & Chr(AscFirstChar) & intRand1% - 1
'Alternate the case
AlternateCase Temp$
'Now if the key we made matches the one we were supplyed, it is
'a legit key, and set the function equal to true
If strKey$ = Temp$ Then
    CheckKey = True
    Exit Function
'if not, set the function false
ElseIf Temp <> strKey Then
    CheckKey = False
End If
'Force the prog to save whatever WtC is in label1
WtC = Label1.Caption
End Function

Private Function AlternateCase(strAlter$) As String
    'Function to alternate the case of a string
    '(ex: tHiSiSaTeSt)
    'Variables
    Dim iLength%, i%
    'Get the length of the string
    iLength% = Len(strAlter$)
    'Lower the case eof str alter
    strAlter = LCase(strAlter$)
    'Loop through every other character and make it uppercase
    For i% = 1 To iLength% Step 2
        Mid$(strAlter$, i, 1) = UCase(Mid$(strAlter$, i, 1))
    Next
    'Set the function to what we made
    AlternateCase = strAlter$
End Function

Private Function ScrambleKey(strScram$, intRand1%) As String
    'The function to scramble the key
    'Variables
    Dim iLength%, i%, Temp1$, Temp2$, Temp3$
    'Get the strength of the scrambler
    Call SetScramStr
    'Get the length of the string to scramble
    iLength = Len(strScram$)
    'Get the first 1/3rd of the string, and store it
    'in temp1
    Temp1$ = Left(strScram$, Int(iLength / 3))
    'Get the last half of the string, and store it
    'in temp2
    Temp2$ = Right(strScram$, Int(iLength / 2))
    'And then get the middle of the string, starting in the
    '1/3rd in spot, and then the length of the
    'full length of the string, minus the length of it divided
    'by 3, and then that total divided by 2
    Temp3$ = Mid(strScram$, Int(iLength / 3), Int((iLength - _
        Int(iLength / 3)) / 2))
    'Now to mess it up
    'Take the length of Temp2, multiply that by intrand1, and then take that
    'total and multiply it by the length of temp2, and take that total and
    'multiply the the Str of the key, and that times the length of Temp3.
    'Next take the length of temp3, multiply it by the str of the key, and
    'then the strength of the scramblestr.
    'Last, take the length of temp1, multiply that by intRand1, and then
    'take that total and multiply it by the scramble str, and set the function
    'to be all the junk we just did.
    ScrambleKey = (Len(Temp2$) * intRand1%) * Len(Temp2$) * Str _
        * Len(Temp3$) & (Len(Temp3$) * Str * ScrambleStr) & _
        Len(Temp1$) * intRand1% * ScrambleStr
End Function

