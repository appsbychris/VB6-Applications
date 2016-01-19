Attribute VB_Name = "modWeaponsAndArmor"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modWeaponsAndArmor
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Function PlayerCanUseWeapon(dbIndex As Long, dbItemIndex As Long) As Boolean
Select Case dbPlayers(dbIndex).iWeapons
    Case 0
        Select Case dbItems(dbItemIndex).iType
            Case 1, 2
                PlayerCanUseWeapon = True
        End Select
    Case 1
        Select Case dbItems(dbItemIndex).iType
            Case 4, 5
                PlayerCanUseWeapon = True
        End Select
    Case 2
        Select Case dbItems(dbItemIndex).iType
            Case 1, 2, 4, 5
                PlayerCanUseWeapon = True
        End Select
    Case 3
        Select Case dbItems(dbItemIndex).iType
            Case 1 To 5
                PlayerCanUseWeapon = True
        End Select
    Case 4
        Select Case dbItems(dbItemIndex).iType
            Case 1, 2, 4, 5, 10
                PlayerCanUseWeapon = True
        End Select
    Case 5
        Select Case dbItems(dbItemIndex).iType
            Case 1, 4, 6, 8
                PlayerCanUseWeapon = True
        End Select
    Case 6
        Select Case dbItems(dbItemIndex).iType
            Case 1, 3
                PlayerCanUseWeapon = True
        End Select
    Case 7
        Select Case dbItems(dbItemIndex).iType
            Case 6, 7
                PlayerCanUseWeapon = True
        End Select
    Case 8
        Select Case dbItems(dbItemIndex).iType
            Case 8, 9
                PlayerCanUseWeapon = True
        End Select
    Case 9
        Select Case dbItems(dbItemIndex).iType
            Case 3, 6 To 10
                PlayerCanUseWeapon = True
        End Select
    Case 10
        Select Case dbItems(dbItemIndex).iType
            Case 1, 2, 6, 7
                PlayerCanUseWeapon = True
        End Select
    Case 11
        Select Case dbItems(dbItemIndex).iType
            Case 4, 5, 8, 9
                PlayerCanUseWeapon = True
        End Select
    Case 12
        Select Case dbItems(dbItemIndex).iType
            Case 2, 5, 7, 9
                PlayerCanUseWeapon = True
        End Select
    Case 13
        PlayerCanUseWeapon = True
End Select
End Function

Public Function GenericCanUseWeapon(iiWeapon As Long, dbItemIndex As Long) As Boolean
Select Case iiWeapons
    Case 0
        Select Case dbItems(dbItemIndex).iType
            Case 1, 2
                GenericCanUseWeapon = True
        End Select
    Case 1
        Select Case dbItems(dbItemIndex).iType
            Case 4, 5
                GenericCanUseWeapon = True
        End Select
    Case 2
        Select Case dbItems(dbItemIndex).iType
            Case 1, 2, 4, 5
                GenericCanUseWeapon = True
        End Select
    Case 3
        Select Case dbItems(dbItemIndex).iType
            Case 1 To 5
                GenericCanUseWeapon = True
        End Select
    Case 4
        Select Case dbItems(dbItemIndex).iType
            Case 1, 2, 4, 5, 10
                GenericCanUseWeapon = True
        End Select
    Case 5
        Select Case dbItems(dbItemIndex).iType
            Case 1, 4, 6, 8
                GenericCanUseWeapon = True
        End Select
    Case 6
        Select Case dbItems(dbItemIndex).iType
            Case 1, 3
                GenericCanUseWeapon = True
        End Select
    Case 7
        Select Case dbItems(dbItemIndex).iType
            Case 6, 7
                GenericCanUseWeapon = True
        End Select
    Case 8
        Select Case dbItems(dbItemIndex).iType
            Case 8, 9
                GenericCanUseWeapon = True
        End Select
    Case 9
        Select Case dbItems(dbItemIndex).iType
            Case 3, 6 To 10
                GenericCanUseWeapon = True
        End Select
    Case 10
        Select Case dbItems(dbItemIndex).iType
            Case 1, 2, 6, 7
                GenericCanUseWeapon = True
        End Select
    Case 11
        Select Case dbItems(dbItemIndex).iType
            Case 4, 5, 8, 9
                GenericCanUseWeapon = True
        End Select
    Case 12
        Select Case dbItems(dbItemIndex).iType
            Case 2, 5, 7, 9
                GenericCanUseWeapon = True
        End Select
    Case 13
        GenericCanUseWeapon = True
End Select
End Function

Public Function PlayerCanUseArmor(dbIndex As Long, dbItemIndex As Long) As Boolean
Select Case dbPlayers(dbIndex).iArmorType
    Case Is <= 11
        If dbPlayers(dbIndex).iArmorType >= dbItems(dbItemIndex).iArmorType Then
            PlayerCanUseArmor = True
        Else
            PlayerCanUseArmor = False
        End If
    Case 12
        '2,3
        Select Case dbItems(dbItemIndex).iArmorType
            Case 2, 3
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 13
        '5 to 7
        Select Case dbItems(dbItemIndex).iArmorType
            Case 5 To 7
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 14
        '8,9
        Select Case dbItems(dbItemIndex).iArmorType
            Case 8, 9
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 15
        '3,8,9
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 8, 9
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 16
        '3,10
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 10
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 17
        '3,10,11
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 10, 11
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 18
        '5 to 9
        Select Case dbItems(dbItemIndex).iArmorType
            Case 5 To 9
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 19
        '11
        Select Case dbItems(dbItemIndex).iArmorType
            Case 11
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 20
        '3,11
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 11
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 21
        '2,5 to 7
        Select Case dbItems(dbItemIndex).iArmorType
            Case 2, 5 To 7
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 22
        '4
        Select Case dbItems(dbItemIndex).iArmorType
            Case 4
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 23
        '2,4
        Select Case dbItems(dbItemIndex).iArmorType
            Case 2, 4
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
    Case 24
        '3,4
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 4
                PlayerCanUseArmor = True
            Case Else
                PlayerCanUseArmor = False
        End Select
End Select
Select Case dbItems(dbItemIndex).sWorn
    Case "arms"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Arms]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "back"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Back]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "body"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Body]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "ears"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Ears]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "face"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Face]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "hands"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Hands]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "head"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Head]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "legs"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Legs]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "neck"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Neck]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "shield"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Shield]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "waist"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Waist]) = 1 Then
            PlayerCanUseArmor = False
        End If
    Case "weapon"
        If modMiscFlag.GetMiscFlag(dbIndex, [Can Eq Weapon]) = 1 Then
            PlayerCanUseArmor = False
        End If
End Select
End Function

Public Function GenericCanUseArmor(iArmor As Long, dbItemIndex As Long) As Boolean
Select Case iArmor
    Case Is <= 11
        If iArmor >= dbItems(dbItemIndex).iArmorType Then
            GenericCanUseArmor = True
        Else
            GenericCanUseArmor = False
        End If
    Case 12
        '2,3
        Select Case dbItems(dbItemIndex).iArmorType
            Case 2, 3
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 13
        '5 to 7
        Select Case dbItems(dbItemIndex).iArmorType
            Case 5 To 7
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 14
        '8,9
        Select Case dbItems(dbItemIndex).iArmorType
            Case 8, 9
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 15
        '3,8,9
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 8, 9
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 16
        '3,10
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 10
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 17
        '3,10,11
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 10, 11
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 18
        '5 to 9
        Select Case dbItems(dbItemIndex).iArmorType
            Case 5 To 9
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 19
        '11
        Select Case dbItems(dbItemIndex).iArmorType
            Case 11
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 20
        '3,11
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 11
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 21
        '2,5 to 7
        Select Case dbItems(dbItemIndex).iArmorType
            Case 2, 5 To 7
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 22
        '4
        Select Case dbItems(dbItemIndex).iArmorType
            Case 4
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 23
        '2,4
        Select Case dbItems(dbItemIndex).iArmorType
            Case 2, 4
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
    Case 24
        '3,4
        Select Case dbItems(dbItemIndex).iArmorType
            Case 3, 4
                GenericCanUseArmor = True
            Case Else
                GenericCanUseArmor = False
        End Select
End Select

End Function
