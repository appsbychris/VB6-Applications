Attribute VB_Name = "modDoSubs"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modDoSubs
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'

Public Function DoCommands(lngIndex As Long, Optional IsDropped As Boolean = False, Optional dbIndex As Long) As Boolean
If IsDropped = False Then
    If MoveMentCommands(lngIndex) = True Then DoCommands = True: Exit Function
    If AttackCommands(lngIndex, dbIndex) = True Then DoCommands = True: Exit Function
    If RemDropGetEQ(lngIndex) = True Then DoCommands = True: Exit Function 'check for equiping/droping/etc commands
    If Sneak(lngIndex) = True Then DoCommands = True: Exit Function
    If pRest(lngIndex) = True Then DoCommands = True: Exit Function 'check for the 'rest' command
    If LookAround(lngIndex) = True Then DoCommands = True: Exit Function 'check for the 'l' command
    If PlayerStats(lngIndex) = True Then DoCommands = True: Exit Function 'check for the 'stats'command
    If AidPlayer(lngIndex) = True Then DoCommands = True: Exit Function
    If Reload(lngIndex) = True Then DoCommands = True: Exit Function
    If UnloadWeapon(lngIndex) = True Then DoCommands = True: Exit Function
    If UseItem(lngIndex) = True Then DoCommands = True: Exit Function 'check for the 'use' command
    If PartyCommands(lngIndex) = True Then DoCommands = True: Exit Function 'check for party commands
    If StatsExtended(lngIndex) = True Then DoCommands = True: Exit Function
    If FamStats(lngIndex) = True Then DoCommands = True: Exit Function
    If Hunger(lngIndex) = True Then DoCommands = True: Exit Function
    If Stamina(lngIndex) = True Then DoCommands = True: Exit Function
    If Steal(lngIndex) = True Then DoCommands = True: Exit Function
    If Mug(lngIndex) = True Then DoCommands = True: Exit Function
    If ListCommands(lngIndex) = True Then DoCommands = True: Exit Function
    If RideFam(lngIndex) = True Then DoCommands = True: Exit Function
    If GetOffFam(lngIndex) = True Then DoCommands = True: Exit Function
    If Eat(lngIndex) = True Then DoCommands = True: Exit Function
    If Tame(lngIndex) = True Then DoCommands = True: Exit Function
    If LetterSubs(lngIndex) = True Then DoCommands = True: Exit Function
    If TimeSubs(lngIndex) = True Then DoCommands = True: Exit Function
    If DeathCommands(lngIndex) = True Then DoCommands = True: Exit Function
    If Train(lngIndex) = True Then DoCommands = True: Exit Function 'check for the 'train' command
    If TrainStats(lngIndex) = True Then DoCommands = True: Exit Function
    If TrainClass(lngIndex) = True Then DoCommands = True: Exit Function
    If Map(lngIndex) = True Then DoCommands = True: Exit Function
    If NameFam(lngIndex) = True Then DoCommands = True: Exit Function
    If GuildCommands(lngIndex) = True Then DoCommands = True: Exit Function
    If BankOptions(lngIndex, dbIndex) = True Then DoCommands = True: Exit Function 'Banking commands
    If Emotes(lngIndex, dbIndex) = True Then DoCommands = True: Exit Function 'check for emotion command
    If SetCommand(lngIndex) = True Then DoCommands = True: Exit Function
    If IsSysCommand(lngIndex) = True Then DoCommands = True: Exit Function 'check for sysop commands
    If Speaking(lngIndex) = True Then DoCommands = True: Exit Function 'default any other command here
Else
    If PlayerStats(lngIndex) = True Then DoCommands = True: Exit Function
    If IsSysCommand(lngIndex) = True Then DoCommands = True: Exit Function
    If DeathCommands(lngIndex) = True Then DoCommands = True: Exit Function
    If Hunger(lngIndex) = True Then DoCommands = True: Exit Function
    If Stamina(lngIndex) = True Then DoCommands = True: Exit Function
    If LookAround(lngIndex) = True Then DoCommands = True: Exit Function
    If TimeSubs(lngIndex) = True Then DoCommands = True: Exit Function
    If StatsExtended(lngIndex) = True Then DoCommands = True: Exit Function
    If Emotes(lngIndex) = True Then DoCommands = True: Exit Function 'check for emotion command
    If SetCommand(lngIndex) = True Then DoCommands = True: Exit Function
    If FamStats(lngIndex) = True Then DoCommands = True: Exit Function
    If Speaking(lngIndex) = True Then DoCommands = True: Exit Function
End If
End Function
