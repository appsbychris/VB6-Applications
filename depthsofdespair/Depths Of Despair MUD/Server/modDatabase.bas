Attribute VB_Name = "modDatabase"
'
'----------------------------------------------------------
' Project   : DoDMudServer
' Module    : modDatabase
' Author    : Chris Van Hooser
' Copyright : 2004, Spike Technologies, Chris Van Hooser
' Email     : spike.spikey@comcast.net
'----------------------------------------------------------
'
Public Sub OpenDatabaseConnection()
modSec.dB_set_and_load "spike technolog", App.Path & "\data.mud", "Standard Jet DB"
modSec.dB_set_and_load "MUD", App.Path & "\data.mud", "4.0"
Set db = OpenDatabase(App.Path & "\data.mud", False, False, modSec.uJunkIt(";√pd¶wb>wd•ÉLa=6K©UÇ") & modSec.uJunkIt(sValue))
End Sub

Public Sub CloseDatabase()
db.Close
Set db = Nothing
modSec.dB_set_and_load "Standard Jet DB", App.Path & "\data.mud", "spike technolog"
modSec.dB_set_and_load "4.0", App.Path & "\data.mud", "MUD"
End Sub

Public Sub InitRecordsets()
Set MRS = db.OpenRecordset("SELECT * FROM Players")
Set MRSITEM = db.OpenRecordset("SELECT * FROM Items")
Set MRSMAP = db.OpenRecordset("SELECT * FROM Map")
Set MRSMONSTER = db.OpenRecordset("SELECT * FROM Monsters")
Set MRSCLASS = db.OpenRecordset("SELECT * FROM Class")
Set MRSRACE = db.OpenRecordset("SELECT * FROM Races")
Set MRSEMOTIONS = db.OpenRecordset("SELECT * FROM Emotions")
Set MRSSPELLS = db.OpenRecordset("SELECT * FROM Spells")
Set MRSFAMILIARS = db.OpenRecordset("SELECT * FROM Familiars")
Set MRSSHOPS = db.OpenRecordset("SELECT * FROM Shops")
Set MRSEVENTS = db.OpenRecordset("SELECT * FROM Events")
End Sub

Public Sub CloseRecordsets()
Set MRS = Nothing
Set MRSITEM = Nothing
Set MRSMAP = Nothing
Set MRSMONSTER = Nothing
Set MRSCLASS = Nothing
Set MRSRACE = Nothing
Set MRSEMOTIONS = Nothing
Set MRSSPELLS = Nothing
Set MRSEVENTS = Nothing
Set MRSSHOPS = Nothing
Set MRSFAMILIARS = Nothing
End Sub

Public Sub ValidateDatabase()
modSec.dB_set_and_load "spike technolog", App.Path & "\data.mud", "Standard Jet DB"
modSec.dB_set_and_load "MUD", App.Path & "\data.mud", "4.0"
End Sub
