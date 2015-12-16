'------------------------------------------------------------------------------------------------
' Filename    : modMCE.vb
' Purpose     : This is the module that provides stores function or variables specific to MCE
' Created By  : Felix Kang - I-CAT Computing (21 MAY 2007)
' Note        : 
' Assumptions : - Code is based on Visual Basic .NET (Visual Studio 2003)
'               - mySQL Connector is installed and referenced properly
'------------------------------------------------------------------------------------------------
' History
' - 21 MAY 2007 : Creation date of the module
'------------------------------------------------------------------------------------------------

#Region " System Imports "

'Imports all the components we need
Imports System
Imports System.Data
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Security.Cryptography
'mySQL DB library
Imports MySql.Data.MySqlClient

#End Region

Module modMCE

#Region " Public Variables "

'----------------- Database Related -----------------
Public strConnectionString As String
Public strCurrentDBServer As String
Public sngGSTRate As Single

#End Region

#Region " Constants "

'Database Related  
Public MCE_DEFAULT_DB_NAME = "Mangos"
Public MCE_MYSQL_DEF_PORT = "3306"
Public MCE_MYSQL_DEF_SERVER = "localhost"
Public MCE_MYSQL_DEF_TIMEOUT = "120"
'Config file
Public MCE_CONFIG_FILE = "MCE Config.xml"

'Character Table Field numbers http://wiki.udbforums.org/index.php/Character_data (current as of 2.2.3)
Public DB_UNIT_FIELD_CURR_HEALTH = "22"     'Current Health
Public DB_UNIT_FIELD_CURR_MANA = "23"       'Current Mana
Public DB_UNIT_FIELD_MAX_HEALTH = "28"      'Max Health
Public DB_UNIT_FIELD_MAX_MANA = "29"        'Max Mana
Public DB_UNIT_FIELD_LEVEL = "34"           'Current Level
Public DB_UNIT_FIELD_STR = "164"            'Strength
Public DB_UNIT_FIELD_AGILITY = "165"        'Agility
Public DB_UNIT_FIELD_STAMINA = "166"        'Stamina
Public DB_UNIT_FIELD_IQ = "167"             'Intellect
Public DB_UNIT_FIELD_SPIRIT = "168"         'Spirit
Public DB_PLAYER_XP = "856"                 'Current Player XP
Public DB_PLAYER_NEXT_LEVEL_XP = "857"      'Amount of XP needed to reach next level
Public DB_PLAYER_BLOCK_PERCENTAGE = "1246"  'Block Percentage
Public DB_PLAYER_DODGE_PERCENTAGE = "1247"  'Dodge Percentage
Public DB_PLAYER_PARRY_PERCENTAGE = "1248"  'Dodge Percentage

#End Region

#Region " DLL definition "

#End Region

#Region " Tooltip Text "

#End Region

#Region " SQL Statements "


#End Region

#Region " Functions and Procedures "

Public Function RaceIDToString(ByVal strRaceID As String) As String
'-------------------------------------------------------------------------------
'Purpose     : This procedure is to translate the race ID into readable string
'Assumptions :
'Input       :
'   - strRaceID, a string consisting the race id, as defined in
'     https://svn.mangosproject.org/trac/MaNGOS/wiki/Database/character
'Returns     :
'Note        : 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandling


  Select Case strRaceID
    Case "1"
      RaceIDToString = "Human"
    Case "2"
      RaceIDToString = "Orc"
    Case "3"
      RaceIDToString = "Dwarf"
    Case "4"
      RaceIDToString = "Night Elf"
    Case "5"
      RaceIDToString = "Undead"
    Case "6"
      RaceIDToString = "Tauren"
    Case "7"
      RaceIDToString = "Gnome"
    Case "8"
   RaceIDToString = "Troll"
  Case "10"
   RaceIDToString = "Blood Elf"
  Case "11"
   RaceIDToString = "Dranei"
  Case Else
   RaceIDToString = "Unknown [" & strRaceID & "]"
 End Select

  Exit Function

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("modMCE.vb", "RaceIDToString", Err.Number, Err.Source, Err.Description)
End Function

Public Function ClassIDToString(ByVal strClassID As String) As String
'-------------------------------------------------------------------------------
'Purpose     : This procedure is to translate the race ID into readable string
'Assumptions :
'Input       :
'   - strRaceID, a string consisting the race id, as defined in
'     https://svn.mangosproject.org/trac/MaNGOS/wiki/Database/character
'Returns     :
'Note        : 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandling


  Select Case strClassID
    Case "1"
      ClassIDToString = "Warrior"
    Case "2"
      ClassIDToString = "Paladin"
    Case "3"
      ClassIDToString = "Hunter"
    Case "4"
      ClassIDToString = "Rogue"
    Case "5"
      ClassIDToString = "Priest"
    Case "6"
      ClassIDToString = "Unk0"
    Case "7"
      ClassIDToString = "Shaman"
    Case "8"
      ClassIDToString = "Mage"
    Case "9"
      ClassIDToString = "Warlock"
    Case "10"
      ClassIDToString = "Unk1"
    Case "11"
      ClassIDToString = "Druid"
    Case Else
      ClassIDToString = "Unknown Class [" & ClassIDToString & "]"
  End Select

  Exit Function

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("modMCE.vb", "RaceIDToString", Err.Number, Err.Source, Err.Description)
End Function

#End Region

End Module
