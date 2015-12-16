'------------------------------------------------------------------------------------------------
' Filename    : frmDBSettings.vb
' Purpose     : This form shows database settings that can be configured for MCE
' Created By  : Felix Kang - I-CAT Computing (21 MAY 2007)
' Note        :
' Assumptions : - Code is based on Visual Basic .NET (Visual Studio 2003)
'------------------------------------------------------------------------------------------------
' History
' - 21 MAY 2007 : Form creation
'------------------------------------------------------------------------------------------------

#Region " System Imports "

'mySQL DB library
Imports MySql.Data.MySqlClient

#End Region

Public Class frmDBSettings
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
  Friend WithEvents txtPortNumber As System.Windows.Forms.TextBox
  Friend WithEvents lblPortNo As System.Windows.Forms.Label
  Friend WithEvents lblSecTimeout As System.Windows.Forms.Label
  Friend WithEvents lblTimeout As System.Windows.Forms.Label
  Friend WithEvents lblServerInfo As System.Windows.Forms.Label
  Friend WithEvents cmdTestConnection As System.Windows.Forms.Button
  Friend WithEvents txtmySQLServer As System.Windows.Forms.TextBox
  Friend WithEvents lblmySQLServer As System.Windows.Forms.Label
  Friend WithEvents cmdSaveDBSetting As System.Windows.Forms.Button
  Friend WithEvents txtMySQLInfo As System.Windows.Forms.TextBox
  Friend WithEvents txtDatabaseName As System.Windows.Forms.TextBox
  Friend WithEvents lblDatabaseName As System.Windows.Forms.Label
  Friend WithEvents txtTimeout As System.Windows.Forms.TextBox
  Friend WithEvents lblUsername As System.Windows.Forms.Label
  Friend WithEvents txtUsername As System.Windows.Forms.TextBox
  Friend WithEvents lblPassword As System.Windows.Forms.Label
  Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.txtPortNumber = New System.Windows.Forms.TextBox
Me.lblPortNo = New System.Windows.Forms.Label
Me.lblSecTimeout = New System.Windows.Forms.Label
Me.lblTimeout = New System.Windows.Forms.Label
Me.lblServerInfo = New System.Windows.Forms.Label
Me.cmdTestConnection = New System.Windows.Forms.Button
Me.txtmySQLServer = New System.Windows.Forms.TextBox
Me.lblmySQLServer = New System.Windows.Forms.Label
Me.lblDatabaseName = New System.Windows.Forms.Label
Me.txtDatabaseName = New System.Windows.Forms.TextBox
Me.txtTimeout = New System.Windows.Forms.TextBox
Me.txtMySQLInfo = New System.Windows.Forms.TextBox
Me.cmdSaveDBSetting = New System.Windows.Forms.Button
Me.lblUsername = New System.Windows.Forms.Label
Me.txtUsername = New System.Windows.Forms.TextBox
Me.lblPassword = New System.Windows.Forms.Label
Me.txtPassword = New System.Windows.Forms.TextBox
Me.SuspendLayout()
'
'txtPortNumber
'
Me.txtPortNumber.Location = New System.Drawing.Point(388, 13)
Me.txtPortNumber.Name = "txtPortNumber"
Me.txtPortNumber.Size = New System.Drawing.Size(44, 20)
Me.txtPortNumber.TabIndex = 2
Me.txtPortNumber.Text = "3306"
'
'lblPortNo
'
Me.lblPortNo.Location = New System.Drawing.Point(318, 16)
Me.lblPortNo.Name = "lblPortNo"
Me.lblPortNo.Size = New System.Drawing.Size(68, 16)
Me.lblPortNo.TabIndex = 18
Me.lblPortNo.Text = "Port Number"
'
'lblSecTimeout
'
Me.lblSecTimeout.Location = New System.Drawing.Point(170, 121)
Me.lblSecTimeout.Name = "lblSecTimeout"
Me.lblSecTimeout.Size = New System.Drawing.Size(54, 16)
Me.lblSecTimeout.TabIndex = 17
Me.lblSecTimeout.Text = "second(s)"
'
'lblTimeout
'
Me.lblTimeout.Location = New System.Drawing.Point(7, 120)
Me.lblTimeout.Name = "lblTimeout"
Me.lblTimeout.Size = New System.Drawing.Size(107, 15)
Me.lblTimeout.TabIndex = 15
Me.lblTimeout.Text = "Connection Timeout"
'
'lblServerInfo
'
Me.lblServerInfo.Location = New System.Drawing.Point(7, 94)
Me.lblServerInfo.Name = "lblServerInfo"
Me.lblServerInfo.Size = New System.Drawing.Size(102, 16)
Me.lblServerInfo.TabIndex = 13
Me.lblServerInfo.Text = "mySQL Server Info"
'
'cmdTestConnection
'
Me.cmdTestConnection.Location = New System.Drawing.Point(441, 12)
Me.cmdTestConnection.Name = "cmdTestConnection"
Me.cmdTestConnection.Size = New System.Drawing.Size(115, 22)
Me.cmdTestConnection.TabIndex = 12
Me.cmdTestConnection.Text = "Test Connection"
'
'txtmySQLServer
'
Me.txtmySQLServer.Location = New System.Drawing.Point(115, 13)
Me.txtmySQLServer.Name = "txtmySQLServer"
Me.txtmySQLServer.Size = New System.Drawing.Size(194, 20)
Me.txtmySQLServer.TabIndex = 1
Me.txtmySQLServer.Text = ""
'
'lblmySQLServer
'
Me.lblmySQLServer.Location = New System.Drawing.Point(7, 14)
Me.lblmySQLServer.Name = "lblmySQLServer"
Me.lblmySQLServer.Size = New System.Drawing.Size(84, 16)
Me.lblmySQLServer.TabIndex = 10
Me.lblmySQLServer.Text = "mySQL Server"
'
'lblDatabaseName
'
Me.lblDatabaseName.Location = New System.Drawing.Point(7, 64)
Me.lblDatabaseName.Name = "lblDatabaseName"
Me.lblDatabaseName.Size = New System.Drawing.Size(93, 16)
Me.lblDatabaseName.TabIndex = 21
Me.lblDatabaseName.Text = "Database Name"
'
'txtDatabaseName
'
Me.txtDatabaseName.Location = New System.Drawing.Point(115, 65)
Me.txtDatabaseName.Name = "txtDatabaseName"
Me.txtDatabaseName.Size = New System.Drawing.Size(193, 20)
Me.txtDatabaseName.TabIndex = 5
Me.txtDatabaseName.Text = ""
'
'txtTimeout
'
Me.txtTimeout.Enabled = False
Me.txtTimeout.Location = New System.Drawing.Point(115, 117)
Me.txtTimeout.Name = "txtTimeout"
Me.txtTimeout.Size = New System.Drawing.Size(51, 20)
Me.txtTimeout.TabIndex = 6
Me.txtTimeout.Text = ""
'
'txtMySQLInfo
'
Me.txtMySQLInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
Me.txtMySQLInfo.Location = New System.Drawing.Point(115, 91)
Me.txtMySQLInfo.Name = "txtMySQLInfo"
Me.txtMySQLInfo.ReadOnly = True
Me.txtMySQLInfo.Size = New System.Drawing.Size(441, 20)
Me.txtMySQLInfo.TabIndex = 14
Me.txtMySQLInfo.Text = ""
'
'cmdSaveDBSetting
'
Me.cmdSaveDBSetting.Location = New System.Drawing.Point(442, 117)
Me.cmdSaveDBSetting.Name = "cmdSaveDBSetting"
Me.cmdSaveDBSetting.Size = New System.Drawing.Size(114, 23)
Me.cmdSaveDBSetting.TabIndex = 23
Me.cmdSaveDBSetting.Text = "Save DB Settings"
'
'lblUsername
'
Me.lblUsername.Location = New System.Drawing.Point(8, 39)
Me.lblUsername.Name = "lblUsername"
Me.lblUsername.Size = New System.Drawing.Size(62, 16)
Me.lblUsername.TabIndex = 24
Me.lblUsername.Text = "Username"
'
'txtUsername
'
Me.txtUsername.Location = New System.Drawing.Point(115, 39)
Me.txtUsername.Name = "txtUsername"
Me.txtUsername.Size = New System.Drawing.Size(130, 20)
Me.txtUsername.TabIndex = 3
Me.txtUsername.Text = ""
'
'lblPassword
'
Me.lblPassword.Location = New System.Drawing.Point(264, 42)
Me.lblPassword.Name = "lblPassword"
Me.lblPassword.Size = New System.Drawing.Size(54, 16)
Me.lblPassword.TabIndex = 26
Me.lblPassword.Text = "Password"
'
'txtPassword
'
Me.txtPassword.Location = New System.Drawing.Point(322, 39)
Me.txtPassword.Name = "txtPassword"
Me.txtPassword.Size = New System.Drawing.Size(130, 20)
Me.txtPassword.TabIndex = 4
Me.txtPassword.Text = ""
'
'frmDBSettings
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(564, 153)
Me.Controls.Add(Me.lblPassword)
Me.Controls.Add(Me.txtPassword)
Me.Controls.Add(Me.lblUsername)
Me.Controls.Add(Me.txtUsername)
Me.Controls.Add(Me.cmdSaveDBSetting)
Me.Controls.Add(Me.lblDatabaseName)
Me.Controls.Add(Me.txtPortNumber)
Me.Controls.Add(Me.lblPortNo)
Me.Controls.Add(Me.lblSecTimeout)
Me.Controls.Add(Me.lblTimeout)
Me.Controls.Add(Me.lblServerInfo)
Me.Controls.Add(Me.cmdTestConnection)
Me.Controls.Add(Me.txtmySQLServer)
Me.Controls.Add(Me.lblmySQLServer)
Me.Controls.Add(Me.txtDatabaseName)
Me.Controls.Add(Me.txtTimeout)
Me.Controls.Add(Me.txtMySQLInfo)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
Me.MaximizeBox = False
Me.Name = "frmDBSettings"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "Database Settings"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Local Variables "

  Dim cnDBConnection As MySqlConnection

#End Region

#Region " Custom Properties "

#End Region

#Region " Event Codes "

Private Sub cmdTestConnection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestConnection.Click
  On Error GoTo ErrorHandling

  Dim cnTestmySQL As MySqlConnection
  Dim blnSuccess As Boolean
  Dim strTempServerString As String


  'Clear up the last test result
  txtMySQLInfo.Text = ""

  'Check if user specified a different port
  If Val(txtPortNumber.Text) <> Val(MCE_MYSQL_DEF_PORT) Then
    'If the port is different, recompose the connection string
    strTempServerString = Trim(txtmySQLServer.Text) & ";Port=" & txtPortNumber.Text
  Else
    'If not, just pass the server name
    strTempServerString = Trim(txtmySQLServer.Text)
  End If

  'Create a new instance of the mySQL connection  
  blnSuccess = ConnectToMySQLDB(strTempServerString, Trim(txtUsername.Text), Trim(txtPassword.Text), Trim(txtDatabaseName.Text), cnTestmySQL)

  If blnSuccess = True Then
    'If it's successful, display some information about the server
    txtMySQLInfo.Text = "[Current Connection Timeout : " & cnTestmySQL.ConnectionTimeout & " sec(s)] " & _
    "[Server Version : " & cnTestmySQL.ServerVersion & "] "
    MsgBox("Connection was successful")
  Else
    MsgBox("Cannot connect to server " & txtmySQLServer.Text, MsgBoxStyle.Critical, "Cannot connect to server")
    Exit Sub
  End If

  Exit Sub

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("frmDBSettings.vb", "cmdTestConnection_Click", Err.Number, Err.Source, Err.Description)
End Sub

Private Sub cmdSaveDBSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveDBSetting.Click
  On Error GoTo ErrorHandling

  'Saving settings via INI formatted XML file
  SaveSettingToINIXML(Application.StartupPath & "\" & MCE_CONFIG_FILE)

  'Just notify people that it's done
  MsgBox("Settings saved.", MsgBoxStyle.Information, "Settings saved")

  Exit Sub

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("frmDBSettings.vb", "cmdSaveDBSetting_Click", Err.Number, Err.Source, Err.Description)
End Sub

#End Region

#Region " Procedures / Functions "

Private Sub SaveSettingToINIXML(ByVal strCompleteFilename As String)
'-------------------------------------------------------------------------------
'Purpose     : This procedure is to save all the settings here into INI (XML formatted) file
'Assumptions :
'Input       :
'   - strCompleteFilename, a string consisting the complete filename including 
'Returns     :
'Note        : 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim blnSuccess As Boolean
  Dim strTemp As String


  'Save the database values 
  If Trim(txtmySQLServer.Text) = "" Then
    'To make sure at least we have a value
    blnSuccess = SaveSettingsInXML("DB Settings", "Server Name", Trim(MCE_MYSQL_DEF_SERVER), strCompleteFilename)
  Else
    blnSuccess = SaveSettingsInXML("DB Settings", "Server Name", Trim(txtmySQLServer.Text), strCompleteFilename)
  End If
  If Trim(txtPortNumber.Text) = "" Then
    blnSuccess = SaveSettingsInXML("DB Settings", "Port Number", Trim(MCE_MYSQL_DEF_PORT), strCompleteFilename)
  Else
    blnSuccess = SaveSettingsInXML("DB Settings", "Port Number", Trim(txtPortNumber.Text), strCompleteFilename)
  End If
  blnSuccess = SaveSettingsInXML("DB Settings", "Connection Timeout", Trim(txtTimeout.Text), strCompleteFilename)

  'Save database name
  blnSuccess = SaveSettingsInXML("DB Settings", "Database", Trim(txtDatabaseName.Text), strCompleteFilename)
  'Save Username 
  blnSuccess = SaveSettingsInXML("DB Settings", "Username", Trim(txtUsername.Text), strCompleteFilename)
  'Save Password
  blnSuccess = SaveSettingsInXML("DB Settings", "Password", Trim(txtPassword.Text), strCompleteFilename)

  Exit Sub

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("frmDBSettings.vb", "SaveSettingToINIXML", Err.Number, Err.Source, Err.Description)
End Sub

#End Region

End Class
