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

Public Class frmMCE
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
  Friend WithEvents staMCE As System.Windows.Forms.StatusBar
  Friend WithEvents grpCharList As System.Windows.Forms.GroupBox
  Friend WithEvents lstCharList As System.Windows.Forms.ListBox
  Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
  Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
  Friend WithEvents mnuConnect As System.Windows.Forms.MenuItem
  Friend WithEvents mnuMCE As System.Windows.Forms.MainMenu
  Friend WithEvents mnuSetting As System.Windows.Forms.MenuItem
  Friend WithEvents mnuSeparator21 As System.Windows.Forms.MenuItem
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents grpStats As System.Windows.Forms.GroupBox
  Friend WithEvents lblXPToLvl As System.Windows.Forms.Label
  Friend WithEvents lblCurrentXP As System.Windows.Forms.Label
  Friend WithEvents txtMaxMana As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents txtMaxHealth As System.Windows.Forms.TextBox
  Friend WithEvents lblMaxHealthMana As System.Windows.Forms.Label
  Friend WithEvents txtCurrMana As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents txtCurrHealth As System.Windows.Forms.TextBox
  Friend WithEvents lblCrntHealthMana As System.Windows.Forms.Label
  Friend WithEvents txtSPI As System.Windows.Forms.TextBox
  Friend WithEvents lblSPI As System.Windows.Forms.Label
  Friend WithEvents txtINT As System.Windows.Forms.TextBox
  Friend WithEvents lblINT As System.Windows.Forms.Label
  Friend WithEvents txtSTA As System.Windows.Forms.TextBox
  Friend WithEvents lblSTA As System.Windows.Forms.Label
  Friend WithEvents txtAGI As System.Windows.Forms.TextBox
  Friend WithEvents lblAGI As System.Windows.Forms.Label
  Friend WithEvents txtSTR As System.Windows.Forms.TextBox
  Friend WithEvents lblSTR As System.Windows.Forms.Label
  Friend WithEvents grpDetails As System.Windows.Forms.GroupBox
  Friend WithEvents lblClass As System.Windows.Forms.Label
  Friend WithEvents lblRace As System.Windows.Forms.Label
  Friend WithEvents txtLevel As System.Windows.Forms.TextBox
  Friend WithEvents lblLevel As System.Windows.Forms.Label
  Friend WithEvents txtName As System.Windows.Forms.TextBox
  Friend WithEvents txtGUID As System.Windows.Forms.TextBox
  Friend WithEvents lblName As System.Windows.Forms.Label
  Friend WithEvents tabMCE As System.Windows.Forms.TabControl
  Friend WithEvents tabPet As System.Windows.Forms.TabPage
  Friend WithEvents grpSkillTalent As System.Windows.Forms.GroupBox
  Friend WithEvents lblSkillInfo As System.Windows.Forms.Label
  Friend WithEvents txtSkillInfo As System.Windows.Forms.TextBox
  Friend WithEvents lblCharPoints1 As System.Windows.Forms.Label
  Friend WithEvents lblCharPoints2 As System.Windows.Forms.Label
  Friend WithEvents txtTraPoints As System.Windows.Forms.TextBox
  Friend WithEvents lblTraPoints As System.Windows.Forms.Label
  Friend WithEvents cmdRefreshCharList As System.Windows.Forms.Button
  Friend WithEvents pnlDBStatus As System.Windows.Forms.StatusBarPanel
  Friend WithEvents txtCharPoint2 As System.Windows.Forms.TextBox
  Friend WithEvents txtCharPoint1 As System.Windows.Forms.TextBox
  Friend WithEvents txtXPToLvl As System.Windows.Forms.TextBox
  Friend WithEvents txtCurrXP As System.Windows.Forms.TextBox
  Friend WithEvents txtClass As System.Windows.Forms.TextBox
  Friend WithEvents txtRace As System.Windows.Forms.TextBox
  Friend WithEvents lblGUID As System.Windows.Forms.Label
  Friend WithEvents cmdRefreshCharData As System.Windows.Forms.Button
  Friend WithEvents mnuDatabase As System.Windows.Forms.MenuItem
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents txtXPNeeded As System.Windows.Forms.TextBox
Friend WithEvents tabMiscInfo As System.Windows.Forms.TabPage
Friend WithEvents lblLastLogoutTime As System.Windows.Forms.Label
Friend WithEvents txtLastLogoutTime As System.Windows.Forms.TextBox
Friend WithEvents tabCharacter1 As System.Windows.Forms.TabPage
Friend WithEvents tabCharacter2 As System.Windows.Forms.TabPage
Friend WithEvents grpCombatStat As System.Windows.Forms.GroupBox
Friend WithEvents txtParryRate As System.Windows.Forms.TextBox
Friend WithEvents lblParryRate As System.Windows.Forms.Label
Friend WithEvents txtRngAttackPwr As System.Windows.Forms.TextBox
Friend WithEvents Label7 As System.Windows.Forms.Label
Friend WithEvents txtAttackPwr As System.Windows.Forms.TextBox
Friend WithEvents lblATCKPower As System.Windows.Forms.Label
Friend WithEvents txtDodgeRate As System.Windows.Forms.TextBox
Friend WithEvents lblDodgeRate As System.Windows.Forms.Label
Friend WithEvents txtMaxRngDmg As System.Windows.Forms.TextBox
Friend WithEvents Label5 As System.Windows.Forms.Label
Friend WithEvents txtMinRngDmg As System.Windows.Forms.TextBox
Friend WithEvents lblMinMaxRngDmg As System.Windows.Forms.Label
Friend WithEvents txtBlockRate As System.Windows.Forms.TextBox
Friend WithEvents lblBlockRate As System.Windows.Forms.Label
Friend WithEvents txtMaxDmg As System.Windows.Forms.TextBox
Friend WithEvents Label4 As System.Windows.Forms.Label
Friend WithEvents txtMinDmg As System.Windows.Forms.TextBox
Friend WithEvents lblMinDmg As System.Windows.Forms.Label
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.staMCE = New System.Windows.Forms.StatusBar
Me.pnlDBStatus = New System.Windows.Forms.StatusBarPanel
Me.grpCharList = New System.Windows.Forms.GroupBox
Me.cmdRefreshCharList = New System.Windows.Forms.Button
Me.lstCharList = New System.Windows.Forms.ListBox
Me.mnuMCE = New System.Windows.Forms.MainMenu
Me.mnuFile = New System.Windows.Forms.MenuItem
Me.mnuExit = New System.Windows.Forms.MenuItem
Me.mnuDatabase = New System.Windows.Forms.MenuItem
Me.mnuConnect = New System.Windows.Forms.MenuItem
Me.mnuSeparator21 = New System.Windows.Forms.MenuItem
Me.mnuSetting = New System.Windows.Forms.MenuItem
Me.cmdSave = New System.Windows.Forms.Button
Me.tabMCE = New System.Windows.Forms.TabControl
Me.tabCharacter1 = New System.Windows.Forms.TabPage
Me.grpSkillTalent = New System.Windows.Forms.GroupBox
Me.txtTraPoints = New System.Windows.Forms.TextBox
Me.lblTraPoints = New System.Windows.Forms.Label
Me.txtCharPoint2 = New System.Windows.Forms.TextBox
Me.lblCharPoints2 = New System.Windows.Forms.Label
Me.txtCharPoint1 = New System.Windows.Forms.TextBox
Me.lblCharPoints1 = New System.Windows.Forms.Label
Me.txtSkillInfo = New System.Windows.Forms.TextBox
Me.lblSkillInfo = New System.Windows.Forms.Label
Me.grpStats = New System.Windows.Forms.GroupBox
Me.txtXPNeeded = New System.Windows.Forms.TextBox
Me.Label1 = New System.Windows.Forms.Label
Me.txtXPToLvl = New System.Windows.Forms.TextBox
Me.lblXPToLvl = New System.Windows.Forms.Label
Me.txtCurrXP = New System.Windows.Forms.TextBox
Me.lblCurrentXP = New System.Windows.Forms.Label
Me.txtMaxMana = New System.Windows.Forms.TextBox
Me.Label3 = New System.Windows.Forms.Label
Me.txtMaxHealth = New System.Windows.Forms.TextBox
Me.lblMaxHealthMana = New System.Windows.Forms.Label
Me.txtCurrMana = New System.Windows.Forms.TextBox
Me.Label2 = New System.Windows.Forms.Label
Me.txtCurrHealth = New System.Windows.Forms.TextBox
Me.lblCrntHealthMana = New System.Windows.Forms.Label
Me.txtSPI = New System.Windows.Forms.TextBox
Me.lblSPI = New System.Windows.Forms.Label
Me.txtINT = New System.Windows.Forms.TextBox
Me.lblINT = New System.Windows.Forms.Label
Me.txtSTA = New System.Windows.Forms.TextBox
Me.lblSTA = New System.Windows.Forms.Label
Me.txtAGI = New System.Windows.Forms.TextBox
Me.lblAGI = New System.Windows.Forms.Label
Me.txtSTR = New System.Windows.Forms.TextBox
Me.lblSTR = New System.Windows.Forms.Label
Me.grpDetails = New System.Windows.Forms.GroupBox
Me.txtClass = New System.Windows.Forms.TextBox
Me.lblClass = New System.Windows.Forms.Label
Me.txtRace = New System.Windows.Forms.TextBox
Me.lblRace = New System.Windows.Forms.Label
Me.txtLevel = New System.Windows.Forms.TextBox
Me.lblLevel = New System.Windows.Forms.Label
Me.txtName = New System.Windows.Forms.TextBox
Me.txtGUID = New System.Windows.Forms.TextBox
Me.lblGUID = New System.Windows.Forms.Label
Me.lblName = New System.Windows.Forms.Label
Me.tabPet = New System.Windows.Forms.TabPage
Me.cmdRefreshCharData = New System.Windows.Forms.Button
Me.tabMiscInfo = New System.Windows.Forms.TabPage
Me.lblLastLogoutTime = New System.Windows.Forms.Label
Me.txtLastLogoutTime = New System.Windows.Forms.TextBox
Me.tabCharacter2 = New System.Windows.Forms.TabPage
Me.grpCombatStat = New System.Windows.Forms.GroupBox
Me.txtParryRate = New System.Windows.Forms.TextBox
Me.lblParryRate = New System.Windows.Forms.Label
Me.txtRngAttackPwr = New System.Windows.Forms.TextBox
Me.Label7 = New System.Windows.Forms.Label
Me.txtAttackPwr = New System.Windows.Forms.TextBox
Me.lblATCKPower = New System.Windows.Forms.Label
Me.txtDodgeRate = New System.Windows.Forms.TextBox
Me.lblDodgeRate = New System.Windows.Forms.Label
Me.txtMaxRngDmg = New System.Windows.Forms.TextBox
Me.Label5 = New System.Windows.Forms.Label
Me.txtMinRngDmg = New System.Windows.Forms.TextBox
Me.lblMinMaxRngDmg = New System.Windows.Forms.Label
Me.txtBlockRate = New System.Windows.Forms.TextBox
Me.lblBlockRate = New System.Windows.Forms.Label
Me.txtMaxDmg = New System.Windows.Forms.TextBox
Me.Label4 = New System.Windows.Forms.Label
Me.txtMinDmg = New System.Windows.Forms.TextBox
Me.lblMinDmg = New System.Windows.Forms.Label
CType(Me.pnlDBStatus, System.ComponentModel.ISupportInitialize).BeginInit()
Me.grpCharList.SuspendLayout()
Me.tabMCE.SuspendLayout()
Me.tabCharacter1.SuspendLayout()
Me.grpSkillTalent.SuspendLayout()
Me.grpStats.SuspendLayout()
Me.grpDetails.SuspendLayout()
Me.tabMiscInfo.SuspendLayout()
Me.tabCharacter2.SuspendLayout()
Me.grpCombatStat.SuspendLayout()
Me.SuspendLayout()
'
'staMCE
'
Me.staMCE.Location = New System.Drawing.Point(0, 601)
Me.staMCE.Name = "staMCE"
Me.staMCE.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.pnlDBStatus})
Me.staMCE.ShowPanels = True
Me.staMCE.Size = New System.Drawing.Size(586, 22)
Me.staMCE.TabIndex = 1
'
'pnlDBStatus
'
Me.pnlDBStatus.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
Me.pnlDBStatus.Width = 570
'
'grpCharList
'
Me.grpCharList.Controls.Add(Me.cmdRefreshCharList)
Me.grpCharList.Controls.Add(Me.lstCharList)
Me.grpCharList.Location = New System.Drawing.Point(8, 4)
Me.grpCharList.Name = "grpCharList"
Me.grpCharList.Size = New System.Drawing.Size(153, 593)
Me.grpCharList.TabIndex = 4
Me.grpCharList.TabStop = False
Me.grpCharList.Text = "Character List"
'
'cmdRefreshCharList
'
Me.cmdRefreshCharList.Location = New System.Drawing.Point(7, 563)
Me.cmdRefreshCharList.Name = "cmdRefreshCharList"
Me.cmdRefreshCharList.Size = New System.Drawing.Size(135, 23)
Me.cmdRefreshCharList.TabIndex = 10
Me.cmdRefreshCharList.Text = "Refresh Character List"
'
'lstCharList
'
Me.lstCharList.Location = New System.Drawing.Point(7, 19)
Me.lstCharList.Name = "lstCharList"
Me.lstCharList.ScrollAlwaysVisible = True
Me.lstCharList.Size = New System.Drawing.Size(137, 537)
Me.lstCharList.Sorted = True
Me.lstCharList.TabIndex = 1
'
'mnuMCE
'
Me.mnuMCE.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuDatabase})
'
'mnuFile
'
Me.mnuFile.Index = 0
Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExit})
Me.mnuFile.Text = "&File"
'
'mnuExit
'
Me.mnuExit.Index = 0
Me.mnuExit.Text = "E&xit"
'
'mnuDatabase
'
Me.mnuDatabase.Index = 1
Me.mnuDatabase.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuConnect, Me.mnuSeparator21, Me.mnuSetting})
Me.mnuDatabase.Text = "&Database"
'
'mnuConnect
'
Me.mnuConnect.Index = 0
Me.mnuConnect.Text = "&Connect to Database"
'
'mnuSeparator21
'
Me.mnuSeparator21.Index = 1
Me.mnuSeparator21.Text = "-"
'
'mnuSetting
'
Me.mnuSetting.Index = 2
Me.mnuSetting.Text = "&Settings"
'
'cmdSave
'
Me.cmdSave.Location = New System.Drawing.Point(329, 572)
Me.cmdSave.Name = "cmdSave"
Me.cmdSave.Size = New System.Drawing.Size(248, 23)
Me.cmdSave.TabIndex = 7
Me.cmdSave.Text = "Save Changes to Mangos Database"
'
'tabMCE
'
Me.tabMCE.Controls.Add(Me.tabCharacter1)
Me.tabMCE.Controls.Add(Me.tabCharacter2)
Me.tabMCE.Controls.Add(Me.tabPet)
Me.tabMCE.Controls.Add(Me.tabMiscInfo)
Me.tabMCE.Location = New System.Drawing.Point(166, 9)
Me.tabMCE.Name = "tabMCE"
Me.tabMCE.SelectedIndex = 0
Me.tabMCE.Size = New System.Drawing.Size(413, 560)
Me.tabMCE.TabIndex = 8
'
'tabCharacter1
'
Me.tabCharacter1.Controls.Add(Me.grpSkillTalent)
Me.tabCharacter1.Controls.Add(Me.grpStats)
Me.tabCharacter1.Controls.Add(Me.grpDetails)
Me.tabCharacter1.Location = New System.Drawing.Point(4, 22)
Me.tabCharacter1.Name = "tabCharacter1"
Me.tabCharacter1.Size = New System.Drawing.Size(405, 534)
Me.tabCharacter1.TabIndex = 0
Me.tabCharacter1.Text = "Character Data (1/2)"
'
'grpSkillTalent
'
Me.grpSkillTalent.Controls.Add(Me.txtTraPoints)
Me.grpSkillTalent.Controls.Add(Me.lblTraPoints)
Me.grpSkillTalent.Controls.Add(Me.txtCharPoint2)
Me.grpSkillTalent.Controls.Add(Me.lblCharPoints2)
Me.grpSkillTalent.Controls.Add(Me.txtCharPoint1)
Me.grpSkillTalent.Controls.Add(Me.lblCharPoints1)
Me.grpSkillTalent.Controls.Add(Me.txtSkillInfo)
Me.grpSkillTalent.Controls.Add(Me.lblSkillInfo)
Me.grpSkillTalent.Location = New System.Drawing.Point(8, 377)
Me.grpSkillTalent.Name = "grpSkillTalent"
Me.grpSkillTalent.Size = New System.Drawing.Size(393, 139)
Me.grpSkillTalent.TabIndex = 10
Me.grpSkillTalent.TabStop = False
Me.grpSkillTalent.Text = "Skill / Talents"
'
'txtTraPoints
'
Me.txtTraPoints.Location = New System.Drawing.Point(80, 98)
Me.txtTraPoints.Name = "txtTraPoints"
Me.txtTraPoints.Size = New System.Drawing.Size(55, 20)
Me.txtTraPoints.TabIndex = 17
Me.txtTraPoints.Text = ""
'
'lblTraPoints
'
Me.lblTraPoints.Location = New System.Drawing.Point(9, 97)
Me.lblTraPoints.Name = "lblTraPoints"
Me.lblTraPoints.Size = New System.Drawing.Size(58, 25)
Me.lblTraPoints.TabIndex = 16
Me.lblTraPoints.Text = "Training Points"
'
'txtCharPoint2
'
Me.txtCharPoint2.Location = New System.Drawing.Point(80, 72)
Me.txtCharPoint2.Name = "txtCharPoint2"
Me.txtCharPoint2.Size = New System.Drawing.Size(306, 20)
Me.txtCharPoint2.TabIndex = 5
Me.txtCharPoint2.Text = ""
'
'lblCharPoints2
'
Me.lblCharPoints2.Location = New System.Drawing.Point(8, 73)
Me.lblCharPoints2.Name = "lblCharPoints2"
Me.lblCharPoints2.Size = New System.Drawing.Size(73, 15)
Me.lblCharPoints2.TabIndex = 4
Me.lblCharPoints2.Text = "Char Points 2"
'
'txtCharPoint1
'
Me.txtCharPoint1.Location = New System.Drawing.Point(80, 46)
Me.txtCharPoint1.Name = "txtCharPoint1"
Me.txtCharPoint1.Size = New System.Drawing.Size(306, 20)
Me.txtCharPoint1.TabIndex = 3
Me.txtCharPoint1.Text = ""
'
'lblCharPoints1
'
Me.lblCharPoints1.Location = New System.Drawing.Point(8, 47)
Me.lblCharPoints1.Name = "lblCharPoints1"
Me.lblCharPoints1.Size = New System.Drawing.Size(73, 15)
Me.lblCharPoints1.TabIndex = 2
Me.lblCharPoints1.Text = "Char Points 1"
'
'txtSkillInfo
'
Me.txtSkillInfo.Location = New System.Drawing.Point(80, 20)
Me.txtSkillInfo.Name = "txtSkillInfo"
Me.txtSkillInfo.Size = New System.Drawing.Size(306, 20)
Me.txtSkillInfo.TabIndex = 1
Me.txtSkillInfo.Text = ""
'
'lblSkillInfo
'
Me.lblSkillInfo.Location = New System.Drawing.Point(8, 22)
Me.lblSkillInfo.Name = "lblSkillInfo"
Me.lblSkillInfo.Size = New System.Drawing.Size(51, 15)
Me.lblSkillInfo.TabIndex = 0
Me.lblSkillInfo.Text = "Skill Info"
'
'grpStats
'
Me.grpStats.Controls.Add(Me.txtXPNeeded)
Me.grpStats.Controls.Add(Me.Label1)
Me.grpStats.Controls.Add(Me.txtXPToLvl)
Me.grpStats.Controls.Add(Me.lblXPToLvl)
Me.grpStats.Controls.Add(Me.txtCurrXP)
Me.grpStats.Controls.Add(Me.lblCurrentXP)
Me.grpStats.Controls.Add(Me.txtMaxMana)
Me.grpStats.Controls.Add(Me.Label3)
Me.grpStats.Controls.Add(Me.txtMaxHealth)
Me.grpStats.Controls.Add(Me.lblMaxHealthMana)
Me.grpStats.Controls.Add(Me.txtCurrMana)
Me.grpStats.Controls.Add(Me.Label2)
Me.grpStats.Controls.Add(Me.txtCurrHealth)
Me.grpStats.Controls.Add(Me.lblCrntHealthMana)
Me.grpStats.Controls.Add(Me.txtSPI)
Me.grpStats.Controls.Add(Me.lblSPI)
Me.grpStats.Controls.Add(Me.txtINT)
Me.grpStats.Controls.Add(Me.lblINT)
Me.grpStats.Controls.Add(Me.txtSTA)
Me.grpStats.Controls.Add(Me.lblSTA)
Me.grpStats.Controls.Add(Me.txtAGI)
Me.grpStats.Controls.Add(Me.lblAGI)
Me.grpStats.Controls.Add(Me.txtSTR)
Me.grpStats.Controls.Add(Me.lblSTR)
Me.grpStats.Location = New System.Drawing.Point(7, 87)
Me.grpStats.Name = "grpStats"
Me.grpStats.Size = New System.Drawing.Size(393, 280)
Me.grpStats.TabIndex = 8
Me.grpStats.TabStop = False
Me.grpStats.Text = "Base Stats"
'
'txtXPNeeded
'
Me.txtXPNeeded.Location = New System.Drawing.Point(273, 74)
Me.txtXPNeeded.Name = "txtXPNeeded"
Me.txtXPNeeded.ReadOnly = True
Me.txtXPNeeded.Size = New System.Drawing.Size(90, 20)
Me.txtXPNeeded.TabIndex = 25
Me.txtXPNeeded.Text = ""
'
'Label1
'
Me.Label1.Location = New System.Drawing.Point(211, 77)
Me.Label1.Name = "Label1"
Me.Label1.Size = New System.Drawing.Size(60, 14)
Me.Label1.TabIndex = 24
Me.Label1.Text = "XP needed"
'
'txtXPToLvl
'
Me.txtXPToLvl.Location = New System.Drawing.Point(100, 74)
Me.txtXPToLvl.Name = "txtXPToLvl"
Me.txtXPToLvl.Size = New System.Drawing.Size(90, 20)
Me.txtXPToLvl.TabIndex = 23
Me.txtXPToLvl.Text = ""
'
'lblXPToLvl
'
Me.lblXPToLvl.Location = New System.Drawing.Point(11, 77)
Me.lblXPToLvl.Name = "lblXPToLvl"
Me.lblXPToLvl.Size = New System.Drawing.Size(89, 18)
Me.lblXPToLvl.TabIndex = 22
Me.lblXPToLvl.Text = "XP for next level"
'
'txtCurrXP
'
Me.txtCurrXP.Location = New System.Drawing.Point(158, 47)
Me.txtCurrXP.Name = "txtCurrXP"
Me.txtCurrXP.Size = New System.Drawing.Size(90, 20)
Me.txtCurrXP.TabIndex = 21
Me.txtCurrXP.Text = ""
'
'lblCurrentXP
'
Me.lblCurrentXP.Location = New System.Drawing.Point(99, 51)
Me.lblCurrentXP.Name = "lblCurrentXP"
Me.lblCurrentXP.Size = New System.Drawing.Size(60, 14)
Me.lblCurrentXP.TabIndex = 20
Me.lblCurrentXP.Text = "Current XP"
'
'txtMaxMana
'
Me.txtMaxMana.Location = New System.Drawing.Point(193, 134)
Me.txtMaxMana.Name = "txtMaxMana"
Me.txtMaxMana.Size = New System.Drawing.Size(55, 20)
Me.txtMaxMana.TabIndex = 17
Me.txtMaxMana.Text = ""
'
'Label3
'
Me.Label3.Location = New System.Drawing.Point(175, 138)
Me.Label3.Name = "Label3"
Me.Label3.Size = New System.Drawing.Size(12, 16)
Me.Label3.TabIndex = 16
Me.Label3.Text = "/"
'
'txtMaxHealth
'
Me.txtMaxHealth.Location = New System.Drawing.Point(193, 103)
Me.txtMaxHealth.Name = "txtMaxHealth"
Me.txtMaxHealth.Size = New System.Drawing.Size(55, 20)
Me.txtMaxHealth.TabIndex = 15
Me.txtMaxHealth.Text = ""
'
'lblMaxHealthMana
'
Me.lblMaxHealthMana.Location = New System.Drawing.Point(8, 135)
Me.lblMaxHealthMana.Name = "lblMaxHealthMana"
Me.lblMaxHealthMana.Size = New System.Drawing.Size(98, 15)
Me.lblMaxHealthMana.TabIndex = 14
Me.lblMaxHealthMana.Text = "Current/Max Mana"
'
'txtCurrMana
'
Me.txtCurrMana.Location = New System.Drawing.Point(112, 134)
Me.txtCurrMana.Name = "txtCurrMana"
Me.txtCurrMana.Size = New System.Drawing.Size(55, 20)
Me.txtCurrMana.TabIndex = 13
Me.txtCurrMana.Text = ""
'
'Label2
'
Me.Label2.Location = New System.Drawing.Point(175, 108)
Me.Label2.Name = "Label2"
Me.Label2.Size = New System.Drawing.Size(12, 16)
Me.Label2.TabIndex = 12
Me.Label2.Text = "/"
'
'txtCurrHealth
'
Me.txtCurrHealth.Location = New System.Drawing.Point(112, 103)
Me.txtCurrHealth.Name = "txtCurrHealth"
Me.txtCurrHealth.Size = New System.Drawing.Size(55, 20)
Me.txtCurrHealth.TabIndex = 11
Me.txtCurrHealth.Text = ""
'
'lblCrntHealthMana
'
Me.lblCrntHealthMana.Location = New System.Drawing.Point(8, 107)
Me.lblCrntHealthMana.Name = "lblCrntHealthMana"
Me.lblCrntHealthMana.Size = New System.Drawing.Size(102, 15)
Me.lblCrntHealthMana.TabIndex = 10
Me.lblCrntHealthMana.Text = "Current/Max Health"
'
'txtSPI
'
Me.txtSPI.Location = New System.Drawing.Point(39, 47)
Me.txtSPI.Name = "txtSPI"
Me.txtSPI.Size = New System.Drawing.Size(45, 20)
Me.txtSPI.TabIndex = 9
Me.txtSPI.Text = ""
'
'lblSPI
'
Me.lblSPI.Location = New System.Drawing.Point(8, 50)
Me.lblSPI.Name = "lblSPI"
Me.lblSPI.Size = New System.Drawing.Size(27, 14)
Me.lblSPI.TabIndex = 8
Me.lblSPI.Text = "SPI"
'
'txtINT
'
Me.txtINT.Location = New System.Drawing.Point(317, 19)
Me.txtINT.Name = "txtINT"
Me.txtINT.Size = New System.Drawing.Size(45, 20)
Me.txtINT.TabIndex = 7
Me.txtINT.Text = ""
'
'lblINT
'
Me.lblINT.Location = New System.Drawing.Point(291, 22)
Me.lblINT.Name = "lblINT"
Me.lblINT.Size = New System.Drawing.Size(27, 14)
Me.lblINT.TabIndex = 6
Me.lblINT.Text = "INT"
'
'txtSTA
'
Me.txtSTA.Location = New System.Drawing.Point(225, 19)
Me.txtSTA.Name = "txtSTA"
Me.txtSTA.Size = New System.Drawing.Size(45, 20)
Me.txtSTA.TabIndex = 5
Me.txtSTA.Text = ""
'
'lblSTA
'
Me.lblSTA.Location = New System.Drawing.Point(195, 22)
Me.lblSTA.Name = "lblSTA"
Me.lblSTA.Size = New System.Drawing.Size(27, 14)
Me.lblSTA.TabIndex = 4
Me.lblSTA.Text = "STA"
'
'txtAGI
'
Me.txtAGI.Location = New System.Drawing.Point(131, 19)
Me.txtAGI.Name = "txtAGI"
Me.txtAGI.Size = New System.Drawing.Size(45, 20)
Me.txtAGI.TabIndex = 3
Me.txtAGI.Text = ""
'
'lblAGI
'
Me.lblAGI.Location = New System.Drawing.Point(103, 22)
Me.lblAGI.Name = "lblAGI"
Me.lblAGI.Size = New System.Drawing.Size(27, 14)
Me.lblAGI.TabIndex = 2
Me.lblAGI.Text = "AGI"
'
'txtSTR
'
Me.txtSTR.Location = New System.Drawing.Point(39, 19)
Me.txtSTR.Name = "txtSTR"
Me.txtSTR.Size = New System.Drawing.Size(45, 20)
Me.txtSTR.TabIndex = 1
Me.txtSTR.Text = ""
'
'lblSTR
'
Me.lblSTR.Location = New System.Drawing.Point(8, 22)
Me.lblSTR.Name = "lblSTR"
Me.lblSTR.Size = New System.Drawing.Size(27, 14)
Me.lblSTR.TabIndex = 0
Me.lblSTR.Text = "STR"
'
'grpDetails
'
Me.grpDetails.Controls.Add(Me.txtClass)
Me.grpDetails.Controls.Add(Me.lblClass)
Me.grpDetails.Controls.Add(Me.txtRace)
Me.grpDetails.Controls.Add(Me.lblRace)
Me.grpDetails.Controls.Add(Me.txtLevel)
Me.grpDetails.Controls.Add(Me.lblLevel)
Me.grpDetails.Controls.Add(Me.txtName)
Me.grpDetails.Controls.Add(Me.txtGUID)
Me.grpDetails.Controls.Add(Me.lblGUID)
Me.grpDetails.Controls.Add(Me.lblName)
Me.grpDetails.Location = New System.Drawing.Point(7, 1)
Me.grpDetails.Name = "grpDetails"
Me.grpDetails.Size = New System.Drawing.Size(393, 79)
Me.grpDetails.TabIndex = 7
Me.grpDetails.TabStop = False
Me.grpDetails.Text = "Details"
'
'txtClass
'
Me.txtClass.BackColor = System.Drawing.Color.White
Me.txtClass.Location = New System.Drawing.Point(229, 47)
Me.txtClass.Name = "txtClass"
Me.txtClass.ReadOnly = True
Me.txtClass.Size = New System.Drawing.Size(134, 20)
Me.txtClass.TabIndex = 14
Me.txtClass.Text = ""
'
'lblClass
'
Me.lblClass.Location = New System.Drawing.Point(197, 50)
Me.lblClass.Name = "lblClass"
Me.lblClass.Size = New System.Drawing.Size(35, 16)
Me.lblClass.TabIndex = 13
Me.lblClass.Text = "Class"
'
'txtRace
'
Me.txtRace.BackColor = System.Drawing.Color.White
Me.txtRace.Location = New System.Drawing.Point(49, 47)
Me.txtRace.Name = "txtRace"
Me.txtRace.ReadOnly = True
Me.txtRace.Size = New System.Drawing.Size(134, 20)
Me.txtRace.TabIndex = 12
Me.txtRace.Text = ""
'
'lblRace
'
Me.lblRace.Location = New System.Drawing.Point(7, 50)
Me.lblRace.Name = "lblRace"
Me.lblRace.Size = New System.Drawing.Size(32, 16)
Me.lblRace.TabIndex = 11
Me.lblRace.Text = "Race"
'
'txtLevel
'
Me.txtLevel.BackColor = System.Drawing.Color.White
Me.txtLevel.Location = New System.Drawing.Point(333, 18)
Me.txtLevel.Name = "txtLevel"
Me.txtLevel.ReadOnly = True
Me.txtLevel.Size = New System.Drawing.Size(49, 20)
Me.txtLevel.TabIndex = 10
Me.txtLevel.Text = ""
'
'lblLevel
'
Me.lblLevel.Location = New System.Drawing.Point(296, 21)
Me.lblLevel.Name = "lblLevel"
Me.lblLevel.Size = New System.Drawing.Size(32, 16)
Me.lblLevel.TabIndex = 9
Me.lblLevel.Text = "Level"
'
'txtName
'
Me.txtName.BackColor = System.Drawing.Color.White
Me.txtName.Location = New System.Drawing.Point(152, 18)
Me.txtName.Name = "txtName"
Me.txtName.ReadOnly = True
Me.txtName.Size = New System.Drawing.Size(133, 20)
Me.txtName.TabIndex = 8
Me.txtName.Text = ""
'
'txtGUID
'
Me.txtGUID.BackColor = System.Drawing.Color.White
Me.txtGUID.Location = New System.Drawing.Point(49, 18)
Me.txtGUID.Name = "txtGUID"
Me.txtGUID.ReadOnly = True
Me.txtGUID.Size = New System.Drawing.Size(49, 20)
Me.txtGUID.TabIndex = 7
Me.txtGUID.Text = ""
'
'lblGUID
'
Me.lblGUID.Location = New System.Drawing.Point(7, 21)
Me.lblGUID.Name = "lblGUID"
Me.lblGUID.Size = New System.Drawing.Size(43, 16)
Me.lblGUID.TabIndex = 6
Me.lblGUID.Text = "GUID :"
'
'lblName
'
Me.lblName.Location = New System.Drawing.Point(112, 22)
Me.lblName.Name = "lblName"
Me.lblName.Size = New System.Drawing.Size(38, 15)
Me.lblName.TabIndex = 0
Me.lblName.Text = "Name"
'
'tabPet
'
Me.tabPet.Location = New System.Drawing.Point(4, 22)
Me.tabPet.Name = "tabPet"
Me.tabPet.Size = New System.Drawing.Size(405, 534)
Me.tabPet.TabIndex = 1
Me.tabPet.Text = "Pet"
'
'cmdRefreshCharData
'
Me.cmdRefreshCharData.Location = New System.Drawing.Point(166, 572)
Me.cmdRefreshCharData.Name = "cmdRefreshCharData"
Me.cmdRefreshCharData.Size = New System.Drawing.Size(157, 23)
Me.cmdRefreshCharData.TabIndex = 9
Me.cmdRefreshCharData.Text = "Refresh Character Data"
'
'tabMiscInfo
'
Me.tabMiscInfo.Controls.Add(Me.txtLastLogoutTime)
Me.tabMiscInfo.Controls.Add(Me.lblLastLogoutTime)
Me.tabMiscInfo.Location = New System.Drawing.Point(4, 22)
Me.tabMiscInfo.Name = "tabMiscInfo"
Me.tabMiscInfo.Size = New System.Drawing.Size(405, 534)
Me.tabMiscInfo.TabIndex = 2
Me.tabMiscInfo.Text = "Miscellaneous Info"
'
'lblLastLogoutTime
'
Me.lblLastLogoutTime.Location = New System.Drawing.Point(8, 10)
Me.lblLastLogoutTime.Name = "lblLastLogoutTime"
Me.lblLastLogoutTime.Size = New System.Drawing.Size(88, 16)
Me.lblLastLogoutTime.TabIndex = 0
Me.lblLastLogoutTime.Text = "Last logout time"
'
'txtLastLogoutTime
'
Me.txtLastLogoutTime.Location = New System.Drawing.Point(96, 8)
Me.txtLastLogoutTime.Name = "txtLastLogoutTime"
Me.txtLastLogoutTime.ReadOnly = True
Me.txtLastLogoutTime.Size = New System.Drawing.Size(232, 20)
Me.txtLastLogoutTime.TabIndex = 1
Me.txtLastLogoutTime.Text = ""
'
'tabCharacter2
'
Me.tabCharacter2.Controls.Add(Me.grpCombatStat)
Me.tabCharacter2.Location = New System.Drawing.Point(4, 22)
Me.tabCharacter2.Name = "tabCharacter2"
Me.tabCharacter2.Size = New System.Drawing.Size(405, 534)
Me.tabCharacter2.TabIndex = 3
Me.tabCharacter2.Text = "Character Data (2/2)"
'
'grpCombatStat
'
Me.grpCombatStat.Controls.Add(Me.txtParryRate)
Me.grpCombatStat.Controls.Add(Me.lblParryRate)
Me.grpCombatStat.Controls.Add(Me.txtRngAttackPwr)
Me.grpCombatStat.Controls.Add(Me.Label7)
Me.grpCombatStat.Controls.Add(Me.txtAttackPwr)
Me.grpCombatStat.Controls.Add(Me.lblATCKPower)
Me.grpCombatStat.Controls.Add(Me.txtDodgeRate)
Me.grpCombatStat.Controls.Add(Me.lblDodgeRate)
Me.grpCombatStat.Controls.Add(Me.txtMaxRngDmg)
Me.grpCombatStat.Controls.Add(Me.Label5)
Me.grpCombatStat.Controls.Add(Me.txtMinRngDmg)
Me.grpCombatStat.Controls.Add(Me.lblMinMaxRngDmg)
Me.grpCombatStat.Controls.Add(Me.txtBlockRate)
Me.grpCombatStat.Controls.Add(Me.lblBlockRate)
Me.grpCombatStat.Controls.Add(Me.txtMaxDmg)
Me.grpCombatStat.Controls.Add(Me.Label4)
Me.grpCombatStat.Controls.Add(Me.txtMinDmg)
Me.grpCombatStat.Controls.Add(Me.lblMinDmg)
Me.grpCombatStat.Location = New System.Drawing.Point(7, 8)
Me.grpCombatStat.Name = "grpCombatStat"
Me.grpCombatStat.Size = New System.Drawing.Size(393, 118)
Me.grpCombatStat.TabIndex = 10
Me.grpCombatStat.TabStop = False
Me.grpCombatStat.Text = "Combat Stats"
'
'txtParryRate
'
Me.txtParryRate.Location = New System.Drawing.Point(342, 78)
Me.txtParryRate.Name = "txtParryRate"
Me.txtParryRate.Size = New System.Drawing.Size(44, 20)
Me.txtParryRate.TabIndex = 37
Me.txtParryRate.Text = ""
'
'lblParryRate
'
Me.lblParryRate.Location = New System.Drawing.Point(287, 79)
Me.lblParryRate.Name = "lblParryRate"
Me.lblParryRate.Size = New System.Drawing.Size(52, 14)
Me.lblParryRate.TabIndex = 36
Me.lblParryRate.Text = "Parry %"
Me.lblParryRate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'txtRngAttackPwr
'
Me.txtRngAttackPwr.Location = New System.Drawing.Point(199, 78)
Me.txtRngAttackPwr.Name = "txtRngAttackPwr"
Me.txtRngAttackPwr.Size = New System.Drawing.Size(69, 20)
Me.txtRngAttackPwr.TabIndex = 35
Me.txtRngAttackPwr.Text = ""
'
'Label7
'
Me.Label7.Location = New System.Drawing.Point(179, 82)
Me.Label7.Name = "Label7"
Me.Label7.Size = New System.Drawing.Size(12, 16)
Me.Label7.TabIndex = 34
Me.Label7.Text = "/"
'
'txtAttackPwr
'
Me.txtAttackPwr.Location = New System.Drawing.Point(100, 78)
Me.txtAttackPwr.Name = "txtAttackPwr"
Me.txtAttackPwr.Size = New System.Drawing.Size(69, 20)
Me.txtAttackPwr.TabIndex = 33
Me.txtAttackPwr.Text = ""
'
'lblATCKPower
'
Me.lblATCKPower.Location = New System.Drawing.Point(8, 77)
Me.lblATCKPower.Name = "lblATCKPower"
Me.lblATCKPower.Size = New System.Drawing.Size(81, 26)
Me.lblATCKPower.TabIndex = 32
Me.lblATCKPower.Text = "Attack/Ranged Attack Power"
'
'txtDodgeRate
'
Me.txtDodgeRate.Location = New System.Drawing.Point(342, 48)
Me.txtDodgeRate.Name = "txtDodgeRate"
Me.txtDodgeRate.Size = New System.Drawing.Size(44, 20)
Me.txtDodgeRate.TabIndex = 31
Me.txtDodgeRate.Text = ""
'
'lblDodgeRate
'
Me.lblDodgeRate.Location = New System.Drawing.Point(287, 48)
Me.lblDodgeRate.Name = "lblDodgeRate"
Me.lblDodgeRate.Size = New System.Drawing.Size(52, 14)
Me.lblDodgeRate.TabIndex = 30
Me.lblDodgeRate.Text = "Dodge %"
Me.lblDodgeRate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'txtMaxRngDmg
'
Me.txtMaxRngDmg.Location = New System.Drawing.Point(199, 48)
Me.txtMaxRngDmg.Name = "txtMaxRngDmg"
Me.txtMaxRngDmg.Size = New System.Drawing.Size(69, 20)
Me.txtMaxRngDmg.TabIndex = 29
Me.txtMaxRngDmg.Text = ""
'
'Label5
'
Me.Label5.Location = New System.Drawing.Point(179, 50)
Me.Label5.Name = "Label5"
Me.Label5.Size = New System.Drawing.Size(12, 16)
Me.Label5.TabIndex = 28
Me.Label5.Text = "/"
'
'txtMinRngDmg
'
Me.txtMinRngDmg.Location = New System.Drawing.Point(100, 48)
Me.txtMinRngDmg.Name = "txtMinRngDmg"
Me.txtMinRngDmg.Size = New System.Drawing.Size(69, 20)
Me.txtMinRngDmg.TabIndex = 27
Me.txtMinRngDmg.Text = ""
'
'lblMinMaxRngDmg
'
Me.lblMinMaxRngDmg.Location = New System.Drawing.Point(8, 44)
Me.lblMinMaxRngDmg.Name = "lblMinMaxRngDmg"
Me.lblMinMaxRngDmg.Size = New System.Drawing.Size(93, 28)
Me.lblMinMaxRngDmg.TabIndex = 26
Me.lblMinMaxRngDmg.Text = "Min/Max Ranged Damage"
'
'txtBlockRate
'
Me.txtBlockRate.Location = New System.Drawing.Point(342, 16)
Me.txtBlockRate.Name = "txtBlockRate"
Me.txtBlockRate.Size = New System.Drawing.Size(44, 20)
Me.txtBlockRate.TabIndex = 25
Me.txtBlockRate.Text = ""
'
'lblBlockRate
'
Me.lblBlockRate.Location = New System.Drawing.Point(293, 18)
Me.lblBlockRate.Name = "lblBlockRate"
Me.lblBlockRate.Size = New System.Drawing.Size(46, 14)
Me.lblBlockRate.TabIndex = 24
Me.lblBlockRate.Text = "Block %"
Me.lblBlockRate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'txtMaxDmg
'
Me.txtMaxDmg.Location = New System.Drawing.Point(199, 18)
Me.txtMaxDmg.Name = "txtMaxDmg"
Me.txtMaxDmg.Size = New System.Drawing.Size(69, 20)
Me.txtMaxDmg.TabIndex = 23
Me.txtMaxDmg.Text = ""
'
'Label4
'
Me.Label4.Location = New System.Drawing.Point(179, 22)
Me.Label4.Name = "Label4"
Me.Label4.Size = New System.Drawing.Size(12, 16)
Me.Label4.TabIndex = 22
Me.Label4.Text = "/"
'
'txtMinDmg
'
Me.txtMinDmg.Location = New System.Drawing.Point(100, 18)
Me.txtMinDmg.Name = "txtMinDmg"
Me.txtMinDmg.Size = New System.Drawing.Size(69, 20)
Me.txtMinDmg.TabIndex = 21
Me.txtMinDmg.Text = ""
'
'lblMinDmg
'
Me.lblMinDmg.Location = New System.Drawing.Point(8, 21)
Me.lblMinDmg.Name = "lblMinDmg"
Me.lblMinDmg.Size = New System.Drawing.Size(93, 15)
Me.lblMinDmg.TabIndex = 20
Me.lblMinDmg.Text = "Min/Max Damage"
'
'frmMCE
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(586, 623)
Me.Controls.Add(Me.cmdRefreshCharData)
Me.Controls.Add(Me.tabMCE)
Me.Controls.Add(Me.cmdSave)
Me.Controls.Add(Me.grpCharList)
Me.Controls.Add(Me.staMCE)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
Me.MaximizeBox = False
Me.Menu = Me.mnuMCE
Me.Name = "frmMCE"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "MCE - Mangos Character Editor/Viewer"
CType(Me.pnlDBStatus, System.ComponentModel.ISupportInitialize).EndInit()
Me.grpCharList.ResumeLayout(False)
Me.tabMCE.ResumeLayout(False)
Me.tabCharacter1.ResumeLayout(False)
Me.grpSkillTalent.ResumeLayout(False)
Me.grpStats.ResumeLayout(False)
Me.grpDetails.ResumeLayout(False)
Me.tabMiscInfo.ResumeLayout(False)
Me.tabCharacter2.ResumeLayout(False)
Me.grpCombatStat.ResumeLayout(False)
Me.ResumeLayout(False)

  End Sub

#End Region

#Region " Local Variables "

  Dim cnDBConnection As MySqlConnection
	Dim strDBName As String
	Dim strCharFields As String()

#End Region

#Region " Custom Properties "

#End Region

#Region " Event Codes "

Private Sub mnuSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetting.Click
  On Error GoTo ErrorHandling

  Dim frmDBSettings_New As frmDBSettings


  'Create a new instance of the CRM Module interface
  frmDBSettings_New = New frmDBSettings
  'Show the form
  frmDBSettings_New.ShowDialog()

  Exit Sub

ErrorHandling:
  GenericErrorHandler("frmMCE.vb", "mnuSetting_Click", Err.Number, Err.Source, Err.Description)
End Sub

Private Sub mnuConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuConnect.Click
  On Error GoTo ErrorHandling

  Dim blnSuccess As Boolean


  'Load database settings and try to connect to it
  blnSuccess = LoadDBSettingAndConnect(Application.StartupPath & "\" & MCE_CONFIG_FILE, cnDBConnection, strDBName)

  If blnSuccess = True Then
    pnlDBStatus.Text = "Successfully connected to database"
    'Start Populating the character list
    PopulateCharList(strDBName)
  Else
    pnlDBStatus.Text = "Cannot connect to database"
  End If

  Exit Sub

ErrorHandling:
  GenericErrorHandler("frmMCE.vb", "mnuSetting_Click", Err.Number, Err.Source, Err.Description)
End Sub

Private Sub lstCharList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstCharList.SelectedIndexChanged
  On Error GoTo ErrorHandling

  Dim blnSuccess As Boolean
  Dim strTempSplit As String()
  Dim strTempChar As String


  'Check to make sure we don't get -1 index
  If lstCharList.SelectedIndex < 0 Then
    Exit Sub
  End If

  'Get the GUID from the highlighted character
  strTempSplit = Split(lstCharList.Items(lstCharList.SelectedIndex), "[")
  'We want the GUID which is the second character
  strTempChar = strTempSplit(1)
  strTempChar = Replace(strTempChar, "]", "")

  'Populate the character tab
  PopulateCharacterTab(strDBName, strTempChar)

  Exit Sub

ErrorHandling:
  GenericErrorHandler("frmMCE.vb", "lstCharList_SelectedIndexChanged", Err.Number, Err.Source, Err.Description)
End Sub

Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

	Dim strTemp As String
	Dim blnSuccess As Boolean
	Dim StrSQL As String


	strCharFields(1244) = txtCharPoint1.Text


	strTemp = Join(strCharFields, " ")
	StrSQL = "UPDATE " & strDBName & ".CHARACTER SET DATA = '" & strTemp & "' " & _
	"WHERE  GUID = " & txtGUID.Text

	'Update just that field
	blnSuccess = ExecuteMySQLStatement(StrSQL, cnDBConnection)


End Sub

#End Region

#Region " Functions and Procedures "

Private Function LoadDBSettingAndConnect(ByVal strConfigFilename As String, ByRef cnDBConnection As MySqlConnection, _
  Optional ByRef strDBName As String = "") As Boolean
'-------------------------------------------------------------------------------
'Purpose     : This procedure is to save load all settings from the INI (XML formatted) file
'Assumptions :
'Input       :
'   - strConfigFilename, a string consisting the complete config filename including path
'   - cnDBConnection, mySQL database connection object
'   - strDBName, a string containing the database, we're going to need this
'Returns     :
'Note        : 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim blnSuccess As Boolean
  Dim strServerName As String
  Dim strConnectionTimeout As String
  Dim strPortNumber As String
  Dim strUsername As String
  Dim strPassword As String
  Dim strDatabaseName As String


  'First of all, check whether there's an existing file there
  If DoesFileExist(strConfigFilename) = False Then Exit Function

  'We only need to proceed if there's an existing file to work on
  'Load setup values for Database settings tab
  strServerName = ReadSettingsInXML("DB Settings", "Server Name", strConfigFilename, MCE_MYSQL_DEF_SERVER)
  strConnectionTimeout = ReadSettingsInXML("DB Settings", "Connection Timeout", strConfigFilename, MCE_MYSQL_DEF_TIMEOUT)
  strPortNumber = ReadSettingsInXML("DB Settings", "Port Number", strConfigFilename, MCE_MYSQL_DEF_PORT)
  strUsername = ReadSettingsInXML("DB Settings", "Username", strConfigFilename, "")
  strDatabaseName = ReadSettingsInXML("DB Settings", "Database", strConfigFilename, "mangos")
  strPassword = ReadSettingsInXML("DB Settings", "Password", strConfigFilename, "mangos")
  'Pass the database name back, we're gonna need it, trust me
  strDBName = strDatabaseName

  'Connect to datbase
  LoadDBSettingAndConnect = ConnectToMySQLDB(strServerName, strUsername, strPassword, strDatabaseName, cnDBConnection)

  Exit Function

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("frmMCE.vb", "LoadDBSettingAndConnect", Err.Number, Err.Source, Err.Description)
End Function

Private Sub PopulateCharList(ByVal strDatabasename As String)
'-------------------------------------------------------------------------------
'Purpose     : This procedure is to populate the listbox with character names and GUID
'Assumptions :
'Input       :
'   - strDatabasename, a string consisting the database name 
'Returns     :
'Note        : 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim dsCharList As DataSet
  Dim blnSuccess As Boolean
  Dim intRecordNo As Integer
  Dim drCharList As DataRow
  Dim intLoopCounter As Integer
  Dim intGUID As UInt32


  'Get the character list from database
  blnSuccess = RetrieveMySQLData("SELECT GUID, NAME FROM " & strDatabasename & ".CHARACTER", cnDBConnection, dsCharList, intRecordNo)

  'Check whether we have any record to display
  If intRecordNo < 1 Then
    MsgBox("Character table is empty", MsgBoxStyle.Information, "Empty Character table")
    Exit Sub
  End If

  'If we're here that means we have records
  For intLoopCounter = 0 To dsCharList.Tables(0).Rows.Count - 1
    'Get the record - one by one
    drCharList = dsCharList.Tables(0).Rows(intLoopCounter)
    'Store GUID into an Unsigned 32bit Integer
    intGUID = drCharList("GUID")
    'Display the char name with the GUID in brackets
    lstCharList.Items.Add(DBNullToString(drCharList("NAME") & " [" & intGUID.ToString & "]"))
  Next

  Exit Sub

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("frmMCE.vb", "PopulateCharList", Err.Number, Err.Source, Err.Description)
End Sub

Private Sub PopulateCharacterTab(ByVal strDatabasename As String, ByVal strGUID As String)
'-------------------------------------------------------------------------------
'Purpose     : This procedure is to populate the character tab
'Assumptions :
'Input       :
'   - strDatabasename, a string consisting the database name 
'Returns     :
'Note        : 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim dsCharacter As DataSet
  Dim drCharacter As DataRow
  Dim blnSuccess As Boolean
  Dim intRecordNo As Integer
  Dim intGUID As UInt32
  Dim strSQL As String
  Dim dtTemp As DateTime

  Dim uintTempSecs As System.UInt32
  Dim dblTempConv As Double


  'Get everything from the character table for that char
  strSQL = "SELECT * FROM " & strDatabasename & ".CHARACTER " & _
  "WHERE GUID = " & strGUID
  'Get the record
  blnSuccess = RetrieveMySQLData(strSQL, cnDBConnection, dsCharacter, intRecordNo)

  'Make sure that we get a record back
  If intRecordNo < 1 Then
    MsgBox("Cannot found character record with GUID " & strGUID, MsgBoxStyle.Exclamation, "Empty record")
    Exit Sub
  End If

  'Assign the record explicitly, since it should come back with only 1 character anyway 
  drCharacter = dsCharacter.Tables(0).Rows(0)

  'Start populating the fields
  txtGUID.Text = strGUID
  txtName.Text = DBNullToString(drCharacter("NAME"))
  txtRace.Text = RaceIDToString(DBNullToString(drCharacter("RACE")))
  txtClass.Text = ClassIDToString(DBNullToString(drCharacter("CLASS")))
  'Read Character data field and populate it to the screen
  ReadCharacterData(drCharacter)
  'Read logout time data
  uintTempSecs = drCharacter("LOGOUT_TIME")
  'Convert the logout time from UNIX time stamp into .NET date/time  
  dtTemp = ConvertFromUnixTimeStamp(CDbl(uintTempSecs.ToString), True)
  txtLastLogoutTime.Text = Format(dtTemp, "ddd, dd/MMM/yyyy hh:mm tt")



  Exit Sub

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("frmMCE.vb", "PopulateCharacterTab", Err.Number, Err.Source, Err.Description)
End Sub

Private Sub ReadCharacterData(ByVal drCharacter As DataRow)
'-------------------------------------------------------------------------------
'Purpose     : This procedure is to populate the character tab
'Assumptions :
'Input       :
'   - drCharacter, a data row object containing the character's data
'Returns     :
'Note        : 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim strCharData As String  
  Dim intTempUint32 As UInt32
  Dim isdsd As Integer

  'Get the character data
  strCharData = DBNullToString(drCharacter("DATA"))
  'Start splitting the data
  strCharFields = Split(strCharData, " ")

  'Populate to the screen based on the layout defined at
  'https://svn.mangosproject.org/trac/MaNGOS/wiki/Database/character/CharacterData
  txtCurrHealth.Text = strCharFields(DB_UNIT_FIELD_CURR_HEALTH)
  txtMaxHealth.Text = strCharFields(DB_UNIT_FIELD_MAX_HEALTH)
  txtCurrMana.Text = strCharFields(DB_UNIT_FIELD_MAX_HEALTH)
  txtMaxMana.Text = strCharFields(DB_UNIT_FIELD_MAX_MANA)
	txtLevel.Text = strCharFields(DB_UNIT_FIELD_LEVEL)
	txtSTR.Text = strCharFields(DB_UNIT_FIELD_STR)
  txtAGI.Text = strCharFields(DB_UNIT_FIELD_AGILITY)
	txtSTA.Text = strCharFields(DB_UNIT_FIELD_STAMINA)
  txtINT.Text = strCharFields(DB_UNIT_FIELD_IQ)
  txtSPI.Text = strCharFields(DB_UNIT_FIELD_SPIRIT)
	txtCurrXP.Text = Format(CInt(strCharFields(DB_PLAYER_XP)), "#,###,###")
	txtXPToLvl.Text = Format(CInt(strCharFields(DB_PLAYER_NEXT_LEVEL_XP)), "#,###,###")
	txtXPNeeded.Text = Format(CInt(strCharFields(DB_PLAYER_NEXT_LEVEL_XP)) - CInt(strCharFields(DB_PLAYER_XP)), "#,###,###")
  txtDodgeRate.Text = strCharFields(DB_PLAYER_DODGE_PERCENTAGE)
  txtBlockRate.Text = strCharFields(DB_PLAYER_BLOCK_PERCENTAGE)
  txtParryRate.Text = strCharFields(DB_PLAYER_PARRY_PERCENTAGE)

  Exit Sub

ErrorHandling:
  'Report it to screen
  GenericErrorHandler("frmMCE.vb", "ReadCharacterData", Err.Number, Err.Source, Err.Description)
End Sub


#End Region

End Class
