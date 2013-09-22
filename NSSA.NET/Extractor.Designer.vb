<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdFinish As ctlButton
	Public WithEvents cmdNext As ctlButton
	Public WithEvents cmdBack As ctlButton
	Public WithEvents cmdExit As ctlButton
	Public WithEvents optIAgree As System.Windows.Forms.RadioButton
	Public WithEvents optDisagree As System.Windows.Forms.RadioButton
	Public WithEvents txtLicenseAgreement As System.Windows.Forms.TextBox
	Public WithEvents lblAgree As System.Windows.Forms.Label
	Public WithEvents Line2 As System.Windows.Forms.Label
	Public WithEvents Line3 As System.Windows.Forms.Label
	Public WithEvents fraLicenseAgreement As System.Windows.Forms.GroupBox
	Public WithEvents Image3 As System.Windows.Forms.PictureBox
	Public WithEvents lblYouRan As System.Windows.Forms.Label
	Public WithEvents _fraSetup_0 As System.Windows.Forms.Panel
	Public WithEvents cmdRunFile As ctlButton
	Public WithEvents lstFiles As System.Windows.Forms.ListBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Image4 As System.Windows.Forms.PictureBox
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents _fraSetup_4 As System.Windows.Forms.Panel
	Public WithEvents XP_ProgressBar1 As ctlProgressBar
	Public WithEvents Image5 As System.Windows.Forms.PictureBox
	Public WithEvents TaskLbl As System.Windows.Forms.Label
	Public WithEvents FileProgressLbl As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents _fraSetup_3 As System.Windows.Forms.Panel
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	Public WithEvents lblReady As System.Windows.Forms.Label
	Public WithEvents _fraSetup_2 As System.Windows.Forms.Panel
	Public WithEvents cmdChangeDir As ctlButton
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Image2 As System.Windows.Forms.PictureBox
	Public WithEvents lblPath As System.Windows.Forms.Label
	Public WithEvents lblInformation As System.Windows.Forms.Label
	Public WithEvents _fraSetup_1 As System.Windows.Forms.Panel
	Public WithEvents Line1 As System.Windows.Forms.Label
	Public WithEvents fraSetup As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdFinish = New ctlButton
		Me.cmdNext = New ctlButton
		Me.cmdBack = New ctlButton
		Me.cmdExit = New ctlButton
		Me._fraSetup_0 = New System.Windows.Forms.Panel
		Me.fraLicenseAgreement = New System.Windows.Forms.GroupBox
		Me.optIAgree = New System.Windows.Forms.RadioButton
		Me.optDisagree = New System.Windows.Forms.RadioButton
		Me.txtLicenseAgreement = New System.Windows.Forms.TextBox
		Me.lblAgree = New System.Windows.Forms.Label
		Me.Line2 = New System.Windows.Forms.Label
		Me.Line3 = New System.Windows.Forms.Label
		Me.Image3 = New System.Windows.Forms.PictureBox
		Me.lblYouRan = New System.Windows.Forms.Label
		Me._fraSetup_4 = New System.Windows.Forms.Panel
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cmdRunFile = New ctlButton
		Me.lstFiles = New System.Windows.Forms.ListBox
		Me.Image4 = New System.Windows.Forms.PictureBox
		Me.Label6 = New System.Windows.Forms.Label
		Me._fraSetup_3 = New System.Windows.Forms.Panel
		Me.XP_ProgressBar1 = New ctlProgressBar
		Me.Image5 = New System.Windows.Forms.PictureBox
		Me.TaskLbl = New System.Windows.Forms.Label
		Me.FileProgressLbl = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me._fraSetup_2 = New System.Windows.Forms.Panel
		Me.Image1 = New System.Windows.Forms.PictureBox
		Me.lblReady = New System.Windows.Forms.Label
		Me._fraSetup_1 = New System.Windows.Forms.Panel
		Me.cmdChangeDir = New ctlButton
		Me.Label1 = New System.Windows.Forms.Label
		Me.Image2 = New System.Windows.Forms.PictureBox
		Me.lblPath = New System.Windows.Forms.Label
		Me.lblInformation = New System.Windows.Forms.Label
		Me.Line1 = New System.Windows.Forms.Label
		Me.fraSetup = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(components)
		Me._fraSetup_0.SuspendLayout()
		Me.fraLicenseAgreement.SuspendLayout()
		Me._fraSetup_4.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me._fraSetup_3.SuspendLayout()
		Me._fraSetup_2.SuspendLayout()
		Me._fraSetup_1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.fraSetup, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Text = "Nexgen Self Extractor"
		Me.ClientSize = New System.Drawing.Size(391, 235)
		Me.Location = New System.Drawing.Point(3, 14)
		Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("frmMain.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Visible = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmMain"
		Me.cmdFinish.Size = New System.Drawing.Size(65, 21)
		Me.cmdFinish.Location = New System.Drawing.Point(320, 208)
		Me.cmdFinish.TabIndex = 14
		Me.cmdFinish.iNonThemeStyle = 0
		Me.cmdFinish.ttForeColor = 0
		Me.cmdFinish.Name = "cmdFinish"
		Me.cmdNext.Size = New System.Drawing.Size(65, 21)
		Me.cmdNext.Location = New System.Drawing.Point(248, 208)
		Me.cmdNext.TabIndex = 15
		Me.cmdNext.iNonThemeStyle = 0
		Me.cmdNext.ttForeColor = 0
		Me.cmdNext.Name = "cmdNext"
		Me.cmdBack.Size = New System.Drawing.Size(65, 21)
		Me.cmdBack.Location = New System.Drawing.Point(176, 208)
		Me.cmdBack.TabIndex = 16
		Me.cmdBack.iNonThemeStyle = 0
		Me.cmdBack.ttForeColor = 0
		Me.cmdBack.Name = "cmdBack"
		Me.cmdExit.Size = New System.Drawing.Size(65, 21)
		Me.cmdExit.Location = New System.Drawing.Point(8, 208)
		Me.cmdExit.TabIndex = 17
		Me.cmdExit.iNonThemeStyle = 0
		Me.cmdExit.ttForeColor = 0
		Me.cmdExit.Name = "cmdExit"
		Me._fraSetup_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._fraSetup_0.Text = "Welcome"
		Me._fraSetup_0.Size = New System.Drawing.Size(377, 193)
		Me._fraSetup_0.Location = New System.Drawing.Point(8, 8)
		Me._fraSetup_0.TabIndex = 0
		Me._fraSetup_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraSetup_0.BackColor = System.Drawing.SystemColors.Control
		Me._fraSetup_0.Enabled = True
		Me._fraSetup_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraSetup_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._fraSetup_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraSetup_0.Visible = True
		Me._fraSetup_0.Name = "_fraSetup_0"
		Me.fraLicenseAgreement.Text = "License Agreement"
		Me.fraLicenseAgreement.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraLicenseAgreement.Size = New System.Drawing.Size(345, 153)
		Me.fraLicenseAgreement.Location = New System.Drawing.Point(24, 32)
		Me.fraLicenseAgreement.TabIndex = 19
		Me.fraLicenseAgreement.BackColor = System.Drawing.SystemColors.Control
		Me.fraLicenseAgreement.Enabled = True
		Me.fraLicenseAgreement.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraLicenseAgreement.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraLicenseAgreement.Visible = True
		Me.fraLicenseAgreement.Name = "fraLicenseAgreement"
		Me.optIAgree.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optIAgree.Text = "I agree"
		Me.optIAgree.ForeColor = System.Drawing.SystemColors.WindowText
		Me.optIAgree.Size = New System.Drawing.Size(57, 17)
		Me.optIAgree.Location = New System.Drawing.Point(280, 128)
		Me.optIAgree.TabIndex = 22
		Me.optIAgree.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optIAgree.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optIAgree.BackColor = System.Drawing.SystemColors.Control
		Me.optIAgree.CausesValidation = True
		Me.optIAgree.Enabled = True
		Me.optIAgree.Cursor = System.Windows.Forms.Cursors.Default
		Me.optIAgree.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optIAgree.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optIAgree.TabStop = True
		Me.optIAgree.Checked = False
		Me.optIAgree.Visible = True
		Me.optIAgree.Name = "optIAgree"
		Me.optDisagree.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDisagree.Text = "I disagree"
		Me.optDisagree.ForeColor = System.Drawing.SystemColors.WindowText
		Me.optDisagree.Size = New System.Drawing.Size(73, 17)
		Me.optDisagree.Location = New System.Drawing.Point(208, 128)
		Me.optDisagree.TabIndex = 21
		Me.optDisagree.Checked = True
		Me.optDisagree.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optDisagree.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDisagree.BackColor = System.Drawing.SystemColors.Control
		Me.optDisagree.CausesValidation = True
		Me.optDisagree.Enabled = True
		Me.optDisagree.Cursor = System.Windows.Forms.Cursors.Default
		Me.optDisagree.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optDisagree.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optDisagree.TabStop = True
		Me.optDisagree.Visible = True
		Me.optDisagree.Name = "optDisagree"
		Me.txtLicenseAgreement.AutoSize = False
		Me.txtLicenseAgreement.BackColor = System.Drawing.SystemColors.Control
		Me.txtLicenseAgreement.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.txtLicenseAgreement.Size = New System.Drawing.Size(333, 105)
		Me.txtLicenseAgreement.Location = New System.Drawing.Point(8, 14)
		Me.txtLicenseAgreement.ReadOnly = True
		Me.txtLicenseAgreement.MultiLine = True
		Me.txtLicenseAgreement.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtLicenseAgreement.TabIndex = 20
		Me.txtLicenseAgreement.Text = "No license agreement" & Chr(13) & Chr(10)
		Me.txtLicenseAgreement.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLicenseAgreement.AcceptsReturn = True
		Me.txtLicenseAgreement.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLicenseAgreement.CausesValidation = True
		Me.txtLicenseAgreement.Enabled = True
		Me.txtLicenseAgreement.HideSelection = True
		Me.txtLicenseAgreement.Maxlength = 0
		Me.txtLicenseAgreement.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLicenseAgreement.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLicenseAgreement.TabStop = True
		Me.txtLicenseAgreement.Visible = True
		Me.txtLicenseAgreement.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtLicenseAgreement.Name = "txtLicenseAgreement"
		Me.lblAgree.Text = "Do you agree with these terms?"
		Me.lblAgree.Size = New System.Drawing.Size(193, 17)
		Me.lblAgree.Location = New System.Drawing.Point(8, 128)
		Me.lblAgree.TabIndex = 26
		Me.lblAgree.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblAgree.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblAgree.BackColor = System.Drawing.SystemColors.Control
		Me.lblAgree.Enabled = True
		Me.lblAgree.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblAgree.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblAgree.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAgree.UseMnemonic = True
		Me.lblAgree.Visible = True
		Me.lblAgree.AutoSize = False
		Me.lblAgree.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAgree.Name = "lblAgree"
		Me.Line2.BackColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.Line2.Visible = True
		Me.Line2.Location = New System.Drawing.Point(0, 120)
		Me.Line2.Size = New System.Drawing.Size(344, 1)
		Me.Line2.Name = "Line2"
		Me.Line3.BackColor = System.Drawing.Color.White
		Me.Line3.Visible = True
		Me.Line3.Location = New System.Drawing.Point(0, 121)
		Me.Line3.Size = New System.Drawing.Size(344, 1)
		Me.Line3.Name = "Line3"
		Me.Image3.Size = New System.Drawing.Size(16, 16)
		Me.Image3.Location = New System.Drawing.Point(0, 0)
		Me.Image3.Enabled = True
		Me.Image3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image3.Visible = True
		Me.Image3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image3.Name = "Image3"
		Me.lblYouRan.Text = "Not Initialized"
		Me.lblYouRan.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblYouRan.Size = New System.Drawing.Size(353, 33)
		Me.lblYouRan.Location = New System.Drawing.Point(24, 0)
		Me.lblYouRan.TabIndex = 5
		Me.lblYouRan.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblYouRan.BackColor = System.Drawing.SystemColors.Control
		Me.lblYouRan.Enabled = True
		Me.lblYouRan.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblYouRan.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblYouRan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblYouRan.UseMnemonic = True
		Me.lblYouRan.Visible = True
		Me.lblYouRan.AutoSize = False
		Me.lblYouRan.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblYouRan.Name = "lblYouRan"
		Me._fraSetup_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._fraSetup_4.Text = "Finished"
		Me._fraSetup_4.Size = New System.Drawing.Size(377, 193)
		Me._fraSetup_4.Location = New System.Drawing.Point(8, 8)
		Me._fraSetup_4.TabIndex = 4
		Me._fraSetup_4.Visible = False
		Me._fraSetup_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraSetup_4.BackColor = System.Drawing.SystemColors.Control
		Me._fraSetup_4.Enabled = True
		Me._fraSetup_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraSetup_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._fraSetup_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraSetup_4.Name = "_fraSetup_4"
		Me.Frame1.Text = "Run a file"
		Me.Frame1.Size = New System.Drawing.Size(353, 169)
		Me.Frame1.Location = New System.Drawing.Point(24, 16)
		Me.Frame1.TabIndex = 23
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.cmdRunFile.Size = New System.Drawing.Size(57, 21)
		Me.cmdRunFile.Location = New System.Drawing.Point(288, 136)
		Me.cmdRunFile.TabIndex = 25
		Me.cmdRunFile.iNonThemeStyle = 0
		Me.cmdRunFile.ttForeColor = 0
		Me.cmdRunFile.Name = "cmdRunFile"
		Me.lstFiles.BackColor = System.Drawing.SystemColors.Control
		Me.lstFiles.Size = New System.Drawing.Size(335, 116)
		Me.lstFiles.IntegralHeight = False
		Me.lstFiles.Location = New System.Drawing.Point(8, 16)
		Me.lstFiles.TabIndex = 24
		Me.lstFiles.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstFiles.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstFiles.CausesValidation = True
		Me.lstFiles.Enabled = True
		Me.lstFiles.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstFiles.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstFiles.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstFiles.Sorted = False
		Me.lstFiles.TabStop = True
		Me.lstFiles.Visible = True
		Me.lstFiles.MultiColumn = False
		Me.lstFiles.Name = "lstFiles"
		Me.Image4.Size = New System.Drawing.Size(16, 16)
		Me.Image4.Location = New System.Drawing.Point(0, 0)
		Me.Image4.Enabled = True
		Me.Image4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image4.Visible = True
		Me.Image4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image4.Name = "Image4"
		Me.Label6.Text = "Self Extract is complete, click 'Finish'"
		Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Size = New System.Drawing.Size(345, 17)
		Me.Label6.Location = New System.Drawing.Point(24, 0)
		Me.Label6.TabIndex = 13
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.BackColor = System.Drawing.SystemColors.Control
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me._fraSetup_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._fraSetup_3.Text = "Progress"
		Me._fraSetup_3.Size = New System.Drawing.Size(377, 193)
		Me._fraSetup_3.Location = New System.Drawing.Point(8, 8)
		Me._fraSetup_3.TabIndex = 3
		Me._fraSetup_3.Visible = False
		Me._fraSetup_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraSetup_3.BackColor = System.Drawing.SystemColors.Control
		Me._fraSetup_3.Enabled = True
		Me._fraSetup_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraSetup_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._fraSetup_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraSetup_3.Name = "_fraSetup_3"
		Me.XP_ProgressBar1.Size = New System.Drawing.Size(313, 17)
		Me.XP_ProgressBar1.Location = New System.Drawing.Point(8, 168)
		Me.XP_ProgressBar1.TabIndex = 9
		Me.XP_ProgressBar1.Name = "XP_ProgressBar1"
		Me.Image5.Size = New System.Drawing.Size(16, 16)
		Me.Image5.Location = New System.Drawing.Point(0, 0)
		Me.Image5.Enabled = True
		Me.Image5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image5.Visible = True
		Me.Image5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image5.Name = "Image5"
		Me.TaskLbl.Text = "..."
		Me.TaskLbl.Size = New System.Drawing.Size(345, 41)
		Me.TaskLbl.Location = New System.Drawing.Point(24, 16)
		Me.TaskLbl.TabIndex = 12
		Me.TaskLbl.Visible = False
		Me.TaskLbl.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TaskLbl.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.TaskLbl.BackColor = System.Drawing.SystemColors.Control
		Me.TaskLbl.Enabled = True
		Me.TaskLbl.ForeColor = System.Drawing.SystemColors.ControlText
		Me.TaskLbl.Cursor = System.Windows.Forms.Cursors.Default
		Me.TaskLbl.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TaskLbl.UseMnemonic = True
		Me.TaskLbl.AutoSize = False
		Me.TaskLbl.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.TaskLbl.Name = "TaskLbl"
		Me.FileProgressLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.FileProgressLbl.Text = "0%"
		Me.FileProgressLbl.Size = New System.Drawing.Size(41, 17)
		Me.FileProgressLbl.Location = New System.Drawing.Point(328, 168)
		Me.FileProgressLbl.TabIndex = 11
		Me.FileProgressLbl.Visible = False
		Me.FileProgressLbl.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FileProgressLbl.BackColor = System.Drawing.SystemColors.Control
		Me.FileProgressLbl.Enabled = True
		Me.FileProgressLbl.ForeColor = System.Drawing.SystemColors.ControlText
		Me.FileProgressLbl.Cursor = System.Windows.Forms.Cursors.Default
		Me.FileProgressLbl.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.FileProgressLbl.UseMnemonic = True
		Me.FileProgressLbl.AutoSize = False
		Me.FileProgressLbl.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.FileProgressLbl.Name = "FileProgressLbl"
		Me.Label5.Text = "Files are being extracted"
		Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Size = New System.Drawing.Size(345, 17)
		Me.Label5.Location = New System.Drawing.Point(24, 0)
		Me.Label5.TabIndex = 10
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me._fraSetup_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._fraSetup_2.Text = "Ready"
		Me._fraSetup_2.Size = New System.Drawing.Size(377, 193)
		Me._fraSetup_2.Location = New System.Drawing.Point(8, 8)
		Me._fraSetup_2.TabIndex = 2
		Me._fraSetup_2.Visible = False
		Me._fraSetup_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraSetup_2.BackColor = System.Drawing.SystemColors.Control
		Me._fraSetup_2.Enabled = True
		Me._fraSetup_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraSetup_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._fraSetup_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraSetup_2.Name = "_fraSetup_2"
		Me.Image1.Size = New System.Drawing.Size(16, 16)
		Me.Image1.Location = New System.Drawing.Point(0, 0)
		Me.Image1.Enabled = True
		Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image1.Visible = True
		Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image1.Name = "Image1"
		Me.lblReady.Text = "The destination path has been set and you are ready to extract. Click 'Next'"
		Me.lblReady.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblReady.Size = New System.Drawing.Size(345, 49)
		Me.lblReady.Location = New System.Drawing.Point(24, 0)
		Me.lblReady.TabIndex = 8
		Me.lblReady.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblReady.BackColor = System.Drawing.SystemColors.Control
		Me.lblReady.Enabled = True
		Me.lblReady.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblReady.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblReady.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblReady.UseMnemonic = True
		Me.lblReady.Visible = True
		Me.lblReady.AutoSize = False
		Me.lblReady.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblReady.Name = "lblReady"
		Me._fraSetup_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._fraSetup_1.Text = "Location"
		Me._fraSetup_1.Size = New System.Drawing.Size(377, 193)
		Me._fraSetup_1.Location = New System.Drawing.Point(8, 8)
		Me._fraSetup_1.TabIndex = 1
		Me._fraSetup_1.Visible = False
		Me._fraSetup_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraSetup_1.BackColor = System.Drawing.SystemColors.Control
		Me._fraSetup_1.Enabled = True
		Me._fraSetup_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraSetup_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._fraSetup_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraSetup_1.Name = "_fraSetup_1"
		Me.cmdChangeDir.Size = New System.Drawing.Size(65, 21)
		Me.cmdChangeDir.Location = New System.Drawing.Point(312, 40)
		Me.cmdChangeDir.TabIndex = 18
		Me.cmdChangeDir.iNonThemeStyle = 0
		Me.cmdChangeDir.ttForeColor = 0
		Me.cmdChangeDir.Name = "cmdChangeDir"
		Me.Label1.Text = "To change the location, click 'Change'"
		Me.Label1.Size = New System.Drawing.Size(265, 17)
		Me.Label1.Location = New System.Drawing.Point(24, 43)
		Me.Label1.TabIndex = 27
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Image2.Size = New System.Drawing.Size(16, 16)
		Me.Image2.Location = New System.Drawing.Point(0, 0)
		Me.Image2.Enabled = True
		Me.Image2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image2.Visible = True
		Me.Image2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image2.Name = "Image2"
		Me.lblPath.Text = "<Path>"
		Me.lblPath.Size = New System.Drawing.Size(337, 17)
		Me.lblPath.Location = New System.Drawing.Point(24, 16)
		Me.lblPath.TabIndex = 7
		Me.lblPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPath.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPath.BackColor = System.Drawing.SystemColors.Control
		Me.lblPath.Enabled = True
		Me.lblPath.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPath.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPath.UseMnemonic = True
		Me.lblPath.Visible = True
		Me.lblPath.AutoSize = False
		Me.lblPath.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPath.Name = "lblPath"
		Me.lblInformation.Text = "<Info>"
		Me.lblInformation.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblInformation.Size = New System.Drawing.Size(353, 17)
		Me.lblInformation.Location = New System.Drawing.Point(24, 0)
		Me.lblInformation.TabIndex = 6
		Me.lblInformation.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblInformation.BackColor = System.Drawing.SystemColors.Control
		Me.lblInformation.Enabled = True
		Me.lblInformation.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblInformation.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblInformation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblInformation.UseMnemonic = True
		Me.lblInformation.Visible = True
		Me.lblInformation.AutoSize = False
		Me.lblInformation.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblInformation.Name = "lblInformation"
		Me.Line1.BackColor = System.Drawing.Color.White
		Me.Line1.Visible = True
		Me.Line1.Location = New System.Drawing.Point(0, 202)
		Me.Line1.Size = New System.Drawing.Size(392, 1)
		Me.Line1.Name = "Line1"
		Me.Controls.Add(cmdFinish)
		Me.Controls.Add(cmdNext)
		Me.Controls.Add(cmdBack)
		Me.Controls.Add(cmdExit)
		Me.Controls.Add(_fraSetup_0)
		Me.Controls.Add(_fraSetup_4)
		Me.Controls.Add(_fraSetup_3)
		Me.Controls.Add(_fraSetup_2)
		Me.Controls.Add(_fraSetup_1)
		Me.Controls.Add(Line1)
		Me._fraSetup_0.Controls.Add(fraLicenseAgreement)
		Me._fraSetup_0.Controls.Add(Image3)
		Me._fraSetup_0.Controls.Add(lblYouRan)
		Me.fraLicenseAgreement.Controls.Add(optIAgree)
		Me.fraLicenseAgreement.Controls.Add(optDisagree)
		Me.fraLicenseAgreement.Controls.Add(txtLicenseAgreement)
		Me.fraLicenseAgreement.Controls.Add(lblAgree)
		Me.fraLicenseAgreement.Controls.Add(Line2)
		Me.fraLicenseAgreement.Controls.Add(Line3)
		Me._fraSetup_4.Controls.Add(Frame1)
		Me._fraSetup_4.Controls.Add(Image4)
		Me._fraSetup_4.Controls.Add(Label6)
		Me.Frame1.Controls.Add(cmdRunFile)
		Me.Frame1.Controls.Add(lstFiles)
		Me._fraSetup_3.Controls.Add(XP_ProgressBar1)
		Me._fraSetup_3.Controls.Add(Image5)
		Me._fraSetup_3.Controls.Add(TaskLbl)
		Me._fraSetup_3.Controls.Add(FileProgressLbl)
		Me._fraSetup_3.Controls.Add(Label5)
		Me._fraSetup_2.Controls.Add(Image1)
		Me._fraSetup_2.Controls.Add(lblReady)
		Me._fraSetup_1.Controls.Add(cmdChangeDir)
		Me._fraSetup_1.Controls.Add(Label1)
		Me._fraSetup_1.Controls.Add(Image2)
		Me._fraSetup_1.Controls.Add(lblPath)
		Me._fraSetup_1.Controls.Add(lblInformation)
		Me.fraSetup.SetIndex(_fraSetup_0, CType(0, Short))
		Me.fraSetup.SetIndex(_fraSetup_4, CType(4, Short))
		Me.fraSetup.SetIndex(_fraSetup_3, CType(3, Short))
		Me.fraSetup.SetIndex(_fraSetup_2, CType(2, Short))
		Me.fraSetup.SetIndex(_fraSetup_1, CType(1, Short))
		CType(Me.fraSetup, System.ComponentModel.ISupportInitialize).EndInit()
		Me._fraSetup_0.ResumeLayout(False)
		Me.fraLicenseAgreement.ResumeLayout(False)
		Me._fraSetup_4.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me._fraSetup_3.ResumeLayout(False)
		Me._fraSetup_2.ResumeLayout(False)
		Me._fraSetup_1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class