<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated(), ToolboxBitmap(GetType(ctlButton), "ctlButton.ToolboxBitmap")> Partial Class ctlButton
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			UserControl_Terminate()
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Friend WithEvents m_About As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ctlButton))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.m_About = New System.Windows.Forms.PictureBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ClientSize = New System.Drawing.Size(229, 97)
		MyBase.Location = New System.Drawing.Point(0, 0)
		MyBase.Name = "ctlButton"
		Me.m_About.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.m_About.Size = New System.Drawing.Size(377, 145)
		Me.m_About.Location = New System.Drawing.Point(496, 416)
		Me.m_About.TabIndex = 0
		Me.m_About.TabStop = False
		Me.m_About.Dock = System.Windows.Forms.DockStyle.None
		Me.m_About.BackColor = System.Drawing.SystemColors.Control
		Me.m_About.CausesValidation = True
		Me.m_About.Enabled = True
		Me.m_About.ForeColor = System.Drawing.SystemColors.ControlText
		Me.m_About.Cursor = System.Windows.Forms.Cursors.Default
		Me.m_About.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.m_About.Visible = True
		Me.m_About.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.m_About.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.m_About.Name = "m_About"
		Me.Controls.Add(m_About)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class