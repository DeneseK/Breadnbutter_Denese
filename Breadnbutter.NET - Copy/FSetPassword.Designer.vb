<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FSetPassword
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
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtVerifyPwd As System.Windows.Forms.TextBox
	Public WithEvents txtNewPwd As System.Windows.Forms.TextBox
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FSetPassword))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Command1 = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtVerifyPwd = New System.Windows.Forms.TextBox
		Me.txtNewPwd = New System.Windows.Forms.TextBox
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Set Password"
		Me.ClientSize = New System.Drawing.Size(322, 145)
		Me.Location = New System.Drawing.Point(4, 28)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FSetPassword"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.Command1
		Me.Command1.Text = "&Cancel"
		Me.Command1.Size = New System.Drawing.Size(117, 32)
		Me.Command1.Location = New System.Drawing.Point(165, 98)
		Me.Command1.TabIndex = 5
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "&OK"
		Me.cmdOK.Size = New System.Drawing.Size(117, 32)
		Me.cmdOK.Location = New System.Drawing.Point(35, 98)
		Me.cmdOK.TabIndex = 4
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.txtVerifyPwd.AutoSize = False
		Me.txtVerifyPwd.Size = New System.Drawing.Size(157, 27)
		Me.txtVerifyPwd.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtVerifyPwd.Location = New System.Drawing.Point(140, 50)
		Me.txtVerifyPwd.PasswordChar = ChrW(42)
		Me.txtVerifyPwd.TabIndex = 1
		Me.txtVerifyPwd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtVerifyPwd.AcceptsReturn = True
		Me.txtVerifyPwd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtVerifyPwd.BackColor = System.Drawing.SystemColors.Window
		Me.txtVerifyPwd.CausesValidation = True
		Me.txtVerifyPwd.Enabled = True
		Me.txtVerifyPwd.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtVerifyPwd.HideSelection = True
		Me.txtVerifyPwd.ReadOnly = False
		Me.txtVerifyPwd.Maxlength = 0
		Me.txtVerifyPwd.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtVerifyPwd.MultiLine = False
		Me.txtVerifyPwd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtVerifyPwd.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtVerifyPwd.TabStop = True
		Me.txtVerifyPwd.Visible = True
		Me.txtVerifyPwd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtVerifyPwd.Name = "txtVerifyPwd"
		Me.txtNewPwd.AutoSize = False
		Me.txtNewPwd.Size = New System.Drawing.Size(157, 27)
		Me.txtNewPwd.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtNewPwd.Location = New System.Drawing.Point(140, 18)
		Me.txtNewPwd.PasswordChar = ChrW(42)
		Me.txtNewPwd.TabIndex = 0
		Me.txtNewPwd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNewPwd.AcceptsReturn = True
		Me.txtNewPwd.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNewPwd.BackColor = System.Drawing.SystemColors.Window
		Me.txtNewPwd.CausesValidation = True
		Me.txtNewPwd.Enabled = True
		Me.txtNewPwd.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNewPwd.HideSelection = True
		Me.txtNewPwd.ReadOnly = False
		Me.txtNewPwd.Maxlength = 0
		Me.txtNewPwd.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNewPwd.MultiLine = False
		Me.txtNewPwd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNewPwd.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNewPwd.TabStop = True
		Me.txtNewPwd.Visible = True
		Me.txtNewPwd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNewPwd.Name = "txtNewPwd"
		Me._Label1_2.Text = "Verify Password:"
		Me._Label1_2.Size = New System.Drawing.Size(102, 24)
		Me._Label1_2.Location = New System.Drawing.Point(25, 53)
		Me._Label1_2.TabIndex = 3
		Me._Label1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_2.Enabled = True
		Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me._Label1_1.Text = "New Password:"
		Me._Label1_1.Size = New System.Drawing.Size(102, 24)
		Me._Label1_1.Location = New System.Drawing.Point(25, 20)
		Me._Label1_1.TabIndex = 2
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me.Controls.Add(Command1)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtVerifyPwd)
		Me.Controls.Add(txtNewPwd)
		Me.Controls.Add(_Label1_2)
		Me.Controls.Add(_Label1_1)
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class