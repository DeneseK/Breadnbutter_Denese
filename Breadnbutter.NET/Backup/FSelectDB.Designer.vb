<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FSelectDB
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
	Public WithEvents chkLogin As System.Windows.Forms.CheckBox
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents txtUserID As System.Windows.Forms.TextBox
	Public WithEvents cmdSelectDB As System.Windows.Forms.Button
	Public WithEvents txtDatabase As System.Windows.Forms.TextBox
	Public WithEvents cboDatabase As System.Windows.Forms.ComboBox
	Public WithEvents cboServer As System.Windows.Forms.ComboBox
	Public WithEvents _optDBType_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optDBType_0 As System.Windows.Forms.RadioButton
	Public WithEvents CancelButton_Renamed As System.Windows.Forms.Button
	Public WithEvents OKButton As System.Windows.Forms.Button
	Public WithEvents lblPassword As System.Windows.Forms.Label
	Public WithEvents lblUserID As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents optDBType As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FSelectDB))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkLogin = New System.Windows.Forms.CheckBox
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.txtUserID = New System.Windows.Forms.TextBox
		Me.cmdSelectDB = New System.Windows.Forms.Button
		Me.txtDatabase = New System.Windows.Forms.TextBox
		Me.cboDatabase = New System.Windows.Forms.ComboBox
		Me.cboServer = New System.Windows.Forms.ComboBox
		Me._optDBType_1 = New System.Windows.Forms.RadioButton
		Me._optDBType_0 = New System.Windows.Forms.RadioButton
		Me.CancelButton_Renamed = New System.Windows.Forms.Button
		Me.OKButton = New System.Windows.Forms.Button
		Me.lblPassword = New System.Windows.Forms.Label
		Me.lblUserID = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.optDBType = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.optDBType, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Select Database"
		Me.ClientSize = New System.Drawing.Size(522, 309)
		Me.Location = New System.Drawing.Point(230, 313)
		Me.Icon = CType(resources.GetObject("FSelectDB.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
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
		Me.Name = "FSelectDB"
		Me.chkLogin.Text = "Login"
		Me.chkLogin.Size = New System.Drawing.Size(322, 32)
		Me.chkLogin.Location = New System.Drawing.Point(25, 110)
		Me.chkLogin.TabIndex = 15
		Me.chkLogin.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLogin.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLogin.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLogin.BackColor = System.Drawing.SystemColors.Control
		Me.chkLogin.CausesValidation = True
		Me.chkLogin.Enabled = True
		Me.chkLogin.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLogin.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLogin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLogin.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLogin.TabStop = True
		Me.chkLogin.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLogin.Visible = True
		Me.chkLogin.Name = "chkLogin"
		Me.txtPassword.AutoSize = False
		Me.txtPassword.Size = New System.Drawing.Size(262, 24)
		Me.txtPassword.Location = New System.Drawing.Point(120, 180)
		Me.txtPassword.TabIndex = 14
		Me.txtPassword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPassword.AcceptsReturn = True
		Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
		Me.txtPassword.CausesValidation = True
		Me.txtPassword.Enabled = True
		Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPassword.HideSelection = True
		Me.txtPassword.ReadOnly = False
		Me.txtPassword.Maxlength = 0
		Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPassword.MultiLine = False
		Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPassword.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPassword.TabStop = True
		Me.txtPassword.Visible = True
		Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPassword.Name = "txtPassword"
		Me.txtUserID.AutoSize = False
		Me.txtUserID.Size = New System.Drawing.Size(262, 24)
		Me.txtUserID.Location = New System.Drawing.Point(120, 150)
		Me.txtUserID.TabIndex = 13
		Me.txtUserID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUserID.AcceptsReturn = True
		Me.txtUserID.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUserID.BackColor = System.Drawing.SystemColors.Window
		Me.txtUserID.CausesValidation = True
		Me.txtUserID.Enabled = True
		Me.txtUserID.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUserID.HideSelection = True
		Me.txtUserID.ReadOnly = False
		Me.txtUserID.Maxlength = 0
		Me.txtUserID.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUserID.MultiLine = False
		Me.txtUserID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUserID.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUserID.TabStop = True
		Me.txtUserID.Visible = True
		Me.txtUserID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUserID.Name = "txtUserID"
		Me.cmdSelectDB.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdSelectDB.Size = New System.Drawing.Size(27, 27)
		Me.cmdSelectDB.Location = New System.Drawing.Point(358, 263)
		Me.cmdSelectDB.Image = CType(resources.GetObject("cmdSelectDB.Image"), System.Drawing.Image)
		Me.cmdSelectDB.TabIndex = 10
		Me.cmdSelectDB.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSelectDB.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSelectDB.CausesValidation = True
		Me.cmdSelectDB.Enabled = True
		Me.cmdSelectDB.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSelectDB.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSelectDB.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSelectDB.TabStop = True
		Me.cmdSelectDB.Name = "cmdSelectDB"
		Me.txtDatabase.AutoSize = False
		Me.txtDatabase.Size = New System.Drawing.Size(232, 27)
		Me.txtDatabase.Location = New System.Drawing.Point(123, 260)
		Me.txtDatabase.TabIndex = 8
		Me.txtDatabase.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDatabase.AcceptsReturn = True
		Me.txtDatabase.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDatabase.BackColor = System.Drawing.SystemColors.Window
		Me.txtDatabase.CausesValidation = True
		Me.txtDatabase.Enabled = True
		Me.txtDatabase.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDatabase.HideSelection = True
		Me.txtDatabase.ReadOnly = False
		Me.txtDatabase.Maxlength = 0
		Me.txtDatabase.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDatabase.MultiLine = False
		Me.txtDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDatabase.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDatabase.TabStop = True
		Me.txtDatabase.Visible = True
		Me.txtDatabase.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDatabase.Name = "txtDatabase"
		Me.cboDatabase.Size = New System.Drawing.Size(264, 27)
		Me.cboDatabase.Location = New System.Drawing.Point(120, 80)
		Me.cboDatabase.Items.AddRange(New Object(){"BNB_DATA"})
		Me.cboDatabase.TabIndex = 6
		Me.cboDatabase.Text = "cboDatabase"
		Me.cboDatabase.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboDatabase.BackColor = System.Drawing.SystemColors.Window
		Me.cboDatabase.CausesValidation = True
		Me.cboDatabase.Enabled = True
		Me.cboDatabase.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboDatabase.IntegralHeight = True
		Me.cboDatabase.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboDatabase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboDatabase.Sorted = False
		Me.cboDatabase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboDatabase.TabStop = True
		Me.cboDatabase.Visible = True
		Me.cboDatabase.Name = "cboDatabase"
		Me.cboServer.Size = New System.Drawing.Size(264, 27)
		Me.cboServer.Location = New System.Drawing.Point(120, 48)
		Me.cboServer.Items.AddRange(New Object(){"HAWKINS-MAIN"})
		Me.cboServer.TabIndex = 4
		Me.cboServer.Text = "cboServer"
		Me.cboServer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboServer.BackColor = System.Drawing.SystemColors.Window
		Me.cboServer.CausesValidation = True
		Me.cboServer.Enabled = True
		Me.cboServer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboServer.IntegralHeight = True
		Me.cboServer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboServer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboServer.Sorted = False
		Me.cboServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboServer.TabStop = True
		Me.cboServer.Visible = True
		Me.cboServer.Name = "cboServer"
		Me._optDBType_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDBType_1.Text = "Access"
		Me._optDBType_1.Size = New System.Drawing.Size(297, 27)
		Me._optDBType_1.Location = New System.Drawing.Point(25, 220)
		Me._optDBType_1.TabIndex = 3
		Me._optDBType_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optDBType_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDBType_1.BackColor = System.Drawing.SystemColors.Control
		Me._optDBType_1.CausesValidation = True
		Me._optDBType_1.Enabled = True
		Me._optDBType_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optDBType_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optDBType_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optDBType_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optDBType_1.TabStop = True
		Me._optDBType_1.Checked = False
		Me._optDBType_1.Visible = True
		Me._optDBType_1.Name = "_optDBType_1"
		Me._optDBType_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDBType_0.Text = "SQL Server"
		Me._optDBType_0.Size = New System.Drawing.Size(297, 27)
		Me._optDBType_0.Location = New System.Drawing.Point(25, 13)
		Me._optDBType_0.TabIndex = 2
		Me._optDBType_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optDBType_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDBType_0.BackColor = System.Drawing.SystemColors.Control
		Me._optDBType_0.CausesValidation = True
		Me._optDBType_0.Enabled = True
		Me._optDBType_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optDBType_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optDBType_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optDBType_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optDBType_0.TabStop = True
		Me._optDBType_0.Checked = False
		Me._optDBType_0.Visible = True
		Me._optDBType_0.Name = "_optDBType_0"
		Me.CancelButton_Renamed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton_Renamed.Text = "Cancel"
		Me.CancelButton_Renamed.Size = New System.Drawing.Size(102, 32)
		Me.CancelButton_Renamed.Location = New System.Drawing.Point(405, 50)
		Me.CancelButton_Renamed.TabIndex = 1
		Me.CancelButton_Renamed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CancelButton_Renamed.BackColor = System.Drawing.SystemColors.Control
		Me.CancelButton_Renamed.CausesValidation = True
		Me.CancelButton_Renamed.Enabled = True
		Me.CancelButton_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CancelButton_Renamed.Cursor = System.Windows.Forms.Cursors.Default
		Me.CancelButton_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CancelButton_Renamed.TabStop = True
		Me.CancelButton_Renamed.Name = "CancelButton_Renamed"
		Me.OKButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.OKButton.Text = "OK"
		Me.OKButton.Size = New System.Drawing.Size(102, 32)
		Me.OKButton.Location = New System.Drawing.Point(405, 10)
		Me.OKButton.TabIndex = 0
		Me.OKButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OKButton.BackColor = System.Drawing.SystemColors.Control
		Me.OKButton.CausesValidation = True
		Me.OKButton.Enabled = True
		Me.OKButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OKButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.OKButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OKButton.TabStop = True
		Me.OKButton.Name = "OKButton"
		Me.lblPassword.Text = "Password"
		Me.lblPassword.Size = New System.Drawing.Size(67, 24)
		Me.lblPassword.Location = New System.Drawing.Point(50, 180)
		Me.lblPassword.TabIndex = 12
		Me.lblPassword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPassword.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPassword.BackColor = System.Drawing.SystemColors.Control
		Me.lblPassword.Enabled = True
		Me.lblPassword.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPassword.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPassword.UseMnemonic = True
		Me.lblPassword.Visible = True
		Me.lblPassword.AutoSize = False
		Me.lblPassword.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPassword.Name = "lblPassword"
		Me.lblUserID.Text = "User ID"
		Me.lblUserID.Size = New System.Drawing.Size(67, 24)
		Me.lblUserID.Location = New System.Drawing.Point(50, 150)
		Me.lblUserID.TabIndex = 11
		Me.lblUserID.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUserID.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblUserID.BackColor = System.Drawing.SystemColors.Control
		Me.lblUserID.Enabled = True
		Me.lblUserID.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblUserID.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblUserID.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUserID.UseMnemonic = True
		Me.lblUserID.Visible = True
		Me.lblUserID.AutoSize = False
		Me.lblUserID.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblUserID.Name = "lblUserID"
		Me.Label3.Text = "Database:"
		Me.Label3.Size = New System.Drawing.Size(67, 24)
		Me.Label3.Location = New System.Drawing.Point(55, 263)
		Me.Label3.TabIndex = 9
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.Text = "Database:"
		Me.Label2.Size = New System.Drawing.Size(67, 24)
		Me.Label2.Location = New System.Drawing.Point(50, 83)
		Me.Label2.TabIndex = 7
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "Server:"
		Me.Label1.Size = New System.Drawing.Size(54, 24)
		Me.Label1.Location = New System.Drawing.Point(50, 50)
		Me.Label1.TabIndex = 5
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
		Me.Controls.Add(chkLogin)
		Me.Controls.Add(txtPassword)
		Me.Controls.Add(txtUserID)
		Me.Controls.Add(cmdSelectDB)
		Me.Controls.Add(txtDatabase)
		Me.Controls.Add(cboDatabase)
		Me.Controls.Add(cboServer)
		Me.Controls.Add(_optDBType_1)
		Me.Controls.Add(_optDBType_0)
		Me.Controls.Add(CancelButton_Renamed)
		Me.Controls.Add(OKButton)
		Me.Controls.Add(lblPassword)
		Me.Controls.Add(lblUserID)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.optDBType.SetIndex(_optDBType_1, CType(1, Short))
		Me.optDBType.SetIndex(_optDBType_0, CType(0, Short))
		CType(Me.optDBType, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class