<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FLogon
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
	Public WithEvents cboServer As System.Windows.Forms.ComboBox
	Public WithEvents cboDatabase As System.Windows.Forms.ComboBox
	Public WithEvents cmdContinue As System.Windows.Forms.Button
	Public WithEvents txtName As System.Windows.Forms.TextBox
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdLog As System.Windows.Forms.Button
	Public WithEvents tmrClock As System.Windows.Forms.Timer
	Public WithEvents ttmTime As AxTDBTime6.AxTDBTime
	Public WithEvents tdtDate As AxTDBDate6.AxTDBDate
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents _Label3_0 As System.Windows.Forms.Label
	Public WithEvents lblWelcomeMessage As System.Windows.Forms.Label
	Public WithEvents Label3 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FLogon))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cboServer = New System.Windows.Forms.ComboBox
		Me.cboDatabase = New System.Windows.Forms.ComboBox
		Me.cmdContinue = New System.Windows.Forms.Button
		Me.txtName = New System.Windows.Forms.TextBox
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdLog = New System.Windows.Forms.Button
		Me.tmrClock = New System.Windows.Forms.Timer(components)
		Me.ttmTime = New AxTDBTime6.AxTDBTime
		Me.tdtDate = New AxTDBDate6.AxTDBDate
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me._Label3_0 = New System.Windows.Forms.Label
		Me.lblWelcomeMessage = New System.Windows.Forms.Label
		Me.Label3 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.ttmTime, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.tdtDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label3, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(409, 352)
		Me.Location = New System.Drawing.Point(640, 217)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Name = "FLogon"
		Me.cboServer.Size = New System.Drawing.Size(249, 27)
		Me.cboServer.Location = New System.Drawing.Point(125, 70)
		Me.cboServer.Items.AddRange(New Object(){"HRI-SVR-03"})
		Me.cboServer.TabIndex = 6
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
		Me.cboDatabase.Size = New System.Drawing.Size(249, 27)
		Me.cboDatabase.Location = New System.Drawing.Point(125, 103)
		Me.cboDatabase.Items.AddRange(New Object(){"BNB_DATA"})
		Me.cboDatabase.TabIndex = 7
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
		Me.cmdContinue.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdContinue.Text = "Continue"
		Me.cmdContinue.Size = New System.Drawing.Size(97, 84)
		Me.cmdContinue.Location = New System.Drawing.Point(153, 260)
		Me.cmdContinue.TabIndex = 2
		Me.cmdContinue.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdContinue.BackColor = System.Drawing.SystemColors.Control
		Me.cmdContinue.CausesValidation = True
		Me.cmdContinue.Enabled = True
		Me.cmdContinue.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdContinue.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdContinue.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdContinue.TabStop = True
		Me.cmdContinue.Name = "cmdContinue"
		Me.txtName.AutoSize = False
		Me.txtName.Size = New System.Drawing.Size(249, 27)
		Me.txtName.Location = New System.Drawing.Point(125, 143)
		Me.txtName.Maxlength = 50
		Me.txtName.TabIndex = 8
		Me.txtName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtName.AcceptsReturn = True
		Me.txtName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtName.BackColor = System.Drawing.SystemColors.Window
		Me.txtName.CausesValidation = True
		Me.txtName.Enabled = True
		Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtName.HideSelection = True
		Me.txtName.ReadOnly = False
		Me.txtName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtName.MultiLine = False
		Me.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtName.TabStop = True
		Me.txtName.Visible = True
		Me.txtName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtName.Name = "txtName"
		Me.txtPassword.AutoSize = False
		Me.txtPassword.Size = New System.Drawing.Size(249, 27)
		Me.txtPassword.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtPassword.Location = New System.Drawing.Point(125, 175)
		Me.txtPassword.PasswordChar = ChrW(42)
		Me.txtPassword.TabIndex = 0
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(104, 84)
		Me.cmdCancel.Location = New System.Drawing.Point(273, 260)
		Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdLog.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdLog.Text = "Log"
		Me.AcceptButton = Me.cmdLog
		Me.cmdLog.Size = New System.Drawing.Size(99, 84)
		Me.cmdLog.Location = New System.Drawing.Point(30, 260)
		Me.cmdLog.TabIndex = 1
		Me.cmdLog.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdLog.BackColor = System.Drawing.SystemColors.Control
		Me.cmdLog.CausesValidation = True
		Me.cmdLog.Enabled = True
		Me.cmdLog.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdLog.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdLog.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdLog.TabStop = True
		Me.cmdLog.Name = "cmdLog"
		Me.tmrClock.Interval = 60000
		Me.tmrClock.Enabled = True
		ttmTime.OcxState = CType(resources.GetObject("ttmTime.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ttmTime.Size = New System.Drawing.Size(124, 27)
		Me.ttmTime.Location = New System.Drawing.Point(250, 210)
		Me.ttmTime.TabIndex = 5
		Me.ttmTime.Name = "ttmTime"
		tdtDate.OcxState = CType(resources.GetObject("tdtDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tdtDate.Size = New System.Drawing.Size(119, 27)
		Me.tdtDate.Location = New System.Drawing.Point(125, 210)
		Me.tdtDate.TabIndex = 4
		Me.tdtDate.Name = "tdtDate"
		Me.Label5.Text = "Server:"
		Me.Label5.Size = New System.Drawing.Size(54, 24)
		Me.Label5.Location = New System.Drawing.Point(30, 73)
		Me.Label5.TabIndex = 14
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label4.Text = "Database:"
		Me.Label4.Size = New System.Drawing.Size(67, 24)
		Me.Label4.Location = New System.Drawing.Point(30, 105)
		Me.Label4.TabIndex = 13
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label1.Text = "Name:"
		Me.Label1.Size = New System.Drawing.Size(87, 22)
		Me.Label1.Location = New System.Drawing.Point(33, 145)
		Me.Label1.TabIndex = 12
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
		Me.Label2.Text = "Password:"
		Me.Label2.Size = New System.Drawing.Size(82, 22)
		Me.Label2.Location = New System.Drawing.Point(30, 180)
		Me.Label2.TabIndex = 11
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
		Me._Label3_0.Text = "Date/Time:"
		Me._Label3_0.Size = New System.Drawing.Size(79, 22)
		Me._Label3_0.Location = New System.Drawing.Point(30, 215)
		Me._Label3_0.TabIndex = 10
		Me._Label3_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label3_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label3_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label3_0.Enabled = True
		Me._Label3_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label3_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label3_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label3_0.UseMnemonic = True
		Me._Label3_0.Visible = True
		Me._Label3_0.AutoSize = False
		Me._Label3_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label3_0.Name = "_Label3_0"
		Me.lblWelcomeMessage.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblWelcomeMessage.Size = New System.Drawing.Size(347, 57)
		Me.lblWelcomeMessage.Location = New System.Drawing.Point(33, 0)
		Me.lblWelcomeMessage.TabIndex = 9
		Me.lblWelcomeMessage.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblWelcomeMessage.BackColor = System.Drawing.SystemColors.Control
		Me.lblWelcomeMessage.Enabled = True
		Me.lblWelcomeMessage.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblWelcomeMessage.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblWelcomeMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblWelcomeMessage.UseMnemonic = True
		Me.lblWelcomeMessage.Visible = True
		Me.lblWelcomeMessage.AutoSize = False
		Me.lblWelcomeMessage.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblWelcomeMessage.Name = "lblWelcomeMessage"
		Me.Controls.Add(cboServer)
		Me.Controls.Add(cboDatabase)
		Me.Controls.Add(cmdContinue)
		Me.Controls.Add(txtName)
		Me.Controls.Add(txtPassword)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdLog)
		Me.Controls.Add(ttmTime)
		Me.Controls.Add(tdtDate)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.Controls.Add(_Label3_0)
		Me.Controls.Add(lblWelcomeMessage)
		Me.Label3.SetIndex(_Label3_0, CType(0, Short))
		CType(Me.Label3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tdtDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.ttmTime, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class