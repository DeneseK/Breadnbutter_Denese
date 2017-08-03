<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FEmployeeLog
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
	Public WithEvents ttmTime As AxTDBTime6.AxTDBTime
	Public WithEvents tdtDate As AxTDBDate6.AxTDBDate
	Public WithEvents tmrClock As System.Windows.Forms.Timer
	Public WithEvents cmdLog As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents txtName As System.Windows.Forms.TextBox
	Public WithEvents cmdContinue As System.Windows.Forms.Button
	Public WithEvents lblWelcomeMessage As System.Windows.Forms.Label
	Public WithEvents _Label3_0 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label3 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FEmployeeLog))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ttmTime = New AxTDBTime6.AxTDBTime
		Me.tdtDate = New AxTDBDate6.AxTDBDate
		Me.tmrClock = New System.Windows.Forms.Timer(components)
		Me.cmdLog = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.txtName = New System.Windows.Forms.TextBox
		Me.cmdContinue = New System.Windows.Forms.Button
		Me.lblWelcomeMessage = New System.Windows.Forms.Label
		Me._Label3_0 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label3 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.ttmTime, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.tdtDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label3, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = " "
		Me.ClientSize = New System.Drawing.Size(390, 280)
		Me.Location = New System.Drawing.Point(207, 187)
		Me.ControlBox = False
		Me.MaximizeBox = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FEmployeeLog"
		ttmTime.OcxState = CType(resources.GetObject("ttmTime.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ttmTime.Size = New System.Drawing.Size(124, 27)
		Me.ttmTime.Location = New System.Drawing.Point(240, 138)
		Me.ttmTime.TabIndex = 2
		Me.ttmTime.Name = "ttmTime"
		tdtDate.OcxState = CType(resources.GetObject("tdtDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tdtDate.Size = New System.Drawing.Size(119, 27)
		Me.tdtDate.Location = New System.Drawing.Point(115, 138)
		Me.tdtDate.TabIndex = 1
		Me.tdtDate.Name = "tdtDate"
		Me.tmrClock.Interval = 60000
		Me.tmrClock.Enabled = True
		Me.cmdLog.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdLog.Text = "Log"
		Me.AcceptButton = Me.cmdLog
		Me.cmdLog.Size = New System.Drawing.Size(99, 84)
		Me.cmdLog.Location = New System.Drawing.Point(20, 188)
		Me.cmdLog.TabIndex = 5
		Me.cmdLog.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdLog.BackColor = System.Drawing.SystemColors.Control
		Me.cmdLog.CausesValidation = True
		Me.cmdLog.Enabled = True
		Me.cmdLog.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdLog.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdLog.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdLog.TabStop = True
		Me.cmdLog.Name = "cmdLog"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(104, 84)
		Me.cmdCancel.Location = New System.Drawing.Point(263, 188)
		Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
		Me.cmdCancel.TabIndex = 4
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.txtPassword.AutoSize = False
		Me.txtPassword.Size = New System.Drawing.Size(249, 27)
		Me.txtPassword.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtPassword.Location = New System.Drawing.Point(115, 103)
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
		Me.txtName.AutoSize = False
		Me.txtName.Size = New System.Drawing.Size(249, 27)
		Me.txtName.Location = New System.Drawing.Point(115, 70)
		Me.txtName.Maxlength = 50
		Me.txtName.TabIndex = 6
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
		Me.cmdContinue.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdContinue.Text = "Continue"
		Me.cmdContinue.Size = New System.Drawing.Size(97, 84)
		Me.cmdContinue.Location = New System.Drawing.Point(143, 188)
		Me.cmdContinue.TabIndex = 3
		Me.cmdContinue.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdContinue.BackColor = System.Drawing.SystemColors.Control
		Me.cmdContinue.CausesValidation = True
		Me.cmdContinue.Enabled = True
		Me.cmdContinue.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdContinue.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdContinue.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdContinue.TabStop = True
		Me.cmdContinue.Name = "cmdContinue"
		Me.lblWelcomeMessage.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblWelcomeMessage.Size = New System.Drawing.Size(347, 57)
		Me.lblWelcomeMessage.Location = New System.Drawing.Point(23, 8)
		Me.lblWelcomeMessage.TabIndex = 10
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
		Me._Label3_0.Text = "Date/Time:"
		Me._Label3_0.Size = New System.Drawing.Size(79, 22)
		Me._Label3_0.Location = New System.Drawing.Point(20, 143)
		Me._Label3_0.TabIndex = 9
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
		Me.Label2.Text = "Password:"
		Me.Label2.Size = New System.Drawing.Size(82, 22)
		Me.Label2.Location = New System.Drawing.Point(20, 108)
		Me.Label2.TabIndex = 8
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
		Me.Label1.Text = "Name:"
		Me.Label1.Size = New System.Drawing.Size(87, 22)
		Me.Label1.Location = New System.Drawing.Point(23, 73)
		Me.Label1.TabIndex = 7
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
		Me.Controls.Add(ttmTime)
		Me.Controls.Add(tdtDate)
		Me.Controls.Add(cmdLog)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(txtPassword)
		Me.Controls.Add(txtName)
		Me.Controls.Add(cmdContinue)
		Me.Controls.Add(lblWelcomeMessage)
		Me.Controls.Add(_Label3_0)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Label3.SetIndex(_Label3_0, CType(0, Short))
		CType(Me.Label3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tdtDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.ttmTime, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class