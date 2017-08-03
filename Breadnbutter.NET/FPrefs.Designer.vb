<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FPrefs
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
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cboStatus As AxSSDataWidgets_B.AxSSDBCombo
	Public WithEvents cboAuthStatus As AxSSDataWidgets_B.AxSSDBCombo
	Public WithEvents cboShipStatus As AxSSDataWidgets_B.AxSSDBCombo
	Public WithEvents _Label2_1 As System.Windows.Forms.Label
	Public WithEvents _Label2_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FPrefs))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cboStatus = New AxSSDataWidgets_B.AxSSDBCombo
		Me.cboAuthStatus = New AxSSDataWidgets_B.AxSSDBCombo
		Me.cboShipStatus = New AxSSDataWidgets_B.AxSSDBCombo
		Me._Label2_1 = New System.Windows.Forms.Label
		Me._Label2_0 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cboStatus, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cboAuthStatus, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cboShipStatus, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Preferences"
		Me.ClientSize = New System.Drawing.Size(449, 180)
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
		Me.Name = "FPrefs"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(117, 29)
		Me.cmdCancel.Location = New System.Drawing.Point(220, 138)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "&OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(117, 29)
		Me.cmdOK.Location = New System.Drawing.Point(88, 138)
		Me.cmdOK.TabIndex = 2
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		cboStatus.OcxState = CType(resources.GetObject("cboStatus.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cboStatus.Size = New System.Drawing.Size(204, 27)
		Me.cboStatus.Location = New System.Drawing.Point(218, 23)
		Me.cboStatus.TabIndex = 0
		Me.cboStatus.Name = "cboStatus"
		cboAuthStatus.OcxState = CType(resources.GetObject("cboAuthStatus.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cboAuthStatus.Size = New System.Drawing.Size(204, 27)
		Me.cboAuthStatus.Location = New System.Drawing.Point(218, 88)
		Me.cboAuthStatus.TabIndex = 4
		Me.cboAuthStatus.Name = "cboAuthStatus"
		cboShipStatus.OcxState = CType(resources.GetObject("cboShipStatus.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cboShipStatus.Size = New System.Drawing.Size(204, 27)
		Me.cboShipStatus.Location = New System.Drawing.Point(218, 55)
		Me.cboShipStatus.TabIndex = 6
		Me.cboShipStatus.Name = "cboShipStatus"
		Me._Label2_1.Text = "Initial Shipping Status:"
		Me._Label2_1.Size = New System.Drawing.Size(169, 24)
		Me._Label2_1.Location = New System.Drawing.Point(35, 58)
		Me._Label2_1.TabIndex = 7
		Me._Label2_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_1.Enabled = True
		Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_1.UseMnemonic = True
		Me._Label2_1.Visible = True
		Me._Label2_1.AutoSize = False
		Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_1.Name = "_Label2_1"
		Me._Label2_0.Text = "Initial Authorization Status:"
		Me._Label2_0.Size = New System.Drawing.Size(169, 24)
		Me._Label2_0.Location = New System.Drawing.Point(35, 90)
		Me._Label2_0.TabIndex = 5
		Me._Label2_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_0.Enabled = True
		Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_0.UseMnemonic = True
		Me._Label2_0.Visible = True
		Me._Label2_0.AutoSize = False
		Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_0.Name = "_Label2_0"
		Me.Label1.Text = "Initial Customer Status:"
		Me.Label1.Size = New System.Drawing.Size(174, 24)
		Me.Label1.Location = New System.Drawing.Point(35, 25)
		Me.Label1.TabIndex = 1
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
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cboStatus)
		Me.Controls.Add(cboAuthStatus)
		Me.Controls.Add(cboShipStatus)
		Me.Controls.Add(_Label2_1)
		Me.Controls.Add(_Label2_0)
		Me.Controls.Add(Label1)
		Me.Label2.SetIndex(_Label2_1, CType(1, Short))
		Me.Label2.SetIndex(_Label2_0, CType(0, Short))
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboShipStatus, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboAuthStatus, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboStatus, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class