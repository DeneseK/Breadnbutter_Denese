<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FEditDetail
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
	Public WithEvents chkOpenCall As System.Windows.Forms.CheckBox
	Public WithEvents cboCase As System.Windows.Forms.ComboBox
	Public WithEvents cboType As AxSSDataWidgets_B_OLEDB.AxSSOleDBCombo
	Public WithEvents ttmTime As AxTDBTime6.AxTDBTime
	Public WithEvents cmdEdit As System.Windows.Forms.Button
	Public WithEvents txtResults As System.Windows.Forms.TextBox
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtSubject As System.Windows.Forms.TextBox
	Public WithEvents mskDate As AxGTMaskDate.AxGTMaskDate
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblUser As System.Windows.Forms.Label
	Public WithEvents _Label3_1 As System.Windows.Forms.Label
	Public WithEvents _Label5_4 As System.Windows.Forms.Label
	Public WithEvents _Label5_5 As System.Windows.Forms.Label
	Public WithEvents _Label5_6 As System.Windows.Forms.Label
	Public WithEvents _Label5_7 As System.Windows.Forms.Label
	Public WithEvents _Label3_3 As System.Windows.Forms.Label
	Public WithEvents Label3 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Label5 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FEditDetail))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkOpenCall = New System.Windows.Forms.CheckBox
		Me.cboCase = New System.Windows.Forms.ComboBox
		Me.cboType = New AxSSDataWidgets_B_OLEDB.AxSSOleDBCombo
		Me.ttmTime = New AxTDBTime6.AxTDBTime
		Me.cmdEdit = New System.Windows.Forms.Button
		Me.txtResults = New System.Windows.Forms.TextBox
		Me.cmdSave = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtSubject = New System.Windows.Forms.TextBox
		Me.mskDate = New AxGTMaskDate.AxGTMaskDate
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblUser = New System.Windows.Forms.Label
		Me._Label3_1 = New System.Windows.Forms.Label
		Me._Label5_4 = New System.Windows.Forms.Label
		Me._Label5_5 = New System.Windows.Forms.Label
		Me._Label5_6 = New System.Windows.Forms.Label
		Me._Label5_7 = New System.Windows.Forms.Label
		Me._Label3_3 = New System.Windows.Forms.Label
		Me.Label3 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Label5 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cboType, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.ttmTime, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label3, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label5, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Call"
		Me.ClientSize = New System.Drawing.Size(780, 217)
		Me.Location = New System.Drawing.Point(227, 343)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
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
		Me.Name = "FEditDetail"
		Me.chkOpenCall.Text = "Open Call"
		Me.chkOpenCall.Enabled = False
		Me.chkOpenCall.Size = New System.Drawing.Size(92, 22)
		Me.chkOpenCall.Location = New System.Drawing.Point(13, 180)
		Me.chkOpenCall.TabIndex = 17
		Me.chkOpenCall.Visible = False
		Me.chkOpenCall.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkOpenCall.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkOpenCall.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkOpenCall.BackColor = System.Drawing.SystemColors.Control
		Me.chkOpenCall.CausesValidation = True
		Me.chkOpenCall.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkOpenCall.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkOpenCall.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkOpenCall.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkOpenCall.TabStop = True
		Me.chkOpenCall.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkOpenCall.Name = "chkOpenCall"
		Me.cboCase.Enabled = False
		Me.cboCase.Size = New System.Drawing.Size(172, 27)
		Me.cboCase.Location = New System.Drawing.Point(60, 100)
		Me.cboCase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboCase.TabIndex = 15
		Me.cboCase.Visible = False
		Me.cboCase.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboCase.BackColor = System.Drawing.SystemColors.Window
		Me.cboCase.CausesValidation = True
		Me.cboCase.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboCase.IntegralHeight = True
		Me.cboCase.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboCase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboCase.Sorted = False
		Me.cboCase.TabStop = True
		Me.cboCase.Name = "cboCase"
		cboType.OcxState = CType(resources.GetObject("cboType.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cboType.Size = New System.Drawing.Size(167, 27)
		Me.cboType.Location = New System.Drawing.Point(60, 10)
		Me.cboType.TabIndex = 1
		Me.cboType.Name = "cboType"
		ttmTime.OcxState = CType(resources.GetObject("ttmTime.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ttmTime.Size = New System.Drawing.Size(107, 27)
		Me.ttmTime.Location = New System.Drawing.Point(60, 70)
		Me.ttmTime.TabIndex = 14
		Me.ttmTime.Name = "ttmTime"
		Me.cmdEdit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdEdit.Text = "Edit"
		Me.cmdEdit.Size = New System.Drawing.Size(87, 27)
		Me.cmdEdit.Location = New System.Drawing.Point(530, 178)
		Me.cmdEdit.TabIndex = 0
		Me.cmdEdit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdEdit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdEdit.CausesValidation = True
		Me.cmdEdit.Enabled = True
		Me.cmdEdit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdEdit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdEdit.TabStop = True
		Me.cmdEdit.Name = "cmdEdit"
		Me.txtResults.AutoSize = False
		Me.txtResults.Enabled = False
		Me.txtResults.Size = New System.Drawing.Size(314, 134)
		Me.txtResults.Location = New System.Drawing.Point(300, 38)
		Me.txtResults.MultiLine = True
		Me.txtResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtResults.TabIndex = 3
		Me.txtResults.Tag = "1"
		Me.txtResults.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtResults.AcceptsReturn = True
		Me.txtResults.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtResults.BackColor = System.Drawing.SystemColors.Window
		Me.txtResults.CausesValidation = True
		Me.txtResults.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtResults.HideSelection = True
		Me.txtResults.ReadOnly = False
		Me.txtResults.Maxlength = 0
		Me.txtResults.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtResults.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtResults.TabStop = True
		Me.txtResults.Visible = True
		Me.txtResults.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtResults.Name = "txtResults"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdSave.Text = "Save"
		Me.cmdSave.Size = New System.Drawing.Size(97, 97)
		Me.cmdSave.Location = New System.Drawing.Point(650, 10)
		Me.cmdSave.Image = CType(resources.GetObject("cmdSave.Image"), System.Drawing.Image)
		Me.cmdSave.TabIndex = 4
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(97, 97)
		Me.cmdCancel.Location = New System.Drawing.Point(650, 110)
		Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
		Me.cmdCancel.TabIndex = 5
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.txtSubject.AutoSize = False
		Me.txtSubject.Enabled = False
		Me.txtSubject.Size = New System.Drawing.Size(314, 27)
		Me.txtSubject.Location = New System.Drawing.Point(300, 8)
		Me.txtSubject.TabIndex = 2
		Me.txtSubject.Tag = "1"
		Me.txtSubject.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSubject.AcceptsReturn = True
		Me.txtSubject.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSubject.BackColor = System.Drawing.SystemColors.Window
		Me.txtSubject.CausesValidation = True
		Me.txtSubject.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSubject.HideSelection = True
		Me.txtSubject.ReadOnly = False
		Me.txtSubject.Maxlength = 0
		Me.txtSubject.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSubject.MultiLine = False
		Me.txtSubject.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSubject.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSubject.TabStop = True
		Me.txtSubject.Visible = True
		Me.txtSubject.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSubject.Name = "txtSubject"
		mskDate.OcxState = CType(resources.GetObject("mskDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskDate.Size = New System.Drawing.Size(107, 27)
		Me.mskDate.Location = New System.Drawing.Point(60, 40)
		Me.mskDate.TabIndex = 6
		Me.mskDate.Name = "mskDate"
		Me.Label1.Text = "Case:"
		Me.Label1.Size = New System.Drawing.Size(62, 22)
		Me.Label1.Location = New System.Drawing.Point(10, 104)
		Me.Label1.TabIndex = 16
		Me.Label1.Visible = False
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.lblUser.Size = New System.Drawing.Size(167, 27)
		Me.lblUser.Location = New System.Drawing.Point(65, 143)
		Me.lblUser.TabIndex = 13
		Me.lblUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUser.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblUser.BackColor = System.Drawing.SystemColors.Control
		Me.lblUser.Enabled = True
		Me.lblUser.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUser.UseMnemonic = True
		Me.lblUser.Visible = True
		Me.lblUser.AutoSize = False
		Me.lblUser.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblUser.Name = "lblUser"
		Me._Label3_1.Text = "Results:"
		Me._Label3_1.Size = New System.Drawing.Size(54, 19)
		Me._Label3_1.Location = New System.Drawing.Point(248, 38)
		Me._Label3_1.TabIndex = 12
		Me._Label3_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label3_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label3_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label3_1.Enabled = True
		Me._Label3_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label3_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label3_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label3_1.UseMnemonic = True
		Me._Label3_1.Visible = True
		Me._Label3_1.AutoSize = False
		Me._Label3_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label3_1.Name = "_Label3_1"
		Me._Label5_4.Text = "Type:"
		Me._Label5_4.Size = New System.Drawing.Size(42, 24)
		Me._Label5_4.Location = New System.Drawing.Point(10, 10)
		Me._Label5_4.TabIndex = 11
		Me._Label5_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label5_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label5_4.BackColor = System.Drawing.SystemColors.Control
		Me._Label5_4.Enabled = True
		Me._Label5_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label5_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label5_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label5_4.UseMnemonic = True
		Me._Label5_4.Visible = True
		Me._Label5_4.AutoSize = False
		Me._Label5_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label5_4.Name = "_Label5_4"
		Me._Label5_5.Text = "User:"
		Me._Label5_5.Size = New System.Drawing.Size(42, 24)
		Me._Label5_5.Location = New System.Drawing.Point(13, 145)
		Me._Label5_5.TabIndex = 10
		Me._Label5_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label5_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label5_5.BackColor = System.Drawing.SystemColors.Control
		Me._Label5_5.Enabled = True
		Me._Label5_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label5_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label5_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label5_5.UseMnemonic = True
		Me._Label5_5.Visible = True
		Me._Label5_5.AutoSize = False
		Me._Label5_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label5_5.Name = "_Label5_5"
		Me._Label5_6.Text = "Date:"
		Me._Label5_6.Size = New System.Drawing.Size(39, 24)
		Me._Label5_6.Location = New System.Drawing.Point(10, 43)
		Me._Label5_6.TabIndex = 9
		Me._Label5_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label5_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label5_6.BackColor = System.Drawing.SystemColors.Control
		Me._Label5_6.Enabled = True
		Me._Label5_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label5_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label5_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label5_6.UseMnemonic = True
		Me._Label5_6.Visible = True
		Me._Label5_6.AutoSize = False
		Me._Label5_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label5_6.Name = "_Label5_6"
		Me._Label5_7.Text = "Time:"
		Me._Label5_7.Size = New System.Drawing.Size(42, 24)
		Me._Label5_7.Location = New System.Drawing.Point(10, 75)
		Me._Label5_7.TabIndex = 8
		Me._Label5_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label5_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label5_7.BackColor = System.Drawing.SystemColors.Control
		Me._Label5_7.Enabled = True
		Me._Label5_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label5_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label5_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label5_7.UseMnemonic = True
		Me._Label5_7.Visible = True
		Me._Label5_7.AutoSize = False
		Me._Label5_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label5_7.Name = "_Label5_7"
		Me._Label3_3.Text = "Subject:"
		Me._Label3_3.Size = New System.Drawing.Size(54, 19)
		Me._Label3_3.Location = New System.Drawing.Point(245, 10)
		Me._Label3_3.TabIndex = 7
		Me._Label3_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label3_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label3_3.BackColor = System.Drawing.SystemColors.Control
		Me._Label3_3.Enabled = True
		Me._Label3_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label3_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label3_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label3_3.UseMnemonic = True
		Me._Label3_3.Visible = True
		Me._Label3_3.AutoSize = False
		Me._Label3_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label3_3.Name = "_Label3_3"
		Me.Controls.Add(chkOpenCall)
		Me.Controls.Add(cboCase)
		Me.Controls.Add(cboType)
		Me.Controls.Add(ttmTime)
		Me.Controls.Add(cmdEdit)
		Me.Controls.Add(txtResults)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(txtSubject)
		Me.Controls.Add(mskDate)
		Me.Controls.Add(Label1)
		Me.Controls.Add(lblUser)
		Me.Controls.Add(_Label3_1)
		Me.Controls.Add(_Label5_4)
		Me.Controls.Add(_Label5_5)
		Me.Controls.Add(_Label5_6)
		Me.Controls.Add(_Label5_7)
		Me.Controls.Add(_Label3_3)
		Me.Label3.SetIndex(_Label3_1, CType(1, Short))
		Me.Label3.SetIndex(_Label3_3, CType(3, Short))
		Me.Label5.SetIndex(_Label5_4, CType(4, Short))
		Me.Label5.SetIndex(_Label5_5, CType(5, Short))
		Me.Label5.SetIndex(_Label5_6, CType(6, Short))
		Me.Label5.SetIndex(_Label5_7, CType(7, Short))
		CType(Me.Label5, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.ttmTime, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboType, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class