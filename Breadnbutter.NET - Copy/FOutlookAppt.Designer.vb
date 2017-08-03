<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FOutlookAppt
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
	Public WithEvents chkSales As System.Windows.Forms.CheckBox
	Public WithEvents txtSubject As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtNote As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents mskTime As AxGTMaskDate.AxGTMaskDate
	Public WithEvents mskDate As AxGTMaskDate.AxGTMaskDate
	Public WithEvents mskTime2 As AxGTMaskDate.AxGTMaskDate
	Public WithEvents lblCompany As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _lblDate_3 As System.Windows.Forms.Label
	Public WithEvents lblName As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _lblDate_2 As System.Windows.Forms.Label
	Public WithEvents _lblDate_1 As System.Windows.Forms.Label
	Public WithEvents _lblDate_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblDate As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FOutlookAppt))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkSales = New System.Windows.Forms.CheckBox
		Me.txtSubject = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtNote = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.mskTime = New AxGTMaskDate.AxGTMaskDate
		Me.mskDate = New AxGTMaskDate.AxGTMaskDate
		Me.mskTime2 = New AxGTMaskDate.AxGTMaskDate
		Me.lblCompany = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._lblDate_3 = New System.Windows.Forms.Label
		Me.lblName = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._lblDate_2 = New System.Windows.Forms.Label
		Me._lblDate_1 = New System.Windows.Forms.Label
		Me._lblDate_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblDate = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.mskTime, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskTime2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblDate, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Appointment"
		Me.ClientSize = New System.Drawing.Size(387, 283)
		Me.Location = New System.Drawing.Point(344, 260)
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
		Me.Name = "FOutlookAppt"
		Me.chkSales.Text = "Add to Sales Contact"
		Me.chkSales.Size = New System.Drawing.Size(162, 17)
		Me.chkSales.Location = New System.Drawing.Point(220, 250)
		Me.chkSales.TabIndex = 15
		Me.chkSales.Visible = False
		Me.chkSales.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSales.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSales.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSales.BackColor = System.Drawing.SystemColors.Control
		Me.chkSales.CausesValidation = True
		Me.chkSales.Enabled = True
		Me.chkSales.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSales.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSales.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSales.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSales.TabStop = True
		Me.chkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSales.Name = "chkSales"
		Me.txtSubject.AutoSize = False
		Me.txtSubject.Size = New System.Drawing.Size(227, 27)
		Me.txtSubject.Location = New System.Drawing.Point(123, 135)
		Me.txtSubject.TabIndex = 2
		Me.txtSubject.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSubject.AcceptsReturn = True
		Me.txtSubject.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSubject.BackColor = System.Drawing.SystemColors.Window
		Me.txtSubject.CausesValidation = True
		Me.txtSubject.Enabled = True
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(89, 34)
		Me.cmdCancel.Location = New System.Drawing.Point(120, 240)
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
		Me.txtNote.AutoSize = False
		Me.txtNote.Size = New System.Drawing.Size(227, 59)
		Me.txtNote.Location = New System.Drawing.Point(123, 168)
		Me.txtNote.MultiLine = True
		Me.txtNote.TabIndex = 3
		Me.txtNote.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNote.AcceptsReturn = True
		Me.txtNote.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNote.BackColor = System.Drawing.SystemColors.Window
		Me.txtNote.CausesValidation = True
		Me.txtNote.Enabled = True
		Me.txtNote.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNote.HideSelection = True
		Me.txtNote.ReadOnly = False
		Me.txtNote.Maxlength = 0
		Me.txtNote.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNote.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNote.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNote.TabStop = True
		Me.txtNote.Visible = True
		Me.txtNote.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNote.Name = "txtNote"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "&OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(92, 34)
		Me.cmdOK.Location = New System.Drawing.Point(18, 240)
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
		mskTime.OcxState = CType(resources.GetObject("mskTime.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskTime.Size = New System.Drawing.Size(117, 29)
		Me.mskTime.Location = New System.Drawing.Point(123, 100)
		Me.mskTime.TabIndex = 1
		Me.mskTime.Name = "mskTime"
		mskDate.OcxState = CType(resources.GetObject("mskDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskDate.Size = New System.Drawing.Size(114, 29)
		Me.mskDate.Location = New System.Drawing.Point(125, 65)
		Me.mskDate.TabIndex = 0
		Me.mskDate.Name = "mskDate"
		mskTime2.OcxState = CType(resources.GetObject("mskTime2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskTime2.Size = New System.Drawing.Size(117, 29)
		Me.mskTime2.Location = New System.Drawing.Point(140, 100)
		Me.mskTime2.TabIndex = 14
		Me.mskTime2.Name = "mskTime2"
		Me.lblCompany.Size = New System.Drawing.Size(239, 24)
		Me.lblCompany.Location = New System.Drawing.Point(128, 33)
		Me.lblCompany.TabIndex = 13
		Me.lblCompany.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCompany.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCompany.BackColor = System.Drawing.SystemColors.Control
		Me.lblCompany.Enabled = True
		Me.lblCompany.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCompany.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCompany.UseMnemonic = True
		Me.lblCompany.Visible = True
		Me.lblCompany.AutoSize = False
		Me.lblCompany.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCompany.Name = "lblCompany"
		Me._Label1_1.Text = "Company:"
		Me._Label1_1.Size = New System.Drawing.Size(72, 24)
		Me._Label1_1.Location = New System.Drawing.Point(48, 33)
		Me._Label1_1.TabIndex = 12
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
		Me._lblDate_3.Text = "Subject:"
		Me._lblDate_3.Size = New System.Drawing.Size(67, 24)
		Me._lblDate_3.Location = New System.Drawing.Point(48, 138)
		Me._lblDate_3.TabIndex = 11
		Me._lblDate_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblDate_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblDate_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblDate_3.Enabled = True
		Me._lblDate_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblDate_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblDate_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblDate_3.UseMnemonic = True
		Me._lblDate_3.Visible = True
		Me._lblDate_3.AutoSize = False
		Me._lblDate_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblDate_3.Name = "_lblDate_3"
		Me.lblName.Size = New System.Drawing.Size(239, 24)
		Me.lblName.Location = New System.Drawing.Point(128, 3)
		Me.lblName.TabIndex = 10
		Me.lblName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblName.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblName.BackColor = System.Drawing.SystemColors.Control
		Me.lblName.Enabled = True
		Me.lblName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblName.UseMnemonic = True
		Me.lblName.Visible = True
		Me.lblName.AutoSize = False
		Me.lblName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblName.Name = "lblName"
		Me._Label1_0.Text = "Name:"
		Me._Label1_0.Size = New System.Drawing.Size(72, 24)
		Me._Label1_0.Location = New System.Drawing.Point(48, 3)
		Me._Label1_0.TabIndex = 9
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._lblDate_2.Text = "Note:"
		Me._lblDate_2.Size = New System.Drawing.Size(67, 24)
		Me._lblDate_2.Location = New System.Drawing.Point(45, 173)
		Me._lblDate_2.TabIndex = 8
		Me._lblDate_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblDate_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblDate_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblDate_2.Enabled = True
		Me._lblDate_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblDate_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblDate_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblDate_2.UseMnemonic = True
		Me._lblDate_2.Visible = True
		Me._lblDate_2.AutoSize = False
		Me._lblDate_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblDate_2.Name = "_lblDate_2"
		Me._lblDate_1.Text = "Time:"
		Me._lblDate_1.Size = New System.Drawing.Size(64, 24)
		Me._lblDate_1.Location = New System.Drawing.Point(48, 103)
		Me._lblDate_1.TabIndex = 7
		Me._lblDate_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblDate_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblDate_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblDate_1.Enabled = True
		Me._lblDate_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblDate_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblDate_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblDate_1.UseMnemonic = True
		Me._lblDate_1.Visible = True
		Me._lblDate_1.AutoSize = False
		Me._lblDate_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblDate_1.Name = "_lblDate_1"
		Me._lblDate_0.Text = "Date:"
		Me._lblDate_0.Size = New System.Drawing.Size(67, 24)
		Me._lblDate_0.Location = New System.Drawing.Point(48, 68)
		Me._lblDate_0.TabIndex = 6
		Me._lblDate_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblDate_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblDate_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblDate_0.Enabled = True
		Me._lblDate_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblDate_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblDate_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblDate_0.UseMnemonic = True
		Me._lblDate_0.Visible = True
		Me._lblDate_0.AutoSize = False
		Me._lblDate_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblDate_0.Name = "_lblDate_0"
		Me.Controls.Add(chkSales)
		Me.Controls.Add(txtSubject)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(txtNote)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(mskTime)
		Me.Controls.Add(mskDate)
		Me.Controls.Add(mskTime2)
		Me.Controls.Add(lblCompany)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_lblDate_3)
		Me.Controls.Add(lblName)
		Me.Controls.Add(_Label1_0)
		Me.Controls.Add(_lblDate_2)
		Me.Controls.Add(_lblDate_1)
		Me.Controls.Add(_lblDate_0)
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.lblDate.SetIndex(_lblDate_3, CType(3, Short))
		Me.lblDate.SetIndex(_lblDate_2, CType(2, Short))
		Me.lblDate.SetIndex(_lblDate_1, CType(1, Short))
		Me.lblDate.SetIndex(_lblDate_0, CType(0, Short))
		CType(Me.lblDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskTime2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskTime, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class