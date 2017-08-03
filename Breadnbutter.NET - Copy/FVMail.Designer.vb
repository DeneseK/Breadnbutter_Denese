<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FVMail
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		'This form is an MDI child.
		'This code simulates the VB6 
		' functionality of automatically
		' loading and showing an MDI
		' child's parent.
		Me.MDIParent = Breadnbutter.FMain
		Breadnbutter.FMain.Show
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
	Public WithEvents cmdCompleted As System.Windows.Forms.Button
	Public WithEvents cmdRefresh As System.Windows.Forms.Button
	Public WithEvents File1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
	Public WithEvents cmdDetails As System.Windows.Forms.Button
	Public WithEvents webBody As System.Windows.Forms.WebBrowser
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents txtBody As System.Windows.Forms.TextBox
	Public WithEvents cmdContactInfo As System.Windows.Forms.Button
	Public WithEvents cmbCaller As System.Windows.Forms.ComboBox
	Public WithEvents cmdGetNames As System.Windows.Forms.Button
	Public WithEvents txtPhone As System.Windows.Forms.TextBox
	Public WithEvents txtsubject As System.Windows.Forms.TextBox
	Public WithEvents cmdForward As System.Windows.Forms.Button
	Public WithEvents cmdBrowser As System.Windows.Forms.Button
	Public WithEvents cmbComment As System.Windows.Forms.ComboBox
	Public WithEvents chkComp As System.Windows.Forms.CheckBox
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents fraDetails As System.Windows.Forms.GroupBox
	Public WithEvents cmdDelete As System.Windows.Forms.Button
	Public WithEvents cmbMessageGroup As System.Windows.Forms.ComboBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents cmdAll As System.Windows.Forms.Button
	Public WithEvents cmdOld As System.Windows.Forms.Button
	Public WithEvents cmdNew As System.Windows.Forms.Button
	Public WithEvents cmdExit As System.Windows.Forms.Button
	Public WithEvents cmdEditGroups As System.Windows.Forms.Button
	Public WithEvents cmdPlay As System.Windows.Forms.Button
	Public WithEvents _ListView1_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListView1_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents ListView1 As System.Windows.Forms.ListView
	Public WithEvents Shape1 As System.Windows.Forms.Label
	Public WithEvents lblGroups As System.Windows.Forms.Label
	Public WithEvents lblMessageGroup As System.Windows.Forms.Label
	Public WithEvents lblLastClient As System.Windows.Forms.Label
	Public WithEvents lblLastServer As System.Windows.Forms.Label
	Public WithEvents lblcount As System.Windows.Forms.Label
	Public WithEvents lblShow As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FVMail))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCompleted = New System.Windows.Forms.Button
		Me.cmdRefresh = New System.Windows.Forms.Button
		Me.File1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
		Me.cmdDetails = New System.Windows.Forms.Button
		Me.fraDetails = New System.Windows.Forms.GroupBox
		Me.webBody = New System.Windows.Forms.WebBrowser
		Me.cmdSave = New System.Windows.Forms.Button
		Me.txtBody = New System.Windows.Forms.TextBox
		Me.cmdContactInfo = New System.Windows.Forms.Button
		Me.cmbCaller = New System.Windows.Forms.ComboBox
		Me.cmdGetNames = New System.Windows.Forms.Button
		Me.txtPhone = New System.Windows.Forms.TextBox
		Me.txtsubject = New System.Windows.Forms.TextBox
		Me.cmdForward = New System.Windows.Forms.Button
		Me.cmdBrowser = New System.Windows.Forms.Button
		Me.cmbComment = New System.Windows.Forms.ComboBox
		Me.chkComp = New System.Windows.Forms.CheckBox
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.cmdDelete = New System.Windows.Forms.Button
		Me.cmbMessageGroup = New System.Windows.Forms.ComboBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.cmdAll = New System.Windows.Forms.Button
		Me.cmdOld = New System.Windows.Forms.Button
		Me.cmdNew = New System.Windows.Forms.Button
		Me.cmdExit = New System.Windows.Forms.Button
		Me.cmdEditGroups = New System.Windows.Forms.Button
		Me.cmdPlay = New System.Windows.Forms.Button
		Me.ListView1 = New System.Windows.Forms.ListView
		Me._ListView1_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._ListView1_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me.Shape1 = New System.Windows.Forms.Label
		Me.lblGroups = New System.Windows.Forms.Label
		Me.lblMessageGroup = New System.Windows.Forms.Label
		Me.lblLastClient = New System.Windows.Forms.Label
		Me.lblLastServer = New System.Windows.Forms.Label
		Me.lblcount = New System.Windows.Forms.Label
		Me.lblShow = New System.Windows.Forms.Label
		Me.fraDetails.SuspendLayout()
		Me.ListView1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "VMailClient"
		Me.ClientSize = New System.Drawing.Size(950, 717)
		Me.Location = New System.Drawing.Point(193, 114)
		Me.ControlBox = False
		Me.Icon = CType(resources.GetObject("FVMail.Icon"), System.Drawing.Icon)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.Visible = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.Name = "FVMail"
		Me.cmdCompleted.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCompleted.Text = "&Completed"
		Me.cmdCompleted.Size = New System.Drawing.Size(172, 42)
		Me.cmdCompleted.Location = New System.Drawing.Point(420, 320)
		Me.cmdCompleted.TabIndex = 35
		Me.cmdCompleted.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCompleted.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCompleted.CausesValidation = True
		Me.cmdCompleted.Enabled = True
		Me.cmdCompleted.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCompleted.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCompleted.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCompleted.TabStop = True
		Me.cmdCompleted.Name = "cmdCompleted"
		Me.cmdRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRefresh.Text = "&Refresh"
		Me.cmdRefresh.Size = New System.Drawing.Size(112, 52)
		Me.cmdRefresh.Location = New System.Drawing.Point(160, 290)
		Me.cmdRefresh.TabIndex = 33
		Me.cmdRefresh.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRefresh.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRefresh.CausesValidation = True
		Me.cmdRefresh.Enabled = True
		Me.cmdRefresh.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRefresh.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRefresh.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRefresh.TabStop = True
		Me.cmdRefresh.Name = "cmdRefresh"
		Me.File1.Size = New System.Drawing.Size(152, 24)
		Me.File1.Location = New System.Drawing.Point(-833, 460)
		Me.File1.TabIndex = 31
		Me.File1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.File1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.File1.Archive = True
		Me.File1.BackColor = System.Drawing.SystemColors.Window
		Me.File1.CausesValidation = True
		Me.File1.Enabled = True
		Me.File1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.File1.Hidden = False
		Me.File1.Cursor = System.Windows.Forms.Cursors.Default
		Me.File1.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.File1.Normal = True
		Me.File1.Pattern = "*.*"
		Me.File1.ReadOnly = True
		Me.File1.System = False
		Me.File1.TabStop = True
		Me.File1.TopIndex = 0
		Me.File1.Visible = True
		Me.File1.Name = "File1"
		Me.cmdDetails.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDetails.Text = "&Details"
		Me.cmdDetails.Size = New System.Drawing.Size(187, 57)
		Me.cmdDetails.Location = New System.Drawing.Point(10, 285)
		Me.cmdDetails.TabIndex = 16
		Me.cmdDetails.Visible = False
		Me.cmdDetails.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDetails.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDetails.CausesValidation = True
		Me.cmdDetails.Enabled = True
		Me.cmdDetails.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDetails.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDetails.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDetails.TabStop = True
		Me.cmdDetails.Name = "cmdDetails"
		Me.fraDetails.Text = "Details"
		Me.fraDetails.ForeColor = System.Drawing.SystemColors.WindowText
		Me.fraDetails.Size = New System.Drawing.Size(962, 312)
		Me.fraDetails.Location = New System.Drawing.Point(10, 390)
		Me.fraDetails.TabIndex = 15
		Me.fraDetails.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraDetails.BackColor = System.Drawing.SystemColors.Control
		Me.fraDetails.Enabled = True
		Me.fraDetails.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraDetails.Visible = True
		Me.fraDetails.Name = "fraDetails"
		Me.webBody.Size = New System.Drawing.Size(582, 122)
		Me.webBody.Location = New System.Drawing.Point(370, 50)
		Me.webBody.TabIndex = 36
		Me.webBody.AllowWebBrowserDrop = True
		Me.webBody.Name = "webBody"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save Changes"
		Me.cmdSave.Size = New System.Drawing.Size(122, 32)
		Me.cmdSave.Location = New System.Drawing.Point(470, 180)
		Me.cmdSave.TabIndex = 34
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.txtBody.AutoSize = False
		Me.txtBody.Size = New System.Drawing.Size(582, 122)
		Me.txtBody.Location = New System.Drawing.Point(370, 50)
		Me.txtBody.ReadOnly = True
		Me.txtBody.MultiLine = True
		Me.txtBody.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtBody.TabIndex = 32
		Me.txtBody.Text = "Text1"
		Me.txtBody.Visible = False
		Me.txtBody.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBody.AcceptsReturn = True
		Me.txtBody.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBody.BackColor = System.Drawing.SystemColors.Window
		Me.txtBody.CausesValidation = True
		Me.txtBody.Enabled = True
		Me.txtBody.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtBody.HideSelection = True
		Me.txtBody.Maxlength = 0
		Me.txtBody.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBody.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBody.TabStop = True
		Me.txtBody.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBody.Name = "txtBody"
		Me.cmdContactInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdContactInfo.Text = "Contact Info."
		Me.cmdContactInfo.Enabled = False
		Me.cmdContactInfo.Size = New System.Drawing.Size(112, 32)
		Me.cmdContactInfo.Location = New System.Drawing.Point(250, 160)
		Me.cmdContactInfo.TabIndex = 30
		Me.cmdContactInfo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdContactInfo.BackColor = System.Drawing.SystemColors.Control
		Me.cmdContactInfo.CausesValidation = True
		Me.cmdContactInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdContactInfo.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdContactInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdContactInfo.TabStop = True
		Me.cmdContactInfo.Name = "cmdContactInfo"
		Me.cmbCaller.Size = New System.Drawing.Size(232, 27)
		Me.cmbCaller.Location = New System.Drawing.Point(10, 160)
		Me.cmbCaller.TabIndex = 29
		Me.cmbCaller.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbCaller.BackColor = System.Drawing.SystemColors.Window
		Me.cmbCaller.CausesValidation = True
		Me.cmbCaller.Enabled = True
		Me.cmbCaller.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbCaller.IntegralHeight = True
		Me.cmbCaller.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbCaller.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbCaller.Sorted = False
		Me.cmbCaller.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbCaller.TabStop = True
		Me.cmbCaller.Visible = True
		Me.cmbCaller.Name = "cmbCaller"
		Me.cmdGetNames.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdGetNames.Text = "Get Names"
		Me.cmdGetNames.Size = New System.Drawing.Size(112, 32)
		Me.cmdGetNames.Location = New System.Drawing.Point(250, 100)
		Me.cmdGetNames.TabIndex = 28
		Me.cmdGetNames.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdGetNames.BackColor = System.Drawing.SystemColors.Control
		Me.cmdGetNames.CausesValidation = True
		Me.cmdGetNames.Enabled = True
		Me.cmdGetNames.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdGetNames.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdGetNames.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdGetNames.TabStop = True
		Me.cmdGetNames.Name = "cmdGetNames"
		Me.txtPhone.AutoSize = False
		Me.txtPhone.Size = New System.Drawing.Size(230, 29)
		Me.txtPhone.Location = New System.Drawing.Point(10, 100)
		Me.txtPhone.ReadOnly = True
		Me.txtPhone.TabIndex = 27
		Me.txtPhone.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPhone.AcceptsReturn = True
		Me.txtPhone.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPhone.BackColor = System.Drawing.SystemColors.Window
		Me.txtPhone.CausesValidation = True
		Me.txtPhone.Enabled = True
		Me.txtPhone.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPhone.HideSelection = True
		Me.txtPhone.Maxlength = 0
		Me.txtPhone.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPhone.MultiLine = False
		Me.txtPhone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPhone.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPhone.TabStop = True
		Me.txtPhone.Visible = True
		Me.txtPhone.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPhone.Name = "txtPhone"
		Me.txtsubject.AutoSize = False
		Me.txtsubject.Size = New System.Drawing.Size(294, 29)
		Me.txtsubject.Location = New System.Drawing.Point(10, 40)
		Me.txtsubject.ReadOnly = True
		Me.txtsubject.TabIndex = 26
		Me.txtsubject.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtsubject.AcceptsReturn = True
		Me.txtsubject.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtsubject.BackColor = System.Drawing.SystemColors.Window
		Me.txtsubject.CausesValidation = True
		Me.txtsubject.Enabled = True
		Me.txtsubject.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtsubject.HideSelection = True
		Me.txtsubject.Maxlength = 0
		Me.txtsubject.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtsubject.MultiLine = False
		Me.txtsubject.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtsubject.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtsubject.TabStop = True
		Me.txtsubject.Visible = True
		Me.txtsubject.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtsubject.Name = "txtsubject"
		Me.cmdForward.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdForward.Text = "Forward Message"
		Me.cmdForward.Size = New System.Drawing.Size(125, 29)
		Me.cmdForward.Location = New System.Drawing.Point(813, 17)
		Me.cmdForward.TabIndex = 25
		Me.cmdForward.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdForward.BackColor = System.Drawing.SystemColors.Control
		Me.cmdForward.CausesValidation = True
		Me.cmdForward.Enabled = True
		Me.cmdForward.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdForward.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdForward.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdForward.TabStop = True
		Me.cmdForward.Name = "cmdForward"
		Me.cmdBrowser.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdBrowser.Text = "View Body In Browser"
		Me.cmdBrowser.Size = New System.Drawing.Size(152, 29)
		Me.cmdBrowser.Location = New System.Drawing.Point(653, 17)
		Me.cmdBrowser.TabIndex = 24
		Me.cmdBrowser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBrowser.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBrowser.CausesValidation = True
		Me.cmdBrowser.Enabled = True
		Me.cmdBrowser.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBrowser.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBrowser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBrowser.TabStop = True
		Me.cmdBrowser.Name = "cmdBrowser"
		Me.cmbComment.Size = New System.Drawing.Size(282, 27)
		Me.cmbComment.Location = New System.Drawing.Point(670, 180)
		Me.cmbComment.TabIndex = 23
		Me.cmbComment.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbComment.BackColor = System.Drawing.SystemColors.Window
		Me.cmbComment.CausesValidation = True
		Me.cmbComment.Enabled = True
		Me.cmbComment.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbComment.IntegralHeight = True
		Me.cmbComment.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbComment.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbComment.Sorted = False
		Me.cmbComment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbComment.TabStop = True
		Me.cmbComment.Visible = True
		Me.cmbComment.Name = "cmbComment"
		Me.chkComp.Text = "Completed"
		Me.chkComp.Size = New System.Drawing.Size(122, 32)
		Me.chkComp.Location = New System.Drawing.Point(370, 180)
		Me.chkComp.TabIndex = 22
		Me.chkComp.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkComp.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkComp.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkComp.BackColor = System.Drawing.SystemColors.Control
		Me.chkComp.CausesValidation = True
		Me.chkComp.Enabled = True
		Me.chkComp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkComp.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkComp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkComp.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkComp.TabStop = True
		Me.chkComp.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkComp.Visible = True
		Me.chkComp.Name = "chkComp"
		Me.Label9.Text = "Comment:"
		Me.Label9.Size = New System.Drawing.Size(112, 22)
		Me.Label9.Location = New System.Drawing.Point(603, 184)
		Me.Label9.TabIndex = 21
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label9.BackColor = System.Drawing.SystemColors.Control
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = False
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label7.Text = "Caller:"
		Me.Label7.Size = New System.Drawing.Size(92, 22)
		Me.Label7.Location = New System.Drawing.Point(10, 140)
		Me.Label7.TabIndex = 20
		Me.Label7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.BackColor = System.Drawing.SystemColors.Control
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.Label5.Text = "Phone Number:"
		Me.Label5.Size = New System.Drawing.Size(112, 22)
		Me.Label5.Location = New System.Drawing.Point(10, 80)
		Me.Label5.TabIndex = 19
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
		Me.Label1.Text = "Subject:"
		Me.Label1.Size = New System.Drawing.Size(82, 22)
		Me.Label1.Location = New System.Drawing.Point(10, 20)
		Me.Label1.TabIndex = 18
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
		Me.Label4.Text = "Body:"
		Me.Label4.Size = New System.Drawing.Size(102, 22)
		Me.Label4.Location = New System.Drawing.Point(370, 30)
		Me.Label4.TabIndex = 17
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
		Me.cmdDelete.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDelete.Text = "D&elete"
		Me.cmdDelete.Size = New System.Drawing.Size(102, 62)
		Me.cmdDelete.Location = New System.Drawing.Point(790, 340)
		Me.cmdDelete.TabIndex = 13
		Me.cmdDelete.Visible = False
		Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDelete.CausesValidation = True
		Me.cmdDelete.Enabled = True
		Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDelete.TabStop = True
		Me.cmdDelete.Name = "cmdDelete"
		Me.cmbMessageGroup.Size = New System.Drawing.Size(147, 27)
		Me.cmbMessageGroup.Location = New System.Drawing.Point(735, 300)
		Me.cmbMessageGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbMessageGroup.TabIndex = 11
		Me.cmbMessageGroup.Visible = False
		Me.cmbMessageGroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbMessageGroup.BackColor = System.Drawing.SystemColors.Window
		Me.cmbMessageGroup.CausesValidation = True
		Me.cmbMessageGroup.Enabled = True
		Me.cmbMessageGroup.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbMessageGroup.IntegralHeight = True
		Me.cmbMessageGroup.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbMessageGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbMessageGroup.Sorted = False
		Me.cmbMessageGroup.TabStop = True
		Me.cmbMessageGroup.Name = "cmbMessageGroup"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.cmdAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAll.Text = "&All"
		Me.cmdAll.Size = New System.Drawing.Size(134, 59)
		Me.cmdAll.Location = New System.Drawing.Point(318, 345)
		Me.cmdAll.TabIndex = 5
		Me.cmdAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAll.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAll.CausesValidation = True
		Me.cmdAll.Enabled = True
		Me.cmdAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAll.TabStop = True
		Me.cmdAll.Name = "cmdAll"
		Me.cmdOld.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOld.Text = "O&ld"
		Me.cmdOld.Size = New System.Drawing.Size(92, 54)
		Me.cmdOld.Location = New System.Drawing.Point(215, 350)
		Me.cmdOld.TabIndex = 4
		Me.cmdOld.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOld.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOld.CausesValidation = True
		Me.cmdOld.Enabled = True
		Me.cmdOld.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOld.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOld.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOld.TabStop = True
		Me.cmdOld.Name = "cmdOld"
		Me.cmdNew.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdNew.Text = "&New"
		Me.cmdNew.Size = New System.Drawing.Size(102, 47)
		Me.cmdNew.Location = New System.Drawing.Point(105, 353)
		Me.cmdNew.TabIndex = 3
		Me.cmdNew.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNew.BackColor = System.Drawing.SystemColors.Control
		Me.cmdNew.CausesValidation = True
		Me.cmdNew.Enabled = True
		Me.cmdNew.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdNew.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdNew.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdNew.TabStop = True
		Me.cmdNew.Name = "cmdNew"
		Me.cmdExit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdExit.Text = "&Exit"
		Me.cmdExit.Size = New System.Drawing.Size(207, 59)
		Me.cmdExit.Location = New System.Drawing.Point(568, 338)
		Me.cmdExit.TabIndex = 2
		Me.cmdExit.Visible = False
		Me.cmdExit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdExit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdExit.CausesValidation = True
		Me.cmdExit.Enabled = True
		Me.cmdExit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdExit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdExit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdExit.TabStop = True
		Me.cmdExit.Name = "cmdExit"
		Me.cmdEditGroups.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdEditGroups.Text = "Edit Groups"
		Me.cmdEditGroups.Size = New System.Drawing.Size(182, 52)
		Me.cmdEditGroups.Location = New System.Drawing.Point(430, 290)
		Me.cmdEditGroups.TabIndex = 1
		Me.cmdEditGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdEditGroups.BackColor = System.Drawing.SystemColors.Control
		Me.cmdEditGroups.CausesValidation = True
		Me.cmdEditGroups.Enabled = True
		Me.cmdEditGroups.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdEditGroups.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdEditGroups.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdEditGroups.TabStop = True
		Me.cmdEditGroups.Name = "cmdEditGroups"
		Me.cmdPlay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPlay.Text = "&Play"
		Me.cmdPlay.Size = New System.Drawing.Size(182, 54)
		Me.cmdPlay.Location = New System.Drawing.Point(228, 288)
		Me.cmdPlay.TabIndex = 0
		Me.cmdPlay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPlay.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPlay.CausesValidation = True
		Me.cmdPlay.Enabled = True
		Me.cmdPlay.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPlay.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPlay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPlay.TabStop = True
		Me.cmdPlay.Name = "cmdPlay"
		Me.ListView1.Size = New System.Drawing.Size(769, 274)
		Me.ListView1.Location = New System.Drawing.Point(10, 0)
		Me.ListView1.TabIndex = 6
		Me.ListView1.TabStop = 0
		Me.ListView1.View = System.Windows.Forms.View.Details
		Me.ListView1.LabelEdit = False
		Me.ListView1.LabelWrap = True
		Me.ListView1.HideSelection = False
		Me.ListView1.FullRowSelect = True
		Me.ListView1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ListView1.BackColor = System.Drawing.SystemColors.Window
		Me.ListView1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ListView1.Name = "ListView1"
		Me._ListView1_ColumnHeader_1.Text = "Group"
		Me._ListView1_ColumnHeader_1.Width = 72247
		Me._ListView1_ColumnHeader_2.Width = 212
		Me.Shape1.BackColor = System.Drawing.SystemColors.Control
		Me.Shape1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Shape1.Size = New System.Drawing.Size(102, 62)
		Me.Shape1.Location = New System.Drawing.Point(800, 130)
		Me.Shape1.Visible = True
		Me.Shape1.Name = "Shape1"
		Me.lblGroups.Text = "ALL"
		Me.lblGroups.Size = New System.Drawing.Size(102, 22)
		Me.lblGroups.Location = New System.Drawing.Point(790, 250)
		Me.lblGroups.TabIndex = 14
		Me.lblGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblGroups.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblGroups.BackColor = System.Drawing.SystemColors.Control
		Me.lblGroups.Enabled = True
		Me.lblGroups.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblGroups.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblGroups.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblGroups.UseMnemonic = True
		Me.lblGroups.Visible = True
		Me.lblGroups.AutoSize = False
		Me.lblGroups.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblGroups.Name = "lblGroups"
		Me.lblMessageGroup.Text = "Message Group:"
		Me.lblMessageGroup.Size = New System.Drawing.Size(102, 22)
		Me.lblMessageGroup.Location = New System.Drawing.Point(795, 200)
		Me.lblMessageGroup.TabIndex = 12
		Me.lblMessageGroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMessageGroup.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblMessageGroup.BackColor = System.Drawing.SystemColors.Control
		Me.lblMessageGroup.Enabled = True
		Me.lblMessageGroup.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblMessageGroup.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblMessageGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMessageGroup.UseMnemonic = True
		Me.lblMessageGroup.Visible = True
		Me.lblMessageGroup.AutoSize = False
		Me.lblMessageGroup.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblMessageGroup.Name = "lblMessageGroup"
		Me.lblLastClient.Text = "Label1"
		Me.lblLastClient.Size = New System.Drawing.Size(297, 17)
		Me.lblLastClient.Location = New System.Drawing.Point(215, 428)
		Me.lblLastClient.TabIndex = 10
		Me.lblLastClient.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLastClient.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLastClient.BackColor = System.Drawing.SystemColors.Control
		Me.lblLastClient.Enabled = True
		Me.lblLastClient.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLastClient.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLastClient.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLastClient.UseMnemonic = True
		Me.lblLastClient.Visible = True
		Me.lblLastClient.AutoSize = False
		Me.lblLastClient.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLastClient.Name = "lblLastClient"
		Me.lblLastServer.Text = "Label1"
		Me.lblLastServer.Size = New System.Drawing.Size(294, 19)
		Me.lblLastServer.Location = New System.Drawing.Point(620, 428)
		Me.lblLastServer.TabIndex = 9
		Me.lblLastServer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLastServer.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLastServer.BackColor = System.Drawing.SystemColors.Control
		Me.lblLastServer.Enabled = True
		Me.lblLastServer.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLastServer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLastServer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLastServer.UseMnemonic = True
		Me.lblLastServer.Visible = True
		Me.lblLastServer.AutoSize = False
		Me.lblLastServer.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLastServer.Name = "lblLastServer"
		Me.lblcount.Text = "Label1"
		Me.lblcount.Size = New System.Drawing.Size(214, 19)
		Me.lblcount.Location = New System.Drawing.Point(0, 428)
		Me.lblcount.TabIndex = 8
		Me.lblcount.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblcount.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblcount.BackColor = System.Drawing.SystemColors.Control
		Me.lblcount.Enabled = True
		Me.lblcount.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblcount.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblcount.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblcount.UseMnemonic = True
		Me.lblcount.Visible = True
		Me.lblcount.AutoSize = False
		Me.lblcount.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblcount.Name = "lblcount"
		Me.lblShow.Text = "Show:"
		Me.lblShow.Size = New System.Drawing.Size(42, 22)
		Me.lblShow.Location = New System.Drawing.Point(48, 348)
		Me.lblShow.TabIndex = 7
		Me.lblShow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblShow.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblShow.BackColor = System.Drawing.SystemColors.Control
		Me.lblShow.Enabled = True
		Me.lblShow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblShow.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblShow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblShow.UseMnemonic = True
		Me.lblShow.Visible = True
		Me.lblShow.AutoSize = False
		Me.lblShow.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblShow.Name = "lblShow"
		Me.Controls.Add(cmdCompleted)
		Me.Controls.Add(cmdRefresh)
		Me.Controls.Add(File1)
		Me.Controls.Add(cmdDetails)
		Me.Controls.Add(fraDetails)
		Me.Controls.Add(cmdDelete)
		Me.Controls.Add(cmbMessageGroup)
		Me.Controls.Add(cmdAll)
		Me.Controls.Add(cmdOld)
		Me.Controls.Add(cmdNew)
		Me.Controls.Add(cmdExit)
		Me.Controls.Add(cmdEditGroups)
		Me.Controls.Add(cmdPlay)
		Me.Controls.Add(ListView1)
		Me.Controls.Add(Shape1)
		Me.Controls.Add(lblGroups)
		Me.Controls.Add(lblMessageGroup)
		Me.Controls.Add(lblLastClient)
		Me.Controls.Add(lblLastServer)
		Me.Controls.Add(lblcount)
		Me.Controls.Add(lblShow)
		Me.fraDetails.Controls.Add(webBody)
		Me.fraDetails.Controls.Add(cmdSave)
		Me.fraDetails.Controls.Add(txtBody)
		Me.fraDetails.Controls.Add(cmdContactInfo)
		Me.fraDetails.Controls.Add(cmbCaller)
		Me.fraDetails.Controls.Add(cmdGetNames)
		Me.fraDetails.Controls.Add(txtPhone)
		Me.fraDetails.Controls.Add(txtsubject)
		Me.fraDetails.Controls.Add(cmdForward)
		Me.fraDetails.Controls.Add(cmdBrowser)
		Me.fraDetails.Controls.Add(cmbComment)
		Me.fraDetails.Controls.Add(chkComp)
		Me.fraDetails.Controls.Add(Label9)
		Me.fraDetails.Controls.Add(Label7)
		Me.fraDetails.Controls.Add(Label5)
		Me.fraDetails.Controls.Add(Label1)
		Me.fraDetails.Controls.Add(Label4)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_1)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_2)
		Me.fraDetails.ResumeLayout(False)
		Me.ListView1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class