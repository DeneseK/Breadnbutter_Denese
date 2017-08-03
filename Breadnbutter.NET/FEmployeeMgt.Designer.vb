<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FEmployeeMgt
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
	Public WithEvents txtIcon As System.Windows.Forms.TextBox
	Public WithEvents cmdEdit As System.Windows.Forms.Button
	Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents cmdDelete As System.Windows.Forms.Button
	Public WithEvents txtMail As System.Windows.Forms.TextBox
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents txtFirst As System.Windows.Forms.TextBox
	Public WithEvents txtMid As System.Windows.Forms.TextBox
	Public WithEvents txtLast As System.Windows.Forms.TextBox
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents chkAuthorizations As System.Windows.Forms.CheckBox
	Public WithEvents chkSales As System.Windows.Forms.CheckBox
	Public WithEvents chkSupport As System.Windows.Forms.CheckBox
	Public WithEvents chkOperator As System.Windows.Forms.CheckBox
	Public WithEvents fmeGroups As System.Windows.Forms.GroupBox
	Public WithEvents chkDev As System.Windows.Forms.CheckBox
	Public WithEvents chkWorkSupport As System.Windows.Forms.CheckBox
	Public WithEvents chkWorkSales As System.Windows.Forms.CheckBox
	Public WithEvents chkManagement As System.Windows.Forms.CheckBox
	Public WithEvents fmeWorkGroups As System.Windows.Forms.GroupBox
	Public WithEvents txtExt As System.Windows.Forms.TextBox
	Public WithEvents optLow As System.Windows.Forms.RadioButton
	Public WithEvents optHigh As System.Windows.Forms.RadioButton
	Public WithEvents fmeSecurity As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents lblMail As System.Windows.Forms.Label
	Public WithEvents lblPass As System.Windows.Forms.Label
	Public WithEvents lblFirst As System.Windows.Forms.Label
	Public WithEvents lblLast As System.Windows.Forms.Label
	Public WithEvents lblMid As System.Windows.Forms.Label
	Public WithEvents lblExt As System.Windows.Forms.Label
	Public WithEvents fmeAddEdit As System.Windows.Forms.GroupBox
	Public WithEvents _ListView1_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListView1_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListView1_ColumnHeader_3 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListView1_ColumnHeader_4 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListView1_ColumnHeader_5 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListView1_ColumnHeader_6 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListView1_ColumnHeader_7 As System.Windows.Forms.ColumnHeader
	Public WithEvents ListView1 As System.Windows.Forms.ListView
	Public WithEvents Frame1 As System.Windows.Forms.Panel
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FEmployeeMgt))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame1 = New System.Windows.Forms.Panel
		Me.txtIcon = New System.Windows.Forms.TextBox
		Me.cmdEdit = New System.Windows.Forms.Button
		Me.cmdAdd = New System.Windows.Forms.Button
		Me.cmdDelete = New System.Windows.Forms.Button
		Me.fmeAddEdit = New System.Windows.Forms.GroupBox
		Me.txtMail = New System.Windows.Forms.TextBox
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.txtFirst = New System.Windows.Forms.TextBox
		Me.txtMid = New System.Windows.Forms.TextBox
		Me.txtLast = New System.Windows.Forms.TextBox
		Me.cmdSave = New System.Windows.Forms.Button
		Me.fmeGroups = New System.Windows.Forms.GroupBox
		Me.chkAuthorizations = New System.Windows.Forms.CheckBox
		Me.chkSales = New System.Windows.Forms.CheckBox
		Me.chkSupport = New System.Windows.Forms.CheckBox
		Me.chkOperator = New System.Windows.Forms.CheckBox
		Me.fmeWorkGroups = New System.Windows.Forms.GroupBox
		Me.chkDev = New System.Windows.Forms.CheckBox
		Me.chkWorkSupport = New System.Windows.Forms.CheckBox
		Me.chkWorkSales = New System.Windows.Forms.CheckBox
		Me.chkManagement = New System.Windows.Forms.CheckBox
		Me.txtExt = New System.Windows.Forms.TextBox
		Me.fmeSecurity = New System.Windows.Forms.GroupBox
		Me.optLow = New System.Windows.Forms.RadioButton
		Me.optHigh = New System.Windows.Forms.RadioButton
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.lblMail = New System.Windows.Forms.Label
		Me.lblPass = New System.Windows.Forms.Label
		Me.lblFirst = New System.Windows.Forms.Label
		Me.lblLast = New System.Windows.Forms.Label
		Me.lblMid = New System.Windows.Forms.Label
		Me.lblExt = New System.Windows.Forms.Label
		Me.ListView1 = New System.Windows.Forms.ListView
		Me._ListView1_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._ListView1_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me._ListView1_ColumnHeader_3 = New System.Windows.Forms.ColumnHeader
		Me._ListView1_ColumnHeader_4 = New System.Windows.Forms.ColumnHeader
		Me._ListView1_ColumnHeader_5 = New System.Windows.Forms.ColumnHeader
		Me._ListView1_ColumnHeader_6 = New System.Windows.Forms.ColumnHeader
		Me._ListView1_ColumnHeader_7 = New System.Windows.Forms.ColumnHeader
		Me.Frame1.SuspendLayout()
		Me.fmeAddEdit.SuspendLayout()
		Me.fmeGroups.SuspendLayout()
		Me.fmeWorkGroups.SuspendLayout()
		Me.fmeSecurity.SuspendLayout()
		Me.ListView1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.SystemColors.AppWorkspace
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Employee Management"
		Me.ClientSize = New System.Drawing.Size(854, 712)
		Me.Location = New System.Drawing.Point(205, 207)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FEmployeeMgt"
		Me.Frame1.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.Frame1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Frame1.Size = New System.Drawing.Size(852, 602)
		Me.Frame1.Location = New System.Drawing.Point(0, 0)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.txtIcon.AutoSize = False
		Me.txtIcon.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.txtIcon.Size = New System.Drawing.Size(22, 24)
		Me.txtIcon.Location = New System.Drawing.Point(820, 320)
		Me.txtIcon.TabIndex = 33
		Me.txtIcon.TabStop = False
		Me.txtIcon.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtIcon.AcceptsReturn = True
		Me.txtIcon.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtIcon.CausesValidation = True
		Me.txtIcon.Enabled = True
		Me.txtIcon.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtIcon.HideSelection = True
		Me.txtIcon.ReadOnly = False
		Me.txtIcon.Maxlength = 0
		Me.txtIcon.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtIcon.MultiLine = False
		Me.txtIcon.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtIcon.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtIcon.Visible = True
		Me.txtIcon.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtIcon.Name = "txtIcon"
		Me.cmdEdit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdEdit.Text = "Edit"
		Me.cmdEdit.Size = New System.Drawing.Size(122, 32)
		Me.cmdEdit.Location = New System.Drawing.Point(370, 320)
		Me.cmdEdit.TabIndex = 3
		Me.cmdEdit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdEdit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdEdit.CausesValidation = True
		Me.cmdEdit.Enabled = True
		Me.cmdEdit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdEdit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdEdit.TabStop = True
		Me.cmdEdit.Name = "cmdEdit"
		Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAdd.Text = "Add New"
		Me.cmdAdd.Size = New System.Drawing.Size(112, 32)
		Me.cmdAdd.Location = New System.Drawing.Point(220, 320)
		Me.cmdAdd.TabIndex = 2
		Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAdd.CausesValidation = True
		Me.cmdAdd.Enabled = True
		Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAdd.TabStop = True
		Me.cmdAdd.Name = "cmdAdd"
		Me.cmdDelete.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDelete.Text = "Delete"
		Me.cmdDelete.Size = New System.Drawing.Size(112, 32)
		Me.cmdDelete.Location = New System.Drawing.Point(530, 320)
		Me.cmdDelete.TabIndex = 4
		Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDelete.CausesValidation = True
		Me.cmdDelete.Enabled = True
		Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDelete.TabStop = True
		Me.cmdDelete.Name = "cmdDelete"
		Me.fmeAddEdit.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.fmeAddEdit.Size = New System.Drawing.Size(792, 212)
		Me.fmeAddEdit.Location = New System.Drawing.Point(30, 360)
		Me.fmeAddEdit.TabIndex = 10
		Me.fmeAddEdit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fmeAddEdit.Enabled = True
		Me.fmeAddEdit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fmeAddEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fmeAddEdit.Visible = True
		Me.fmeAddEdit.Name = "fmeAddEdit"
		Me.txtMail.AutoSize = False
		Me.txtMail.Size = New System.Drawing.Size(222, 24)
		Me.txtMail.Location = New System.Drawing.Point(560, 40)
		Me.txtMail.TabIndex = 9
		Me.txtMail.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMail.AcceptsReturn = True
		Me.txtMail.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMail.BackColor = System.Drawing.SystemColors.Window
		Me.txtMail.CausesValidation = True
		Me.txtMail.Enabled = True
		Me.txtMail.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMail.HideSelection = True
		Me.txtMail.ReadOnly = False
		Me.txtMail.Maxlength = 0
		Me.txtMail.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMail.MultiLine = False
		Me.txtMail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMail.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMail.TabStop = True
		Me.txtMail.Visible = True
		Me.txtMail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMail.Name = "txtMail"
		Me.txtPassword.AutoSize = False
		Me.txtPassword.Size = New System.Drawing.Size(112, 24)
		Me.txtPassword.Location = New System.Drawing.Point(600, 100)
		Me.txtPassword.TabIndex = 21
		Me.txtPassword.Visible = False
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
		Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPassword.Name = "txtPassword"
		Me.txtFirst.AutoSize = False
		Me.txtFirst.Size = New System.Drawing.Size(142, 24)
		Me.txtFirst.Location = New System.Drawing.Point(10, 40)
		Me.txtFirst.TabIndex = 5
		Me.txtFirst.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFirst.AcceptsReturn = True
		Me.txtFirst.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFirst.BackColor = System.Drawing.SystemColors.Window
		Me.txtFirst.CausesValidation = True
		Me.txtFirst.Enabled = True
		Me.txtFirst.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFirst.HideSelection = True
		Me.txtFirst.ReadOnly = False
		Me.txtFirst.Maxlength = 0
		Me.txtFirst.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFirst.MultiLine = False
		Me.txtFirst.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFirst.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFirst.TabStop = True
		Me.txtFirst.Visible = True
		Me.txtFirst.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFirst.Name = "txtFirst"
		Me.txtMid.AutoSize = False
		Me.txtMid.Size = New System.Drawing.Size(142, 24)
		Me.txtMid.Location = New System.Drawing.Point(160, 40)
		Me.txtMid.TabIndex = 6
		Me.txtMid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMid.AcceptsReturn = True
		Me.txtMid.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMid.BackColor = System.Drawing.SystemColors.Window
		Me.txtMid.CausesValidation = True
		Me.txtMid.Enabled = True
		Me.txtMid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMid.HideSelection = True
		Me.txtMid.ReadOnly = False
		Me.txtMid.Maxlength = 0
		Me.txtMid.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMid.MultiLine = False
		Me.txtMid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMid.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMid.TabStop = True
		Me.txtMid.Visible = True
		Me.txtMid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMid.Name = "txtMid"
		Me.txtLast.AutoSize = False
		Me.txtLast.Size = New System.Drawing.Size(172, 24)
		Me.txtLast.Location = New System.Drawing.Point(310, 40)
		Me.txtLast.TabIndex = 7
		Me.txtLast.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLast.AcceptsReturn = True
		Me.txtLast.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLast.BackColor = System.Drawing.SystemColors.Window
		Me.txtLast.CausesValidation = True
		Me.txtLast.Enabled = True
		Me.txtLast.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLast.HideSelection = True
		Me.txtLast.ReadOnly = False
		Me.txtLast.Maxlength = 0
		Me.txtLast.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLast.MultiLine = False
		Me.txtLast.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLast.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLast.TabStop = True
		Me.txtLast.Visible = True
		Me.txtLast.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLast.Name = "txtLast"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save"
		Me.cmdSave.Size = New System.Drawing.Size(112, 25)
		Me.cmdSave.Location = New System.Drawing.Point(600, 130)
		Me.cmdSave.TabIndex = 22
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.fmeGroups.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.fmeGroups.Text = "Groups"
		Me.fmeGroups.Size = New System.Drawing.Size(132, 112)
		Me.fmeGroups.Location = New System.Drawing.Point(30, 80)
		Me.fmeGroups.TabIndex = 26
		Me.fmeGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fmeGroups.Enabled = True
		Me.fmeGroups.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fmeGroups.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fmeGroups.Visible = True
		Me.fmeGroups.Name = "fmeGroups"
		Me.chkAuthorizations.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.chkAuthorizations.Text = "Authorizations"
		Me.chkAuthorizations.Size = New System.Drawing.Size(112, 22)
		Me.chkAuthorizations.Location = New System.Drawing.Point(10, 20)
		Me.chkAuthorizations.TabIndex = 11
		Me.chkAuthorizations.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAuthorizations.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAuthorizations.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAuthorizations.CausesValidation = True
		Me.chkAuthorizations.Enabled = True
		Me.chkAuthorizations.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAuthorizations.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAuthorizations.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAuthorizations.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAuthorizations.TabStop = True
		Me.chkAuthorizations.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkAuthorizations.Visible = True
		Me.chkAuthorizations.Name = "chkAuthorizations"
		Me.chkSales.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.chkSales.Text = "Sales"
		Me.chkSales.Size = New System.Drawing.Size(82, 22)
		Me.chkSales.Location = New System.Drawing.Point(10, 40)
		Me.chkSales.TabIndex = 12
		Me.chkSales.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSales.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSales.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSales.CausesValidation = True
		Me.chkSales.Enabled = True
		Me.chkSales.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSales.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSales.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSales.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSales.TabStop = True
		Me.chkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSales.Visible = True
		Me.chkSales.Name = "chkSales"
		Me.chkSupport.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.chkSupport.Text = "Support"
		Me.chkSupport.Size = New System.Drawing.Size(82, 22)
		Me.chkSupport.Location = New System.Drawing.Point(10, 60)
		Me.chkSupport.TabIndex = 13
		Me.chkSupport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSupport.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSupport.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSupport.CausesValidation = True
		Me.chkSupport.Enabled = True
		Me.chkSupport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSupport.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSupport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSupport.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSupport.TabStop = True
		Me.chkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSupport.Visible = True
		Me.chkSupport.Name = "chkSupport"
		Me.chkOperator.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.chkOperator.Text = "Operator"
		Me.chkOperator.Size = New System.Drawing.Size(92, 22)
		Me.chkOperator.Location = New System.Drawing.Point(10, 80)
		Me.chkOperator.TabIndex = 14
		Me.chkOperator.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkOperator.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkOperator.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkOperator.CausesValidation = True
		Me.chkOperator.Enabled = True
		Me.chkOperator.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkOperator.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkOperator.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkOperator.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkOperator.TabStop = True
		Me.chkOperator.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkOperator.Visible = True
		Me.chkOperator.Name = "chkOperator"
		Me.fmeWorkGroups.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.fmeWorkGroups.Text = "WorkGroups"
		Me.fmeWorkGroups.Size = New System.Drawing.Size(132, 112)
		Me.fmeWorkGroups.Location = New System.Drawing.Point(230, 80)
		Me.fmeWorkGroups.TabIndex = 25
		Me.fmeWorkGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fmeWorkGroups.Enabled = True
		Me.fmeWorkGroups.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fmeWorkGroups.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fmeWorkGroups.Visible = True
		Me.fmeWorkGroups.Name = "fmeWorkGroups"
		Me.chkDev.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.chkDev.Text = "Development"
		Me.chkDev.Size = New System.Drawing.Size(112, 22)
		Me.chkDev.Location = New System.Drawing.Point(10, 80)
		Me.chkDev.TabIndex = 18
		Me.chkDev.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDev.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDev.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDev.CausesValidation = True
		Me.chkDev.Enabled = True
		Me.chkDev.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkDev.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDev.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDev.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDev.TabStop = True
		Me.chkDev.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDev.Visible = True
		Me.chkDev.Name = "chkDev"
		Me.chkWorkSupport.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.chkWorkSupport.Text = "Support"
		Me.chkWorkSupport.Size = New System.Drawing.Size(82, 22)
		Me.chkWorkSupport.Location = New System.Drawing.Point(10, 60)
		Me.chkWorkSupport.TabIndex = 17
		Me.chkWorkSupport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkWorkSupport.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkWorkSupport.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkWorkSupport.CausesValidation = True
		Me.chkWorkSupport.Enabled = True
		Me.chkWorkSupport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkWorkSupport.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkWorkSupport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkWorkSupport.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkWorkSupport.TabStop = True
		Me.chkWorkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkWorkSupport.Visible = True
		Me.chkWorkSupport.Name = "chkWorkSupport"
		Me.chkWorkSales.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.chkWorkSales.Text = "Sales"
		Me.chkWorkSales.Size = New System.Drawing.Size(82, 22)
		Me.chkWorkSales.Location = New System.Drawing.Point(10, 40)
		Me.chkWorkSales.TabIndex = 16
		Me.chkWorkSales.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkWorkSales.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkWorkSales.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkWorkSales.CausesValidation = True
		Me.chkWorkSales.Enabled = True
		Me.chkWorkSales.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkWorkSales.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkWorkSales.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkWorkSales.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkWorkSales.TabStop = True
		Me.chkWorkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkWorkSales.Visible = True
		Me.chkWorkSales.Name = "chkWorkSales"
		Me.chkManagement.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.chkManagement.Text = "Management"
		Me.chkManagement.Size = New System.Drawing.Size(112, 22)
		Me.chkManagement.Location = New System.Drawing.Point(10, 20)
		Me.chkManagement.TabIndex = 15
		Me.chkManagement.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkManagement.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkManagement.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkManagement.CausesValidation = True
		Me.chkManagement.Enabled = True
		Me.chkManagement.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkManagement.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkManagement.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkManagement.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkManagement.TabStop = True
		Me.chkManagement.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkManagement.Visible = True
		Me.chkManagement.Name = "chkManagement"
		Me.txtExt.AutoSize = False
		Me.txtExt.Size = New System.Drawing.Size(62, 24)
		Me.txtExt.Location = New System.Drawing.Point(490, 40)
		Me.txtExt.TabIndex = 8
		Me.txtExt.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtExt.AcceptsReturn = True
		Me.txtExt.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtExt.BackColor = System.Drawing.SystemColors.Window
		Me.txtExt.CausesValidation = True
		Me.txtExt.Enabled = True
		Me.txtExt.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtExt.HideSelection = True
		Me.txtExt.ReadOnly = False
		Me.txtExt.Maxlength = 0
		Me.txtExt.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtExt.MultiLine = False
		Me.txtExt.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtExt.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtExt.TabStop = True
		Me.txtExt.Visible = True
		Me.txtExt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtExt.Name = "txtExt"
		Me.fmeSecurity.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.fmeSecurity.Text = "Security Level"
		Me.fmeSecurity.Size = New System.Drawing.Size(112, 112)
		Me.fmeSecurity.Location = New System.Drawing.Point(420, 80)
		Me.fmeSecurity.TabIndex = 24
		Me.fmeSecurity.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fmeSecurity.Enabled = True
		Me.fmeSecurity.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fmeSecurity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fmeSecurity.Visible = True
		Me.fmeSecurity.Name = "fmeSecurity"
		Me.optLow.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optLow.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.optLow.Text = "Low"
		Me.optLow.Size = New System.Drawing.Size(72, 22)
		Me.optLow.Location = New System.Drawing.Point(20, 60)
		Me.optLow.TabIndex = 20
		Me.optLow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optLow.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optLow.CausesValidation = True
		Me.optLow.Enabled = True
		Me.optLow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optLow.Cursor = System.Windows.Forms.Cursors.Default
		Me.optLow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optLow.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optLow.TabStop = True
		Me.optLow.Checked = False
		Me.optLow.Visible = True
		Me.optLow.Name = "optLow"
		Me.optHigh.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optHigh.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.optHigh.Text = "High"
		Me.optHigh.Size = New System.Drawing.Size(72, 22)
		Me.optHigh.Location = New System.Drawing.Point(20, 30)
		Me.optHigh.TabIndex = 19
		Me.optHigh.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optHigh.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optHigh.CausesValidation = True
		Me.optHigh.Enabled = True
		Me.optHigh.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optHigh.Cursor = System.Windows.Forms.Cursors.Default
		Me.optHigh.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optHigh.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optHigh.TabStop = True
		Me.optHigh.Checked = False
		Me.optHigh.Visible = True
		Me.optHigh.Name = "optHigh"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(112, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(600, 160)
		Me.cmdCancel.TabIndex = 23
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.lblMail.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.lblMail.Text = "E-Mail  Address"
		Me.lblMail.Size = New System.Drawing.Size(132, 22)
		Me.lblMail.Location = New System.Drawing.Point(560, 20)
		Me.lblMail.TabIndex = 32
		Me.lblMail.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMail.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblMail.Enabled = True
		Me.lblMail.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblMail.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblMail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMail.UseMnemonic = True
		Me.lblMail.Visible = True
		Me.lblMail.AutoSize = False
		Me.lblMail.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblMail.Name = "lblMail"
		Me.lblPass.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.lblPass.Text = "Password"
		Me.lblPass.Size = New System.Drawing.Size(58, 17)
		Me.lblPass.Location = New System.Drawing.Point(600, 80)
		Me.lblPass.TabIndex = 31
		Me.lblPass.Visible = False
		Me.lblPass.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPass.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPass.Enabled = True
		Me.lblPass.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPass.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPass.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPass.UseMnemonic = True
		Me.lblPass.AutoSize = True
		Me.lblPass.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPass.Name = "lblPass"
		Me.lblFirst.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.lblFirst.Text = "First Name"
		Me.lblFirst.Size = New System.Drawing.Size(63, 17)
		Me.lblFirst.Location = New System.Drawing.Point(10, 20)
		Me.lblFirst.TabIndex = 30
		Me.lblFirst.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFirst.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFirst.Enabled = True
		Me.lblFirst.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFirst.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFirst.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFirst.UseMnemonic = True
		Me.lblFirst.Visible = True
		Me.lblFirst.AutoSize = True
		Me.lblFirst.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFirst.Name = "lblFirst"
		Me.lblLast.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.lblLast.Text = "Last Name"
		Me.lblLast.Size = New System.Drawing.Size(64, 17)
		Me.lblLast.Location = New System.Drawing.Point(310, 20)
		Me.lblLast.TabIndex = 29
		Me.lblLast.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLast.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLast.Enabled = True
		Me.lblLast.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLast.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLast.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLast.UseMnemonic = True
		Me.lblLast.Visible = True
		Me.lblLast.AutoSize = True
		Me.lblLast.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLast.Name = "lblLast"
		Me.lblMid.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.lblMid.Text = "Middle Name"
		Me.lblMid.Size = New System.Drawing.Size(78, 17)
		Me.lblMid.Location = New System.Drawing.Point(160, 20)
		Me.lblMid.TabIndex = 28
		Me.lblMid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMid.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblMid.Enabled = True
		Me.lblMid.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblMid.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblMid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMid.UseMnemonic = True
		Me.lblMid.Visible = True
		Me.lblMid.AutoSize = True
		Me.lblMid.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblMid.Name = "lblMid"
		Me.lblExt.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.lblExt.Text = "Ext."
		Me.lblExt.Size = New System.Drawing.Size(23, 17)
		Me.lblExt.Location = New System.Drawing.Point(490, 20)
		Me.lblExt.TabIndex = 27
		Me.lblExt.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblExt.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblExt.Enabled = True
		Me.lblExt.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblExt.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblExt.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblExt.UseMnemonic = True
		Me.lblExt.Visible = True
		Me.lblExt.AutoSize = True
		Me.lblExt.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblExt.Name = "lblExt"
		Me.ListView1.Size = New System.Drawing.Size(832, 292)
		Me.ListView1.Location = New System.Drawing.Point(10, 20)
		Me.ListView1.TabIndex = 1
		Me.ListView1.View = System.Windows.Forms.View.Details
		Me.ListView1.LabelEdit = False
		Me.ListView1.LabelWrap = True
		Me.ListView1.HideSelection = False
		Me.ListView1.FullRowSelect = True
		Me.ListView1.GridLines = True
		Me.ListView1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ListView1.BackColor = System.Drawing.SystemColors.Window
		Me.ListView1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ListView1.Name = "ListView1"
		Me._ListView1_ColumnHeader_1.Text = "ID"
		Me._ListView1_ColumnHeader_1.Width = 177
		Me._ListView1_ColumnHeader_2.Text = "Name"
		Me._ListView1_ColumnHeader_2.Width = 212
		Me._ListView1_ColumnHeader_3.Text = "Ext"
		Me._ListView1_ColumnHeader_3.Width = 177
		Me._ListView1_ColumnHeader_4.Text = "Security Level"
		Me._ListView1_ColumnHeader_4.Width = 212
		Me._ListView1_ColumnHeader_5.Text = "Groups"
		Me._ListView1_ColumnHeader_5.Width = 212
		Me._ListView1_ColumnHeader_6.Text = "WorkGroups"
		Me._ListView1_ColumnHeader_6.Width = 212
		Me._ListView1_ColumnHeader_7.Text = "E-Mail"
		Me._ListView1_ColumnHeader_7.Width = 212
		Me.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(txtIcon)
		Me.Frame1.Controls.Add(cmdEdit)
		Me.Frame1.Controls.Add(cmdAdd)
		Me.Frame1.Controls.Add(cmdDelete)
		Me.Frame1.Controls.Add(fmeAddEdit)
		Me.Frame1.Controls.Add(ListView1)
		Me.fmeAddEdit.Controls.Add(txtMail)
		Me.fmeAddEdit.Controls.Add(txtPassword)
		Me.fmeAddEdit.Controls.Add(txtFirst)
		Me.fmeAddEdit.Controls.Add(txtMid)
		Me.fmeAddEdit.Controls.Add(txtLast)
		Me.fmeAddEdit.Controls.Add(cmdSave)
		Me.fmeAddEdit.Controls.Add(fmeGroups)
		Me.fmeAddEdit.Controls.Add(fmeWorkGroups)
		Me.fmeAddEdit.Controls.Add(txtExt)
		Me.fmeAddEdit.Controls.Add(fmeSecurity)
		Me.fmeAddEdit.Controls.Add(cmdCancel)
		Me.fmeAddEdit.Controls.Add(lblMail)
		Me.fmeAddEdit.Controls.Add(lblPass)
		Me.fmeAddEdit.Controls.Add(lblFirst)
		Me.fmeAddEdit.Controls.Add(lblLast)
		Me.fmeAddEdit.Controls.Add(lblMid)
		Me.fmeAddEdit.Controls.Add(lblExt)
		Me.fmeGroups.Controls.Add(chkAuthorizations)
		Me.fmeGroups.Controls.Add(chkSales)
		Me.fmeGroups.Controls.Add(chkSupport)
		Me.fmeGroups.Controls.Add(chkOperator)
		Me.fmeWorkGroups.Controls.Add(chkDev)
		Me.fmeWorkGroups.Controls.Add(chkWorkSupport)
		Me.fmeWorkGroups.Controls.Add(chkWorkSales)
		Me.fmeWorkGroups.Controls.Add(chkManagement)
		Me.fmeSecurity.Controls.Add(optLow)
		Me.fmeSecurity.Controls.Add(optHigh)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_1)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_2)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_3)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_4)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_5)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_6)
		Me.ListView1.Columns.Add(_ListView1_ColumnHeader_7)
		Me.Frame1.ResumeLayout(False)
		Me.fmeAddEdit.ResumeLayout(False)
		Me.fmeGroups.ResumeLayout(False)
		Me.fmeWorkGroups.ResumeLayout(False)
		Me.fmeSecurity.ResumeLayout(False)
		Me.ListView1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class