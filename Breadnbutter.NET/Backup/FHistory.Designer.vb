<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FHistory
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
	Public WithEvents chkLimit As System.Windows.Forms.CheckBox
	Public WithEvents txtLimit As System.Windows.Forms.TextBox
	Public WithEvents Frame10 As System.Windows.Forms.GroupBox
	Public WithEvents cboOrder As System.Windows.Forms.ComboBox
	Public WithEvents Frame9 As System.Windows.Forms.GroupBox
	Public WithEvents cmdCopy As System.Windows.Forms.Button
	Public WithEvents cboProduct As System.Windows.Forms.ComboBox
	Public WithEvents Frame8 As System.Windows.Forms.GroupBox
	Public WithEvents cmdShowResults As System.Windows.Forms.Button
	Public WithEvents cmdPreviewReport As System.Windows.Forms.Button
	Public WithEvents cboCategory As System.Windows.Forms.ComboBox
	Public WithEvents Frame7 As System.Windows.Forms.GroupBox
	Public WithEvents cmdDateSet1 As System.Windows.Forms.Button
	Public WithEvents cmdDateSet2 As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents lblDate1 As System.Windows.Forms.Label
	Public WithEvents lblDate2 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents lstReport As System.Windows.Forms.ComboBox
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents cboUser As System.Windows.Forms.ComboBox
	Public WithEvents User As System.Windows.Forms.GroupBox
	Public WithEvents lstStatus As System.Windows.Forms.ComboBox
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	Public WithEvents txtBranch As System.Windows.Forms.TextBox
	Public WithEvents cboType As System.Windows.Forms.ComboBox
	Public WithEvents txtFirstName As System.Windows.Forms.TextBox
	Public WithEvents txtLastName As System.Windows.Forms.TextBox
	Public WithEvents txtCompany As System.Windows.Forms.TextBox
	Public WithEvents cboState As System.Windows.Forms.ComboBox
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Frame5 As System.Windows.Forms.GroupBox
	Public WithEvents txtHistory As System.Windows.Forms.TextBox
	Public WithEvents Frame6 As System.Windows.Forms.GroupBox
	Public WithEvents cmdPrint As System.Windows.Forms.Button
	Public WithEvents grdHistory As AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents LblCount As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.Panel
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FHistory))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame1 = New System.Windows.Forms.Panel
		Me.Frame10 = New System.Windows.Forms.GroupBox
		Me.chkLimit = New System.Windows.Forms.CheckBox
		Me.txtLimit = New System.Windows.Forms.TextBox
		Me.Frame9 = New System.Windows.Forms.GroupBox
		Me.cboOrder = New System.Windows.Forms.ComboBox
		Me.cmdCopy = New System.Windows.Forms.Button
		Me.Frame8 = New System.Windows.Forms.GroupBox
		Me.cboProduct = New System.Windows.Forms.ComboBox
		Me.cmdShowResults = New System.Windows.Forms.Button
		Me.cmdPreviewReport = New System.Windows.Forms.Button
		Me.Frame7 = New System.Windows.Forms.GroupBox
		Me.cboCategory = New System.Windows.Forms.ComboBox
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.cmdDateSet1 = New System.Windows.Forms.Button
		Me.cmdDateSet2 = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.lblDate1 = New System.Windows.Forms.Label
		Me.lblDate2 = New System.Windows.Forms.Label
		Me.Frame3 = New System.Windows.Forms.GroupBox
		Me.lstReport = New System.Windows.Forms.ComboBox
		Me.User = New System.Windows.Forms.GroupBox
		Me.cboUser = New System.Windows.Forms.ComboBox
		Me.Frame4 = New System.Windows.Forms.GroupBox
		Me.lstStatus = New System.Windows.Forms.ComboBox
		Me.Frame5 = New System.Windows.Forms.GroupBox
		Me.txtBranch = New System.Windows.Forms.TextBox
		Me.cboType = New System.Windows.Forms.ComboBox
		Me.txtFirstName = New System.Windows.Forms.TextBox
		Me.txtLastName = New System.Windows.Forms.TextBox
		Me.txtCompany = New System.Windows.Forms.TextBox
		Me.cboState = New System.Windows.Forms.ComboBox
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Frame6 = New System.Windows.Forms.GroupBox
		Me.txtHistory = New System.Windows.Forms.TextBox
		Me.cmdPrint = New System.Windows.Forms.Button
		Me.grdHistory = New AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
		Me.Label7 = New System.Windows.Forms.Label
		Me.LblCount = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.Frame10.SuspendLayout()
		Me.Frame9.SuspendLayout()
		Me.Frame8.SuspendLayout()
		Me.Frame7.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.Frame3.SuspendLayout()
		Me.User.SuspendLayout()
		Me.Frame4.SuspendLayout()
		Me.Frame5.SuspendLayout()
		Me.Frame6.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdHistory, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.AppWorkspace
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "History Report"
		Me.ClientSize = New System.Drawing.Size(975, 628)
		Me.Location = New System.Drawing.Point(160, 147)
		Me.Icon = CType(resources.GetObject("FHistory.Icon"), System.Drawing.Icon)
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
		Me.Name = "FHistory"
		Me.Frame1.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Frame1.Text = "Frame1"
		Me.Frame1.Size = New System.Drawing.Size(964, 612)
		Me.Frame1.Location = New System.Drawing.Point(0, 0)
		Me.Frame1.TabIndex = 19
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.Frame10.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame10.Text = "Record Limit"
		Me.Frame10.Size = New System.Drawing.Size(172, 62)
		Me.Frame10.Location = New System.Drawing.Point(770, 540)
		Me.Frame10.TabIndex = 42
		Me.Frame10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame10.Enabled = True
		Me.Frame10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame10.Visible = True
		Me.Frame10.Name = "Frame10"
		Me.chkLimit.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.chkLimit.Text = "Check1"
		Me.chkLimit.Size = New System.Drawing.Size(22, 22)
		Me.chkLimit.Location = New System.Drawing.Point(10, 25)
		Me.chkLimit.TabIndex = 44
		Me.chkLimit.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkLimit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLimit.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLimit.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLimit.CausesValidation = True
		Me.chkLimit.Enabled = True
		Me.chkLimit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLimit.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLimit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLimit.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLimit.TabStop = True
		Me.chkLimit.Visible = True
		Me.chkLimit.Name = "chkLimit"
		Me.txtLimit.AutoSize = False
		Me.txtLimit.Size = New System.Drawing.Size(122, 24)
		Me.txtLimit.Location = New System.Drawing.Point(40, 25)
		Me.txtLimit.Maxlength = 6
		Me.txtLimit.TabIndex = 43
		Me.txtLimit.Text = "1000"
		Me.txtLimit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLimit.AcceptsReturn = True
		Me.txtLimit.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLimit.BackColor = System.Drawing.SystemColors.Window
		Me.txtLimit.CausesValidation = True
		Me.txtLimit.Enabled = True
		Me.txtLimit.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLimit.HideSelection = True
		Me.txtLimit.ReadOnly = False
		Me.txtLimit.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLimit.MultiLine = False
		Me.txtLimit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLimit.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLimit.TabStop = True
		Me.txtLimit.Visible = True
		Me.txtLimit.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLimit.Name = "txtLimit"
		Me.Frame9.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame9.Text = "Order By"
		Me.Frame9.Size = New System.Drawing.Size(172, 62)
		Me.Frame9.Location = New System.Drawing.Point(580, 540)
		Me.Frame9.TabIndex = 41
		Me.Frame9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame9.Enabled = True
		Me.Frame9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame9.Visible = True
		Me.Frame9.Name = "Frame9"
		Me.cboOrder.Size = New System.Drawing.Size(152, 27)
		Me.cboOrder.Location = New System.Drawing.Point(10, 24)
		Me.cboOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboOrder.TabIndex = 18
		Me.cboOrder.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboOrder.BackColor = System.Drawing.SystemColors.Window
		Me.cboOrder.CausesValidation = True
		Me.cboOrder.Enabled = True
		Me.cboOrder.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboOrder.IntegralHeight = True
		Me.cboOrder.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboOrder.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboOrder.Sorted = False
		Me.cboOrder.TabStop = True
		Me.cboOrder.Visible = True
		Me.cboOrder.Name = "cboOrder"
		Me.cmdCopy.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCopy.Text = "Copy Results to Clipboad"
		Me.cmdCopy.Size = New System.Drawing.Size(167, 27)
		Me.cmdCopy.Location = New System.Drawing.Point(120, 578)
		Me.cmdCopy.TabIndex = 16
		Me.cmdCopy.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCopy.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCopy.CausesValidation = True
		Me.cmdCopy.Enabled = True
		Me.cmdCopy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCopy.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCopy.TabStop = True
		Me.cmdCopy.Name = "cmdCopy"
		Me.Frame8.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame8.Text = "Product"
		Me.Frame8.Size = New System.Drawing.Size(182, 52)
		Me.Frame8.Location = New System.Drawing.Point(200, 428)
		Me.Frame8.TabIndex = 39
		Me.Frame8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame8.Enabled = True
		Me.Frame8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame8.Visible = True
		Me.Frame8.Name = "Frame8"
		Me.cboProduct.Size = New System.Drawing.Size(162, 27)
		Me.cboProduct.Location = New System.Drawing.Point(9, 19)
		Me.cboProduct.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboProduct.TabIndex = 4
		Me.cboProduct.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboProduct.BackColor = System.Drawing.SystemColors.Window
		Me.cboProduct.CausesValidation = True
		Me.cboProduct.Enabled = True
		Me.cboProduct.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboProduct.IntegralHeight = True
		Me.cboProduct.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboProduct.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboProduct.Sorted = False
		Me.cboProduct.TabStop = True
		Me.cboProduct.Visible = True
		Me.cboProduct.Name = "cboProduct"
		Me.cmdShowResults.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdShowResults.Text = "Show Results"
		Me.cmdShowResults.Size = New System.Drawing.Size(167, 27)
		Me.cmdShowResults.Location = New System.Drawing.Point(120, 545)
		Me.cmdShowResults.TabIndex = 14
		Me.cmdShowResults.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdShowResults.BackColor = System.Drawing.SystemColors.Control
		Me.cmdShowResults.CausesValidation = True
		Me.cmdShowResults.Enabled = True
		Me.cmdShowResults.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdShowResults.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdShowResults.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdShowResults.TabStop = True
		Me.cmdShowResults.Name = "cmdShowResults"
		Me.cmdPreviewReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPreviewReport.Text = "Preview Report"
		Me.cmdPreviewReport.Size = New System.Drawing.Size(167, 27)
		Me.cmdPreviewReport.Location = New System.Drawing.Point(298, 545)
		Me.cmdPreviewReport.TabIndex = 15
		Me.cmdPreviewReport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPreviewReport.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPreviewReport.CausesValidation = True
		Me.cmdPreviewReport.Enabled = True
		Me.cmdPreviewReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPreviewReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPreviewReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPreviewReport.TabStop = True
		Me.cmdPreviewReport.Name = "cmdPreviewReport"
		Me.Frame7.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame7.Text = "Category"
		Me.Frame7.Size = New System.Drawing.Size(177, 57)
		Me.Frame7.Location = New System.Drawing.Point(10, 350)
		Me.Frame7.TabIndex = 34
		Me.Frame7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame7.Enabled = True
		Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame7.Visible = True
		Me.Frame7.Name = "Frame7"
		Me.cboCategory.Size = New System.Drawing.Size(162, 27)
		Me.cboCategory.Location = New System.Drawing.Point(8, 20)
		Me.cboCategory.TabIndex = 0
		Me.cboCategory.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboCategory.BackColor = System.Drawing.SystemColors.Window
		Me.cboCategory.CausesValidation = True
		Me.cboCategory.Enabled = True
		Me.cboCategory.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboCategory.IntegralHeight = True
		Me.cboCategory.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboCategory.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboCategory.Sorted = False
		Me.cboCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboCategory.TabStop = True
		Me.cboCategory.Visible = True
		Me.cboCategory.Name = "cboCategory"
		Me.Frame2.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame2.Text = "Date"
		Me.Frame2.Size = New System.Drawing.Size(180, 72)
		Me.Frame2.Location = New System.Drawing.Point(203, 350)
		Me.Frame2.TabIndex = 29
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Name = "Frame2"
		Me.cmdDateSet1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDateSet1.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDateSet1.Text = "Set"
		Me.cmdDateSet1.Size = New System.Drawing.Size(42, 19)
		Me.cmdDateSet1.Location = New System.Drawing.Point(130, 18)
		Me.cmdDateSet1.TabIndex = 2
		Me.cmdDateSet1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDateSet1.CausesValidation = True
		Me.cmdDateSet1.Enabled = True
		Me.cmdDateSet1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDateSet1.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDateSet1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDateSet1.TabStop = True
		Me.cmdDateSet1.Name = "cmdDateSet1"
		Me.cmdDateSet2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDateSet2.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDateSet2.Text = "Set"
		Me.cmdDateSet2.Size = New System.Drawing.Size(42, 19)
		Me.cmdDateSet2.Location = New System.Drawing.Point(130, 45)
		Me.cmdDateSet2.TabIndex = 3
		Me.cmdDateSet2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDateSet2.CausesValidation = True
		Me.cmdDateSet2.Enabled = True
		Me.cmdDateSet2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDateSet2.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDateSet2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDateSet2.TabStop = True
		Me.cmdDateSet2.Name = "cmdDateSet2"
		Me.Label1.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label1.Text = "From:"
		Me.Label1.Size = New System.Drawing.Size(42, 22)
		Me.Label1.Location = New System.Drawing.Point(7, 17)
		Me.Label1.TabIndex = 33
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Label2.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label2.Text = "To:"
		Me.Label2.Size = New System.Drawing.Size(27, 22)
		Me.Label2.Location = New System.Drawing.Point(18, 44)
		Me.Label2.TabIndex = 32
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.lblDate1.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.lblDate1.Text = "Label7"
		Me.lblDate1.ForeColor = System.Drawing.Color.White
		Me.lblDate1.Size = New System.Drawing.Size(82, 27)
		Me.lblDate1.Location = New System.Drawing.Point(47, 17)
		Me.lblDate1.TabIndex = 31
		Me.lblDate1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDate1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDate1.Enabled = True
		Me.lblDate1.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDate1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDate1.UseMnemonic = True
		Me.lblDate1.Visible = True
		Me.lblDate1.AutoSize = False
		Me.lblDate1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDate1.Name = "lblDate1"
		Me.lblDate2.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.lblDate2.Text = "Label8"
		Me.lblDate2.ForeColor = System.Drawing.Color.White
		Me.lblDate2.Size = New System.Drawing.Size(79, 27)
		Me.lblDate2.Location = New System.Drawing.Point(45, 44)
		Me.lblDate2.TabIndex = 30
		Me.lblDate2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDate2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDate2.Enabled = True
		Me.lblDate2.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDate2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDate2.UseMnemonic = True
		Me.lblDate2.Visible = True
		Me.lblDate2.AutoSize = False
		Me.lblDate2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDate2.Name = "lblDate2"
		Me.Frame3.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame3.Text = "Report Type"
		Me.Frame3.Size = New System.Drawing.Size(159, 54)
		Me.Frame3.Location = New System.Drawing.Point(403, 350)
		Me.Frame3.TabIndex = 28
		Me.Frame3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame3.Enabled = True
		Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame3.Visible = True
		Me.Frame3.Name = "Frame3"
		Me.lstReport.Size = New System.Drawing.Size(139, 27)
		Me.lstReport.Location = New System.Drawing.Point(10, 20)
		Me.lstReport.TabIndex = 5
		Me.lstReport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstReport.BackColor = System.Drawing.SystemColors.Window
		Me.lstReport.CausesValidation = True
		Me.lstReport.Enabled = True
		Me.lstReport.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstReport.IntegralHeight = True
		Me.lstReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstReport.Sorted = False
		Me.lstReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.lstReport.TabStop = True
		Me.lstReport.Visible = True
		Me.lstReport.Name = "lstReport"
		Me.User.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.User.Text = "User"
		Me.User.Size = New System.Drawing.Size(177, 52)
		Me.User.Location = New System.Drawing.Point(10, 415)
		Me.User.TabIndex = 27
		Me.User.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.User.Enabled = True
		Me.User.ForeColor = System.Drawing.SystemColors.ControlText
		Me.User.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.User.Visible = True
		Me.User.Name = "User"
		Me.cboUser.Size = New System.Drawing.Size(162, 27)
		Me.cboUser.Location = New System.Drawing.Point(8, 18)
		Me.cboUser.TabIndex = 1
		Me.cboUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboUser.BackColor = System.Drawing.SystemColors.Window
		Me.cboUser.CausesValidation = True
		Me.cboUser.Enabled = True
		Me.cboUser.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboUser.IntegralHeight = True
		Me.cboUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboUser.Sorted = False
		Me.cboUser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboUser.TabStop = True
		Me.cboUser.Visible = True
		Me.cboUser.Name = "cboUser"
		Me.Frame4.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame4.Text = "Contact Status"
		Me.Frame4.Size = New System.Drawing.Size(159, 52)
		Me.Frame4.Location = New System.Drawing.Point(403, 413)
		Me.Frame4.TabIndex = 26
		Me.Frame4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame4.Enabled = True
		Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame4.Visible = True
		Me.Frame4.Name = "Frame4"
		Me.lstStatus.Size = New System.Drawing.Size(139, 27)
		Me.lstStatus.Location = New System.Drawing.Point(10, 18)
		Me.lstStatus.TabIndex = 6
		Me.lstStatus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstStatus.BackColor = System.Drawing.SystemColors.Window
		Me.lstStatus.CausesValidation = True
		Me.lstStatus.Enabled = True
		Me.lstStatus.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstStatus.IntegralHeight = True
		Me.lstStatus.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstStatus.Sorted = False
		Me.lstStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.lstStatus.TabStop = True
		Me.lstStatus.Visible = True
		Me.lstStatus.Name = "lstStatus"
		Me.Frame5.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame5.Text = "Other Criteria"
		Me.Frame5.Size = New System.Drawing.Size(364, 179)
		Me.Frame5.Location = New System.Drawing.Point(580, 350)
		Me.Frame5.TabIndex = 21
		Me.Frame5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame5.Enabled = True
		Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame5.Visible = True
		Me.Frame5.Name = "Frame5"
		Me.txtBranch.AutoSize = False
		Me.txtBranch.Size = New System.Drawing.Size(272, 27)
		Me.txtBranch.Location = New System.Drawing.Point(83, 110)
		Me.txtBranch.TabIndex = 11
		Me.txtBranch.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBranch.AcceptsReturn = True
		Me.txtBranch.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBranch.BackColor = System.Drawing.SystemColors.Window
		Me.txtBranch.CausesValidation = True
		Me.txtBranch.Enabled = True
		Me.txtBranch.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtBranch.HideSelection = True
		Me.txtBranch.ReadOnly = False
		Me.txtBranch.Maxlength = 0
		Me.txtBranch.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBranch.MultiLine = False
		Me.txtBranch.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBranch.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBranch.TabStop = True
		Me.txtBranch.Visible = True
		Me.txtBranch.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBranch.Name = "txtBranch"
		Me.cboType.Size = New System.Drawing.Size(142, 27)
		Me.cboType.Location = New System.Drawing.Point(213, 140)
		Me.cboType.TabIndex = 13
		Me.cboType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboType.BackColor = System.Drawing.SystemColors.Window
		Me.cboType.CausesValidation = True
		Me.cboType.Enabled = True
		Me.cboType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboType.IntegralHeight = True
		Me.cboType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboType.Sorted = False
		Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboType.TabStop = True
		Me.cboType.Visible = True
		Me.cboType.Name = "cboType"
		Me.txtFirstName.AutoSize = False
		Me.txtFirstName.Size = New System.Drawing.Size(272, 27)
		Me.txtFirstName.Location = New System.Drawing.Point(83, 20)
		Me.txtFirstName.TabIndex = 8
		Me.txtFirstName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFirstName.AcceptsReturn = True
		Me.txtFirstName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFirstName.BackColor = System.Drawing.SystemColors.Window
		Me.txtFirstName.CausesValidation = True
		Me.txtFirstName.Enabled = True
		Me.txtFirstName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFirstName.HideSelection = True
		Me.txtFirstName.ReadOnly = False
		Me.txtFirstName.Maxlength = 0
		Me.txtFirstName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFirstName.MultiLine = False
		Me.txtFirstName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFirstName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFirstName.TabStop = True
		Me.txtFirstName.Visible = True
		Me.txtFirstName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFirstName.Name = "txtFirstName"
		Me.txtLastName.AutoSize = False
		Me.txtLastName.Size = New System.Drawing.Size(272, 27)
		Me.txtLastName.Location = New System.Drawing.Point(83, 50)
		Me.txtLastName.TabIndex = 9
		Me.txtLastName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLastName.AcceptsReturn = True
		Me.txtLastName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLastName.BackColor = System.Drawing.SystemColors.Window
		Me.txtLastName.CausesValidation = True
		Me.txtLastName.Enabled = True
		Me.txtLastName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLastName.HideSelection = True
		Me.txtLastName.ReadOnly = False
		Me.txtLastName.Maxlength = 0
		Me.txtLastName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLastName.MultiLine = False
		Me.txtLastName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLastName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLastName.TabStop = True
		Me.txtLastName.Visible = True
		Me.txtLastName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLastName.Name = "txtLastName"
		Me.txtCompany.AutoSize = False
		Me.txtCompany.Size = New System.Drawing.Size(272, 27)
		Me.txtCompany.Location = New System.Drawing.Point(83, 80)
		Me.txtCompany.TabIndex = 10
		Me.txtCompany.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCompany.AcceptsReturn = True
		Me.txtCompany.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCompany.BackColor = System.Drawing.SystemColors.Window
		Me.txtCompany.CausesValidation = True
		Me.txtCompany.Enabled = True
		Me.txtCompany.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCompany.HideSelection = True
		Me.txtCompany.ReadOnly = False
		Me.txtCompany.Maxlength = 0
		Me.txtCompany.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCompany.MultiLine = False
		Me.txtCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCompany.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCompany.TabStop = True
		Me.txtCompany.Visible = True
		Me.txtCompany.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCompany.Name = "txtCompany"
		Me.cboState.Size = New System.Drawing.Size(59, 27)
		Me.cboState.Location = New System.Drawing.Point(83, 140)
		Me.cboState.TabIndex = 12
		Me.cboState.Text = "All"
		Me.cboState.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboState.BackColor = System.Drawing.SystemColors.Window
		Me.cboState.CausesValidation = True
		Me.cboState.Enabled = True
		Me.cboState.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboState.IntegralHeight = True
		Me.cboState.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboState.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboState.Sorted = False
		Me.cboState.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboState.TabStop = True
		Me.cboState.Visible = True
		Me.cboState.Name = "cboState"
		Me.Label9.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label9.Text = "Branch:"
		Me.Label9.Size = New System.Drawing.Size(47, 17)
		Me.Label9.Location = New System.Drawing.Point(30, 113)
		Me.Label9.TabIndex = 40
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = True
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label8.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label8.Text = "Type:"
		Me.Label8.Size = New System.Drawing.Size(39, 22)
		Me.Label8.Location = New System.Drawing.Point(170, 143)
		Me.Label8.TabIndex = 38
		Me.Label8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label8.Enabled = True
		Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label8.UseMnemonic = True
		Me.Label8.Visible = True
		Me.Label8.AutoSize = False
		Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label8.Name = "Label8"
		Me.Label3.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label3.Text = "First Name:"
		Me.Label3.Size = New System.Drawing.Size(69, 22)
		Me.Label3.Location = New System.Drawing.Point(10, 25)
		Me.Label3.TabIndex = 25
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label4.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label4.Text = "Last Name:"
		Me.Label4.Size = New System.Drawing.Size(69, 22)
		Me.Label4.Location = New System.Drawing.Point(10, 55)
		Me.Label4.TabIndex = 24
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label5.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label5.Text = "Company:"
		Me.Label5.Size = New System.Drawing.Size(64, 22)
		Me.Label5.Location = New System.Drawing.Point(18, 83)
		Me.Label5.TabIndex = 23
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label6.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label6.Text = "State:"
		Me.Label6.Size = New System.Drawing.Size(39, 22)
		Me.Label6.Location = New System.Drawing.Point(40, 143)
		Me.Label6.TabIndex = 22
		Me.Label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.Frame6.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame6.Text = "Search History Notes"
		Me.Frame6.Size = New System.Drawing.Size(559, 54)
		Me.Frame6.Location = New System.Drawing.Point(10, 475)
		Me.Frame6.TabIndex = 20
		Me.Frame6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame6.Enabled = True
		Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame6.Visible = True
		Me.Frame6.Name = "Frame6"
		Me.txtHistory.AutoSize = False
		Me.txtHistory.Size = New System.Drawing.Size(542, 27)
		Me.txtHistory.Location = New System.Drawing.Point(10, 20)
		Me.txtHistory.TabIndex = 7
		Me.txtHistory.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtHistory.AcceptsReturn = True
		Me.txtHistory.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtHistory.BackColor = System.Drawing.SystemColors.Window
		Me.txtHistory.CausesValidation = True
		Me.txtHistory.Enabled = True
		Me.txtHistory.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtHistory.HideSelection = True
		Me.txtHistory.ReadOnly = False
		Me.txtHistory.Maxlength = 0
		Me.txtHistory.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHistory.MultiLine = False
		Me.txtHistory.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHistory.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtHistory.TabStop = True
		Me.txtHistory.Visible = True
		Me.txtHistory.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtHistory.Name = "txtHistory"
		Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPrint.Text = "Print Report"
		Me.cmdPrint.Size = New System.Drawing.Size(167, 27)
		Me.cmdPrint.Location = New System.Drawing.Point(298, 578)
		Me.cmdPrint.TabIndex = 17
		Me.cmdPrint.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPrint.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPrint.CausesValidation = True
		Me.cmdPrint.Enabled = True
		Me.cmdPrint.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPrint.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPrint.TabStop = True
		Me.cmdPrint.Name = "cmdPrint"
		grdHistory.OcxState = CType(resources.GetObject("grdHistory.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdHistory.Size = New System.Drawing.Size(964, 324)
		Me.grdHistory.Location = New System.Drawing.Point(0, 23)
		Me.grdHistory.TabIndex = 35
		Me.grdHistory.Name = "grdHistory"
		Me.Label7.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label7.Text = "Notes Found:"
		Me.Label7.Size = New System.Drawing.Size(82, 22)
		Me.Label7.Location = New System.Drawing.Point(10, 0)
		Me.Label7.TabIndex = 37
		Me.Label7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.LblCount.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.LblCount.Text = "0"
		Me.LblCount.Size = New System.Drawing.Size(84, 22)
		Me.LblCount.Location = New System.Drawing.Point(100, 0)
		Me.LblCount.TabIndex = 36
		Me.LblCount.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblCount.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.LblCount.Enabled = True
		Me.LblCount.ForeColor = System.Drawing.SystemColors.ControlText
		Me.LblCount.Cursor = System.Windows.Forms.Cursors.Default
		Me.LblCount.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LblCount.UseMnemonic = True
		Me.LblCount.Visible = True
		Me.LblCount.AutoSize = False
		Me.LblCount.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LblCount.Name = "LblCount"
		CType(Me.grdHistory, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(Frame10)
		Me.Frame1.Controls.Add(Frame9)
		Me.Frame1.Controls.Add(cmdCopy)
		Me.Frame1.Controls.Add(Frame8)
		Me.Frame1.Controls.Add(cmdShowResults)
		Me.Frame1.Controls.Add(cmdPreviewReport)
		Me.Frame1.Controls.Add(Frame7)
		Me.Frame1.Controls.Add(Frame2)
		Me.Frame1.Controls.Add(Frame3)
		Me.Frame1.Controls.Add(User)
		Me.Frame1.Controls.Add(Frame4)
		Me.Frame1.Controls.Add(Frame5)
		Me.Frame1.Controls.Add(Frame6)
		Me.Frame1.Controls.Add(cmdPrint)
		Me.Frame1.Controls.Add(grdHistory)
		Me.Frame1.Controls.Add(Label7)
		Me.Frame1.Controls.Add(LblCount)
		Me.Frame10.Controls.Add(chkLimit)
		Me.Frame10.Controls.Add(txtLimit)
		Me.Frame9.Controls.Add(cboOrder)
		Me.Frame8.Controls.Add(cboProduct)
		Me.Frame7.Controls.Add(cboCategory)
		Me.Frame2.Controls.Add(cmdDateSet1)
		Me.Frame2.Controls.Add(cmdDateSet2)
		Me.Frame2.Controls.Add(Label1)
		Me.Frame2.Controls.Add(Label2)
		Me.Frame2.Controls.Add(lblDate1)
		Me.Frame2.Controls.Add(lblDate2)
		Me.Frame3.Controls.Add(lstReport)
		Me.User.Controls.Add(cboUser)
		Me.Frame4.Controls.Add(lstStatus)
		Me.Frame5.Controls.Add(txtBranch)
		Me.Frame5.Controls.Add(cboType)
		Me.Frame5.Controls.Add(txtFirstName)
		Me.Frame5.Controls.Add(txtLastName)
		Me.Frame5.Controls.Add(txtCompany)
		Me.Frame5.Controls.Add(cboState)
		Me.Frame5.Controls.Add(Label9)
		Me.Frame5.Controls.Add(Label8)
		Me.Frame5.Controls.Add(Label3)
		Me.Frame5.Controls.Add(Label4)
		Me.Frame5.Controls.Add(Label5)
		Me.Frame5.Controls.Add(Label6)
		Me.Frame6.Controls.Add(txtHistory)
		Me.Frame1.ResumeLayout(False)
		Me.Frame10.ResumeLayout(False)
		Me.Frame9.ResumeLayout(False)
		Me.Frame8.ResumeLayout(False)
		Me.Frame7.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.Frame3.ResumeLayout(False)
		Me.User.ResumeLayout(False)
		Me.Frame4.ResumeLayout(False)
		Me.Frame5.ResumeLayout(False)
		Me.Frame6.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class