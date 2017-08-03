<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FReport
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
	Public WithEvents cmdCopyToClipBoard As System.Windows.Forms.Button
	Public WithEvents grdHistory As AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
	Public WithEvents ShowResults As System.Windows.Forms.Button
	Public WithEvents PreviewReport As System.Windows.Forms.Button
	Public WithEvents txtDays As System.Windows.Forms.TextBox
	Public WithEvents _optChoice_4 As System.Windows.Forms.RadioButton
	Public WithEvents chkTtls As System.Windows.Forms.CheckBox
	Public WithEvents _chkFilter_1 As System.Windows.Forms.CheckBox
	Public WithEvents _chkFilter_0 As System.Windows.Forms.CheckBox
	Public WithEvents lstGroups As System.Windows.Forms.ListBox
	Public WithEvents _optChoice_5 As System.Windows.Forms.RadioButton
	Public WithEvents cboProduct As System.Windows.Forms.ComboBox
	Public WithEvents _optChoice_3 As System.Windows.Forms.RadioButton
	Public WithEvents _optChoice_2 As System.Windows.Forms.RadioButton
	Public WithEvents _optChoice_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optChoice_0 As System.Windows.Forms.RadioButton
	Public WithEvents TextDaysMax As System.Windows.Forms.TextBox
	Public WithEvents TextDaysMin As System.Windows.Forms.TextBox
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label22 As System.Windows.Forms.Label
	Public WithEvents Label23 As System.Windows.Forms.Label
	Public WithEvents Shape3 As System.Windows.Forms.Label
	Public WithEvents Days As System.Windows.Forms.Panel
	Public WithEvents textNotes As System.Windows.Forms.ComboBox
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents Shape5 As System.Windows.Forms.Label
	Public WithEvents Text_Renamed As System.Windows.Forms.Panel
	Public WithEvents _Picture1_1 As System.Windows.Forms.Panel
	Public WithEvents chkAlpha As System.Windows.Forms.CheckBox
	Public WithEvents _Picture1_2 As System.Windows.Forms.Panel
	Public WithEvents cboAction As System.Windows.Forms.ComboBox
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents _Picture1_4 As System.Windows.Forms.Panel
	Public WithEvents _Picture1_5 As System.Windows.Forms.Panel
	Public WithEvents Frame4 As System.Windows.Forms.Panel
	Public WithEvents TextSource As System.Windows.Forms.TextBox
	Public WithEvents ComboStatus As System.Windows.Forms.ComboBox
	Public WithEvents TextZip As System.Windows.Forms.TextBox
	Public WithEvents TextCity As System.Windows.Forms.TextBox
	Public WithEvents TextCompany As System.Windows.Forms.TextBox
	Public WithEvents TextLastName As System.Windows.Forms.TextBox
	Public WithEvents TextFirstName As System.Windows.Forms.TextBox
	Public WithEvents StateCombo As System.Windows.Forms.ComboBox
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents frmCriteria As System.Windows.Forms.Panel
	Public WithEvents PrintButton As System.Windows.Forms.Button
	Public WithEvents ListView1 As System.Windows.Forms.ListView
	Public WithEvents LabelResults As System.Windows.Forms.Label
	Public WithEvents frame1 As System.Windows.Forms.Panel
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Picture1 As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
	Public WithEvents chkFilter As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents optChoice As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FReport))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.frame1 = New System.Windows.Forms.Panel
		Me.cmdCopyToClipBoard = New System.Windows.Forms.Button
		Me.grdHistory = New AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
		Me.ShowResults = New System.Windows.Forms.Button
		Me.PreviewReport = New System.Windows.Forms.Button
		Me.Frame4 = New System.Windows.Forms.Panel
		Me.txtDays = New System.Windows.Forms.TextBox
		Me._optChoice_4 = New System.Windows.Forms.RadioButton
		Me.chkTtls = New System.Windows.Forms.CheckBox
		Me._chkFilter_1 = New System.Windows.Forms.CheckBox
		Me._chkFilter_0 = New System.Windows.Forms.CheckBox
		Me.lstGroups = New System.Windows.Forms.ListBox
		Me._optChoice_5 = New System.Windows.Forms.RadioButton
		Me.cboProduct = New System.Windows.Forms.ComboBox
		Me._optChoice_3 = New System.Windows.Forms.RadioButton
		Me._optChoice_2 = New System.Windows.Forms.RadioButton
		Me._optChoice_1 = New System.Windows.Forms.RadioButton
		Me._optChoice_0 = New System.Windows.Forms.RadioButton
		Me.Days = New System.Windows.Forms.Panel
		Me.TextDaysMax = New System.Windows.Forms.TextBox
		Me.TextDaysMin = New System.Windows.Forms.TextBox
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label22 = New System.Windows.Forms.Label
		Me.Label23 = New System.Windows.Forms.Label
		Me.Shape3 = New System.Windows.Forms.Label
		Me.Text_Renamed = New System.Windows.Forms.Panel
		Me.textNotes = New System.Windows.Forms.ComboBox
		Me.Label12 = New System.Windows.Forms.Label
		Me.Shape5 = New System.Windows.Forms.Label
		Me._Picture1_1 = New System.Windows.Forms.Panel
		Me._Picture1_2 = New System.Windows.Forms.Panel
		Me.chkAlpha = New System.Windows.Forms.CheckBox
		Me._Picture1_4 = New System.Windows.Forms.Panel
		Me.cboAction = New System.Windows.Forms.ComboBox
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me._Picture1_5 = New System.Windows.Forms.Panel
		Me.frmCriteria = New System.Windows.Forms.Panel
		Me.TextSource = New System.Windows.Forms.TextBox
		Me.ComboStatus = New System.Windows.Forms.ComboBox
		Me.TextZip = New System.Windows.Forms.TextBox
		Me.TextCity = New System.Windows.Forms.TextBox
		Me.TextCompany = New System.Windows.Forms.TextBox
		Me.TextLastName = New System.Windows.Forms.TextBox
		Me.TextFirstName = New System.Windows.Forms.TextBox
		Me.StateCombo = New System.Windows.Forms.ComboBox
		Me.Label13 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me.PrintButton = New System.Windows.Forms.Button
		Me.ListView1 = New System.Windows.Forms.ListView
		Me.LabelResults = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Picture1 = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(components)
		Me.chkFilter = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.optChoice = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.frame1.SuspendLayout()
		Me.Frame4.SuspendLayout()
		Me.Days.SuspendLayout()
		Me.Text_Renamed.SuspendLayout()
		Me._Picture1_2.SuspendLayout()
		Me._Picture1_4.SuspendLayout()
		Me.frmCriteria.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdHistory, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkFilter, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optChoice, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.AppWorkspace
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Report Printer"
		Me.ClientSize = New System.Drawing.Size(997, 703)
		Me.Location = New System.Drawing.Point(178, 187)
		Me.Icon = CType(resources.GetObject("FReport.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.Visible = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FReport"
		Me.frame1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frame1.Size = New System.Drawing.Size(974, 672)
		Me.frame1.Location = New System.Drawing.Point(0, 0)
		Me.frame1.TabIndex = 25
		Me.frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frame1.BackColor = System.Drawing.SystemColors.Control
		Me.frame1.Enabled = True
		Me.frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frame1.Cursor = System.Windows.Forms.Cursors.Default
		Me.frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frame1.Visible = True
		Me.frame1.Name = "frame1"
		Me.cmdCopyToClipBoard.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCopyToClipBoard.Text = "Copy To Clipboard"
		Me.cmdCopyToClipBoard.Size = New System.Drawing.Size(164, 29)
		Me.cmdCopyToClipBoard.Location = New System.Drawing.Point(600, 630)
		Me.cmdCopyToClipBoard.TabIndex = 49
		Me.cmdCopyToClipBoard.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCopyToClipBoard.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCopyToClipBoard.CausesValidation = True
		Me.cmdCopyToClipBoard.Enabled = True
		Me.cmdCopyToClipBoard.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCopyToClipBoard.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCopyToClipBoard.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCopyToClipBoard.TabStop = True
		Me.cmdCopyToClipBoard.Name = "cmdCopyToClipBoard"
		grdHistory.OcxState = CType(resources.GetObject("grdHistory.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdHistory.Size = New System.Drawing.Size(589, 144)
		Me.grdHistory.Location = New System.Drawing.Point(380, 470)
		Me.grdHistory.TabIndex = 48
		Me.grdHistory.Name = "grdHistory"
		Me.ShowResults.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.ShowResults.Text = "Show Results"
		Me.ShowResults.Size = New System.Drawing.Size(164, 29)
		Me.ShowResults.Location = New System.Drawing.Point(240, 630)
		Me.ShowResults.TabIndex = 22
		Me.ShowResults.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ShowResults.BackColor = System.Drawing.SystemColors.Control
		Me.ShowResults.CausesValidation = True
		Me.ShowResults.Enabled = True
		Me.ShowResults.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ShowResults.Cursor = System.Windows.Forms.Cursors.Default
		Me.ShowResults.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowResults.TabStop = True
		Me.ShowResults.Name = "ShowResults"
		Me.PreviewReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.PreviewReport.Text = "Old Report"
		Me.PreviewReport.Size = New System.Drawing.Size(172, 29)
		Me.PreviewReport.Location = New System.Drawing.Point(40, 630)
		Me.PreviewReport.TabIndex = 23
		Me.PreviewReport.Visible = False
		Me.PreviewReport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PreviewReport.BackColor = System.Drawing.SystemColors.Control
		Me.PreviewReport.CausesValidation = True
		Me.PreviewReport.Enabled = True
		Me.PreviewReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.PreviewReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.PreviewReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PreviewReport.TabStop = True
		Me.PreviewReport.Name = "PreviewReport"
		Me.Frame4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Frame4.Text = "Choose Search"
		Me.Frame4.ForeColor = System.Drawing.Color.Black
		Me.Frame4.Size = New System.Drawing.Size(360, 609)
		Me.Frame4.Location = New System.Drawing.Point(10, 10)
		Me.Frame4.TabIndex = 34
		Me.Frame4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame4.BackColor = System.Drawing.SystemColors.Control
		Me.Frame4.Enabled = True
		Me.Frame4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame4.Visible = True
		Me.Frame4.Name = "Frame4"
		Me.txtDays.AutoSize = False
		Me.txtDays.Size = New System.Drawing.Size(42, 24)
		Me.txtDays.Location = New System.Drawing.Point(30, 127)
		Me.txtDays.TabIndex = 53
		Me.txtDays.Text = "90"
		Me.txtDays.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDays.AcceptsReturn = True
		Me.txtDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDays.BackColor = System.Drawing.SystemColors.Window
		Me.txtDays.CausesValidation = True
		Me.txtDays.Enabled = True
		Me.txtDays.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDays.HideSelection = True
		Me.txtDays.ReadOnly = False
		Me.txtDays.Maxlength = 0
		Me.txtDays.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDays.MultiLine = False
		Me.txtDays.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDays.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDays.TabStop = True
		Me.txtDays.Visible = True
		Me.txtDays.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDays.Name = "txtDays"
		Me._optChoice_4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_4.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me._optChoice_4.ForeColor = System.Drawing.Color.Black
		Me._optChoice_4.Size = New System.Drawing.Size(24, 27)
		Me._optChoice_4.Location = New System.Drawing.Point(10, 127)
		Me._optChoice_4.TabIndex = 52
		Me._optChoice_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optChoice_4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_4.CausesValidation = True
		Me._optChoice_4.Enabled = True
		Me._optChoice_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._optChoice_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optChoice_4.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optChoice_4.TabStop = True
		Me._optChoice_4.Checked = False
		Me._optChoice_4.Visible = True
		Me._optChoice_4.Name = "_optChoice_4"
		Me.chkTtls.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me.chkTtls.Text = "Show &Totals"
		Me.chkTtls.Size = New System.Drawing.Size(99, 19)
		Me.chkTtls.Location = New System.Drawing.Point(183, 554)
		Me.chkTtls.TabIndex = 12
		Me.chkTtls.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTtls.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTtls.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTtls.CausesValidation = True
		Me.chkTtls.Enabled = True
		Me.chkTtls.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkTtls.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkTtls.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkTtls.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkTtls.TabStop = True
		Me.chkTtls.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkTtls.Visible = True
		Me.chkTtls.Name = "chkTtls"
		Me._chkFilter_1.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me._chkFilter_1.Text = "AM &Best"
		Me._chkFilter_1.Size = New System.Drawing.Size(77, 19)
		Me._chkFilter_1.Location = New System.Drawing.Point(100, 554)
		Me._chkFilter_1.TabIndex = 11
		Me._chkFilter_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkFilter_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkFilter_1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkFilter_1.CausesValidation = True
		Me._chkFilter_1.Enabled = True
		Me._chkFilter_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkFilter_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkFilter_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkFilter_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkFilter_1.TabStop = True
		Me._chkFilter_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkFilter_1.Visible = True
		Me._chkFilter_1.Name = "_chkFilter_1"
		Me._chkFilter_0.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me._chkFilter_0.Text = "&Standard"
		Me._chkFilter_0.Size = New System.Drawing.Size(82, 19)
		Me._chkFilter_0.Location = New System.Drawing.Point(13, 554)
		Me._chkFilter_0.TabIndex = 10
		Me._chkFilter_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkFilter_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkFilter_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkFilter_0.CausesValidation = True
		Me._chkFilter_0.Enabled = True
		Me._chkFilter_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkFilter_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkFilter_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkFilter_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkFilter_0.TabStop = True
		Me._chkFilter_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkFilter_0.Visible = True
		Me._chkFilter_0.Name = "_chkFilter_0"
		Me.lstGroups.Size = New System.Drawing.Size(342, 253)
		Me.lstGroups.Location = New System.Drawing.Point(10, 270)
		Me.lstGroups.TabIndex = 9
		Me.lstGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstGroups.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstGroups.BackColor = System.Drawing.SystemColors.Window
		Me.lstGroups.CausesValidation = True
		Me.lstGroups.Enabled = True
		Me.lstGroups.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstGroups.IntegralHeight = True
		Me.lstGroups.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstGroups.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstGroups.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstGroups.Sorted = False
		Me.lstGroups.TabStop = True
		Me.lstGroups.Visible = True
		Me.lstGroups.MultiColumn = False
		Me.lstGroups.Name = "lstGroups"
		Me._optChoice_5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_5.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me._optChoice_5.Text = "Group"
		Me._optChoice_5.ForeColor = System.Drawing.Color.Black
		Me._optChoice_5.Size = New System.Drawing.Size(79, 24)
		Me._optChoice_5.Location = New System.Drawing.Point(4, 240)
		Me._optChoice_5.TabIndex = 8
		Me._optChoice_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optChoice_5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_5.CausesValidation = True
		Me._optChoice_5.Enabled = True
		Me._optChoice_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._optChoice_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optChoice_5.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optChoice_5.TabStop = True
		Me._optChoice_5.Checked = False
		Me._optChoice_5.Visible = True
		Me._optChoice_5.Name = "_optChoice_5"
		Me.cboProduct.Size = New System.Drawing.Size(142, 27)
		Me.cboProduct.Location = New System.Drawing.Point(15, 93)
		Me.cboProduct.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboProduct.TabIndex = 3
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
		Me._optChoice_3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_3.BackColor = System.Drawing.Color.FromARGB(225, 225, 255)
		Me._optChoice_3.Text = "Notes"
		Me._optChoice_3.ForeColor = System.Drawing.Color.Black
		Me._optChoice_3.Size = New System.Drawing.Size(64, 29)
		Me._optChoice_3.Location = New System.Drawing.Point(10, 170)
		Me._optChoice_3.TabIndex = 6
		Me._optChoice_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optChoice_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_3.CausesValidation = True
		Me._optChoice_3.Enabled = True
		Me._optChoice_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._optChoice_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optChoice_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optChoice_3.TabStop = True
		Me._optChoice_3.Checked = False
		Me._optChoice_3.Visible = True
		Me._optChoice_3.Name = "_optChoice_3"
		Me._optChoice_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_2.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me._optChoice_2.Text = "Days Not Authorized"
		Me._optChoice_2.ForeColor = System.Drawing.Color.Black
		Me._optChoice_2.Size = New System.Drawing.Size(152, 24)
		Me._optChoice_2.Location = New System.Drawing.Point(10, 70)
		Me._optChoice_2.TabIndex = 2
		Me._optChoice_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optChoice_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_2.CausesValidation = True
		Me._optChoice_2.Enabled = True
		Me._optChoice_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._optChoice_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optChoice_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optChoice_2.TabStop = True
		Me._optChoice_2.Checked = False
		Me._optChoice_2.Visible = True
		Me._optChoice_2.Name = "_optChoice_2"
		Me._optChoice_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_1.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me._optChoice_1.Text = "Days Left"
		Me._optChoice_1.ForeColor = System.Drawing.Color.Black
		Me._optChoice_1.Size = New System.Drawing.Size(94, 27)
		Me._optChoice_1.Location = New System.Drawing.Point(10, 40)
		Me._optChoice_1.TabIndex = 1
		Me._optChoice_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optChoice_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_1.CausesValidation = True
		Me._optChoice_1.Enabled = True
		Me._optChoice_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optChoice_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optChoice_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optChoice_1.TabStop = True
		Me._optChoice_1.Checked = False
		Me._optChoice_1.Visible = True
		Me._optChoice_1.Name = "_optChoice_1"
		Me._optChoice_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_0.BackColor = System.Drawing.Color.FromARGB(225, 225, 255)
		Me._optChoice_0.Text = "Basic"
		Me._optChoice_0.ForeColor = System.Drawing.Color.Black
		Me._optChoice_0.Size = New System.Drawing.Size(79, 24)
		Me._optChoice_0.Location = New System.Drawing.Point(10, 8)
		Me._optChoice_0.TabIndex = 0
		Me._optChoice_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optChoice_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optChoice_0.CausesValidation = True
		Me._optChoice_0.Enabled = True
		Me._optChoice_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optChoice_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optChoice_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optChoice_0.TabStop = True
		Me._optChoice_0.Checked = False
		Me._optChoice_0.Visible = True
		Me._optChoice_0.Name = "_optChoice_0"
		Me.Days.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me.Days.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Days.Text = "Days"
		Me.Days.Size = New System.Drawing.Size(187, 82)
		Me.Days.Location = New System.Drawing.Point(160, 43)
		Me.Days.TabIndex = 37
		Me.Days.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Days.Enabled = True
		Me.Days.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Days.Cursor = System.Windows.Forms.Cursors.Default
		Me.Days.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Days.Visible = True
		Me.Days.Name = "Days"
		Me.TextDaysMax.AutoSize = False
		Me.TextDaysMax.Size = New System.Drawing.Size(109, 27)
		Me.TextDaysMax.Location = New System.Drawing.Point(54, 42)
		Me.TextDaysMax.TabIndex = 5
		Me.TextDaysMax.Text = "30"
		Me.TextDaysMax.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextDaysMax.AcceptsReturn = True
		Me.TextDaysMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextDaysMax.BackColor = System.Drawing.SystemColors.Window
		Me.TextDaysMax.CausesValidation = True
		Me.TextDaysMax.Enabled = True
		Me.TextDaysMax.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextDaysMax.HideSelection = True
		Me.TextDaysMax.ReadOnly = False
		Me.TextDaysMax.Maxlength = 0
		Me.TextDaysMax.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextDaysMax.MultiLine = False
		Me.TextDaysMax.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextDaysMax.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextDaysMax.TabStop = True
		Me.TextDaysMax.Visible = True
		Me.TextDaysMax.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextDaysMax.Name = "TextDaysMax"
		Me.TextDaysMin.AutoSize = False
		Me.TextDaysMin.Size = New System.Drawing.Size(109, 27)
		Me.TextDaysMin.Location = New System.Drawing.Point(54, 14)
		Me.TextDaysMin.TabIndex = 4
		Me.TextDaysMin.Text = "0"
		Me.TextDaysMin.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextDaysMin.AcceptsReturn = True
		Me.TextDaysMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextDaysMin.BackColor = System.Drawing.SystemColors.Window
		Me.TextDaysMin.CausesValidation = True
		Me.TextDaysMin.Enabled = True
		Me.TextDaysMin.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextDaysMin.HideSelection = True
		Me.TextDaysMin.ReadOnly = False
		Me.TextDaysMin.Maxlength = 0
		Me.TextDaysMin.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextDaysMin.MultiLine = False
		Me.TextDaysMin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextDaysMin.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextDaysMin.TabStop = True
		Me.TextDaysMin.Visible = True
		Me.TextDaysMin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextDaysMin.Name = "TextDaysMin"
		Me.Label10.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me.Label10.Text = " Days"
		Me.Label10.Size = New System.Drawing.Size(39, 19)
		Me.Label10.Location = New System.Drawing.Point(9, -1)
		Me.Label10.TabIndex = 38
		Me.Label10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label10.Enabled = True
		Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label10.UseMnemonic = True
		Me.Label10.Visible = True
		Me.Label10.AutoSize = False
		Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label10.Name = "Label10"
		Me.Label22.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me.Label22.Text = "Min:"
		Me.Label22.ForeColor = System.Drawing.Color.Black
		Me.Label22.Size = New System.Drawing.Size(32, 24)
		Me.Label22.Location = New System.Drawing.Point(19, 19)
		Me.Label22.TabIndex = 40
		Me.Label22.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label22.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label22.Enabled = True
		Me.Label22.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label22.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label22.UseMnemonic = True
		Me.Label22.Visible = True
		Me.Label22.AutoSize = False
		Me.Label22.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label22.Name = "Label22"
		Me.Label23.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me.Label23.Text = "Max:"
		Me.Label23.ForeColor = System.Drawing.Color.Black
		Me.Label23.Size = New System.Drawing.Size(34, 24)
		Me.Label23.Location = New System.Drawing.Point(17, 47)
		Me.Label23.TabIndex = 39
		Me.Label23.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label23.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label23.Enabled = True
		Me.Label23.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label23.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label23.UseMnemonic = True
		Me.Label23.Visible = True
		Me.Label23.AutoSize = False
		Me.Label23.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label23.Name = "Label23"
		Me.Shape3.BackColor = System.Drawing.Color.Transparent
		Me.Shape3.Size = New System.Drawing.Size(177, 69)
		Me.Shape3.Location = New System.Drawing.Point(3, 7)
		Me.Shape3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape3.Visible = True
		Me.Shape3.Name = "Shape3"
		Me.Text_Renamed.BackColor = System.Drawing.Color.FromARGB(225, 225, 255)
		Me.Text_Renamed.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Text_Renamed.Text = "Text"
		Me.Text_Renamed.Size = New System.Drawing.Size(262, 64)
		Me.Text_Renamed.Location = New System.Drawing.Point(80, 163)
		Me.Text_Renamed.TabIndex = 35
		Me.Text_Renamed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text_Renamed.Enabled = True
		Me.Text_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Text_Renamed.Cursor = System.Windows.Forms.Cursors.Default
		Me.Text_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text_Renamed.Visible = True
		Me.Text_Renamed.Name = "Text_Renamed"
		Me.textNotes.Size = New System.Drawing.Size(247, 27)
		Me.textNotes.Location = New System.Drawing.Point(10, 20)
		Me.textNotes.TabIndex = 7
		Me.textNotes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.textNotes.BackColor = System.Drawing.SystemColors.Window
		Me.textNotes.CausesValidation = True
		Me.textNotes.Enabled = True
		Me.textNotes.ForeColor = System.Drawing.SystemColors.WindowText
		Me.textNotes.IntegralHeight = True
		Me.textNotes.Cursor = System.Windows.Forms.Cursors.Default
		Me.textNotes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.textNotes.Sorted = False
		Me.textNotes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.textNotes.TabStop = True
		Me.textNotes.Visible = True
		Me.textNotes.Name = "textNotes"
		Me.Label12.BackColor = System.Drawing.Color.FromARGB(225, 225, 255)
		Me.Label12.Text = "Text"
		Me.Label12.Size = New System.Drawing.Size(39, 19)
		Me.Label12.Location = New System.Drawing.Point(10, 0)
		Me.Label12.TabIndex = 36
		Me.Label12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label12.Enabled = True
		Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label12.UseMnemonic = True
		Me.Label12.Visible = True
		Me.Label12.AutoSize = False
		Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label12.Name = "Label12"
		Me.Shape5.Size = New System.Drawing.Size(259, 47)
		Me.Shape5.Location = New System.Drawing.Point(0, 10)
		Me.Shape5.BackColor = System.Drawing.Color.Transparent
		Me.Shape5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape5.Visible = True
		Me.Shape5.Name = "Shape5"
		Me._Picture1_1.BackColor = System.Drawing.Color.FromARGB(225, 225, 255)
		Me._Picture1_1.Size = New System.Drawing.Size(362, 72)
		Me._Picture1_1.Location = New System.Drawing.Point(0, 160)
		Me._Picture1_1.TabIndex = 44
		Me._Picture1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Picture1_1.Dock = System.Windows.Forms.DockStyle.None
		Me._Picture1_1.CausesValidation = True
		Me._Picture1_1.Enabled = True
		Me._Picture1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Picture1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Picture1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Picture1_1.TabStop = True
		Me._Picture1_1.Visible = True
		Me._Picture1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Picture1_1.Name = "_Picture1_1"
		Me._Picture1_2.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me._Picture1_2.Size = New System.Drawing.Size(362, 372)
		Me._Picture1_2.Location = New System.Drawing.Point(0, 230)
		Me._Picture1_2.TabIndex = 45
		Me._Picture1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Picture1_2.Dock = System.Windows.Forms.DockStyle.None
		Me._Picture1_2.CausesValidation = True
		Me._Picture1_2.Enabled = True
		Me._Picture1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Picture1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Picture1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Picture1_2.TabStop = True
		Me._Picture1_2.Visible = True
		Me._Picture1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Picture1_2.Name = "_Picture1_2"
		Me.chkAlpha.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me.chkAlpha.Text = "Sort Alphabetically"
		Me.chkAlpha.Size = New System.Drawing.Size(142, 22)
		Me.chkAlpha.Location = New System.Drawing.Point(12, 300)
		Me.chkAlpha.TabIndex = 13
		Me.chkAlpha.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAlpha.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAlpha.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAlpha.CausesValidation = True
		Me.chkAlpha.Enabled = True
		Me.chkAlpha.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAlpha.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAlpha.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAlpha.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAlpha.TabStop = True
		Me.chkAlpha.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkAlpha.Visible = True
		Me.chkAlpha.Name = "chkAlpha"
		Me._Picture1_4.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me._Picture1_4.Size = New System.Drawing.Size(362, 122)
		Me._Picture1_4.Location = New System.Drawing.Point(0, 40)
		Me._Picture1_4.TabIndex = 46
		Me._Picture1_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Picture1_4.Dock = System.Windows.Forms.DockStyle.None
		Me._Picture1_4.CausesValidation = True
		Me._Picture1_4.Enabled = True
		Me._Picture1_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Picture1_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Picture1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Picture1_4.TabStop = True
		Me._Picture1_4.Visible = True
		Me._Picture1_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Picture1_4.Name = "_Picture1_4"
		Me.cboAction.Size = New System.Drawing.Size(142, 27)
		Me.cboAction.Location = New System.Drawing.Point(150, 85)
		Me.cboAction.TabIndex = 50
		Me.cboAction.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboAction.BackColor = System.Drawing.SystemColors.Window
		Me.cboAction.CausesValidation = True
		Me.cboAction.Enabled = True
		Me.cboAction.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboAction.IntegralHeight = True
		Me.cboAction.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboAction.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboAction.Sorted = False
		Me.cboAction.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboAction.TabStop = True
		Me.cboAction.Visible = True
		Me.cboAction.Name = "cboAction"
		Me.Label9.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me.Label9.Text = "Days Since"
		Me.Label9.Size = New System.Drawing.Size(68, 17)
		Me.Label9.Location = New System.Drawing.Point(80, 90)
		Me.Label9.TabIndex = 54
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
		Me.Label8.BackColor = System.Drawing.Color.FromARGB(204, 204, 255)
		Me.Label8.Text = "Contact"
		Me.Label8.Size = New System.Drawing.Size(52, 22)
		Me.Label8.Location = New System.Drawing.Point(300, 90)
		Me.Label8.TabIndex = 51
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
		Me._Picture1_5.BackColor = System.Drawing.Color.FromARGB(225, 225, 255)
		Me._Picture1_5.Size = New System.Drawing.Size(362, 52)
		Me._Picture1_5.Location = New System.Drawing.Point(0, 0)
		Me._Picture1_5.TabIndex = 47
		Me._Picture1_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Picture1_5.Dock = System.Windows.Forms.DockStyle.None
		Me._Picture1_5.CausesValidation = True
		Me._Picture1_5.Enabled = True
		Me._Picture1_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Picture1_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Picture1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Picture1_5.TabStop = True
		Me._Picture1_5.Visible = True
		Me._Picture1_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Picture1_5.Name = "_Picture1_5"
		Me.frmCriteria.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmCriteria.Text = "Other Criteria"
		Me.frmCriteria.ForeColor = System.Drawing.Color.Black
		Me.frmCriteria.Size = New System.Drawing.Size(579, 144)
		Me.frmCriteria.Location = New System.Drawing.Point(380, 470)
		Me.frmCriteria.TabIndex = 26
		Me.frmCriteria.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmCriteria.BackColor = System.Drawing.SystemColors.Control
		Me.frmCriteria.Enabled = True
		Me.frmCriteria.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmCriteria.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmCriteria.Visible = True
		Me.frmCriteria.Name = "frmCriteria"
		Me.TextSource.AutoSize = False
		Me.TextSource.Size = New System.Drawing.Size(173, 24)
		Me.TextSource.Location = New System.Drawing.Point(365, 109)
		Me.TextSource.Maxlength = 100
		Me.TextSource.TabIndex = 21
		Me.TextSource.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextSource.AcceptsReturn = True
		Me.TextSource.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextSource.BackColor = System.Drawing.SystemColors.Window
		Me.TextSource.CausesValidation = True
		Me.TextSource.Enabled = True
		Me.TextSource.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextSource.HideSelection = True
		Me.TextSource.ReadOnly = False
		Me.TextSource.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextSource.MultiLine = False
		Me.TextSource.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextSource.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextSource.TabStop = True
		Me.TextSource.Visible = True
		Me.TextSource.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextSource.Name = "TextSource"
		Me.ComboStatus.Size = New System.Drawing.Size(172, 27)
		Me.ComboStatus.Location = New System.Drawing.Point(132, 18)
		Me.ComboStatus.TabIndex = 14
		Me.ComboStatus.Text = "Customer"
		Me.ComboStatus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ComboStatus.BackColor = System.Drawing.SystemColors.Window
		Me.ComboStatus.CausesValidation = True
		Me.ComboStatus.Enabled = True
		Me.ComboStatus.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ComboStatus.IntegralHeight = True
		Me.ComboStatus.Cursor = System.Windows.Forms.Cursors.Default
		Me.ComboStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ComboStatus.Sorted = False
		Me.ComboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.ComboStatus.TabStop = True
		Me.ComboStatus.Visible = True
		Me.ComboStatus.Name = "ComboStatus"
		Me.TextZip.AutoSize = False
		Me.TextZip.Size = New System.Drawing.Size(172, 27)
		Me.TextZip.Location = New System.Drawing.Point(365, 80)
		Me.TextZip.TabIndex = 20
		Me.TextZip.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextZip.AcceptsReturn = True
		Me.TextZip.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextZip.BackColor = System.Drawing.SystemColors.Window
		Me.TextZip.CausesValidation = True
		Me.TextZip.Enabled = True
		Me.TextZip.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextZip.HideSelection = True
		Me.TextZip.ReadOnly = False
		Me.TextZip.Maxlength = 0
		Me.TextZip.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextZip.MultiLine = False
		Me.TextZip.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextZip.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextZip.TabStop = True
		Me.TextZip.Visible = True
		Me.TextZip.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextZip.Name = "TextZip"
		Me.TextCity.AutoSize = False
		Me.TextCity.Size = New System.Drawing.Size(172, 27)
		Me.TextCity.Location = New System.Drawing.Point(365, 19)
		Me.TextCity.TabIndex = 18
		Me.TextCity.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextCity.AcceptsReturn = True
		Me.TextCity.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextCity.BackColor = System.Drawing.SystemColors.Window
		Me.TextCity.CausesValidation = True
		Me.TextCity.Enabled = True
		Me.TextCity.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextCity.HideSelection = True
		Me.TextCity.ReadOnly = False
		Me.TextCity.Maxlength = 0
		Me.TextCity.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextCity.MultiLine = False
		Me.TextCity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextCity.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextCity.TabStop = True
		Me.TextCity.Visible = True
		Me.TextCity.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextCity.Name = "TextCity"
		Me.TextCompany.AutoSize = False
		Me.TextCompany.Size = New System.Drawing.Size(172, 27)
		Me.TextCompany.Location = New System.Drawing.Point(132, 105)
		Me.TextCompany.TabIndex = 17
		Me.TextCompany.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextCompany.AcceptsReturn = True
		Me.TextCompany.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextCompany.BackColor = System.Drawing.SystemColors.Window
		Me.TextCompany.CausesValidation = True
		Me.TextCompany.Enabled = True
		Me.TextCompany.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextCompany.HideSelection = True
		Me.TextCompany.ReadOnly = False
		Me.TextCompany.Maxlength = 0
		Me.TextCompany.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextCompany.MultiLine = False
		Me.TextCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextCompany.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextCompany.TabStop = True
		Me.TextCompany.Visible = True
		Me.TextCompany.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextCompany.Name = "TextCompany"
		Me.TextLastName.AutoSize = False
		Me.TextLastName.Size = New System.Drawing.Size(172, 27)
		Me.TextLastName.Location = New System.Drawing.Point(132, 75)
		Me.TextLastName.TabIndex = 16
		Me.TextLastName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextLastName.AcceptsReturn = True
		Me.TextLastName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextLastName.BackColor = System.Drawing.SystemColors.Window
		Me.TextLastName.CausesValidation = True
		Me.TextLastName.Enabled = True
		Me.TextLastName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextLastName.HideSelection = True
		Me.TextLastName.ReadOnly = False
		Me.TextLastName.Maxlength = 0
		Me.TextLastName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextLastName.MultiLine = False
		Me.TextLastName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextLastName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextLastName.TabStop = True
		Me.TextLastName.Visible = True
		Me.TextLastName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextLastName.Name = "TextLastName"
		Me.TextFirstName.AutoSize = False
		Me.TextFirstName.Size = New System.Drawing.Size(172, 27)
		Me.TextFirstName.Location = New System.Drawing.Point(132, 45)
		Me.TextFirstName.TabIndex = 15
		Me.TextFirstName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextFirstName.AcceptsReturn = True
		Me.TextFirstName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextFirstName.BackColor = System.Drawing.SystemColors.Window
		Me.TextFirstName.CausesValidation = True
		Me.TextFirstName.Enabled = True
		Me.TextFirstName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextFirstName.HideSelection = True
		Me.TextFirstName.ReadOnly = False
		Me.TextFirstName.Maxlength = 0
		Me.TextFirstName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextFirstName.MultiLine = False
		Me.TextFirstName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextFirstName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextFirstName.TabStop = True
		Me.TextFirstName.Visible = True
		Me.TextFirstName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextFirstName.Name = "TextFirstName"
		Me.StateCombo.Size = New System.Drawing.Size(72, 27)
		Me.StateCombo.Location = New System.Drawing.Point(365, 49)
		Me.StateCombo.TabIndex = 19
		Me.StateCombo.Text = "All"
		Me.StateCombo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StateCombo.BackColor = System.Drawing.SystemColors.Window
		Me.StateCombo.CausesValidation = True
		Me.StateCombo.Enabled = True
		Me.StateCombo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.StateCombo.IntegralHeight = True
		Me.StateCombo.Cursor = System.Windows.Forms.Cursors.Default
		Me.StateCombo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.StateCombo.Sorted = False
		Me.StateCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.StateCombo.TabStop = True
		Me.StateCombo.Visible = True
		Me.StateCombo.Name = "StateCombo"
		Me.Label13.Text = "Source:"
		Me.Label13.Size = New System.Drawing.Size(52, 22)
		Me.Label13.Location = New System.Drawing.Point(308, 110)
		Me.Label13.TabIndex = 43
		Me.Label13.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label13.BackColor = System.Drawing.SystemColors.Control
		Me.Label13.Enabled = True
		Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label13.UseMnemonic = True
		Me.Label13.Visible = True
		Me.Label13.AutoSize = False
		Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label13.Name = "Label13"
		Me.Label7.Text = "Status:"
		Me.Label7.ForeColor = System.Drawing.Color.Black
		Me.Label7.Size = New System.Drawing.Size(44, 22)
		Me.Label7.Location = New System.Drawing.Point(84, 20)
		Me.Label7.TabIndex = 33
		Me.Label7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.BackColor = System.Drawing.SystemColors.Control
		Me.Label7.Enabled = True
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.Label6.Text = "Zip:"
		Me.Label6.ForeColor = System.Drawing.Color.Black
		Me.Label6.Size = New System.Drawing.Size(29, 22)
		Me.Label6.Location = New System.Drawing.Point(333, 82)
		Me.Label6.TabIndex = 32
		Me.Label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.BackColor = System.Drawing.SystemColors.Control
		Me.Label6.Enabled = True
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.Label5.Text = "State:"
		Me.Label5.ForeColor = System.Drawing.Color.Black
		Me.Label5.Size = New System.Drawing.Size(42, 22)
		Me.Label5.Location = New System.Drawing.Point(320, 52)
		Me.Label5.TabIndex = 31
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.Text = "City:"
		Me.Label4.ForeColor = System.Drawing.Color.Black
		Me.Label4.Size = New System.Drawing.Size(32, 22)
		Me.Label4.Location = New System.Drawing.Point(333, 22)
		Me.Label4.TabIndex = 30
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.Text = "Company:"
		Me.Label3.ForeColor = System.Drawing.Color.Black
		Me.Label3.Size = New System.Drawing.Size(64, 22)
		Me.Label3.Location = New System.Drawing.Point(64, 110)
		Me.Label3.TabIndex = 29
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.Text = "Last Name:"
		Me.Label2.ForeColor = System.Drawing.Color.Black
		Me.Label2.Size = New System.Drawing.Size(72, 19)
		Me.Label2.Location = New System.Drawing.Point(57, 80)
		Me.Label2.TabIndex = 28
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me._Label1_1.Text = "First Name:"
		Me._Label1_1.ForeColor = System.Drawing.Color.Black
		Me._Label1_1.Size = New System.Drawing.Size(69, 22)
		Me._Label1_1.Location = New System.Drawing.Point(62, 50)
		Me._Label1_1.TabIndex = 27
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me.PrintButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.PrintButton.Text = "Print Report"
		Me.PrintButton.Size = New System.Drawing.Size(164, 29)
		Me.PrintButton.Location = New System.Drawing.Point(420, 630)
		Me.PrintButton.TabIndex = 24
		Me.PrintButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PrintButton.BackColor = System.Drawing.SystemColors.Control
		Me.PrintButton.CausesValidation = True
		Me.PrintButton.Enabled = True
		Me.PrintButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.PrintButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.PrintButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PrintButton.TabStop = True
		Me.PrintButton.Name = "PrintButton"
		Me.ListView1.Size = New System.Drawing.Size(589, 442)
		Me.ListView1.Location = New System.Drawing.Point(378, 23)
		Me.ListView1.TabIndex = 41
		Me.ListView1.TabStop = 0
		Me.ListView1.View = System.Windows.Forms.View.Details
		Me.ListView1.LabelEdit = False
		Me.ListView1.LabelWrap = True
		Me.ListView1.HideSelection = True
		Me.ListView1.FullRowSelect = True
		Me.ListView1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ListView1.BackColor = System.Drawing.SystemColors.Window
		Me.ListView1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ListView1.Name = "ListView1"
		Me.LabelResults.ForeColor = System.Drawing.Color.FromARGB(64, 64, 64)
		Me.LabelResults.Size = New System.Drawing.Size(104, 17)
		Me.LabelResults.Location = New System.Drawing.Point(380, 5)
		Me.LabelResults.TabIndex = 42
		Me.LabelResults.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LabelResults.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.LabelResults.BackColor = System.Drawing.SystemColors.Control
		Me.LabelResults.Enabled = True
		Me.LabelResults.Cursor = System.Windows.Forms.Cursors.Default
		Me.LabelResults.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.LabelResults.UseMnemonic = True
		Me.LabelResults.Visible = True
		Me.LabelResults.AutoSize = False
		Me.LabelResults.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.LabelResults.Name = "LabelResults"
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Picture1.SetIndex(_Picture1_1, CType(1, Short))
		Me.Picture1.SetIndex(_Picture1_2, CType(2, Short))
		Me.Picture1.SetIndex(_Picture1_4, CType(4, Short))
		Me.Picture1.SetIndex(_Picture1_5, CType(5, Short))
		Me.chkFilter.SetIndex(_chkFilter_1, CType(1, Short))
		Me.chkFilter.SetIndex(_chkFilter_0, CType(0, Short))
		Me.optChoice.SetIndex(_optChoice_4, CType(4, Short))
		Me.optChoice.SetIndex(_optChoice_5, CType(5, Short))
		Me.optChoice.SetIndex(_optChoice_3, CType(3, Short))
		Me.optChoice.SetIndex(_optChoice_2, CType(2, Short))
		Me.optChoice.SetIndex(_optChoice_1, CType(1, Short))
		Me.optChoice.SetIndex(_optChoice_0, CType(0, Short))
		CType(Me.optChoice, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chkFilter, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdHistory, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(frame1)
		Me.frame1.Controls.Add(cmdCopyToClipBoard)
		Me.frame1.Controls.Add(grdHistory)
		Me.frame1.Controls.Add(ShowResults)
		Me.frame1.Controls.Add(PreviewReport)
		Me.frame1.Controls.Add(Frame4)
		Me.frame1.Controls.Add(frmCriteria)
		Me.frame1.Controls.Add(PrintButton)
		Me.frame1.Controls.Add(ListView1)
		Me.frame1.Controls.Add(LabelResults)
		Me.Frame4.Controls.Add(txtDays)
		Me.Frame4.Controls.Add(_optChoice_4)
		Me.Frame4.Controls.Add(chkTtls)
		Me.Frame4.Controls.Add(_chkFilter_1)
		Me.Frame4.Controls.Add(_chkFilter_0)
		Me.Frame4.Controls.Add(lstGroups)
		Me.Frame4.Controls.Add(_optChoice_5)
		Me.Frame4.Controls.Add(cboProduct)
		Me.Frame4.Controls.Add(_optChoice_3)
		Me.Frame4.Controls.Add(_optChoice_2)
		Me.Frame4.Controls.Add(_optChoice_1)
		Me.Frame4.Controls.Add(_optChoice_0)
		Me.Frame4.Controls.Add(Days)
		Me.Frame4.Controls.Add(Text_Renamed)
		Me.Frame4.Controls.Add(_Picture1_1)
		Me.Frame4.Controls.Add(_Picture1_2)
		Me.Frame4.Controls.Add(_Picture1_4)
		Me.Frame4.Controls.Add(_Picture1_5)
		Me.Days.Controls.Add(TextDaysMax)
		Me.Days.Controls.Add(TextDaysMin)
		Me.Days.Controls.Add(Label10)
		Me.Days.Controls.Add(Label22)
		Me.Days.Controls.Add(Label23)
		Me.Days.Controls.Add(Shape3)
		Me.Text_Renamed.Controls.Add(textNotes)
		Me.Text_Renamed.Controls.Add(Label12)
		Me.Text_Renamed.Controls.Add(Shape5)
		Me._Picture1_2.Controls.Add(chkAlpha)
		Me._Picture1_4.Controls.Add(cboAction)
		Me._Picture1_4.Controls.Add(Label9)
		Me._Picture1_4.Controls.Add(Label8)
		Me.frmCriteria.Controls.Add(TextSource)
		Me.frmCriteria.Controls.Add(ComboStatus)
		Me.frmCriteria.Controls.Add(TextZip)
		Me.frmCriteria.Controls.Add(TextCity)
		Me.frmCriteria.Controls.Add(TextCompany)
		Me.frmCriteria.Controls.Add(TextLastName)
		Me.frmCriteria.Controls.Add(TextFirstName)
		Me.frmCriteria.Controls.Add(StateCombo)
		Me.frmCriteria.Controls.Add(Label13)
		Me.frmCriteria.Controls.Add(Label7)
		Me.frmCriteria.Controls.Add(Label6)
		Me.frmCriteria.Controls.Add(Label5)
		Me.frmCriteria.Controls.Add(Label4)
		Me.frmCriteria.Controls.Add(Label3)
		Me.frmCriteria.Controls.Add(Label2)
		Me.frmCriteria.Controls.Add(_Label1_1)
		Me.frame1.ResumeLayout(False)
		Me.Frame4.ResumeLayout(False)
		Me.Days.ResumeLayout(False)
		Me.Text_Renamed.ResumeLayout(False)
		Me._Picture1_2.ResumeLayout(False)
		Me._Picture1_4.ResumeLayout(False)
		Me.frmCriteria.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class