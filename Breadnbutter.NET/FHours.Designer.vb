<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FHours
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
	Public WithEvents txtHours As System.Windows.Forms.TextBox
	Public WithEvents txtLunch As System.Windows.Forms.TextBox
	Public WithEvents txtTotalHours As System.Windows.Forms.TextBox
	Public WithEvents cmdCalcHours As System.Windows.Forms.Button
	Public WithEvents cmdPrint As System.Windows.Forms.Button
	Public WithEvents cmdActual As System.Windows.Forms.Button
	Public WithEvents chkShow As System.Windows.Forms.CheckBox
	Public WithEvents FlexGridHours As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents dcHours As VB6.ADODC
	Public WithEvents cmbEmployee As AxSSDataWidgets_B.AxSSDBCombo
	Public WithEvents mskEndDate As AxTDBDate6.AxTDBDate
	Public WithEvents mskBeginDate As AxTDBDate6.AxTDBDate
	Public WithEvents cmdEndDate As AxThreed.AxSSCommand
	Public WithEvents cmdBeginDate As AxThreed.AxSSCommand
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FHours))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.txtHours = New System.Windows.Forms.TextBox
		Me.txtLunch = New System.Windows.Forms.TextBox
		Me.txtTotalHours = New System.Windows.Forms.TextBox
		Me.cmdCalcHours = New System.Windows.Forms.Button
		Me.cmdPrint = New System.Windows.Forms.Button
		Me.cmdActual = New System.Windows.Forms.Button
		Me.chkShow = New System.Windows.Forms.CheckBox
		Me.FlexGridHours = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.dcHours = New VB6.ADODC
		Me.cmbEmployee = New AxSSDataWidgets_B.AxSSDBCombo
		Me.mskEndDate = New AxTDBDate6.AxTDBDate
		Me.mskBeginDate = New AxTDBDate6.AxTDBDate
		Me.cmdEndDate = New AxThreed.AxSSCommand
		Me.cmdBeginDate = New AxThreed.AxSSCommand
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.FlexGridHours, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmbEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskEndDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskBeginDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdEndDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdBeginDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.AppWorkspace
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Employee Hours"
		Me.ClientSize = New System.Drawing.Size(949, 572)
		Me.Location = New System.Drawing.Point(289, 278)
		Me.Icon = CType(resources.GetObject("FHours.Icon"), System.Drawing.Icon)
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
		Me.Name = "FHours"
		Me.Frame1.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me.Frame1.Size = New System.Drawing.Size(922, 492)
		Me.Frame1.Location = New System.Drawing.Point(0, 0)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.txtHours.AutoSize = False
		Me.txtHours.Size = New System.Drawing.Size(92, 27)
		Me.txtHours.Location = New System.Drawing.Point(805, 385)
		Me.txtHours.TabIndex = 8
		Me.txtHours.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtHours.AcceptsReturn = True
		Me.txtHours.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtHours.BackColor = System.Drawing.SystemColors.Window
		Me.txtHours.CausesValidation = True
		Me.txtHours.Enabled = True
		Me.txtHours.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtHours.HideSelection = True
		Me.txtHours.ReadOnly = False
		Me.txtHours.Maxlength = 0
		Me.txtHours.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHours.MultiLine = False
		Me.txtHours.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHours.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtHours.TabStop = True
		Me.txtHours.Visible = True
		Me.txtHours.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtHours.Name = "txtHours"
		Me.txtLunch.AutoSize = False
		Me.txtLunch.Size = New System.Drawing.Size(92, 27)
		Me.txtLunch.Location = New System.Drawing.Point(805, 415)
		Me.txtLunch.TabIndex = 7
		Me.txtLunch.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLunch.AcceptsReturn = True
		Me.txtLunch.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLunch.BackColor = System.Drawing.SystemColors.Window
		Me.txtLunch.CausesValidation = True
		Me.txtLunch.Enabled = True
		Me.txtLunch.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLunch.HideSelection = True
		Me.txtLunch.ReadOnly = False
		Me.txtLunch.Maxlength = 0
		Me.txtLunch.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLunch.MultiLine = False
		Me.txtLunch.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLunch.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLunch.TabStop = True
		Me.txtLunch.Visible = True
		Me.txtLunch.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLunch.Name = "txtLunch"
		Me.txtTotalHours.AutoSize = False
		Me.txtTotalHours.Size = New System.Drawing.Size(92, 27)
		Me.txtTotalHours.Location = New System.Drawing.Point(805, 445)
		Me.txtTotalHours.TabIndex = 6
		Me.txtTotalHours.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTotalHours.AcceptsReturn = True
		Me.txtTotalHours.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTotalHours.BackColor = System.Drawing.SystemColors.Window
		Me.txtTotalHours.CausesValidation = True
		Me.txtTotalHours.Enabled = True
		Me.txtTotalHours.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTotalHours.HideSelection = True
		Me.txtTotalHours.ReadOnly = False
		Me.txtTotalHours.Maxlength = 0
		Me.txtTotalHours.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTotalHours.MultiLine = False
		Me.txtTotalHours.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTotalHours.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTotalHours.TabStop = True
		Me.txtTotalHours.Visible = True
		Me.txtTotalHours.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTotalHours.Name = "txtTotalHours"
		Me.cmdCalcHours.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCalcHours.Text = "Calculate Hours"
		Me.cmdCalcHours.Size = New System.Drawing.Size(152, 27)
		Me.cmdCalcHours.Location = New System.Drawing.Point(760, 47)
		Me.cmdCalcHours.TabIndex = 5
		Me.cmdCalcHours.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCalcHours.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCalcHours.CausesValidation = True
		Me.cmdCalcHours.Enabled = True
		Me.cmdCalcHours.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCalcHours.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCalcHours.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCalcHours.TabStop = True
		Me.cmdCalcHours.Name = "cmdCalcHours"
		Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPrint.Text = "&Print"
		Me.cmdPrint.Size = New System.Drawing.Size(117, 27)
		Me.cmdPrint.Location = New System.Drawing.Point(20, 420)
		Me.cmdPrint.TabIndex = 4
		Me.cmdPrint.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPrint.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPrint.CausesValidation = True
		Me.cmdPrint.Enabled = True
		Me.cmdPrint.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPrint.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPrint.TabStop = True
		Me.cmdPrint.Name = "cmdPrint"
		Me.cmdActual.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdActual.Text = "Hide/Unhide Actual Times"
		Me.cmdActual.Size = New System.Drawing.Size(192, 27)
		Me.cmdActual.Location = New System.Drawing.Point(340, 425)
		Me.cmdActual.TabIndex = 3
		Me.cmdActual.Visible = False
		Me.cmdActual.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdActual.BackColor = System.Drawing.SystemColors.Control
		Me.cmdActual.CausesValidation = True
		Me.cmdActual.Enabled = True
		Me.cmdActual.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdActual.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdActual.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdActual.TabStop = True
		Me.cmdActual.Name = "cmdActual"
		Me.chkShow.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me.chkShow.Text = "Show Actual Log Times"
		Me.chkShow.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkShow.Size = New System.Drawing.Size(252, 22)
		Me.chkShow.Location = New System.Drawing.Point(500, 47)
		Me.chkShow.TabIndex = 1
		Me.chkShow.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkShow.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkShow.CausesValidation = True
		Me.chkShow.Enabled = True
		Me.chkShow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkShow.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkShow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkShow.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkShow.TabStop = True
		Me.chkShow.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkShow.Visible = True
		Me.chkShow.Name = "chkShow"
		FlexGridHours.OcxState = CType(resources.GetObject("FlexGridHours.OcxState"), System.Windows.Forms.AxHost.State)
		Me.FlexGridHours.Size = New System.Drawing.Size(902, 292)
		Me.FlexGridHours.Location = New System.Drawing.Point(10, 82)
		Me.FlexGridHours.TabIndex = 2
		Me.FlexGridHours.Name = "FlexGridHours"
		Me.dcHours.Size = New System.Drawing.Size(189, 28)
		Me.dcHours.Location = New System.Drawing.Point(330, 385)
		Me.dcHours.Tag = "tblHours"
		Me.dcHours.Visible = 0
		Me.dcHours.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.dcHours.ConnectionTimeout = 15
		Me.dcHours.CommandTimeout = 30
		Me.dcHours.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
		Me.dcHours.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.dcHours.CommandType = ADODB.CommandTypeEnum.adCmdText
		Me.dcHours.CacheSize = 50
		Me.dcHours.MaxRecords = 0
		Me.dcHours.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.dcHours.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.dcHours.BackColor = System.Drawing.SystemColors.Window
		Me.dcHours.ForeColor = System.Drawing.SystemColors.WindowText
		Me.dcHours.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.dcHours.Enabled = True
		Me.dcHours.UserName = ""
		Me.dcHours.RecordSource = ""
		Me.dcHours.Text = "Hours"
		Me.dcHours.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.dcHours.ConnectionString = ""
		Me.dcHours.Name = "dcHours"
		cmbEmployee.OcxState = CType(resources.GetObject("cmbEmployee.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmbEmployee.Size = New System.Drawing.Size(197, 27)
		Me.cmbEmployee.Location = New System.Drawing.Point(285, 45)
		Me.cmbEmployee.TabIndex = 9
		Me.cmbEmployee.Name = "cmbEmployee"
		mskEndDate.OcxState = CType(resources.GetObject("mskEndDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskEndDate.Size = New System.Drawing.Size(87, 27)
		Me.mskEndDate.Location = New System.Drawing.Point(150, 45)
		Me.mskEndDate.TabIndex = 10
		Me.mskEndDate.Name = "mskEndDate"
		mskBeginDate.OcxState = CType(resources.GetObject("mskBeginDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskBeginDate.Size = New System.Drawing.Size(87, 27)
		Me.mskBeginDate.Location = New System.Drawing.Point(10, 45)
		Me.mskBeginDate.TabIndex = 11
		Me.mskBeginDate.Name = "mskBeginDate"
		cmdEndDate.OcxState = CType(resources.GetObject("cmdEndDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdEndDate.Size = New System.Drawing.Size(24, 27)
		Me.cmdEndDate.Location = New System.Drawing.Point(240, 45)
		Me.cmdEndDate.TabIndex = 12
		Me.cmdEndDate.Name = "cmdEndDate"
		cmdBeginDate.OcxState = CType(resources.GetObject("cmdBeginDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdBeginDate.Size = New System.Drawing.Size(24, 27)
		Me.cmdBeginDate.Location = New System.Drawing.Point(100, 45)
		Me.cmdBeginDate.TabIndex = 13
		Me.cmdBeginDate.Name = "cmdBeginDate"
		Me._Label1_0.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me._Label1_0.Text = "Begin Date:"
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.Size = New System.Drawing.Size(107, 22)
		Me._Label1_0.Location = New System.Drawing.Point(15, 20)
		Me._Label1_0.TabIndex = 19
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._Label1_1.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me._Label1_1.Text = "End Date:"
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.Size = New System.Drawing.Size(107, 22)
		Me._Label1_1.Location = New System.Drawing.Point(150, 20)
		Me._Label1_1.TabIndex = 18
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_2.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me._Label1_2.Text = "Employee:"
		Me._Label1_2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.Size = New System.Drawing.Size(187, 22)
		Me._Label1_2.Location = New System.Drawing.Point(285, 20)
		Me._Label1_2.TabIndex = 17
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_2.Enabled = True
		Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me.Label2.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me.Label2.Text = "Hours"
		Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(87, 22)
		Me.Label2.Location = New System.Drawing.Point(715, 390)
		Me.Label2.TabIndex = 16
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
		Me.Label3.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me.Label3.Text = "Lunch"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(92, 22)
		Me.Label3.Location = New System.Drawing.Point(715, 420)
		Me.Label3.TabIndex = 15
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
		Me.Label4.BackColor = System.Drawing.Color.FromARGB(0, 128, 0)
		Me.Label4.Text = "Total"
		Me.Label4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(92, 22)
		Me.Label4.Location = New System.Drawing.Point(715, 450)
		Me.Label4.TabIndex = 14
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
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdBeginDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdEndDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskBeginDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskEndDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmbEmployee, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.FlexGridHours, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(txtHours)
		Me.Frame1.Controls.Add(txtLunch)
		Me.Frame1.Controls.Add(txtTotalHours)
		Me.Frame1.Controls.Add(cmdCalcHours)
		Me.Frame1.Controls.Add(cmdPrint)
		Me.Frame1.Controls.Add(cmdActual)
		Me.Frame1.Controls.Add(chkShow)
		Me.Frame1.Controls.Add(FlexGridHours)
		Me.Frame1.Controls.Add(dcHours)
		Me.Frame1.Controls.Add(cmbEmployee)
		Me.Frame1.Controls.Add(mskEndDate)
		Me.Frame1.Controls.Add(mskBeginDate)
		Me.Frame1.Controls.Add(cmdEndDate)
		Me.Frame1.Controls.Add(cmdBeginDate)
		Me.Frame1.Controls.Add(_Label1_0)
		Me.Frame1.Controls.Add(_Label1_1)
		Me.Frame1.Controls.Add(_Label1_2)
		Me.Frame1.Controls.Add(Label2)
		Me.Frame1.Controls.Add(Label3)
		Me.Frame1.Controls.Add(Label4)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class