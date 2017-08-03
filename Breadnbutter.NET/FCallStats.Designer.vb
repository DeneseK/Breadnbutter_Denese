<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FCallStats
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
		Form_Initialize_renamed()
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
	Public WithEvents ListView1 As System.Windows.Forms.ListView
	Public WithEvents cboMins As System.Windows.Forms.ComboBox
	Public WithEvents chkUpdate As System.Windows.Forms.CheckBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents grdCallData As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents txtAvg As System.Windows.Forms.TextBox
	Public WithEvents txtTotal As System.Windows.Forms.TextBox
	Public WithEvents optDuration As System.Windows.Forms.RadioButton
	Public WithEvents optBNB As System.Windows.Forms.RadioButton
	Public WithEvents cboCallDir As System.Windows.Forms.ComboBox
	Public WithEvents cboGroup As System.Windows.Forms.ComboBox
	Public WithEvents optExt As System.Windows.Forms.RadioButton
	Public WithEvents optAll As System.Windows.Forms.RadioButton
	Public WithEvents lblCallType As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents fraGroup As System.Windows.Forms.GroupBox
	Public WithEvents cmdGenReport As System.Windows.Forms.Button
	Public WithEvents cmdPrintReport As System.Windows.Forms.Button
	Public WithEvents DTPicker2 As AxMSComCtl2.AxDTPicker
	Public WithEvents DTPicker1 As AxMSComCtl2.AxDTPicker
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblStart As System.Windows.Forms.Label
	Public WithEvents lblEnd As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FCallStats))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ListView1 = New System.Windows.Forms.ListView
		Me.cboMins = New System.Windows.Forms.ComboBox
		Me.chkUpdate = New System.Windows.Forms.CheckBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.grdCallData = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.txtAvg = New System.Windows.Forms.TextBox
		Me.txtTotal = New System.Windows.Forms.TextBox
		Me.optDuration = New System.Windows.Forms.RadioButton
		Me.optBNB = New System.Windows.Forms.RadioButton
		Me.fraGroup = New System.Windows.Forms.GroupBox
		Me.cboCallDir = New System.Windows.Forms.ComboBox
		Me.cboGroup = New System.Windows.Forms.ComboBox
		Me.optExt = New System.Windows.Forms.RadioButton
		Me.optAll = New System.Windows.Forms.RadioButton
		Me.lblCallType = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.cmdGenReport = New System.Windows.Forms.Button
		Me.cmdPrintReport = New System.Windows.Forms.Button
		Me.DTPicker2 = New AxMSComCtl2.AxDTPicker
		Me.DTPicker1 = New AxMSComCtl2.AxDTPicker
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblStart = New System.Windows.Forms.Label
		Me.lblEnd = New System.Windows.Forms.Label
		Me.fraGroup.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdCallData, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DTPicker2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DTPicker1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Call / BNB Notes Statistics"
		Me.ClientSize = New System.Drawing.Size(802, 412)
		Me.Location = New System.Drawing.Point(255, 227)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FCallStats"
		Me.ListView1.Size = New System.Drawing.Size(142, 82)
		Me.ListView1.Location = New System.Drawing.Point(40, 170)
		Me.ListView1.TabIndex = 24
		Me.ListView1.View = System.Windows.Forms.View.Details
		Me.ListView1.LabelEdit = False
		Me.ListView1.LabelWrap = True
		Me.ListView1.HideSelection = False
		Me.ListView1.AllowColumnReorder = -1
		Me.ListView1.GridLines = True
		Me.ListView1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ListView1.BackColor = System.Drawing.SystemColors.Window
		Me.ListView1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ListView1.Name = "ListView1"
		Me.cboMins.Size = New System.Drawing.Size(62, 27)
		Me.cboMins.Location = New System.Drawing.Point(270, 95)
		Me.cboMins.Items.AddRange(New Object(){"5", "10", "15", "30", "45", "60"})
		Me.cboMins.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboMins.TabIndex = 22
		Me.cboMins.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboMins.BackColor = System.Drawing.SystemColors.Window
		Me.cboMins.CausesValidation = True
		Me.cboMins.Enabled = True
		Me.cboMins.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboMins.IntegralHeight = True
		Me.cboMins.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboMins.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboMins.Sorted = False
		Me.cboMins.TabStop = True
		Me.cboMins.Visible = True
		Me.cboMins.Name = "cboMins"
		Me.chkUpdate.Text = "Auto Update every"
		Me.chkUpdate.Size = New System.Drawing.Size(142, 22)
		Me.chkUpdate.Location = New System.Drawing.Point(120, 100)
		Me.chkUpdate.TabIndex = 21
		Me.chkUpdate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkUpdate.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUpdate.BackColor = System.Drawing.SystemColors.Control
		Me.chkUpdate.CausesValidation = True
		Me.chkUpdate.Enabled = True
		Me.chkUpdate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkUpdate.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUpdate.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUpdate.TabStop = True
		Me.chkUpdate.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkUpdate.Visible = True
		Me.chkUpdate.Name = "chkUpdate"
		Me.Timer1.Interval = 60000
		Me.Timer1.Enabled = True
		grdCallData.OcxState = CType(resources.GetObject("grdCallData.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdCallData.Size = New System.Drawing.Size(572, 182)
		Me.grdCallData.Location = New System.Drawing.Point(40, 170)
		Me.grdCallData.TabIndex = 2
		Me.grdCallData.Name = "grdCallData"
		Me.txtAvg.AutoSize = False
		Me.txtAvg.Size = New System.Drawing.Size(142, 24)
		Me.txtAvg.Location = New System.Drawing.Point(240, 380)
		Me.txtAvg.TabIndex = 18
		Me.txtAvg.Visible = False
		Me.txtAvg.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtAvg.AcceptsReturn = True
		Me.txtAvg.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAvg.BackColor = System.Drawing.SystemColors.Window
		Me.txtAvg.CausesValidation = True
		Me.txtAvg.Enabled = True
		Me.txtAvg.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtAvg.HideSelection = True
		Me.txtAvg.ReadOnly = False
		Me.txtAvg.Maxlength = 0
		Me.txtAvg.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAvg.MultiLine = False
		Me.txtAvg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAvg.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAvg.TabStop = True
		Me.txtAvg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAvg.Name = "txtAvg"
		Me.txtTotal.AutoSize = False
		Me.txtTotal.Size = New System.Drawing.Size(132, 24)
		Me.txtTotal.Location = New System.Drawing.Point(40, 380)
		Me.txtTotal.TabIndex = 17
		Me.txtTotal.Visible = False
		Me.txtTotal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTotal.AcceptsReturn = True
		Me.txtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTotal.BackColor = System.Drawing.SystemColors.Window
		Me.txtTotal.CausesValidation = True
		Me.txtTotal.Enabled = True
		Me.txtTotal.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTotal.HideSelection = True
		Me.txtTotal.ReadOnly = False
		Me.txtTotal.Maxlength = 0
		Me.txtTotal.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTotal.MultiLine = False
		Me.txtTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTotal.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTotal.TabStop = True
		Me.txtTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTotal.Name = "txtTotal"
		Me.optDuration.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDuration.Text = "Duration"
		Me.optDuration.Size = New System.Drawing.Size(92, 22)
		Me.optDuration.Location = New System.Drawing.Point(10, 100)
		Me.optDuration.TabIndex = 13
		Me.optDuration.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optDuration.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDuration.BackColor = System.Drawing.SystemColors.Control
		Me.optDuration.CausesValidation = True
		Me.optDuration.Enabled = True
		Me.optDuration.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optDuration.Cursor = System.Windows.Forms.Cursors.Default
		Me.optDuration.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optDuration.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optDuration.TabStop = True
		Me.optDuration.Checked = False
		Me.optDuration.Visible = True
		Me.optDuration.Name = "optDuration"
		Me.optBNB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optBNB.Text = "BNB/Call"
		Me.optBNB.Size = New System.Drawing.Size(92, 22)
		Me.optBNB.Location = New System.Drawing.Point(10, 60)
		Me.optBNB.TabIndex = 12
		Me.optBNB.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optBNB.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optBNB.BackColor = System.Drawing.SystemColors.Control
		Me.optBNB.CausesValidation = True
		Me.optBNB.Enabled = True
		Me.optBNB.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optBNB.Cursor = System.Windows.Forms.Cursors.Default
		Me.optBNB.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optBNB.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optBNB.TabStop = True
		Me.optBNB.Checked = False
		Me.optBNB.Visible = True
		Me.optBNB.Name = "optBNB"
		Me.fraGroup.Text = "Ext or Group"
		Me.fraGroup.Size = New System.Drawing.Size(302, 82)
		Me.fraGroup.Location = New System.Drawing.Point(430, 40)
		Me.fraGroup.TabIndex = 3
		Me.fraGroup.Visible = False
		Me.fraGroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraGroup.BackColor = System.Drawing.SystemColors.Control
		Me.fraGroup.Enabled = True
		Me.fraGroup.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraGroup.Name = "fraGroup"
		Me.cboCallDir.Size = New System.Drawing.Size(92, 27)
		Me.cboCallDir.Location = New System.Drawing.Point(200, 40)
		Me.cboCallDir.Items.AddRange(New Object(){"Incoming", "Outgoing", "Both"})
		Me.cboCallDir.TabIndex = 14
		Me.cboCallDir.Text = "Both"
		Me.cboCallDir.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboCallDir.BackColor = System.Drawing.SystemColors.Window
		Me.cboCallDir.CausesValidation = True
		Me.cboCallDir.Enabled = True
		Me.cboCallDir.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboCallDir.IntegralHeight = True
		Me.cboCallDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboCallDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboCallDir.Sorted = False
		Me.cboCallDir.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboCallDir.TabStop = True
		Me.cboCallDir.Visible = True
		Me.cboCallDir.Name = "cboCallDir"
		Me.cboGroup.Size = New System.Drawing.Size(182, 27)
		Me.cboGroup.Location = New System.Drawing.Point(10, 40)
		Me.cboGroup.TabIndex = 4
		Me.cboGroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboGroup.BackColor = System.Drawing.SystemColors.Window
		Me.cboGroup.CausesValidation = True
		Me.cboGroup.Enabled = True
		Me.cboGroup.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboGroup.IntegralHeight = True
		Me.cboGroup.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboGroup.Sorted = False
		Me.cboGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboGroup.TabStop = True
		Me.cboGroup.Visible = True
		Me.cboGroup.Name = "cboGroup"
		Me.optExt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optExt.Text = "by Ext"
		Me.optExt.Size = New System.Drawing.Size(72, 22)
		Me.optExt.Location = New System.Drawing.Point(120, 10)
		Me.optExt.TabIndex = 6
		Me.optExt.Visible = False
		Me.optExt.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optExt.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optExt.BackColor = System.Drawing.SystemColors.Control
		Me.optExt.CausesValidation = True
		Me.optExt.Enabled = True
		Me.optExt.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optExt.Cursor = System.Windows.Forms.Cursors.Default
		Me.optExt.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optExt.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optExt.TabStop = True
		Me.optExt.Checked = False
		Me.optExt.Name = "optExt"
		Me.optAll.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optAll.Text = "All"
		Me.optAll.Size = New System.Drawing.Size(72, 22)
		Me.optAll.Location = New System.Drawing.Point(40, 10)
		Me.optAll.TabIndex = 5
		Me.optAll.Visible = False
		Me.optAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optAll.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optAll.BackColor = System.Drawing.SystemColors.Control
		Me.optAll.CausesValidation = True
		Me.optAll.Enabled = True
		Me.optAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.optAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optAll.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optAll.TabStop = True
		Me.optAll.Checked = False
		Me.optAll.Name = "optAll"
		Me.lblCallType.Text = "Call Direction"
		Me.lblCallType.Size = New System.Drawing.Size(78, 17)
		Me.lblCallType.Location = New System.Drawing.Point(200, 20)
		Me.lblCallType.TabIndex = 15
		Me.lblCallType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCallType.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCallType.BackColor = System.Drawing.SystemColors.Control
		Me.lblCallType.Enabled = True
		Me.lblCallType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCallType.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCallType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCallType.UseMnemonic = True
		Me.lblCallType.Visible = True
		Me.lblCallType.AutoSize = True
		Me.lblCallType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCallType.Name = "lblCallType"
		Me.Label2.Text = "Name"
		Me.Label2.Size = New System.Drawing.Size(42, 22)
		Me.Label2.Location = New System.Drawing.Point(10, 20)
		Me.Label2.TabIndex = 16
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
		Me.cmdGenReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdGenReport.Text = "Run Report"
		Me.cmdGenReport.Size = New System.Drawing.Size(132, 27)
		Me.cmdGenReport.Location = New System.Drawing.Point(120, 130)
		Me.cmdGenReport.TabIndex = 1
		Me.cmdGenReport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdGenReport.BackColor = System.Drawing.SystemColors.Control
		Me.cmdGenReport.CausesValidation = True
		Me.cmdGenReport.Enabled = True
		Me.cmdGenReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdGenReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdGenReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdGenReport.TabStop = True
		Me.cmdGenReport.Name = "cmdGenReport"
		Me.cmdPrintReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPrintReport.Text = "Print Report"
		Me.cmdPrintReport.Enabled = False
		Me.cmdPrintReport.Size = New System.Drawing.Size(132, 27)
		Me.cmdPrintReport.Location = New System.Drawing.Point(270, 130)
		Me.cmdPrintReport.TabIndex = 0
		Me.cmdPrintReport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPrintReport.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPrintReport.CausesValidation = True
		Me.cmdPrintReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPrintReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPrintReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPrintReport.TabStop = True
		Me.cmdPrintReport.Name = "cmdPrintReport"
		DTPicker2.OcxState = CType(resources.GetObject("DTPicker2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DTPicker2.Size = New System.Drawing.Size(132, 27)
		Me.DTPicker2.Location = New System.Drawing.Point(270, 60)
		Me.DTPicker2.TabIndex = 7
		Me.DTPicker2.Name = "DTPicker2"
		DTPicker1.OcxState = CType(resources.GetObject("DTPicker1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DTPicker1.Size = New System.Drawing.Size(132, 27)
		Me.DTPicker1.Location = New System.Drawing.Point(120, 60)
		Me.DTPicker1.TabIndex = 8
		Me.DTPicker1.Name = "DTPicker1"
		Me.Label5.Text = "Mins"
		Me.Label5.Size = New System.Drawing.Size(52, 22)
		Me.Label5.Location = New System.Drawing.Point(340, 100)
		Me.Label5.TabIndex = 23
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
		Me.Label4.Text = "Average Call Time"
		Me.Label4.Size = New System.Drawing.Size(108, 17)
		Me.Label4.Location = New System.Drawing.Point(240, 360)
		Me.Label4.TabIndex = 20
		Me.Label4.Visible = False
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.AutoSize = True
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.Text = "Total Call Time"
		Me.Label3.Size = New System.Drawing.Size(88, 17)
		Me.Label3.Location = New System.Drawing.Point(40, 360)
		Me.Label3.TabIndex = 19
		Me.Label3.Visible = False
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.AutoSize = True
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label1.Text = "Call / BNB Notes Statistics"
		Me.Label1.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(302, 30)
		Me.Label1.Location = New System.Drawing.Point(40, 0)
		Me.Label1.TabIndex = 11
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = True
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.lblStart.Text = "Start Date"
		Me.lblStart.Size = New System.Drawing.Size(60, 17)
		Me.lblStart.Location = New System.Drawing.Point(120, 40)
		Me.lblStart.TabIndex = 10
		Me.lblStart.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblStart.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblStart.BackColor = System.Drawing.SystemColors.Control
		Me.lblStart.Enabled = True
		Me.lblStart.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblStart.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblStart.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblStart.UseMnemonic = True
		Me.lblStart.Visible = True
		Me.lblStart.AutoSize = True
		Me.lblStart.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblStart.Name = "lblStart"
		Me.lblEnd.Text = "End Date"
		Me.lblEnd.Size = New System.Drawing.Size(57, 17)
		Me.lblEnd.Location = New System.Drawing.Point(270, 40)
		Me.lblEnd.TabIndex = 9
		Me.lblEnd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblEnd.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblEnd.BackColor = System.Drawing.SystemColors.Control
		Me.lblEnd.Enabled = True
		Me.lblEnd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblEnd.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEnd.UseMnemonic = True
		Me.lblEnd.Visible = True
		Me.lblEnd.AutoSize = True
		Me.lblEnd.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblEnd.Name = "lblEnd"
		CType(Me.DTPicker1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DTPicker2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdCallData, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(ListView1)
		Me.Controls.Add(cboMins)
		Me.Controls.Add(chkUpdate)
		Me.Controls.Add(grdCallData)
		Me.Controls.Add(txtAvg)
		Me.Controls.Add(txtTotal)
		Me.Controls.Add(optDuration)
		Me.Controls.Add(optBNB)
		Me.Controls.Add(fraGroup)
		Me.Controls.Add(cmdGenReport)
		Me.Controls.Add(cmdPrintReport)
		Me.Controls.Add(DTPicker2)
		Me.Controls.Add(DTPicker1)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label1)
		Me.Controls.Add(lblStart)
		Me.Controls.Add(lblEnd)
		Me.fraGroup.Controls.Add(cboCallDir)
		Me.fraGroup.Controls.Add(cboGroup)
		Me.fraGroup.Controls.Add(optExt)
		Me.fraGroup.Controls.Add(optAll)
		Me.fraGroup.Controls.Add(lblCallType)
		Me.fraGroup.Controls.Add(Label2)
		Me.fraGroup.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class