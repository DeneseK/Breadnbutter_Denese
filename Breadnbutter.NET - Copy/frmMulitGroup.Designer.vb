<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMultiChart
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
	Public WithEvents optVoiceMail As System.Windows.Forms.RadioButton
	Public WithEvents optCalls As System.Windows.Forms.RadioButton
	Public WithEvents optBoth As System.Windows.Forms.RadioButton
	Public WithEvents fraCallType As System.Windows.Forms.GroupBox
	Public WithEvents chkLablePoints As System.Windows.Forms.CheckBox
	Public WithEvents cmdAvg As System.Windows.Forms.Button
	Public WithEvents MonthView1 As AxMSComCtl2.AxMonthView
	Public WithEvents cboChartType As System.Windows.Forms.ComboBox
	Public WithEvents optCallVsVoice As System.Windows.Forms.RadioButton
	Public WithEvents optNone As System.Windows.Forms.RadioButton
	Public WithEvents optDirection As System.Windows.Forms.RadioButton
	Public WithEvents optMultiDates As System.Windows.Forms.RadioButton
	Public WithEvents optMultiGroups As System.Windows.Forms.RadioButton
	Public WithEvents fraMultiLines As System.Windows.Forms.GroupBox
	Public WithEvents cmdChart As System.Windows.Forms.Button
	Public WithEvents cmdPrintChart As System.Windows.Forms.Button
	Public WithEvents _cboGroup_0 As System.Windows.Forms.ComboBox
	Public WithEvents optExt As System.Windows.Forms.RadioButton
	Public WithEvents optWorkgroup As System.Windows.Forms.RadioButton
	Public WithEvents fraGroup As System.Windows.Forms.GroupBox
	Public WithEvents optTotal As System.Windows.Forms.RadioButton
	Public WithEvents optAvg As System.Windows.Forms.RadioButton
	Public WithEvents DTPicker4 As AxMSComCtl2.AxDTPicker
	Public WithEvents DTPicker3 As AxMSComCtl2.AxDTPicker
	Public WithEvents cboCallDir As System.Windows.Forms.ComboBox
	Public WithEvents cboDateType As System.Windows.Forms.ComboBox
	Public WithEvents DTPicker2 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_0 As AxMSComCtl2.AxDTPicker
	Public WithEvents cboYear As System.Windows.Forms.ComboBox
	Public WithEvents cboSunday As System.Windows.Forms.ComboBox
	Public WithEvents cboMonth As System.Windows.Forms.ComboBox
	Public WithEvents lblCallType As System.Windows.Forms.Label
	Public WithEvents lblDateType As System.Windows.Forms.Label
	Public WithEvents lblEnd As System.Windows.Forms.Label
	Public WithEvents _lblStart_0 As System.Windows.Forms.Label
	Public WithEvents lblMon As System.Windows.Forms.Label
	Public WithEvents lblWeek As System.Windows.Forms.Label
	Public WithEvents lblYear As System.Windows.Forms.Label
	Public WithEvents fraVariables As System.Windows.Forms.GroupBox
	Public dlgCommonPrint As System.Windows.Forms.PrintDialog
	Public WithEvents cboNum As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_9 As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_8 As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_7 As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_6 As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_5 As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_4 As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_3 As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_2 As System.Windows.Forms.ComboBox
	Public WithEvents _cboGroup_1 As System.Windows.Forms.ComboBox
	Public WithEvents _lblExtGroup_8 As System.Windows.Forms.Label
	Public WithEvents _lblExtGroup_7 As System.Windows.Forms.Label
	Public WithEvents _lblExtGroup_6 As System.Windows.Forms.Label
	Public WithEvents _lblExtGroup_5 As System.Windows.Forms.Label
	Public WithEvents _lblExtGroup_4 As System.Windows.Forms.Label
	Public WithEvents _lblExtGroup_3 As System.Windows.Forms.Label
	Public WithEvents _lblExtGroup_2 As System.Windows.Forms.Label
	Public WithEvents _lblExtGroup_1 As System.Windows.Forms.Label
	Public WithEvents _lblExtGroup_0 As System.Windows.Forms.Label
	Public WithEvents lblGroupNum As System.Windows.Forms.Label
	Public WithEvents fraGroups As System.Windows.Forms.GroupBox
	Public WithEvents cboDateNum As System.Windows.Forms.ComboBox
	Public WithEvents _DTPicker1_1 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_2 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_3 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_4 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_5 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_6 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_7 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_8 As AxMSComCtl2.AxDTPicker
	Public WithEvents _DTPicker1_9 As AxMSComCtl2.AxDTPicker
	Public WithEvents lblLines As System.Windows.Forms.Label
	Public WithEvents _lblStart_9 As System.Windows.Forms.Label
	Public WithEvents _lblStart_8 As System.Windows.Forms.Label
	Public WithEvents _lblStart_7 As System.Windows.Forms.Label
	Public WithEvents _lblStart_6 As System.Windows.Forms.Label
	Public WithEvents _lblStart_5 As System.Windows.Forms.Label
	Public WithEvents _lblStart_4 As System.Windows.Forms.Label
	Public WithEvents _lblStart_3 As System.Windows.Forms.Label
	Public WithEvents _lblStart_2 As System.Windows.Forms.Label
	Public WithEvents _lblStart_1 As System.Windows.Forms.Label
	Public WithEvents fraMultiDates As System.Windows.Forms.GroupBox
	Public WithEvents MSChart1 As AxMSChart20Lib.AxMSChart
	Public WithEvents lblChartType As System.Windows.Forms.Label
	Public WithEvents DTPicker1 As AxDTPickerArray.AxDTPickerArray
	Public WithEvents cboGroup As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
	Public WithEvents lblExtGroup As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblStart As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMultiChart))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.fraCallType = New System.Windows.Forms.GroupBox
		Me.optVoiceMail = New System.Windows.Forms.RadioButton
		Me.optCalls = New System.Windows.Forms.RadioButton
		Me.optBoth = New System.Windows.Forms.RadioButton
		Me.chkLablePoints = New System.Windows.Forms.CheckBox
		Me.cmdAvg = New System.Windows.Forms.Button
		Me.MonthView1 = New AxMSComCtl2.AxMonthView
		Me.cboChartType = New System.Windows.Forms.ComboBox
		Me.fraMultiLines = New System.Windows.Forms.GroupBox
		Me.optCallVsVoice = New System.Windows.Forms.RadioButton
		Me.optNone = New System.Windows.Forms.RadioButton
		Me.optDirection = New System.Windows.Forms.RadioButton
		Me.optMultiDates = New System.Windows.Forms.RadioButton
		Me.optMultiGroups = New System.Windows.Forms.RadioButton
		Me.cmdChart = New System.Windows.Forms.Button
		Me.cmdPrintChart = New System.Windows.Forms.Button
		Me.fraGroup = New System.Windows.Forms.GroupBox
		Me._cboGroup_0 = New System.Windows.Forms.ComboBox
		Me.optExt = New System.Windows.Forms.RadioButton
		Me.optWorkgroup = New System.Windows.Forms.RadioButton
		Me.fraVariables = New System.Windows.Forms.GroupBox
		Me.optTotal = New System.Windows.Forms.RadioButton
		Me.optAvg = New System.Windows.Forms.RadioButton
		Me.DTPicker4 = New AxMSComCtl2.AxDTPicker
		Me.DTPicker3 = New AxMSComCtl2.AxDTPicker
		Me.cboCallDir = New System.Windows.Forms.ComboBox
		Me.cboDateType = New System.Windows.Forms.ComboBox
		Me.DTPicker2 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_0 = New AxMSComCtl2.AxDTPicker
		Me.cboYear = New System.Windows.Forms.ComboBox
		Me.cboSunday = New System.Windows.Forms.ComboBox
		Me.cboMonth = New System.Windows.Forms.ComboBox
		Me.lblCallType = New System.Windows.Forms.Label
		Me.lblDateType = New System.Windows.Forms.Label
		Me.lblEnd = New System.Windows.Forms.Label
		Me._lblStart_0 = New System.Windows.Forms.Label
		Me.lblMon = New System.Windows.Forms.Label
		Me.lblWeek = New System.Windows.Forms.Label
		Me.lblYear = New System.Windows.Forms.Label
		Me.dlgCommonPrint = New System.Windows.Forms.PrintDialog
		Me.dlgCommonPrint.PrinterSettings = New System.Drawing.Printing.PrinterSettings
		Me.fraGroups = New System.Windows.Forms.GroupBox
		Me.cboNum = New System.Windows.Forms.ComboBox
		Me._cboGroup_9 = New System.Windows.Forms.ComboBox
		Me._cboGroup_8 = New System.Windows.Forms.ComboBox
		Me._cboGroup_7 = New System.Windows.Forms.ComboBox
		Me._cboGroup_6 = New System.Windows.Forms.ComboBox
		Me._cboGroup_5 = New System.Windows.Forms.ComboBox
		Me._cboGroup_4 = New System.Windows.Forms.ComboBox
		Me._cboGroup_3 = New System.Windows.Forms.ComboBox
		Me._cboGroup_2 = New System.Windows.Forms.ComboBox
		Me._cboGroup_1 = New System.Windows.Forms.ComboBox
		Me._lblExtGroup_8 = New System.Windows.Forms.Label
		Me._lblExtGroup_7 = New System.Windows.Forms.Label
		Me._lblExtGroup_6 = New System.Windows.Forms.Label
		Me._lblExtGroup_5 = New System.Windows.Forms.Label
		Me._lblExtGroup_4 = New System.Windows.Forms.Label
		Me._lblExtGroup_3 = New System.Windows.Forms.Label
		Me._lblExtGroup_2 = New System.Windows.Forms.Label
		Me._lblExtGroup_1 = New System.Windows.Forms.Label
		Me._lblExtGroup_0 = New System.Windows.Forms.Label
		Me.lblGroupNum = New System.Windows.Forms.Label
		Me.fraMultiDates = New System.Windows.Forms.GroupBox
		Me.cboDateNum = New System.Windows.Forms.ComboBox
		Me._DTPicker1_1 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_2 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_3 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_4 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_5 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_6 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_7 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_8 = New AxMSComCtl2.AxDTPicker
		Me._DTPicker1_9 = New AxMSComCtl2.AxDTPicker
		Me.lblLines = New System.Windows.Forms.Label
		Me._lblStart_9 = New System.Windows.Forms.Label
		Me._lblStart_8 = New System.Windows.Forms.Label
		Me._lblStart_7 = New System.Windows.Forms.Label
		Me._lblStart_6 = New System.Windows.Forms.Label
		Me._lblStart_5 = New System.Windows.Forms.Label
		Me._lblStart_4 = New System.Windows.Forms.Label
		Me._lblStart_3 = New System.Windows.Forms.Label
		Me._lblStart_2 = New System.Windows.Forms.Label
		Me._lblStart_1 = New System.Windows.Forms.Label
		Me.MSChart1 = New AxMSChart20Lib.AxMSChart
		Me.lblChartType = New System.Windows.Forms.Label
		Me.DTPicker1 = New AxDTPickerArray.AxDTPickerArray(components)
		Me.cboGroup = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(components)
		Me.lblExtGroup = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblStart = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.fraCallType.SuspendLayout()
		Me.fraMultiLines.SuspendLayout()
		Me.fraGroup.SuspendLayout()
		Me.fraVariables.SuspendLayout()
		Me.fraGroups.SuspendLayout()
		Me.fraMultiDates.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.MonthView1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DTPicker4, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DTPicker3, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DTPicker2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_3, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_4, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_5, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_6, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_7, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_8, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._DTPicker1_9, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MSChart1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DTPicker1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cboGroup, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblExtGroup, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblStart, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Multi-Line Chart"
		Me.ClientSize = New System.Drawing.Size(1048, 734)
		Me.Location = New System.Drawing.Point(172, 158)
		Me.Icon = CType(resources.GetObject("frmMultiChart.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmMultiChart"
		Me.fraCallType.Text = "Calls and V-Mail"
		Me.fraCallType.Size = New System.Drawing.Size(132, 152)
		Me.fraCallType.Location = New System.Drawing.Point(150, 10)
		Me.fraCallType.TabIndex = 79
		Me.fraCallType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraCallType.BackColor = System.Drawing.SystemColors.Control
		Me.fraCallType.Enabled = True
		Me.fraCallType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraCallType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraCallType.Visible = True
		Me.fraCallType.Name = "fraCallType"
		Me.optVoiceMail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optVoiceMail.Text = "V-Mail Only"
		Me.optVoiceMail.Size = New System.Drawing.Size(102, 22)
		Me.optVoiceMail.Location = New System.Drawing.Point(20, 90)
		Me.optVoiceMail.TabIndex = 82
		Me.optVoiceMail.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optVoiceMail.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optVoiceMail.BackColor = System.Drawing.SystemColors.Control
		Me.optVoiceMail.CausesValidation = True
		Me.optVoiceMail.Enabled = True
		Me.optVoiceMail.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optVoiceMail.Cursor = System.Windows.Forms.Cursors.Default
		Me.optVoiceMail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optVoiceMail.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optVoiceMail.TabStop = True
		Me.optVoiceMail.Checked = False
		Me.optVoiceMail.Visible = True
		Me.optVoiceMail.Name = "optVoiceMail"
		Me.optCalls.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCalls.Text = "Calls Only"
		Me.optCalls.Size = New System.Drawing.Size(92, 22)
		Me.optCalls.Location = New System.Drawing.Point(20, 60)
		Me.optCalls.TabIndex = 81
		Me.optCalls.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optCalls.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCalls.BackColor = System.Drawing.SystemColors.Control
		Me.optCalls.CausesValidation = True
		Me.optCalls.Enabled = True
		Me.optCalls.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optCalls.Cursor = System.Windows.Forms.Cursors.Default
		Me.optCalls.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optCalls.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optCalls.TabStop = True
		Me.optCalls.Checked = False
		Me.optCalls.Visible = True
		Me.optCalls.Name = "optCalls"
		Me.optBoth.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optBoth.Text = "Both"
		Me.optBoth.Size = New System.Drawing.Size(92, 22)
		Me.optBoth.Location = New System.Drawing.Point(20, 30)
		Me.optBoth.TabIndex = 80
		Me.optBoth.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optBoth.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optBoth.BackColor = System.Drawing.SystemColors.Control
		Me.optBoth.CausesValidation = True
		Me.optBoth.Enabled = True
		Me.optBoth.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optBoth.Cursor = System.Windows.Forms.Cursors.Default
		Me.optBoth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optBoth.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optBoth.TabStop = True
		Me.optBoth.Checked = False
		Me.optBoth.Visible = True
		Me.optBoth.Name = "optBoth"
		Me.chkLablePoints.Text = "Label Points"
		Me.chkLablePoints.Size = New System.Drawing.Size(112, 22)
		Me.chkLablePoints.Location = New System.Drawing.Point(750, 60)
		Me.chkLablePoints.TabIndex = 78
		Me.chkLablePoints.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLablePoints.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLablePoints.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLablePoints.BackColor = System.Drawing.SystemColors.Control
		Me.chkLablePoints.CausesValidation = True
		Me.chkLablePoints.Enabled = True
		Me.chkLablePoints.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLablePoints.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLablePoints.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLablePoints.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLablePoints.TabStop = True
		Me.chkLablePoints.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLablePoints.Visible = True
		Me.chkLablePoints.Name = "chkLablePoints"
		Me.cmdAvg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAvg.Text = "CHART IT!"
		Me.cmdAvg.Size = New System.Drawing.Size(112, 27)
		Me.cmdAvg.Location = New System.Drawing.Point(750, 90)
		Me.cmdAvg.TabIndex = 73
		Me.cmdAvg.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAvg.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAvg.CausesValidation = True
		Me.cmdAvg.Enabled = True
		Me.cmdAvg.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAvg.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAvg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAvg.TabStop = True
		Me.cmdAvg.Name = "cmdAvg"
		MonthView1.OcxState = CType(resources.GetObject("MonthView1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.MonthView1.Size = New System.Drawing.Size(254, 193)
		Me.MonthView1.Location = New System.Drawing.Point(550, 340)
		Me.MonthView1.TabIndex = 66
		Me.MonthView1.Visible = False
		Me.MonthView1.Name = "MonthView1"
		Me.cboChartType.Size = New System.Drawing.Size(112, 27)
		Me.cboChartType.Location = New System.Drawing.Point(750, 30)
		Me.cboChartType.Items.AddRange(New Object(){"2D Line", "3D Line", "2D Bar", "3D Bar", "2D Area", "3D Area", "2D Step", "3D Step", "2D Combo", "3D Combo"})
		Me.cboChartType.TabIndex = 52
		Me.cboChartType.Text = "2D Line"
		Me.cboChartType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboChartType.BackColor = System.Drawing.SystemColors.Window
		Me.cboChartType.CausesValidation = True
		Me.cboChartType.Enabled = True
		Me.cboChartType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboChartType.IntegralHeight = True
		Me.cboChartType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboChartType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboChartType.Sorted = False
		Me.cboChartType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboChartType.TabStop = True
		Me.cboChartType.Visible = True
		Me.cboChartType.Name = "cboChartType"
		Me.fraMultiLines.Text = "Multiple Lines"
		Me.fraMultiLines.Size = New System.Drawing.Size(212, 152)
		Me.fraMultiLines.Location = New System.Drawing.Point(530, 10)
		Me.fraMultiLines.TabIndex = 20
		Me.fraMultiLines.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraMultiLines.BackColor = System.Drawing.SystemColors.Control
		Me.fraMultiLines.Enabled = True
		Me.fraMultiLines.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraMultiLines.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraMultiLines.Visible = True
		Me.fraMultiLines.Name = "fraMultiLines"
		Me.optCallVsVoice.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCallVsVoice.Text = "Calls vs. V-Mails"
		Me.optCallVsVoice.Size = New System.Drawing.Size(162, 22)
		Me.optCallVsVoice.Location = New System.Drawing.Point(20, 110)
		Me.optCallVsVoice.TabIndex = 77
		Me.optCallVsVoice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optCallVsVoice.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCallVsVoice.BackColor = System.Drawing.SystemColors.Control
		Me.optCallVsVoice.CausesValidation = True
		Me.optCallVsVoice.Enabled = True
		Me.optCallVsVoice.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optCallVsVoice.Cursor = System.Windows.Forms.Cursors.Default
		Me.optCallVsVoice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optCallVsVoice.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optCallVsVoice.TabStop = True
		Me.optCallVsVoice.Checked = False
		Me.optCallVsVoice.Visible = True
		Me.optCallVsVoice.Name = "optCallVsVoice"
		Me.optNone.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNone.Text = "None"
		Me.optNone.Size = New System.Drawing.Size(132, 22)
		Me.optNone.Location = New System.Drawing.Point(20, 30)
		Me.optNone.TabIndex = 24
		Me.optNone.Checked = True
		Me.optNone.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optNone.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNone.BackColor = System.Drawing.SystemColors.Control
		Me.optNone.CausesValidation = True
		Me.optNone.Enabled = True
		Me.optNone.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optNone.Cursor = System.Windows.Forms.Cursors.Default
		Me.optNone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optNone.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optNone.TabStop = True
		Me.optNone.Visible = True
		Me.optNone.Name = "optNone"
		Me.optDirection.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDirection.Text = "All Call Directions"
		Me.optDirection.Size = New System.Drawing.Size(142, 22)
		Me.optDirection.Location = New System.Drawing.Point(20, 90)
		Me.optDirection.TabIndex = 23
		Me.optDirection.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optDirection.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDirection.BackColor = System.Drawing.SystemColors.Control
		Me.optDirection.CausesValidation = True
		Me.optDirection.Enabled = True
		Me.optDirection.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optDirection.Cursor = System.Windows.Forms.Cursors.Default
		Me.optDirection.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optDirection.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optDirection.TabStop = True
		Me.optDirection.Checked = False
		Me.optDirection.Visible = True
		Me.optDirection.Name = "optDirection"
		Me.optMultiDates.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optMultiDates.Text = "Multiple Dates"
		Me.optMultiDates.Size = New System.Drawing.Size(162, 22)
		Me.optMultiDates.Location = New System.Drawing.Point(20, 70)
		Me.optMultiDates.TabIndex = 22
		Me.optMultiDates.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optMultiDates.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optMultiDates.BackColor = System.Drawing.SystemColors.Control
		Me.optMultiDates.CausesValidation = True
		Me.optMultiDates.Enabled = True
		Me.optMultiDates.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optMultiDates.Cursor = System.Windows.Forms.Cursors.Default
		Me.optMultiDates.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optMultiDates.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optMultiDates.TabStop = True
		Me.optMultiDates.Checked = False
		Me.optMultiDates.Visible = True
		Me.optMultiDates.Name = "optMultiDates"
		Me.optMultiGroups.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optMultiGroups.Text = "Multiple Ext#s or Group#s"
		Me.optMultiGroups.Size = New System.Drawing.Size(182, 22)
		Me.optMultiGroups.Location = New System.Drawing.Point(20, 50)
		Me.optMultiGroups.TabIndex = 21
		Me.optMultiGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optMultiGroups.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optMultiGroups.BackColor = System.Drawing.SystemColors.Control
		Me.optMultiGroups.CausesValidation = True
		Me.optMultiGroups.Enabled = True
		Me.optMultiGroups.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optMultiGroups.Cursor = System.Windows.Forms.Cursors.Default
		Me.optMultiGroups.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optMultiGroups.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optMultiGroups.TabStop = True
		Me.optMultiGroups.Checked = False
		Me.optMultiGroups.Visible = True
		Me.optMultiGroups.Name = "optMultiGroups"
		Me.cmdChart.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdChart.Text = "CHART IT!"
		Me.cmdChart.Size = New System.Drawing.Size(112, 27)
		Me.cmdChart.Location = New System.Drawing.Point(750, 90)
		Me.cmdChart.TabIndex = 19
		Me.cmdChart.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdChart.BackColor = System.Drawing.SystemColors.Control
		Me.cmdChart.CausesValidation = True
		Me.cmdChart.Enabled = True
		Me.cmdChart.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdChart.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdChart.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdChart.TabStop = True
		Me.cmdChart.Name = "cmdChart"
		Me.cmdPrintChart.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPrintChart.Text = "PRINT CHART!"
		Me.cmdPrintChart.Enabled = False
		Me.cmdPrintChart.Size = New System.Drawing.Size(112, 27)
		Me.cmdPrintChart.Location = New System.Drawing.Point(750, 130)
		Me.cmdPrintChart.TabIndex = 18
		Me.cmdPrintChart.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPrintChart.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPrintChart.CausesValidation = True
		Me.cmdPrintChart.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPrintChart.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPrintChart.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPrintChart.TabStop = True
		Me.cmdPrintChart.Name = "cmdPrintChart"
		Me.fraGroup.Text = "Ext# or Group#"
		Me.fraGroup.Size = New System.Drawing.Size(132, 152)
		Me.fraGroup.Location = New System.Drawing.Point(10, 10)
		Me.fraGroup.TabIndex = 10
		Me.fraGroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraGroup.BackColor = System.Drawing.SystemColors.Control
		Me.fraGroup.Enabled = True
		Me.fraGroup.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraGroup.Visible = True
		Me.fraGroup.Name = "fraGroup"
		Me._cboGroup_0.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_0.Location = New System.Drawing.Point(30, 90)
		Me._cboGroup_0.TabIndex = 17
		Me._cboGroup_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_0.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_0.CausesValidation = True
		Me._cboGroup_0.Enabled = True
		Me._cboGroup_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_0.IntegralHeight = True
		Me._cboGroup_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_0.Sorted = False
		Me._cboGroup_0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_0.TabStop = True
		Me._cboGroup_0.Visible = True
		Me._cboGroup_0.Name = "_cboGroup_0"
		Me.optExt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optExt.Text = "by Ext"
		Me.optExt.Size = New System.Drawing.Size(82, 22)
		Me.optExt.Location = New System.Drawing.Point(10, 30)
		Me.optExt.TabIndex = 12
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
		Me.optExt.Visible = True
		Me.optExt.Name = "optExt"
		Me.optWorkgroup.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optWorkgroup.Text = "by Workgroup"
		Me.optWorkgroup.Size = New System.Drawing.Size(112, 22)
		Me.optWorkgroup.Location = New System.Drawing.Point(10, 60)
		Me.optWorkgroup.TabIndex = 11
		Me.optWorkgroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optWorkgroup.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optWorkgroup.BackColor = System.Drawing.SystemColors.Control
		Me.optWorkgroup.CausesValidation = True
		Me.optWorkgroup.Enabled = True
		Me.optWorkgroup.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optWorkgroup.Cursor = System.Windows.Forms.Cursors.Default
		Me.optWorkgroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optWorkgroup.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optWorkgroup.TabStop = True
		Me.optWorkgroup.Checked = False
		Me.optWorkgroup.Visible = True
		Me.optWorkgroup.Name = "optWorkgroup"
		Me.fraVariables.Text = "Variables"
		Me.fraVariables.Size = New System.Drawing.Size(232, 152)
		Me.fraVariables.Location = New System.Drawing.Point(290, 10)
		Me.fraVariables.TabIndex = 1
		Me.fraVariables.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraVariables.BackColor = System.Drawing.SystemColors.Control
		Me.fraVariables.Enabled = True
		Me.fraVariables.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraVariables.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraVariables.Visible = True
		Me.fraVariables.Name = "fraVariables"
		Me.optTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optTotal.Text = "Total Count"
		Me.optTotal.Size = New System.Drawing.Size(102, 22)
		Me.optTotal.Location = New System.Drawing.Point(110, 120)
		Me.optTotal.TabIndex = 75
		Me.optTotal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optTotal.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optTotal.BackColor = System.Drawing.SystemColors.Control
		Me.optTotal.CausesValidation = True
		Me.optTotal.Enabled = True
		Me.optTotal.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optTotal.Cursor = System.Windows.Forms.Cursors.Default
		Me.optTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optTotal.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optTotal.TabStop = True
		Me.optTotal.Checked = False
		Me.optTotal.Visible = True
		Me.optTotal.Name = "optTotal"
		Me.optAvg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optAvg.Text = "Average"
		Me.optAvg.Size = New System.Drawing.Size(82, 22)
		Me.optAvg.Location = New System.Drawing.Point(10, 120)
		Me.optAvg.TabIndex = 74
		Me.optAvg.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optAvg.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optAvg.BackColor = System.Drawing.SystemColors.Control
		Me.optAvg.CausesValidation = True
		Me.optAvg.Enabled = True
		Me.optAvg.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optAvg.Cursor = System.Windows.Forms.Cursors.Default
		Me.optAvg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optAvg.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optAvg.TabStop = True
		Me.optAvg.Checked = False
		Me.optAvg.Visible = True
		Me.optAvg.Name = "optAvg"
		DTPicker4.OcxState = CType(resources.GetObject("DTPicker4.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DTPicker4.Size = New System.Drawing.Size(112, 27)
		Me.DTPicker4.Location = New System.Drawing.Point(110, 90)
		Me.DTPicker4.TabIndex = 72
		Me.DTPicker4.Visible = False
		Me.DTPicker4.Name = "DTPicker4"
		DTPicker3.OcxState = CType(resources.GetObject("DTPicker3.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DTPicker3.Size = New System.Drawing.Size(112, 27)
		Me.DTPicker3.Location = New System.Drawing.Point(110, 40)
		Me.DTPicker3.TabIndex = 71
		Me.DTPicker3.Visible = False
		Me.DTPicker3.Name = "DTPicker3"
		Me.cboCallDir.Size = New System.Drawing.Size(92, 27)
		Me.cboCallDir.Location = New System.Drawing.Point(10, 90)
		Me.cboCallDir.Items.AddRange(New Object(){"Incoming", "Outgoing", "Both"})
		Me.cboCallDir.TabIndex = 3
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
		Me.cboDateType.Size = New System.Drawing.Size(92, 27)
		Me.cboDateType.Location = New System.Drawing.Point(10, 40)
		Me.cboDateType.Items.AddRange(New Object(){"Hour", "Day", "Week", "Month"})
		Me.cboDateType.TabIndex = 2
		Me.cboDateType.Text = "Hour"
		Me.cboDateType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboDateType.BackColor = System.Drawing.SystemColors.Window
		Me.cboDateType.CausesValidation = True
		Me.cboDateType.Enabled = True
		Me.cboDateType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboDateType.IntegralHeight = True
		Me.cboDateType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboDateType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboDateType.Sorted = False
		Me.cboDateType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboDateType.TabStop = True
		Me.cboDateType.Visible = True
		Me.cboDateType.Name = "cboDateType"
		DTPicker2.OcxState = CType(resources.GetObject("DTPicker2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DTPicker2.Size = New System.Drawing.Size(112, 27)
		Me.DTPicker2.Location = New System.Drawing.Point(110, 90)
		Me.DTPicker2.TabIndex = 4
		Me.DTPicker2.Name = "DTPicker2"
		_DTPicker1_0.OcxState = CType(resources.GetObject("_DTPicker1_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_0.Size = New System.Drawing.Size(112, 27)
		Me._DTPicker1_0.Location = New System.Drawing.Point(110, 40)
		Me._DTPicker1_0.TabIndex = 5
		Me._DTPicker1_0.Name = "_DTPicker1_0"
		Me.cboYear.Size = New System.Drawing.Size(62, 27)
		Me.cboYear.Location = New System.Drawing.Point(110, 90)
		Me.cboYear.TabIndex = 64
		Me.cboYear.Text = "2003"
		Me.cboYear.Visible = False
		Me.cboYear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboYear.BackColor = System.Drawing.SystemColors.Window
		Me.cboYear.CausesValidation = True
		Me.cboYear.Enabled = True
		Me.cboYear.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboYear.IntegralHeight = True
		Me.cboYear.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboYear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboYear.Sorted = False
		Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboYear.TabStop = True
		Me.cboYear.Name = "cboYear"
		Me.cboSunday.Size = New System.Drawing.Size(52, 27)
		Me.cboSunday.Location = New System.Drawing.Point(240, 90)
		Me.cboSunday.TabIndex = 65
		Me.cboSunday.Text = "01"
		Me.cboSunday.Visible = False
		Me.cboSunday.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSunday.BackColor = System.Drawing.SystemColors.Window
		Me.cboSunday.CausesValidation = True
		Me.cboSunday.Enabled = True
		Me.cboSunday.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSunday.IntegralHeight = True
		Me.cboSunday.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSunday.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSunday.Sorted = False
		Me.cboSunday.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSunday.TabStop = True
		Me.cboSunday.Name = "cboSunday"
		Me.cboMonth.Size = New System.Drawing.Size(52, 27)
		Me.cboMonth.Location = New System.Drawing.Point(180, 90)
		Me.cboMonth.TabIndex = 69
		Me.cboMonth.Text = "01"
		Me.cboMonth.Visible = False
		Me.cboMonth.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboMonth.BackColor = System.Drawing.SystemColors.Window
		Me.cboMonth.CausesValidation = True
		Me.cboMonth.Enabled = True
		Me.cboMonth.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboMonth.IntegralHeight = True
		Me.cboMonth.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboMonth.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboMonth.Sorted = False
		Me.cboMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboMonth.TabStop = True
		Me.cboMonth.Name = "cboMonth"
		Me.lblCallType.Text = "Call Direction"
		Me.lblCallType.Size = New System.Drawing.Size(78, 17)
		Me.lblCallType.Location = New System.Drawing.Point(10, 70)
		Me.lblCallType.TabIndex = 9
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
		Me.lblDateType.Text = "By"
		Me.lblDateType.Size = New System.Drawing.Size(15, 17)
		Me.lblDateType.Location = New System.Drawing.Point(10, 20)
		Me.lblDateType.TabIndex = 6
		Me.lblDateType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDateType.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDateType.BackColor = System.Drawing.SystemColors.Control
		Me.lblDateType.Enabled = True
		Me.lblDateType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDateType.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDateType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDateType.UseMnemonic = True
		Me.lblDateType.Visible = True
		Me.lblDateType.AutoSize = True
		Me.lblDateType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDateType.Name = "lblDateType"
		Me.lblEnd.Text = "End Date"
		Me.lblEnd.Size = New System.Drawing.Size(57, 17)
		Me.lblEnd.Location = New System.Drawing.Point(110, 70)
		Me.lblEnd.TabIndex = 8
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
		Me._lblStart_0.Text = "Start Date"
		Me._lblStart_0.Size = New System.Drawing.Size(60, 17)
		Me._lblStart_0.Location = New System.Drawing.Point(110, 20)
		Me._lblStart_0.TabIndex = 7
		Me._lblStart_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_0.Enabled = True
		Me._lblStart_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_0.UseMnemonic = True
		Me._lblStart_0.Visible = True
		Me._lblStart_0.AutoSize = True
		Me._lblStart_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_0.Name = "_lblStart_0"
		Me.lblMon.Text = "Month"
		Me.lblMon.Size = New System.Drawing.Size(52, 22)
		Me.lblMon.Location = New System.Drawing.Point(180, 70)
		Me.lblMon.TabIndex = 70
		Me.lblMon.Visible = False
		Me.lblMon.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMon.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblMon.BackColor = System.Drawing.SystemColors.Control
		Me.lblMon.Enabled = True
		Me.lblMon.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblMon.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblMon.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblMon.UseMnemonic = True
		Me.lblMon.AutoSize = False
		Me.lblMon.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblMon.Name = "lblMon"
		Me.lblWeek.Text = "Day"
		Me.lblWeek.Size = New System.Drawing.Size(24, 17)
		Me.lblWeek.Location = New System.Drawing.Point(240, 70)
		Me.lblWeek.TabIndex = 68
		Me.lblWeek.Visible = False
		Me.lblWeek.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblWeek.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblWeek.BackColor = System.Drawing.SystemColors.Control
		Me.lblWeek.Enabled = True
		Me.lblWeek.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblWeek.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblWeek.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblWeek.UseMnemonic = True
		Me.lblWeek.AutoSize = True
		Me.lblWeek.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblWeek.Name = "lblWeek"
		Me.lblYear.Text = "Year"
		Me.lblYear.Size = New System.Drawing.Size(28, 17)
		Me.lblYear.Location = New System.Drawing.Point(110, 70)
		Me.lblYear.TabIndex = 67
		Me.lblYear.Visible = False
		Me.lblYear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblYear.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblYear.BackColor = System.Drawing.SystemColors.Control
		Me.lblYear.Enabled = True
		Me.lblYear.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblYear.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblYear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblYear.UseMnemonic = True
		Me.lblYear.AutoSize = True
		Me.lblYear.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblYear.Name = "lblYear"
		Me.fraGroups.Text = "Ext#s or Group#s"
		Me.fraGroups.Size = New System.Drawing.Size(172, 712)
		Me.fraGroups.Location = New System.Drawing.Point(870, 10)
		Me.fraGroups.TabIndex = 0
		Me.fraGroups.Visible = False
		Me.fraGroups.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraGroups.BackColor = System.Drawing.SystemColors.Control
		Me.fraGroups.Enabled = True
		Me.fraGroups.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraGroups.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraGroups.Name = "fraGroups"
		Me.cboNum.Size = New System.Drawing.Size(52, 27)
		Me.cboNum.Location = New System.Drawing.Point(110, 50)
		Me.cboNum.Items.AddRange(New Object(){"2", "3", "4", "5", "6", "7", "8", "9", "10"})
		Me.cboNum.TabIndex = 30
		Me.cboNum.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboNum.BackColor = System.Drawing.SystemColors.Window
		Me.cboNum.CausesValidation = True
		Me.cboNum.Enabled = True
		Me.cboNum.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboNum.IntegralHeight = True
		Me.cboNum.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboNum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboNum.Sorted = False
		Me.cboNum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboNum.TabStop = True
		Me.cboNum.Visible = True
		Me.cboNum.Name = "cboNum"
		Me._cboGroup_9.Enabled = False
		Me._cboGroup_9.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_9.Location = New System.Drawing.Point(60, 610)
		Me._cboGroup_9.TabIndex = 29
		Me._cboGroup_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_9.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_9.CausesValidation = True
		Me._cboGroup_9.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_9.IntegralHeight = True
		Me._cboGroup_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_9.Sorted = False
		Me._cboGroup_9.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_9.TabStop = True
		Me._cboGroup_9.Visible = True
		Me._cboGroup_9.Name = "_cboGroup_9"
		Me._cboGroup_8.Enabled = False
		Me._cboGroup_8.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_8.Location = New System.Drawing.Point(60, 550)
		Me._cboGroup_8.TabIndex = 28
		Me._cboGroup_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_8.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_8.CausesValidation = True
		Me._cboGroup_8.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_8.IntegralHeight = True
		Me._cboGroup_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_8.Sorted = False
		Me._cboGroup_8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_8.TabStop = True
		Me._cboGroup_8.Visible = True
		Me._cboGroup_8.Name = "_cboGroup_8"
		Me._cboGroup_7.Enabled = False
		Me._cboGroup_7.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_7.Location = New System.Drawing.Point(60, 490)
		Me._cboGroup_7.TabIndex = 27
		Me._cboGroup_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_7.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_7.CausesValidation = True
		Me._cboGroup_7.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_7.IntegralHeight = True
		Me._cboGroup_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_7.Sorted = False
		Me._cboGroup_7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_7.TabStop = True
		Me._cboGroup_7.Visible = True
		Me._cboGroup_7.Name = "_cboGroup_7"
		Me._cboGroup_6.Enabled = False
		Me._cboGroup_6.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_6.Location = New System.Drawing.Point(60, 430)
		Me._cboGroup_6.TabIndex = 26
		Me._cboGroup_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_6.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_6.CausesValidation = True
		Me._cboGroup_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_6.IntegralHeight = True
		Me._cboGroup_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_6.Sorted = False
		Me._cboGroup_6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_6.TabStop = True
		Me._cboGroup_6.Visible = True
		Me._cboGroup_6.Name = "_cboGroup_6"
		Me._cboGroup_5.Enabled = False
		Me._cboGroup_5.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_5.Location = New System.Drawing.Point(60, 370)
		Me._cboGroup_5.TabIndex = 25
		Me._cboGroup_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_5.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_5.CausesValidation = True
		Me._cboGroup_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_5.IntegralHeight = True
		Me._cboGroup_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_5.Sorted = False
		Me._cboGroup_5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_5.TabStop = True
		Me._cboGroup_5.Visible = True
		Me._cboGroup_5.Name = "_cboGroup_5"
		Me._cboGroup_4.Enabled = False
		Me._cboGroup_4.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_4.Location = New System.Drawing.Point(60, 310)
		Me._cboGroup_4.TabIndex = 16
		Me._cboGroup_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_4.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_4.CausesValidation = True
		Me._cboGroup_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_4.IntegralHeight = True
		Me._cboGroup_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_4.Sorted = False
		Me._cboGroup_4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_4.TabStop = True
		Me._cboGroup_4.Visible = True
		Me._cboGroup_4.Name = "_cboGroup_4"
		Me._cboGroup_3.Enabled = False
		Me._cboGroup_3.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_3.Location = New System.Drawing.Point(60, 250)
		Me._cboGroup_3.TabIndex = 15
		Me._cboGroup_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_3.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_3.CausesValidation = True
		Me._cboGroup_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_3.IntegralHeight = True
		Me._cboGroup_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_3.Sorted = False
		Me._cboGroup_3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_3.TabStop = True
		Me._cboGroup_3.Visible = True
		Me._cboGroup_3.Name = "_cboGroup_3"
		Me._cboGroup_2.Enabled = False
		Me._cboGroup_2.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_2.Location = New System.Drawing.Point(60, 190)
		Me._cboGroup_2.TabIndex = 14
		Me._cboGroup_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_2.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_2.CausesValidation = True
		Me._cboGroup_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_2.IntegralHeight = True
		Me._cboGroup_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_2.Sorted = False
		Me._cboGroup_2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_2.TabStop = True
		Me._cboGroup_2.Visible = True
		Me._cboGroup_2.Name = "_cboGroup_2"
		Me._cboGroup_1.Enabled = False
		Me._cboGroup_1.Size = New System.Drawing.Size(72, 27)
		Me._cboGroup_1.Location = New System.Drawing.Point(60, 130)
		Me._cboGroup_1.TabIndex = 13
		Me._cboGroup_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cboGroup_1.BackColor = System.Drawing.SystemColors.Window
		Me._cboGroup_1.CausesValidation = True
		Me._cboGroup_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cboGroup_1.IntegralHeight = True
		Me._cboGroup_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cboGroup_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cboGroup_1.Sorted = False
		Me._cboGroup_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me._cboGroup_1.TabStop = True
		Me._cboGroup_1.Visible = True
		Me._cboGroup_1.Name = "_cboGroup_1"
		Me._lblExtGroup_8.Text = "Ext or Group#"
		Me._lblExtGroup_8.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_8.Location = New System.Drawing.Point(50, 590)
		Me._lblExtGroup_8.TabIndex = 63
		Me._lblExtGroup_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_8.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_8.Enabled = True
		Me._lblExtGroup_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_8.UseMnemonic = True
		Me._lblExtGroup_8.Visible = True
		Me._lblExtGroup_8.AutoSize = True
		Me._lblExtGroup_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_8.Name = "_lblExtGroup_8"
		Me._lblExtGroup_7.Text = "Ext or Group#"
		Me._lblExtGroup_7.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_7.Location = New System.Drawing.Point(50, 530)
		Me._lblExtGroup_7.TabIndex = 62
		Me._lblExtGroup_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_7.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_7.Enabled = True
		Me._lblExtGroup_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_7.UseMnemonic = True
		Me._lblExtGroup_7.Visible = True
		Me._lblExtGroup_7.AutoSize = True
		Me._lblExtGroup_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_7.Name = "_lblExtGroup_7"
		Me._lblExtGroup_6.Text = "Ext or Group#"
		Me._lblExtGroup_6.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_6.Location = New System.Drawing.Point(50, 470)
		Me._lblExtGroup_6.TabIndex = 61
		Me._lblExtGroup_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_6.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_6.Enabled = True
		Me._lblExtGroup_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_6.UseMnemonic = True
		Me._lblExtGroup_6.Visible = True
		Me._lblExtGroup_6.AutoSize = True
		Me._lblExtGroup_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_6.Name = "_lblExtGroup_6"
		Me._lblExtGroup_5.Text = "Ext or Group#"
		Me._lblExtGroup_5.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_5.Location = New System.Drawing.Point(50, 410)
		Me._lblExtGroup_5.TabIndex = 60
		Me._lblExtGroup_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_5.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_5.Enabled = True
		Me._lblExtGroup_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_5.UseMnemonic = True
		Me._lblExtGroup_5.Visible = True
		Me._lblExtGroup_5.AutoSize = True
		Me._lblExtGroup_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_5.Name = "_lblExtGroup_5"
		Me._lblExtGroup_4.Text = "Ext or Group#"
		Me._lblExtGroup_4.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_4.Location = New System.Drawing.Point(50, 350)
		Me._lblExtGroup_4.TabIndex = 59
		Me._lblExtGroup_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_4.Enabled = True
		Me._lblExtGroup_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_4.UseMnemonic = True
		Me._lblExtGroup_4.Visible = True
		Me._lblExtGroup_4.AutoSize = True
		Me._lblExtGroup_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_4.Name = "_lblExtGroup_4"
		Me._lblExtGroup_3.Text = "Ext or Group#"
		Me._lblExtGroup_3.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_3.Location = New System.Drawing.Point(50, 290)
		Me._lblExtGroup_3.TabIndex = 58
		Me._lblExtGroup_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_3.Enabled = True
		Me._lblExtGroup_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_3.UseMnemonic = True
		Me._lblExtGroup_3.Visible = True
		Me._lblExtGroup_3.AutoSize = True
		Me._lblExtGroup_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_3.Name = "_lblExtGroup_3"
		Me._lblExtGroup_2.Text = "Ext or Group#"
		Me._lblExtGroup_2.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_2.Location = New System.Drawing.Point(50, 230)
		Me._lblExtGroup_2.TabIndex = 57
		Me._lblExtGroup_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_2.Enabled = True
		Me._lblExtGroup_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_2.UseMnemonic = True
		Me._lblExtGroup_2.Visible = True
		Me._lblExtGroup_2.AutoSize = True
		Me._lblExtGroup_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_2.Name = "_lblExtGroup_2"
		Me._lblExtGroup_1.Text = "Ext or Group#"
		Me._lblExtGroup_1.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_1.Location = New System.Drawing.Point(50, 170)
		Me._lblExtGroup_1.TabIndex = 56
		Me._lblExtGroup_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_1.Enabled = True
		Me._lblExtGroup_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_1.UseMnemonic = True
		Me._lblExtGroup_1.Visible = True
		Me._lblExtGroup_1.AutoSize = True
		Me._lblExtGroup_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_1.Name = "_lblExtGroup_1"
		Me._lblExtGroup_0.Text = "Ext or Group#"
		Me._lblExtGroup_0.Size = New System.Drawing.Size(83, 17)
		Me._lblExtGroup_0.Location = New System.Drawing.Point(50, 110)
		Me._lblExtGroup_0.TabIndex = 55
		Me._lblExtGroup_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblExtGroup_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblExtGroup_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblExtGroup_0.Enabled = True
		Me._lblExtGroup_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblExtGroup_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblExtGroup_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblExtGroup_0.UseMnemonic = True
		Me._lblExtGroup_0.Visible = True
		Me._lblExtGroup_0.AutoSize = True
		Me._lblExtGroup_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblExtGroup_0.Name = "_lblExtGroup_0"
		Me.lblGroupNum.Text = "How Many?"
		Me.lblGroupNum.Size = New System.Drawing.Size(72, 17)
		Me.lblGroupNum.Location = New System.Drawing.Point(30, 60)
		Me.lblGroupNum.TabIndex = 54
		Me.lblGroupNum.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblGroupNum.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblGroupNum.BackColor = System.Drawing.SystemColors.Control
		Me.lblGroupNum.Enabled = True
		Me.lblGroupNum.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblGroupNum.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblGroupNum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblGroupNum.UseMnemonic = True
		Me.lblGroupNum.Visible = True
		Me.lblGroupNum.AutoSize = True
		Me.lblGroupNum.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblGroupNum.Name = "lblGroupNum"
		Me.fraMultiDates.Text = "Multiple Dates"
		Me.fraMultiDates.Size = New System.Drawing.Size(172, 712)
		Me.fraMultiDates.Location = New System.Drawing.Point(870, 10)
		Me.fraMultiDates.TabIndex = 31
		Me.fraMultiDates.Visible = False
		Me.fraMultiDates.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraMultiDates.BackColor = System.Drawing.SystemColors.Control
		Me.fraMultiDates.Enabled = True
		Me.fraMultiDates.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraMultiDates.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraMultiDates.Name = "fraMultiDates"
		Me.cboDateNum.Size = New System.Drawing.Size(52, 27)
		Me.cboDateNum.Location = New System.Drawing.Point(50, 30)
		Me.cboDateNum.Items.AddRange(New Object(){"2", "3", "4", "5", "6", "7", "8", "9", "10"})
		Me.cboDateNum.TabIndex = 50
		Me.cboDateNum.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboDateNum.BackColor = System.Drawing.SystemColors.Window
		Me.cboDateNum.CausesValidation = True
		Me.cboDateNum.Enabled = True
		Me.cboDateNum.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboDateNum.IntegralHeight = True
		Me.cboDateNum.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboDateNum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboDateNum.Sorted = False
		Me.cboDateNum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboDateNum.TabStop = True
		Me.cboDateNum.Visible = True
		Me.cboDateNum.Name = "cboDateNum"
		_DTPicker1_1.OcxState = CType(resources.GetObject("_DTPicker1_1.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_1.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_1.Location = New System.Drawing.Point(30, 90)
		Me._DTPicker1_1.TabIndex = 32
		Me._DTPicker1_1.Name = "_DTPicker1_1"
		_DTPicker1_2.OcxState = CType(resources.GetObject("_DTPicker1_2.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_2.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_2.Location = New System.Drawing.Point(30, 160)
		Me._DTPicker1_2.TabIndex = 34
		Me._DTPicker1_2.Name = "_DTPicker1_2"
		_DTPicker1_3.OcxState = CType(resources.GetObject("_DTPicker1_3.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_3.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_3.Location = New System.Drawing.Point(30, 230)
		Me._DTPicker1_3.TabIndex = 36
		Me._DTPicker1_3.Name = "_DTPicker1_3"
		_DTPicker1_4.OcxState = CType(resources.GetObject("_DTPicker1_4.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_4.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_4.Location = New System.Drawing.Point(30, 300)
		Me._DTPicker1_4.TabIndex = 38
		Me._DTPicker1_4.Name = "_DTPicker1_4"
		_DTPicker1_5.OcxState = CType(resources.GetObject("_DTPicker1_5.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_5.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_5.Location = New System.Drawing.Point(30, 370)
		Me._DTPicker1_5.TabIndex = 40
		Me._DTPicker1_5.Name = "_DTPicker1_5"
		_DTPicker1_6.OcxState = CType(resources.GetObject("_DTPicker1_6.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_6.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_6.Location = New System.Drawing.Point(30, 440)
		Me._DTPicker1_6.TabIndex = 42
		Me._DTPicker1_6.Name = "_DTPicker1_6"
		_DTPicker1_7.OcxState = CType(resources.GetObject("_DTPicker1_7.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_7.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_7.Location = New System.Drawing.Point(30, 510)
		Me._DTPicker1_7.TabIndex = 44
		Me._DTPicker1_7.Name = "_DTPicker1_7"
		_DTPicker1_8.OcxState = CType(resources.GetObject("_DTPicker1_8.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_8.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_8.Location = New System.Drawing.Point(30, 580)
		Me._DTPicker1_8.TabIndex = 46
		Me._DTPicker1_8.Name = "_DTPicker1_8"
		_DTPicker1_9.OcxState = CType(resources.GetObject("_DTPicker1_9.OcxState"), System.Windows.Forms.AxHost.State)
		Me._DTPicker1_9.Size = New System.Drawing.Size(132, 27)
		Me._DTPicker1_9.Location = New System.Drawing.Point(30, 650)
		Me._DTPicker1_9.TabIndex = 48
		Me._DTPicker1_9.Name = "_DTPicker1_9"
		Me.lblLines.Text = "Lines"
		Me.lblLines.Size = New System.Drawing.Size(32, 17)
		Me.lblLines.Location = New System.Drawing.Point(110, 40)
		Me.lblLines.TabIndex = 51
		Me.lblLines.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLines.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLines.BackColor = System.Drawing.SystemColors.Control
		Me.lblLines.Enabled = True
		Me.lblLines.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLines.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLines.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLines.UseMnemonic = True
		Me.lblLines.Visible = True
		Me.lblLines.AutoSize = True
		Me.lblLines.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLines.Name = "lblLines"
		Me._lblStart_9.Text = "Start Date# 10"
		Me._lblStart_9.Size = New System.Drawing.Size(88, 17)
		Me._lblStart_9.Location = New System.Drawing.Point(30, 630)
		Me._lblStart_9.TabIndex = 49
		Me._lblStart_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_9.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_9.Enabled = True
		Me._lblStart_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_9.UseMnemonic = True
		Me._lblStart_9.Visible = True
		Me._lblStart_9.AutoSize = True
		Me._lblStart_9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_9.Name = "_lblStart_9"
		Me._lblStart_8.Text = "Start Date #9"
		Me._lblStart_8.Size = New System.Drawing.Size(80, 17)
		Me._lblStart_8.Location = New System.Drawing.Point(30, 560)
		Me._lblStart_8.TabIndex = 47
		Me._lblStart_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_8.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_8.Enabled = True
		Me._lblStart_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_8.UseMnemonic = True
		Me._lblStart_8.Visible = True
		Me._lblStart_8.AutoSize = True
		Me._lblStart_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_8.Name = "_lblStart_8"
		Me._lblStart_7.Text = "Start Date #8"
		Me._lblStart_7.Size = New System.Drawing.Size(80, 17)
		Me._lblStart_7.Location = New System.Drawing.Point(30, 490)
		Me._lblStart_7.TabIndex = 45
		Me._lblStart_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_7.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_7.Enabled = True
		Me._lblStart_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_7.UseMnemonic = True
		Me._lblStart_7.Visible = True
		Me._lblStart_7.AutoSize = True
		Me._lblStart_7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_7.Name = "_lblStart_7"
		Me._lblStart_6.Text = "Start Date #7"
		Me._lblStart_6.Size = New System.Drawing.Size(80, 17)
		Me._lblStart_6.Location = New System.Drawing.Point(30, 420)
		Me._lblStart_6.TabIndex = 43
		Me._lblStart_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_6.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_6.Enabled = True
		Me._lblStart_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_6.UseMnemonic = True
		Me._lblStart_6.Visible = True
		Me._lblStart_6.AutoSize = True
		Me._lblStart_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_6.Name = "_lblStart_6"
		Me._lblStart_5.Text = "Start Date #6"
		Me._lblStart_5.Size = New System.Drawing.Size(80, 17)
		Me._lblStart_5.Location = New System.Drawing.Point(30, 350)
		Me._lblStart_5.TabIndex = 41
		Me._lblStart_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_5.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_5.Enabled = True
		Me._lblStart_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_5.UseMnemonic = True
		Me._lblStart_5.Visible = True
		Me._lblStart_5.AutoSize = True
		Me._lblStart_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_5.Name = "_lblStart_5"
		Me._lblStart_4.Text = "Start Date #5"
		Me._lblStart_4.Size = New System.Drawing.Size(80, 17)
		Me._lblStart_4.Location = New System.Drawing.Point(30, 280)
		Me._lblStart_4.TabIndex = 39
		Me._lblStart_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_4.Enabled = True
		Me._lblStart_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_4.UseMnemonic = True
		Me._lblStart_4.Visible = True
		Me._lblStart_4.AutoSize = True
		Me._lblStart_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_4.Name = "_lblStart_4"
		Me._lblStart_3.Text = "Start Date #4"
		Me._lblStart_3.Size = New System.Drawing.Size(80, 17)
		Me._lblStart_3.Location = New System.Drawing.Point(30, 210)
		Me._lblStart_3.TabIndex = 37
		Me._lblStart_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_3.Enabled = True
		Me._lblStart_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_3.UseMnemonic = True
		Me._lblStart_3.Visible = True
		Me._lblStart_3.AutoSize = True
		Me._lblStart_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_3.Name = "_lblStart_3"
		Me._lblStart_2.Text = "Start Date #3"
		Me._lblStart_2.Size = New System.Drawing.Size(80, 17)
		Me._lblStart_2.Location = New System.Drawing.Point(30, 140)
		Me._lblStart_2.TabIndex = 35
		Me._lblStart_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_2.Enabled = True
		Me._lblStart_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_2.UseMnemonic = True
		Me._lblStart_2.Visible = True
		Me._lblStart_2.AutoSize = True
		Me._lblStart_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_2.Name = "_lblStart_2"
		Me._lblStart_1.Text = "Start Date #2"
		Me._lblStart_1.Size = New System.Drawing.Size(80, 17)
		Me._lblStart_1.Location = New System.Drawing.Point(30, 70)
		Me._lblStart_1.TabIndex = 33
		Me._lblStart_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblStart_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblStart_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblStart_1.Enabled = True
		Me._lblStart_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblStart_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblStart_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblStart_1.UseMnemonic = True
		Me._lblStart_1.Visible = True
		Me._lblStart_1.AutoSize = True
		Me._lblStart_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblStart_1.Name = "_lblStart_1"
		MSChart1.OcxState = CType(resources.GetObject("MSChart1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.MSChart1.Size = New System.Drawing.Size(862, 562)
		Me.MSChart1.Location = New System.Drawing.Point(0, 170)
		Me.MSChart1.TabIndex = 76
		Me.MSChart1.Name = "MSChart1"
		Me.lblChartType.Text = "Chart Type"
		Me.lblChartType.Size = New System.Drawing.Size(65, 17)
		Me.lblChartType.Location = New System.Drawing.Point(750, 10)
		Me.lblChartType.TabIndex = 53
		Me.lblChartType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblChartType.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblChartType.BackColor = System.Drawing.SystemColors.Control
		Me.lblChartType.Enabled = True
		Me.lblChartType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblChartType.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblChartType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblChartType.UseMnemonic = True
		Me.lblChartType.Visible = True
		Me.lblChartType.AutoSize = True
		Me.lblChartType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblChartType.Name = "lblChartType"
		Me.DTPicker1.SetIndex(_DTPicker1_0, CType(0, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_1, CType(1, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_2, CType(2, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_3, CType(3, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_4, CType(4, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_5, CType(5, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_6, CType(6, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_7, CType(7, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_8, CType(8, Short))
		Me.DTPicker1.SetIndex(_DTPicker1_9, CType(9, Short))
		Me.cboGroup.SetIndex(_cboGroup_0, CType(0, Short))
		Me.cboGroup.SetIndex(_cboGroup_9, CType(9, Short))
		Me.cboGroup.SetIndex(_cboGroup_8, CType(8, Short))
		Me.cboGroup.SetIndex(_cboGroup_7, CType(7, Short))
		Me.cboGroup.SetIndex(_cboGroup_6, CType(6, Short))
		Me.cboGroup.SetIndex(_cboGroup_5, CType(5, Short))
		Me.cboGroup.SetIndex(_cboGroup_4, CType(4, Short))
		Me.cboGroup.SetIndex(_cboGroup_3, CType(3, Short))
		Me.cboGroup.SetIndex(_cboGroup_2, CType(2, Short))
		Me.cboGroup.SetIndex(_cboGroup_1, CType(1, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_8, CType(8, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_7, CType(7, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_6, CType(6, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_5, CType(5, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_4, CType(4, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_3, CType(3, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_2, CType(2, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_1, CType(1, Short))
		Me.lblExtGroup.SetIndex(_lblExtGroup_0, CType(0, Short))
		Me.lblStart.SetIndex(_lblStart_0, CType(0, Short))
		Me.lblStart.SetIndex(_lblStart_9, CType(9, Short))
		Me.lblStart.SetIndex(_lblStart_8, CType(8, Short))
		Me.lblStart.SetIndex(_lblStart_7, CType(7, Short))
		Me.lblStart.SetIndex(_lblStart_6, CType(6, Short))
		Me.lblStart.SetIndex(_lblStart_5, CType(5, Short))
		Me.lblStart.SetIndex(_lblStart_4, CType(4, Short))
		Me.lblStart.SetIndex(_lblStart_3, CType(3, Short))
		Me.lblStart.SetIndex(_lblStart_2, CType(2, Short))
		Me.lblStart.SetIndex(_lblStart_1, CType(1, Short))
		CType(Me.lblStart, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblExtGroup, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboGroup, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DTPicker1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MSChart1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_9, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_8, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_7, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_6, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_5, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_4, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._DTPicker1_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DTPicker2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DTPicker3, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DTPicker4, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MonthView1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(fraCallType)
		Me.Controls.Add(chkLablePoints)
		Me.Controls.Add(cmdAvg)
		Me.Controls.Add(MonthView1)
		Me.Controls.Add(cboChartType)
		Me.Controls.Add(fraMultiLines)
		Me.Controls.Add(cmdChart)
		Me.Controls.Add(cmdPrintChart)
		Me.Controls.Add(fraGroup)
		Me.Controls.Add(fraVariables)
		Me.Controls.Add(fraGroups)
		Me.Controls.Add(fraMultiDates)
		Me.Controls.Add(MSChart1)
		Me.Controls.Add(lblChartType)
		Me.fraCallType.Controls.Add(optVoiceMail)
		Me.fraCallType.Controls.Add(optCalls)
		Me.fraCallType.Controls.Add(optBoth)
		Me.fraMultiLines.Controls.Add(optCallVsVoice)
		Me.fraMultiLines.Controls.Add(optNone)
		Me.fraMultiLines.Controls.Add(optDirection)
		Me.fraMultiLines.Controls.Add(optMultiDates)
		Me.fraMultiLines.Controls.Add(optMultiGroups)
		Me.fraGroup.Controls.Add(_cboGroup_0)
		Me.fraGroup.Controls.Add(optExt)
		Me.fraGroup.Controls.Add(optWorkgroup)
		Me.fraVariables.Controls.Add(optTotal)
		Me.fraVariables.Controls.Add(optAvg)
		Me.fraVariables.Controls.Add(DTPicker4)
		Me.fraVariables.Controls.Add(DTPicker3)
		Me.fraVariables.Controls.Add(cboCallDir)
		Me.fraVariables.Controls.Add(cboDateType)
		Me.fraVariables.Controls.Add(DTPicker2)
		Me.fraVariables.Controls.Add(_DTPicker1_0)
		Me.fraVariables.Controls.Add(cboYear)
		Me.fraVariables.Controls.Add(cboSunday)
		Me.fraVariables.Controls.Add(cboMonth)
		Me.fraVariables.Controls.Add(lblCallType)
		Me.fraVariables.Controls.Add(lblDateType)
		Me.fraVariables.Controls.Add(lblEnd)
		Me.fraVariables.Controls.Add(_lblStart_0)
		Me.fraVariables.Controls.Add(lblMon)
		Me.fraVariables.Controls.Add(lblWeek)
		Me.fraVariables.Controls.Add(lblYear)
		Me.fraGroups.Controls.Add(cboNum)
		Me.fraGroups.Controls.Add(_cboGroup_9)
		Me.fraGroups.Controls.Add(_cboGroup_8)
		Me.fraGroups.Controls.Add(_cboGroup_7)
		Me.fraGroups.Controls.Add(_cboGroup_6)
		Me.fraGroups.Controls.Add(_cboGroup_5)
		Me.fraGroups.Controls.Add(_cboGroup_4)
		Me.fraGroups.Controls.Add(_cboGroup_3)
		Me.fraGroups.Controls.Add(_cboGroup_2)
		Me.fraGroups.Controls.Add(_cboGroup_1)
		Me.fraGroups.Controls.Add(_lblExtGroup_8)
		Me.fraGroups.Controls.Add(_lblExtGroup_7)
		Me.fraGroups.Controls.Add(_lblExtGroup_6)
		Me.fraGroups.Controls.Add(_lblExtGroup_5)
		Me.fraGroups.Controls.Add(_lblExtGroup_4)
		Me.fraGroups.Controls.Add(_lblExtGroup_3)
		Me.fraGroups.Controls.Add(_lblExtGroup_2)
		Me.fraGroups.Controls.Add(_lblExtGroup_1)
		Me.fraGroups.Controls.Add(_lblExtGroup_0)
		Me.fraGroups.Controls.Add(lblGroupNum)
		Me.fraMultiDates.Controls.Add(cboDateNum)
		Me.fraMultiDates.Controls.Add(_DTPicker1_1)
		Me.fraMultiDates.Controls.Add(_DTPicker1_2)
		Me.fraMultiDates.Controls.Add(_DTPicker1_3)
		Me.fraMultiDates.Controls.Add(_DTPicker1_4)
		Me.fraMultiDates.Controls.Add(_DTPicker1_5)
		Me.fraMultiDates.Controls.Add(_DTPicker1_6)
		Me.fraMultiDates.Controls.Add(_DTPicker1_7)
		Me.fraMultiDates.Controls.Add(_DTPicker1_8)
		Me.fraMultiDates.Controls.Add(_DTPicker1_9)
		Me.fraMultiDates.Controls.Add(lblLines)
		Me.fraMultiDates.Controls.Add(_lblStart_9)
		Me.fraMultiDates.Controls.Add(_lblStart_8)
		Me.fraMultiDates.Controls.Add(_lblStart_7)
		Me.fraMultiDates.Controls.Add(_lblStart_6)
		Me.fraMultiDates.Controls.Add(_lblStart_5)
		Me.fraMultiDates.Controls.Add(_lblStart_4)
		Me.fraMultiDates.Controls.Add(_lblStart_3)
		Me.fraMultiDates.Controls.Add(_lblStart_2)
		Me.fraMultiDates.Controls.Add(_lblStart_1)
		Me.fraCallType.ResumeLayout(False)
		Me.fraMultiLines.ResumeLayout(False)
		Me.fraGroup.ResumeLayout(False)
		Me.fraVariables.ResumeLayout(False)
		Me.fraGroups.ResumeLayout(False)
		Me.fraMultiDates.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class