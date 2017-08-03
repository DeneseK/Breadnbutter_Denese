<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FSupportOpen
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
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents lstGroup As System.Windows.Forms.CheckedListBox
	Public WithEvents Group As System.Windows.Forms.GroupBox
	Public WithEvents chkDate As System.Windows.Forms.CheckBox
	Public WithEvents DTPicker2 As AxMSComCtl2.AxDTPicker
	Public WithEvents DTPicker1 As AxMSComCtl2.AxDTPicker
	Public WithEvents lblTo As System.Windows.Forms.Label
	Public WithEvents lblFrom As System.Windows.Forms.Label
	Public WithEvents Date_Renamed As System.Windows.Forms.GroupBox
	Public WithEvents optGroup As System.Windows.Forms.RadioButton
	Public WithEvents optUser As System.Windows.Forms.RadioButton
	Public WithEvents lstUsers As System.Windows.Forms.CheckedListBox
	Public WithEvents User As System.Windows.Forms.GroupBox
	Public WithEvents lstCategory As System.Windows.Forms.CheckedListBox
	Public WithEvents Frame7 As System.Windows.Forms.GroupBox
	Public WithEvents grdHistory As AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
	Public WithEvents cmdPrint As System.Windows.Forms.Button
	Public WithEvents cmdShowResults As System.Windows.Forms.Button
	Public WithEvents cmdCopy As System.Windows.Forms.Button
	Public WithEvents lblCategory As System.Windows.Forms.Label
	Public WithEvents LblCount As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.Panel
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FSupportOpen))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame1 = New System.Windows.Forms.Panel
		Me.Command1 = New System.Windows.Forms.Button
		Me.Group = New System.Windows.Forms.GroupBox
		Me.lstGroup = New System.Windows.Forms.CheckedListBox
		Me.chkDate = New System.Windows.Forms.CheckBox
		Me.Date_Renamed = New System.Windows.Forms.GroupBox
		Me.DTPicker2 = New AxMSComCtl2.AxDTPicker
		Me.DTPicker1 = New AxMSComCtl2.AxDTPicker
		Me.lblTo = New System.Windows.Forms.Label
		Me.lblFrom = New System.Windows.Forms.Label
		Me.optGroup = New System.Windows.Forms.RadioButton
		Me.optUser = New System.Windows.Forms.RadioButton
		Me.User = New System.Windows.Forms.GroupBox
		Me.lstUsers = New System.Windows.Forms.CheckedListBox
		Me.Frame7 = New System.Windows.Forms.GroupBox
		Me.lstCategory = New System.Windows.Forms.CheckedListBox
		Me.grdHistory = New AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
		Me.cmdPrint = New System.Windows.Forms.Button
		Me.cmdShowResults = New System.Windows.Forms.Button
		Me.cmdCopy = New System.Windows.Forms.Button
		Me.lblCategory = New System.Windows.Forms.Label
		Me.LblCount = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.Group.SuspendLayout()
		Me.Date_Renamed.SuspendLayout()
		Me.User.SuspendLayout()
		Me.Frame7.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DTPicker2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DTPicker1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.grdHistory, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.AppWorkspace
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Open Calls"
		Me.ClientSize = New System.Drawing.Size(982, 614)
		Me.Location = New System.Drawing.Point(180, 197)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FSupportOpen"
		Me.Frame1.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Frame1.Text = "Frame1"
		Me.Frame1.Size = New System.Drawing.Size(964, 587)
		Me.Frame1.Location = New System.Drawing.Point(0, 0)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "close selected calls"
		Me.Command1.Size = New System.Drawing.Size(167, 27)
		Me.Command1.Location = New System.Drawing.Point(780, 550)
		Me.Command1.TabIndex = 13
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Group.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Group.Size = New System.Drawing.Size(182, 137)
		Me.Group.Location = New System.Drawing.Point(390, 440)
		Me.Group.TabIndex = 22
		Me.Group.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Group.Enabled = True
		Me.Group.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Group.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Group.Visible = True
		Me.Group.Name = "Group"
		Me.lstGroup.Size = New System.Drawing.Size(162, 99)
		Me.lstGroup.Location = New System.Drawing.Point(10, 20)
		Me.lstGroup.TabIndex = 7
		Me.lstGroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstGroup.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstGroup.BackColor = System.Drawing.SystemColors.Window
		Me.lstGroup.CausesValidation = True
		Me.lstGroup.Enabled = True
		Me.lstGroup.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstGroup.IntegralHeight = True
		Me.lstGroup.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstGroup.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstGroup.Sorted = False
		Me.lstGroup.TabStop = True
		Me.lstGroup.Visible = True
		Me.lstGroup.MultiColumn = False
		Me.lstGroup.Name = "lstGroup"
		Me.chkDate.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.chkDate.Text = "Date"
		Me.chkDate.Size = New System.Drawing.Size(112, 22)
		Me.chkDate.Location = New System.Drawing.Point(580, 420)
		Me.chkDate.TabIndex = 4
		Me.chkDate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDate.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDate.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDate.CausesValidation = True
		Me.chkDate.Enabled = True
		Me.chkDate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkDate.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDate.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDate.TabStop = True
		Me.chkDate.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDate.Visible = True
		Me.chkDate.Name = "chkDate"
		Me.Date_Renamed.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Date_Renamed.Size = New System.Drawing.Size(182, 137)
		Me.Date_Renamed.Location = New System.Drawing.Point(580, 440)
		Me.Date_Renamed.TabIndex = 18
		Me.Date_Renamed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Date_Renamed.Enabled = True
		Me.Date_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Date_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Date_Renamed.Visible = True
		Me.Date_Renamed.Name = "Date_Renamed"
		DTPicker2.OcxState = CType(resources.GetObject("DTPicker2.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DTPicker2.Size = New System.Drawing.Size(142, 32)
		Me.DTPicker2.Location = New System.Drawing.Point(20, 90)
		Me.DTPicker2.TabIndex = 9
		Me.DTPicker2.Name = "DTPicker2"
		DTPicker1.OcxState = CType(resources.GetObject("DTPicker1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DTPicker1.Size = New System.Drawing.Size(142, 32)
		Me.DTPicker1.Location = New System.Drawing.Point(20, 30)
		Me.DTPicker1.TabIndex = 8
		Me.DTPicker1.Name = "DTPicker1"
		Me.lblTo.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.lblTo.Text = "To"
		Me.lblTo.Size = New System.Drawing.Size(17, 17)
		Me.lblTo.Location = New System.Drawing.Point(20, 70)
		Me.lblTo.TabIndex = 20
		Me.lblTo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTo.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblTo.Enabled = True
		Me.lblTo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTo.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTo.UseMnemonic = True
		Me.lblTo.Visible = True
		Me.lblTo.AutoSize = True
		Me.lblTo.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTo.Name = "lblTo"
		Me.lblFrom.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.lblFrom.Text = "From"
		Me.lblFrom.Size = New System.Drawing.Size(29, 17)
		Me.lblFrom.Location = New System.Drawing.Point(20, 10)
		Me.lblFrom.TabIndex = 19
		Me.lblFrom.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFrom.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFrom.Enabled = True
		Me.lblFrom.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFrom.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFrom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFrom.UseMnemonic = True
		Me.lblFrom.Visible = True
		Me.lblFrom.AutoSize = True
		Me.lblFrom.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFrom.Name = "lblFrom"
		Me.optGroup.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optGroup.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.optGroup.Text = "Group"
		Me.optGroup.Size = New System.Drawing.Size(102, 22)
		Me.optGroup.Location = New System.Drawing.Point(390, 420)
		Me.optGroup.TabIndex = 3
		Me.optGroup.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optGroup.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optGroup.CausesValidation = True
		Me.optGroup.Enabled = True
		Me.optGroup.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optGroup.Cursor = System.Windows.Forms.Cursors.Default
		Me.optGroup.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optGroup.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optGroup.TabStop = True
		Me.optGroup.Checked = False
		Me.optGroup.Visible = True
		Me.optGroup.Name = "optGroup"
		Me.optUser.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUser.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.optUser.Text = "User"
		Me.optUser.Size = New System.Drawing.Size(92, 22)
		Me.optUser.Location = New System.Drawing.Point(200, 420)
		Me.optUser.TabIndex = 2
		Me.optUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optUser.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUser.CausesValidation = True
		Me.optUser.Enabled = True
		Me.optUser.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.optUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optUser.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optUser.TabStop = True
		Me.optUser.Checked = False
		Me.optUser.Visible = True
		Me.optUser.Name = "optUser"
		Me.User.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.User.Size = New System.Drawing.Size(182, 137)
		Me.User.Location = New System.Drawing.Point(200, 440)
		Me.User.TabIndex = 15
		Me.User.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.User.Enabled = True
		Me.User.ForeColor = System.Drawing.SystemColors.ControlText
		Me.User.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.User.Visible = True
		Me.User.Name = "User"
		Me.lstUsers.Size = New System.Drawing.Size(162, 99)
		Me.lstUsers.Location = New System.Drawing.Point(10, 20)
		Me.lstUsers.TabIndex = 6
		Me.lstUsers.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstUsers.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstUsers.BackColor = System.Drawing.SystemColors.Window
		Me.lstUsers.CausesValidation = True
		Me.lstUsers.Enabled = True
		Me.lstUsers.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstUsers.IntegralHeight = True
		Me.lstUsers.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstUsers.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstUsers.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstUsers.Sorted = False
		Me.lstUsers.TabStop = True
		Me.lstUsers.Visible = True
		Me.lstUsers.MultiColumn = False
		Me.lstUsers.Name = "lstUsers"
		Me.Frame7.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Frame7.Size = New System.Drawing.Size(182, 137)
		Me.Frame7.Location = New System.Drawing.Point(10, 440)
		Me.Frame7.TabIndex = 14
		Me.Frame7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame7.Enabled = True
		Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame7.Visible = True
		Me.Frame7.Name = "Frame7"
		Me.lstCategory.Size = New System.Drawing.Size(162, 99)
		Me.lstCategory.Location = New System.Drawing.Point(10, 20)
		Me.lstCategory.TabIndex = 5
		Me.lstCategory.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstCategory.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstCategory.BackColor = System.Drawing.SystemColors.Window
		Me.lstCategory.CausesValidation = True
		Me.lstCategory.Enabled = True
		Me.lstCategory.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstCategory.IntegralHeight = True
		Me.lstCategory.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstCategory.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstCategory.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstCategory.Sorted = False
		Me.lstCategory.TabStop = True
		Me.lstCategory.Visible = True
		Me.lstCategory.MultiColumn = False
		Me.lstCategory.Name = "lstCategory"
		grdHistory.OcxState = CType(resources.GetObject("grdHistory.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdHistory.Size = New System.Drawing.Size(964, 392)
		Me.grdHistory.Location = New System.Drawing.Point(0, 20)
		Me.grdHistory.TabIndex = 1
		Me.grdHistory.Name = "grdHistory"
		Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPrint.Text = "Print Report"
		Me.cmdPrint.Size = New System.Drawing.Size(167, 27)
		Me.cmdPrint.Location = New System.Drawing.Point(780, 470)
		Me.cmdPrint.TabIndex = 11
		Me.cmdPrint.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPrint.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPrint.CausesValidation = True
		Me.cmdPrint.Enabled = True
		Me.cmdPrint.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPrint.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPrint.TabStop = True
		Me.cmdPrint.Name = "cmdPrint"
		Me.cmdShowResults.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdShowResults.Text = "Show Results"
		Me.cmdShowResults.Size = New System.Drawing.Size(167, 27)
		Me.cmdShowResults.Location = New System.Drawing.Point(780, 430)
		Me.cmdShowResults.TabIndex = 10
		Me.cmdShowResults.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdShowResults.BackColor = System.Drawing.SystemColors.Control
		Me.cmdShowResults.CausesValidation = True
		Me.cmdShowResults.Enabled = True
		Me.cmdShowResults.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdShowResults.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdShowResults.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdShowResults.TabStop = True
		Me.cmdShowResults.Name = "cmdShowResults"
		Me.cmdCopy.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCopy.Text = "Copy Results to Clipboad"
		Me.cmdCopy.Size = New System.Drawing.Size(167, 27)
		Me.cmdCopy.Location = New System.Drawing.Point(780, 510)
		Me.cmdCopy.TabIndex = 12
		Me.cmdCopy.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCopy.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCopy.CausesValidation = True
		Me.cmdCopy.Enabled = True
		Me.cmdCopy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCopy.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCopy.TabStop = True
		Me.cmdCopy.Name = "cmdCopy"
		Me.lblCategory.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.lblCategory.Text = "Category"
		Me.lblCategory.Size = New System.Drawing.Size(53, 17)
		Me.lblCategory.Location = New System.Drawing.Point(10, 420)
		Me.lblCategory.TabIndex = 21
		Me.lblCategory.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCategory.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCategory.Enabled = True
		Me.lblCategory.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCategory.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCategory.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCategory.UseMnemonic = True
		Me.lblCategory.Visible = True
		Me.lblCategory.AutoSize = True
		Me.lblCategory.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCategory.Name = "lblCategory"
		Me.LblCount.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.LblCount.Text = "0"
		Me.LblCount.Size = New System.Drawing.Size(84, 22)
		Me.LblCount.Location = New System.Drawing.Point(100, 0)
		Me.LblCount.TabIndex = 17
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
		Me.Label7.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.Label7.Text = "Notes Found:"
		Me.Label7.Size = New System.Drawing.Size(82, 22)
		Me.Label7.Location = New System.Drawing.Point(10, 0)
		Me.Label7.TabIndex = 16
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
		CType(Me.grdHistory, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DTPicker1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DTPicker2, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(Command1)
		Me.Frame1.Controls.Add(Group)
		Me.Frame1.Controls.Add(chkDate)
		Me.Frame1.Controls.Add(Date_Renamed)
		Me.Frame1.Controls.Add(optGroup)
		Me.Frame1.Controls.Add(optUser)
		Me.Frame1.Controls.Add(User)
		Me.Frame1.Controls.Add(Frame7)
		Me.Frame1.Controls.Add(grdHistory)
		Me.Frame1.Controls.Add(cmdPrint)
		Me.Frame1.Controls.Add(cmdShowResults)
		Me.Frame1.Controls.Add(cmdCopy)
		Me.Frame1.Controls.Add(lblCategory)
		Me.Frame1.Controls.Add(LblCount)
		Me.Frame1.Controls.Add(Label7)
		Me.Group.Controls.Add(lstGroup)
		Me.Date_Renamed.Controls.Add(DTPicker2)
		Me.Date_Renamed.Controls.Add(DTPicker1)
		Me.Date_Renamed.Controls.Add(lblTo)
		Me.Date_Renamed.Controls.Add(lblFrom)
		Me.User.Controls.Add(lstUsers)
		Me.Frame7.Controls.Add(lstCategory)
		Me.Frame1.ResumeLayout(False)
		Me.Group.ResumeLayout(False)
		Me.Date_Renamed.ResumeLayout(False)
		Me.User.ResumeLayout(False)
		Me.Frame7.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class