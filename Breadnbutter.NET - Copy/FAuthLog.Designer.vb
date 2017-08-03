<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FAuthLog
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
	Public WithEvents cboFilter As AxSSDataWidgets_B.AxSSDBCombo
	Public WithEvents Label19 As System.Windows.Forms.Label
	Public WithEvents lblActs As System.Windows.Forms.Label
	Public WithEvents Frame5 As System.Windows.Forms.GroupBox
	Public WithEvents _lvwLog_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvwLog_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvwLog_ColumnHeader_3 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvwLog_ColumnHeader_4 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvwLog_ColumnHeader_5 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvwLog As System.Windows.Forms.ListView
	Public WithEvents ilLog As System.Windows.Forms.ImageList
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	Public WithEvents Label18 As System.Windows.Forms.Label
	Public WithEvents Image2 As System.Windows.Forms.PictureBox
	Public WithEvents Label12 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FAuthLog))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame5 = New System.Windows.Forms.GroupBox
		Me.cboFilter = New AxSSDataWidgets_B.AxSSDBCombo
		Me.Label19 = New System.Windows.Forms.Label
		Me.lblActs = New System.Windows.Forms.Label
		Me.lvwLog = New System.Windows.Forms.ListView
		Me._lvwLog_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._lvwLog_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me._lvwLog_ColumnHeader_3 = New System.Windows.Forms.ColumnHeader
		Me._lvwLog_ColumnHeader_4 = New System.Windows.Forms.ColumnHeader
		Me._lvwLog_ColumnHeader_5 = New System.Windows.Forms.ColumnHeader
		Me.ilLog = New System.Windows.Forms.ImageList
		Me.Image1 = New System.Windows.Forms.PictureBox
		Me.Label18 = New System.Windows.Forms.Label
		Me.Image2 = New System.Windows.Forms.PictureBox
		Me.Label12 = New System.Windows.Forms.Label
		Me.Frame5.SuspendLayout()
		Me.lvwLog.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cboFilter, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(999, 589)
		Me.Location = New System.Drawing.Point(85, 139)
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
		Me.Name = "FAuthLog"
		Me.Frame5.Text = "Options"
		Me.Frame5.Size = New System.Drawing.Size(444, 74)
		Me.Frame5.Location = New System.Drawing.Point(8, 0)
		Me.Frame5.TabIndex = 0
		Me.Frame5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame5.BackColor = System.Drawing.SystemColors.Control
		Me.Frame5.Enabled = True
		Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame5.Visible = True
		Me.Frame5.Name = "Frame5"
		cboFilter.OcxState = CType(resources.GetObject("cboFilter.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cboFilter.Size = New System.Drawing.Size(182, 24)
		Me.cboFilter.Location = New System.Drawing.Point(108, 28)
		Me.cboFilter.TabIndex = 1
		Me.cboFilter.Name = "cboFilter"
		Me.Label19.Text = "Log Filter:"
		Me.Label19.Size = New System.Drawing.Size(59, 19)
		Me.Label19.Location = New System.Drawing.Point(23, 30)
		Me.Label19.TabIndex = 3
		Me.Label19.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label19.BackColor = System.Drawing.SystemColors.Control
		Me.Label19.Enabled = True
		Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label19.UseMnemonic = True
		Me.Label19.Visible = True
		Me.Label19.AutoSize = False
		Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label19.Name = "Label19"
		Me.lblActs.Text = "(0000 of 0000)"
		Me.lblActs.Size = New System.Drawing.Size(97, 17)
		Me.lblActs.Location = New System.Drawing.Point(305, 30)
		Me.lblActs.TabIndex = 2
		Me.lblActs.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblActs.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblActs.BackColor = System.Drawing.SystemColors.Control
		Me.lblActs.Enabled = True
		Me.lblActs.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblActs.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblActs.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblActs.UseMnemonic = True
		Me.lblActs.Visible = True
		Me.lblActs.AutoSize = False
		Me.lblActs.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblActs.Name = "lblActs"
		Me.lvwLog.Size = New System.Drawing.Size(989, 474)
		Me.lvwLog.Location = New System.Drawing.Point(3, 110)
		Me.lvwLog.TabIndex = 4
		Me.lvwLog.View = System.Windows.Forms.View.Details
		Me.lvwLog.LabelWrap = True
		Me.lvwLog.HideSelection = True
		Me.lvwLog.FullRowSelect = True
		Me.lvwLog.LargeImageList = ilLog
		Me.lvwLog.SmallImageList = ilLog
		Me.lvwLog.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lvwLog.BackColor = System.Drawing.SystemColors.Window
		Me.lvwLog.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lvwLog.LabelEdit = True
		Me.lvwLog.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lvwLog.Name = "lvwLog"
		Me._lvwLog_ColumnHeader_1.Text = "Date/Time"
		Me._lvwLog_ColumnHeader_1.Width = 294
		Me._lvwLog_ColumnHeader_2.Text = "Employee"
		Me._lvwLog_ColumnHeader_2.Width = 212
		Me._lvwLog_ColumnHeader_3.Text = "Company"
		Me._lvwLog_ColumnHeader_3.Width = 441
		Me._lvwLog_ColumnHeader_4.Text = "User"
		Me._lvwLog_ColumnHeader_4.Width = 412
		Me._lvwLog_ColumnHeader_5.Text = "Action"
		Me._lvwLog_ColumnHeader_5.Width = 343
		Me.ilLog.ImageSize = New System.Drawing.Size(16, 16)
		Me.ilLog.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.ilLog.ImageStream = CType(resources.GetObject("ilLog.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.ilLog.Images.SetKeyName(0, "imgAscending")
		Me.ilLog.Images.SetKeyName(1, "imgDescending")
		Me.ilLog.Images.SetKeyName(2, "imgAuthorize")
		Me.ilLog.Images.SetKeyName(3, "imgDeauthorize")
		Me.ilLog.Images.SetKeyName(4, "imgRestore")
		Me.Image1.Size = New System.Drawing.Size(20, 20)
		Me.Image1.Location = New System.Drawing.Point(5, 85)
		Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
		Me.Image1.Enabled = True
		Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image1.Visible = True
		Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image1.Name = "Image1"
		Me.Label18.Text = "Activity Log"
		Me.Label18.Size = New System.Drawing.Size(77, 17)
		Me.Label18.Location = New System.Drawing.Point(30, 83)
		Me.Label18.TabIndex = 6
		Me.Label18.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label18.BackColor = System.Drawing.SystemColors.Control
		Me.Label18.Enabled = True
		Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label18.UseMnemonic = True
		Me.Label18.Visible = True
		Me.Label18.AutoSize = False
		Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label18.Name = "Label18"
		Me.Image2.Size = New System.Drawing.Size(25, 25)
		Me.Image2.Location = New System.Drawing.Point(710, 80)
		Me.Image2.Image = CType(resources.GetObject("Image2.Image"), System.Drawing.Image)
		Me.Image2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
		Me.Image2.Enabled = True
		Me.Image2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image2.Visible = True
		Me.Image2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image2.Name = "Image2"
		Me.Label12.Text = "Double-click activity to view full details"
		Me.Label12.Size = New System.Drawing.Size(232, 17)
		Me.Label12.Location = New System.Drawing.Point(740, 83)
		Me.Label12.TabIndex = 5
		Me.Label12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label12.BackColor = System.Drawing.SystemColors.Control
		Me.Label12.Enabled = True
		Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label12.UseMnemonic = True
		Me.Label12.Visible = True
		Me.Label12.AutoSize = False
		Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label12.Name = "Label12"
		CType(Me.cboFilter, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(Frame5)
		Me.Controls.Add(lvwLog)
		Me.Controls.Add(Image1)
		Me.Controls.Add(Label18)
		Me.Controls.Add(Image2)
		Me.Controls.Add(Label12)
		Me.Frame5.Controls.Add(cboFilter)
		Me.Frame5.Controls.Add(Label19)
		Me.Frame5.Controls.Add(lblActs)
		Me.lvwLog.Columns.Add(_lvwLog_ColumnHeader_1)
		Me.lvwLog.Columns.Add(_lvwLog_ColumnHeader_2)
		Me.lvwLog.Columns.Add(_lvwLog_ColumnHeader_3)
		Me.lvwLog.Columns.Add(_lvwLog_ColumnHeader_4)
		Me.lvwLog.Columns.Add(_lvwLog_ColumnHeader_5)
		Me.Frame5.ResumeLayout(False)
		Me.lvwLog.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class