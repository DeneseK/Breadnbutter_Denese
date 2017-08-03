<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		MDIForm_Initialize_renamed()
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
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents TrayIconList3 As System.Windows.Forms.ImageList
	Public WithEvents TrayIconList2 As System.Windows.Forms.ImageList
	Public WithEvents TrayIconList As System.Windows.Forms.ImageList
	Public WithEvents tmrMessages As System.Windows.Forms.Timer
	Public WithEvents tmrSecChk As System.Windows.Forms.Timer
	Public WithEvents tbMain As AxActiveToolBars.AxSSActiveToolBars
	Public WithEvents tmrTray As System.Windows.Forms.Timer
	Public WithEvents License As SKCLLib.LFile
	Public WithEvents mnuTrayOpen As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTrayExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuTray As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FMain))
		Me.IsMDIContainer = True
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Picture1 = New System.Windows.Forms.PictureBox
		Me.TrayIconList3 = New System.Windows.Forms.ImageList
		Me.TrayIconList2 = New System.Windows.Forms.ImageList
		Me.TrayIconList = New System.Windows.Forms.ImageList
		Me.tmrMessages = New System.Windows.Forms.Timer(components)
		Me.tmrSecChk = New System.Windows.Forms.Timer(components)
		Me.tbMain = New AxActiveToolBars.AxSSActiveToolBars
		Me.tmrTray = New System.Windows.Forms.Timer(components)
		Me.License = New SKCLLib.LFile
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuTray = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuTrayOpen = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuTrayExit = New System.Windows.Forms.ToolStripMenuItem
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.tbMain, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.License, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.SystemColors.AppWorkspace
		Me.Text = "Bread 'n' Butter"
		Me.ClientSize = New System.Drawing.Size(762, 555)
		Me.Location = New System.Drawing.Point(272, 223)
		Me.Icon = CType(resources.GetObject("FMain.Icon"), System.Drawing.Icon)
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Name = "FMain"
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.Top
		Me.Picture1.Size = New System.Drawing.Size(762, 72)
		Me.Picture1.Location = New System.Drawing.Point(0, 0)
		Me.Picture1.TabIndex = 0
		Me.Picture1.Visible = False
		Me.Picture1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture1.BackColor = System.Drawing.SystemColors.Control
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Picture1.Name = "Picture1"
		Me.TrayIconList3.ImageSize = New System.Drawing.Size(32, 32)
		Me.TrayIconList3.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.TrayIconList3.ImageStream = CType(resources.GetObject("TrayIconList3.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.TrayIconList3.Images.SetKeyName(0, "")
		Me.TrayIconList3.Images.SetKeyName(1, "")
		Me.TrayIconList3.Images.SetKeyName(2, "")
		Me.TrayIconList3.Images.SetKeyName(3, "")
		Me.TrayIconList3.Images.SetKeyName(4, "")
		Me.TrayIconList3.Images.SetKeyName(5, "")
		Me.TrayIconList3.Images.SetKeyName(6, "")
		Me.TrayIconList3.Images.SetKeyName(7, "")
		Me.TrayIconList3.Images.SetKeyName(8, "")
		Me.TrayIconList3.Images.SetKeyName(9, "")
		Me.TrayIconList3.Images.SetKeyName(10, "")
		Me.TrayIconList3.Images.SetKeyName(11, "")
		Me.TrayIconList3.Images.SetKeyName(12, "")
		Me.TrayIconList3.Images.SetKeyName(13, "")
		Me.TrayIconList3.Images.SetKeyName(14, "")
		Me.TrayIconList3.Images.SetKeyName(15, "")
		Me.TrayIconList3.Images.SetKeyName(16, "")
		Me.TrayIconList3.Images.SetKeyName(17, "")
		Me.TrayIconList3.Images.SetKeyName(18, "")
		Me.TrayIconList3.Images.SetKeyName(19, "")
		Me.TrayIconList3.Images.SetKeyName(20, "")
		Me.TrayIconList3.Images.SetKeyName(21, "")
		Me.TrayIconList3.Images.SetKeyName(22, "")
		Me.TrayIconList3.Images.SetKeyName(23, "")
		Me.TrayIconList2.ImageSize = New System.Drawing.Size(32, 32)
		Me.TrayIconList2.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.TrayIconList2.ImageStream = CType(resources.GetObject("TrayIconList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.TrayIconList2.Images.SetKeyName(0, "")
		Me.TrayIconList2.Images.SetKeyName(1, "")
		Me.TrayIconList2.Images.SetKeyName(2, "")
		Me.TrayIconList2.Images.SetKeyName(3, "")
		Me.TrayIconList2.Images.SetKeyName(4, "")
		Me.TrayIconList2.Images.SetKeyName(5, "")
		Me.TrayIconList2.Images.SetKeyName(6, "")
		Me.TrayIconList2.Images.SetKeyName(7, "")
		Me.TrayIconList2.Images.SetKeyName(8, "")
		Me.TrayIconList2.Images.SetKeyName(9, "")
		Me.TrayIconList2.Images.SetKeyName(10, "")
		Me.TrayIconList2.Images.SetKeyName(11, "")
		Me.TrayIconList2.Images.SetKeyName(12, "")
		Me.TrayIconList2.Images.SetKeyName(13, "")
		Me.TrayIconList2.Images.SetKeyName(14, "")
		Me.TrayIconList2.Images.SetKeyName(15, "")
		Me.TrayIconList2.Images.SetKeyName(16, "")
		Me.TrayIconList2.Images.SetKeyName(17, "")
		Me.TrayIconList2.Images.SetKeyName(18, "")
		Me.TrayIconList2.Images.SetKeyName(19, "")
		Me.TrayIconList2.Images.SetKeyName(20, "")
		Me.TrayIconList2.Images.SetKeyName(21, "")
		Me.TrayIconList2.Images.SetKeyName(22, "")
		Me.TrayIconList2.Images.SetKeyName(23, "")
		Me.TrayIconList2.Images.SetKeyName(24, "")
		Me.TrayIconList2.Images.SetKeyName(25, "")
		Me.TrayIconList2.Images.SetKeyName(26, "")
		Me.TrayIconList2.Images.SetKeyName(27, "")
		Me.TrayIconList2.Images.SetKeyName(28, "")
		Me.TrayIconList2.Images.SetKeyName(29, "")
		Me.TrayIconList2.Images.SetKeyName(30, "")
		Me.TrayIconList2.Images.SetKeyName(31, "")
		Me.TrayIconList2.Images.SetKeyName(32, "")
		Me.TrayIconList2.Images.SetKeyName(33, "")
		Me.TrayIconList2.Images.SetKeyName(34, "")
		Me.TrayIconList2.Images.SetKeyName(35, "")
		Me.TrayIconList2.Images.SetKeyName(36, "")
		Me.TrayIconList2.Images.SetKeyName(37, "")
		Me.TrayIconList2.Images.SetKeyName(38, "")
		Me.TrayIconList2.Images.SetKeyName(39, "")
		Me.TrayIconList2.Images.SetKeyName(40, "")
		Me.TrayIconList2.Images.SetKeyName(41, "")
		Me.TrayIconList2.Images.SetKeyName(42, "")
		Me.TrayIconList2.Images.SetKeyName(43, "")
		Me.TrayIconList2.Images.SetKeyName(44, "")
		Me.TrayIconList2.Images.SetKeyName(45, "")
		Me.TrayIconList2.Images.SetKeyName(46, "")
		Me.TrayIconList2.Images.SetKeyName(47, "")
		Me.TrayIconList.ImageSize = New System.Drawing.Size(16, 16)
		Me.TrayIconList.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.TrayIconList.ImageStream = CType(resources.GetObject("TrayIconList.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.TrayIconList.Images.SetKeyName(0, "")
		Me.TrayIconList.Images.SetKeyName(1, "")
		Me.TrayIconList.Images.SetKeyName(2, "")
		Me.TrayIconList.Images.SetKeyName(3, "")
		Me.TrayIconList.Images.SetKeyName(4, "")
		Me.TrayIconList.Images.SetKeyName(5, "")
		Me.TrayIconList.Images.SetKeyName(6, "")
		Me.TrayIconList.Images.SetKeyName(7, "")
		Me.TrayIconList.Images.SetKeyName(8, "")
		Me.TrayIconList.Images.SetKeyName(9, "")
		Me.TrayIconList.Images.SetKeyName(10, "")
		Me.TrayIconList.Images.SetKeyName(11, "")
		Me.TrayIconList.Images.SetKeyName(12, "")
		Me.TrayIconList.Images.SetKeyName(13, "")
		Me.TrayIconList.Images.SetKeyName(14, "")
		Me.TrayIconList.Images.SetKeyName(15, "")
		Me.TrayIconList.Images.SetKeyName(16, "")
		Me.TrayIconList.Images.SetKeyName(17, "")
		Me.TrayIconList.Images.SetKeyName(18, "")
		Me.TrayIconList.Images.SetKeyName(19, "")
		Me.tmrMessages.Interval = 30000
		Me.tmrMessages.Enabled = True
		Me.tmrSecChk.Interval = 60000
		Me.tmrSecChk.Enabled = True
		tbMain.OcxState = CType(resources.GetObject("tbMain.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tbMain.Location = New System.Drawing.Point(25, 208)
		Me.tbMain.Name = "tbMain"
		Me.tmrTray.Interval = 150
		Me.tmrTray.Enabled = True
		License.OcxState = CType(resources.GetObject("License.OcxState"), System.Windows.Forms.AxHost.State)
		Me.License.Location = New System.Drawing.Point(28, 250)
		Me.License.Name = "License"
		Me.mnuTray.Name = "mnuTray"
		Me.mnuTray.Text = "Tray"
		Me.mnuTray.Visible = False
		Me.mnuTray.Checked = False
		Me.mnuTray.Enabled = True
		Me.mnuTrayOpen.Name = "mnuTrayOpen"
		Me.mnuTrayOpen.Text = "Open"
		Me.mnuTrayOpen.Checked = False
		Me.mnuTrayOpen.Enabled = True
		Me.mnuTrayOpen.Visible = True
		Me.mnuTrayExit.Name = "mnuTrayExit"
		Me.mnuTrayExit.Text = "Exit"
		Me.mnuTrayExit.Checked = False
		Me.mnuTrayExit.Enabled = True
		Me.mnuTrayExit.Visible = True
		Me.Controls.Add(Picture1)
		Me.Controls.Add(tbMain)
		Me.Controls.Add(License)
		CType(Me.License, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tbMain, System.ComponentModel.ISupportInitialize).EndInit()
		Me.mnuTray.MergeAction = System.Windows.Forms.MergeAction.Remove
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuTray})
		mnuTray.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuTrayOpen, Me.mnuTrayExit})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class