<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FLicense
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
	Public WithEvents txtDescript As AxTx4oleLib.AxTXTextControl
	Public WithEvents Picture2 As System.Windows.Forms.Panel
	Public WithEvents tbLic As AxActiveToolBars.AxSSActiveToolBars
	Public WithEvents Picture1 As System.Windows.Forms.Panel
	Public WithEvents cmdDeauthorize As AxThreed.AxSSCommand
	Public WithEvents cmdImprint As AxThreed.AxSSCommand
	Public WithEvents cmdExport As AxThreed.AxSSCommand
	Public WithEvents cmdImport As AxThreed.AxSSCommand
	Public WithEvents mskConfirm As AxTDBMask6.AxTDBMask
	Public WithEvents cmdRefresh As AxThreed.AxSSCommand
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents _Label1_88 As System.Windows.Forms.Label
	Public WithEvents fmeAdv As System.Windows.Forms.GroupBox
	Public WithEvents txtSiteCode As System.Windows.Forms.TextBox
	Public WithEvents cmdAuthorize As AxThreed.AxSSCommand
	Public WithEvents mskSiteKey As AxTDBMask6.AxTDBMask
	Public WithEvents cmdDone As AxThreed.AxSSCommand
	Public WithEvents cmdAdv As AxThreed.AxSSCommand
	Public WithEvents imgSec As System.Windows.Forms.PictureBox
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label1_46 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FLicense))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Picture2 = New System.Windows.Forms.Panel
		Me.txtDescript = New AxTx4oleLib.AxTXTextControl
		Me.Picture1 = New System.Windows.Forms.Panel
		Me.tbLic = New AxActiveToolBars.AxSSActiveToolBars
		Me.fmeAdv = New System.Windows.Forms.GroupBox
		Me.cmdDeauthorize = New AxThreed.AxSSCommand
		Me.cmdImprint = New AxThreed.AxSSCommand
		Me.cmdExport = New AxThreed.AxSSCommand
		Me.cmdImport = New AxThreed.AxSSCommand
		Me.mskConfirm = New AxTDBMask6.AxTDBMask
		Me.cmdRefresh = New AxThreed.AxSSCommand
		Me.Label2 = New System.Windows.Forms.Label
		Me._Label1_88 = New System.Windows.Forms.Label
		Me.txtSiteCode = New System.Windows.Forms.TextBox
		Me.cmdAuthorize = New AxThreed.AxSSCommand
		Me.mskSiteKey = New AxTDBMask6.AxTDBMask
		Me.cmdDone = New AxThreed.AxSSCommand
		Me.cmdAdv = New AxThreed.AxSSCommand
		Me.imgSec = New System.Windows.Forms.PictureBox
		Me._Label1_6 = New System.Windows.Forms.Label
		Me._Label1_46 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Picture2.SuspendLayout()
		Me.Picture1.SuspendLayout()
		Me.fmeAdv.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.txtDescript, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.tbLic, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdDeauthorize, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdImprint, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdExport, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdImport, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskConfirm, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdRefresh, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdAuthorize, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskSiteKey, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdDone, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdAdv, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.ClientSize = New System.Drawing.Size(670, 560)
		Me.Location = New System.Drawing.Point(0, 0)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FLicense"
		Me.Picture2.Size = New System.Drawing.Size(647, 189)
		Me.Picture2.Location = New System.Drawing.Point(13, 13)
		Me.Picture2.TabIndex = 16
		Me.Picture2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture2.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture2.BackColor = System.Drawing.SystemColors.Control
		Me.Picture2.CausesValidation = True
		Me.Picture2.Enabled = True
		Me.Picture2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Picture2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture2.TabStop = True
		Me.Picture2.Visible = True
		Me.Picture2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Picture2.Name = "Picture2"
		txtDescript.OcxState = CType(resources.GetObject("txtDescript.OcxState"), System.Windows.Forms.AxHost.State)
		Me.txtDescript.Size = New System.Drawing.Size(642, 184)
		Me.txtDescript.Location = New System.Drawing.Point(0, 0)
		Me.txtDescript.TabIndex = 17
		Me.txtDescript.Name = "txtDescript"
		Me.Picture1.Size = New System.Drawing.Size(47, 44)
		Me.Picture1.Location = New System.Drawing.Point(10, 490)
		Me.Picture1.TabIndex = 9
		Me.Picture1.Visible = False
		Me.Picture1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture1.BackColor = System.Drawing.SystemColors.Control
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Picture1.Name = "Picture1"
		tbLic.OcxState = CType(resources.GetObject("tbLic.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tbLic.Location = New System.Drawing.Point(3, 8)
		Me.tbLic.Visible = False
		Me.tbLic.Name = "tbLic"
		Me.fmeAdv.Text = "Advanced Operations"
		Me.fmeAdv.Size = New System.Drawing.Size(642, 124)
		Me.fmeAdv.Location = New System.Drawing.Point(15, 350)
		Me.fmeAdv.TabIndex = 3
		Me.fmeAdv.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fmeAdv.BackColor = System.Drawing.SystemColors.Control
		Me.fmeAdv.Enabled = True
		Me.fmeAdv.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fmeAdv.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fmeAdv.Visible = True
		Me.fmeAdv.Name = "fmeAdv"
		cmdDeauthorize.OcxState = CType(resources.GetObject("cmdDeauthorize.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdDeauthorize.Size = New System.Drawing.Size(99, 24)
		Me.cmdDeauthorize.Location = New System.Drawing.Point(23, 85)
		Me.cmdDeauthorize.TabIndex = 10
		Me.cmdDeauthorize.Name = "cmdDeauthorize"
		cmdImprint.OcxState = CType(resources.GetObject("cmdImprint.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdImprint.Size = New System.Drawing.Size(69, 27)
		Me.cmdImprint.Location = New System.Drawing.Point(170, 145)
		Me.cmdImprint.TabIndex = 11
		Me.cmdImprint.Visible = False
		Me.cmdImprint.Name = "cmdImprint"
		cmdExport.OcxState = CType(resources.GetObject("cmdExport.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdExport.Size = New System.Drawing.Size(69, 27)
		Me.cmdExport.Location = New System.Drawing.Point(250, 145)
		Me.cmdExport.TabIndex = 12
		Me.cmdExport.Visible = False
		Me.cmdExport.Name = "cmdExport"
		cmdImport.OcxState = CType(resources.GetObject("cmdImport.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdImport.Size = New System.Drawing.Size(69, 27)
		Me.cmdImport.Location = New System.Drawing.Point(330, 145)
		Me.cmdImport.TabIndex = 13
		Me.cmdImport.Visible = False
		Me.cmdImport.Name = "cmdImport"
		mskConfirm.OcxState = CType(resources.GetObject("mskConfirm.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskConfirm.Size = New System.Drawing.Size(232, 24)
		Me.mskConfirm.Location = New System.Drawing.Point(288, 85)
		Me.mskConfirm.TabIndex = 15
		Me.mskConfirm.Name = "mskConfirm"
		cmdRefresh.OcxState = CType(resources.GetObject("cmdRefresh.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdRefresh.Size = New System.Drawing.Size(139, 24)
		Me.cmdRefresh.Location = New System.Drawing.Point(23, 33)
		Me.cmdRefresh.TabIndex = 18
		Me.cmdRefresh.Name = "cmdRefresh"
		Me.Label2.Text = "Transfer Functions:"
		Me.Label2.Size = New System.Drawing.Size(119, 19)
		Me.Label2.Location = New System.Drawing.Point(43, 150)
		Me.Label2.TabIndex = 14
		Me.Label2.Visible = False
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me._Label1_88.Text = "Confirmation Code:"
		Me._Label1_88.Size = New System.Drawing.Size(114, 17)
		Me._Label1_88.Location = New System.Drawing.Point(168, 88)
		Me._Label1_88.TabIndex = 4
		Me._Label1_88.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_88.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_88.BackColor = System.Drawing.Color.Transparent
		Me._Label1_88.Enabled = True
		Me._Label1_88.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_88.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_88.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_88.UseMnemonic = True
		Me._Label1_88.Visible = True
		Me._Label1_88.AutoSize = False
		Me._Label1_88.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_88.Name = "_Label1_88"
		Me.txtSiteCode.AutoSize = False
		Me.txtSiteCode.BackColor = System.Drawing.Color.Green
		Me.txtSiteCode.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSiteCode.Size = New System.Drawing.Size(394, 29)
		Me.txtSiteCode.Location = New System.Drawing.Point(140, 228)
		Me.txtSiteCode.ReadOnly = True
		Me.txtSiteCode.TabIndex = 0
		Me.txtSiteCode.AcceptsReturn = True
		Me.txtSiteCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSiteCode.CausesValidation = True
		Me.txtSiteCode.Enabled = True
		Me.txtSiteCode.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSiteCode.HideSelection = True
		Me.txtSiteCode.Maxlength = 0
		Me.txtSiteCode.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSiteCode.MultiLine = False
		Me.txtSiteCode.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSiteCode.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSiteCode.TabStop = True
		Me.txtSiteCode.Visible = True
		Me.txtSiteCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSiteCode.Name = "txtSiteCode"
		cmdAuthorize.OcxState = CType(resources.GetObject("cmdAuthorize.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdAuthorize.Size = New System.Drawing.Size(112, 32)
		Me.cmdAuthorize.Location = New System.Drawing.Point(545, 228)
		Me.cmdAuthorize.TabIndex = 1
		Me.cmdAuthorize.Name = "cmdAuthorize"
		mskSiteKey.OcxState = CType(resources.GetObject("mskSiteKey.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskSiteKey.Size = New System.Drawing.Size(394, 29)
		Me.mskSiteKey.Location = New System.Drawing.Point(140, 263)
		Me.mskSiteKey.TabIndex = 2
		Me.mskSiteKey.Name = "mskSiteKey"
		cmdDone.OcxState = CType(resources.GetObject("cmdDone.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdDone.Size = New System.Drawing.Size(112, 32)
		Me.cmdDone.Location = New System.Drawing.Point(545, 263)
		Me.cmdDone.TabIndex = 5
		Me.cmdDone.Name = "cmdDone"
		cmdAdv.OcxState = CType(resources.GetObject("cmdAdv.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdAdv.Size = New System.Drawing.Size(112, 32)
		Me.cmdAdv.Location = New System.Drawing.Point(545, 308)
		Me.cmdAdv.TabIndex = 6
		Me.cmdAdv.Visible = False
		Me.cmdAdv.Name = "cmdAdv"
		Me.imgSec.Size = New System.Drawing.Size(40, 40)
		Me.imgSec.Location = New System.Drawing.Point(20, 238)
		Me.imgSec.Image = CType(resources.GetObject("imgSec.Image"), System.Drawing.Image)
		Me.imgSec.Enabled = True
		Me.imgSec.Cursor = System.Windows.Forms.Cursors.Default
		Me.imgSec.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.imgSec.Visible = True
		Me.imgSec.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.imgSec.Name = "imgSec"
		Me._Label1_6.Text = "Site Key:"
		Me._Label1_6.Size = New System.Drawing.Size(57, 17)
		Me._Label1_6.Location = New System.Drawing.Point(70, 268)
		Me._Label1_6.TabIndex = 8
		Me._Label1_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_6.BackColor = System.Drawing.Color.Transparent
		Me._Label1_6.Enabled = True
		Me._Label1_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_6.UseMnemonic = True
		Me._Label1_6.Visible = True
		Me._Label1_6.AutoSize = False
		Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_6.Name = "_Label1_6"
		Me._Label1_46.Text = "Site Code:"
		Me._Label1_46.Size = New System.Drawing.Size(62, 17)
		Me._Label1_46.Location = New System.Drawing.Point(70, 233)
		Me._Label1_46.TabIndex = 7
		Me._Label1_46.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_46.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_46.BackColor = System.Drawing.Color.Transparent
		Me._Label1_46.Enabled = True
		Me._Label1_46.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_46.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_46.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_46.UseMnemonic = True
		Me._Label1_46.Visible = True
		Me._Label1_46.AutoSize = False
		Me._Label1_46.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_46.Name = "_Label1_46"
		Me.Label1.SetIndex(_Label1_88, CType(88, Short))
		Me.Label1.SetIndex(_Label1_6, CType(6, Short))
		Me.Label1.SetIndex(_Label1_46, CType(46, Short))
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdAdv, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdDone, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskSiteKey, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdAuthorize, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdRefresh, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskConfirm, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdImport, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdExport, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdImprint, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdDeauthorize, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tbLic, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.txtDescript, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(Picture2)
		Me.Controls.Add(Picture1)
		Me.Controls.Add(fmeAdv)
		Me.Controls.Add(txtSiteCode)
		Me.Controls.Add(cmdAuthorize)
		Me.Controls.Add(mskSiteKey)
		Me.Controls.Add(cmdDone)
		Me.Controls.Add(cmdAdv)
		Me.Controls.Add(imgSec)
		Me.Controls.Add(_Label1_6)
		Me.Controls.Add(_Label1_46)
		Me.Picture2.Controls.Add(txtDescript)
		Me.Picture1.Controls.Add(tbLic)
		Me.fmeAdv.Controls.Add(cmdDeauthorize)
		Me.fmeAdv.Controls.Add(cmdImprint)
		Me.fmeAdv.Controls.Add(cmdExport)
		Me.fmeAdv.Controls.Add(cmdImport)
		Me.fmeAdv.Controls.Add(mskConfirm)
		Me.fmeAdv.Controls.Add(cmdRefresh)
		Me.fmeAdv.Controls.Add(Label2)
		Me.fmeAdv.Controls.Add(_Label1_88)
		Me.Picture2.ResumeLayout(False)
		Me.Picture1.ResumeLayout(False)
		Me.fmeAdv.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class