<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FDetails
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
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
	Public WithEvents txtPVAuthDays As System.Windows.Forms.TextBox
	Public WithEvents txtPVVersionShipped As System.Windows.Forms.TextBox
	Public WithEvents txtPVAuths As System.Windows.Forms.TextBox
	Public WithEvents txtPVPendingDays As System.Windows.Forms.TextBox
	Public WithEvents txtPVGraceDays As System.Windows.Forms.TextBox
	Public WithEvents txtPVSaleDays As System.Windows.Forms.TextBox
	Public WithEvents cboPVAuthStatus As AxSSDataWidgets_B.AxSSDBCombo
	Public WithEvents cboPVShipStatus As AxSSDataWidgets_B.AxSSDBCombo
	Public WithEvents mskPVShipDate As AxTDBDate6.AxTDBDate
	Public WithEvents mskPVAuthDate As AxTDBDate6.AxTDBDate
	Public WithEvents cmdPVShipDate As AxThreed.AxSSCommand
	Public WithEvents cmdPVAuthDate As AxThreed.AxSSCommand
	Public WithEvents cboPVDownloadStatus As AxSSDataWidgets_B.AxSSDBCombo
	Public WithEvents mskPVDownloadDate As AxTDBDate6.AxTDBDate
	Public WithEvents cmdPVDownloadDate As AxThreed.AxSSCommand
	Public WithEvents mskPVSaleDate As AxTDBDate6.AxTDBDate
	Public WithEvents cmdPVSalesDate As AxThreed.AxSSCommand
	Public WithEvents _Label10_5 As System.Windows.Forms.Label
	Public WithEvents _Label10_1 As System.Windows.Forms.Label
	Public WithEvents _Label4_6 As System.Windows.Forms.Label
	Public WithEvents _Label10_2 As System.Windows.Forms.Label
	Public WithEvents lblPVAuthRemaining As System.Windows.Forms.Label
	Public WithEvents lblPVExpires As System.Windows.Forms.Label
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label17 As System.Windows.Forms.Label
	Public WithEvents Label14 As System.Windows.Forms.Label
	Public WithEvents _Label10_8 As System.Windows.Forms.Label
	Public WithEvents _Label10_9 As System.Windows.Forms.Label
	Public WithEvents _Label10_12 As System.Windows.Forms.Label
	Public WithEvents fmePVAuthStatus As System.Windows.Forms.Panel
	Public WithEvents Timer4 As System.Windows.Forms.Timer
	Public WithEvents Timer3 As System.Windows.Forms.Timer
	Public WithEvents Timer2 As System.Windows.Forms.Timer
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Image9 As System.Windows.Forms.PictureBox
	Public WithEvents Image8 As System.Windows.Forms.PictureBox
	Public WithEvents Image7 As System.Windows.Forms.PictureBox
	Public WithEvents Image6 As System.Windows.Forms.PictureBox
	Public WithEvents Image5 As System.Windows.Forms.PictureBox
	Public WithEvents Image4 As System.Windows.Forms.PictureBox
	Public WithEvents Image3 As System.Windows.Forms.PictureBox
	Public WithEvents Image2 As System.Windows.Forms.PictureBox
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	Public WithEvents Shape8 As System.Windows.Forms.Label
	Public WithEvents Shape7 As System.Windows.Forms.Label
	Public WithEvents Shape6 As System.Windows.Forms.Label
	Public WithEvents Shape5 As System.Windows.Forms.Label
	Public WithEvents Shape4 As System.Windows.Forms.Label
	Public WithEvents Shape3 As System.Windows.Forms.Label
	Public WithEvents Shape2 As System.Windows.Forms.Label
	Public WithEvents Shape1 As System.Windows.Forms.Label
	Public WithEvents Label10 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Label4 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FDetails))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.fmePVAuthStatus = New System.Windows.Forms.Panel
		Me.txtPVAuthDays = New System.Windows.Forms.TextBox
		Me.txtPVVersionShipped = New System.Windows.Forms.TextBox
		Me.txtPVAuths = New System.Windows.Forms.TextBox
		Me.txtPVPendingDays = New System.Windows.Forms.TextBox
		Me.txtPVGraceDays = New System.Windows.Forms.TextBox
		Me.txtPVSaleDays = New System.Windows.Forms.TextBox
		Me.cboPVAuthStatus = New AxSSDataWidgets_B.AxSSDBCombo
		Me.cboPVShipStatus = New AxSSDataWidgets_B.AxSSDBCombo
		Me.mskPVShipDate = New AxTDBDate6.AxTDBDate
		Me.mskPVAuthDate = New AxTDBDate6.AxTDBDate
		Me.cmdPVShipDate = New AxThreed.AxSSCommand
		Me.cmdPVAuthDate = New AxThreed.AxSSCommand
		Me.cboPVDownloadStatus = New AxSSDataWidgets_B.AxSSDBCombo
		Me.mskPVDownloadDate = New AxTDBDate6.AxTDBDate
		Me.cmdPVDownloadDate = New AxThreed.AxSSCommand
		Me.mskPVSaleDate = New AxTDBDate6.AxTDBDate
		Me.cmdPVSalesDate = New AxThreed.AxSSCommand
		Me._Label10_5 = New System.Windows.Forms.Label
		Me._Label10_1 = New System.Windows.Forms.Label
		Me._Label4_6 = New System.Windows.Forms.Label
		Me._Label10_2 = New System.Windows.Forms.Label
		Me.lblPVAuthRemaining = New System.Windows.Forms.Label
		Me.lblPVExpires = New System.Windows.Forms.Label
		Me.Label11 = New System.Windows.Forms.Label
		Me.Label17 = New System.Windows.Forms.Label
		Me.Label14 = New System.Windows.Forms.Label
		Me._Label10_8 = New System.Windows.Forms.Label
		Me._Label10_9 = New System.Windows.Forms.Label
		Me._Label10_12 = New System.Windows.Forms.Label
		Me.Timer4 = New System.Windows.Forms.Timer(components)
		Me.Timer3 = New System.Windows.Forms.Timer(components)
		Me.Timer2 = New System.Windows.Forms.Timer(components)
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Image9 = New System.Windows.Forms.PictureBox
		Me.Image8 = New System.Windows.Forms.PictureBox
		Me.Image7 = New System.Windows.Forms.PictureBox
		Me.Image6 = New System.Windows.Forms.PictureBox
		Me.Image5 = New System.Windows.Forms.PictureBox
		Me.Image4 = New System.Windows.Forms.PictureBox
		Me.Image3 = New System.Windows.Forms.PictureBox
		Me.Image2 = New System.Windows.Forms.PictureBox
		Me.Image1 = New System.Windows.Forms.PictureBox
		Me.Shape8 = New System.Windows.Forms.Label
		Me.Shape7 = New System.Windows.Forms.Label
		Me.Shape6 = New System.Windows.Forms.Label
		Me.Shape5 = New System.Windows.Forms.Label
		Me.Shape4 = New System.Windows.Forms.Label
		Me.Shape3 = New System.Windows.Forms.Label
		Me.Shape2 = New System.Windows.Forms.Label
		Me.Shape1 = New System.Windows.Forms.Label
		Me.Label10 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Label4 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.fmePVAuthStatus.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cboPVAuthStatus, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cboPVShipStatus, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskPVShipDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskPVAuthDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdPVShipDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdPVAuthDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cboPVDownloadStatus, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskPVDownloadDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdPVDownloadDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.mskPVSaleDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdPVSalesDate, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label10, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label4, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Pac Man"
		Me.ClientSize = New System.Drawing.Size(507, 183)
		Me.Location = New System.Drawing.Point(328, 284)
		Me.ControlBox = False
		Me.Icon = CType(resources.GetObject("FDetails.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FDetails"
		Me.fmePVAuthStatus.BackColor = System.Drawing.Color.FromARGB(255, 192, 255)
		Me.fmePVAuthStatus.Size = New System.Drawing.Size(474, 167)
		Me.fmePVAuthStatus.Location = New System.Drawing.Point(0, 0)
		Me.fmePVAuthStatus.TabIndex = 2
		Me.fmePVAuthStatus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fmePVAuthStatus.Dock = System.Windows.Forms.DockStyle.None
		Me.fmePVAuthStatus.CausesValidation = True
		Me.fmePVAuthStatus.Enabled = True
		Me.fmePVAuthStatus.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fmePVAuthStatus.Cursor = System.Windows.Forms.Cursors.Default
		Me.fmePVAuthStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fmePVAuthStatus.TabStop = True
		Me.fmePVAuthStatus.Visible = True
		Me.fmePVAuthStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.fmePVAuthStatus.Name = "fmePVAuthStatus"
		Me.txtPVAuthDays.AutoSize = False
		Me.txtPVAuthDays.Enabled = False
		Me.txtPVAuthDays.Size = New System.Drawing.Size(57, 27)
		Me.txtPVAuthDays.Location = New System.Drawing.Point(350, 130)
		Me.txtPVAuthDays.TabIndex = 8
		Me.txtPVAuthDays.Tag = "1"
		Me.txtPVAuthDays.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPVAuthDays.AcceptsReturn = True
		Me.txtPVAuthDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPVAuthDays.BackColor = System.Drawing.SystemColors.Window
		Me.txtPVAuthDays.CausesValidation = True
		Me.txtPVAuthDays.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPVAuthDays.HideSelection = True
		Me.txtPVAuthDays.ReadOnly = False
		Me.txtPVAuthDays.Maxlength = 0
		Me.txtPVAuthDays.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPVAuthDays.MultiLine = False
		Me.txtPVAuthDays.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPVAuthDays.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPVAuthDays.TabStop = True
		Me.txtPVAuthDays.Visible = True
		Me.txtPVAuthDays.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPVAuthDays.Name = "txtPVAuthDays"
		Me.txtPVVersionShipped.AutoSize = False
		Me.txtPVVersionShipped.Size = New System.Drawing.Size(57, 27)
		Me.txtPVVersionShipped.Location = New System.Drawing.Point(410, 104)
		Me.txtPVVersionShipped.TabIndex = 7
		Me.txtPVVersionShipped.Tag = "1"
		Me.txtPVVersionShipped.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPVVersionShipped.AcceptsReturn = True
		Me.txtPVVersionShipped.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPVVersionShipped.BackColor = System.Drawing.SystemColors.Window
		Me.txtPVVersionShipped.CausesValidation = True
		Me.txtPVVersionShipped.Enabled = True
		Me.txtPVVersionShipped.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPVVersionShipped.HideSelection = True
		Me.txtPVVersionShipped.ReadOnly = False
		Me.txtPVVersionShipped.Maxlength = 0
		Me.txtPVVersionShipped.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPVVersionShipped.MultiLine = False
		Me.txtPVVersionShipped.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPVVersionShipped.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPVVersionShipped.TabStop = True
		Me.txtPVVersionShipped.Visible = True
		Me.txtPVVersionShipped.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPVVersionShipped.Name = "txtPVVersionShipped"
		Me.txtPVAuths.AutoSize = False
		Me.txtPVAuths.Size = New System.Drawing.Size(32, 27)
		Me.txtPVAuths.Location = New System.Drawing.Point(280, 7)
		Me.txtPVAuths.Maxlength = 1
		Me.txtPVAuths.TabIndex = 6
		Me.txtPVAuths.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPVAuths.AcceptsReturn = True
		Me.txtPVAuths.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPVAuths.BackColor = System.Drawing.SystemColors.Window
		Me.txtPVAuths.CausesValidation = True
		Me.txtPVAuths.Enabled = True
		Me.txtPVAuths.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPVAuths.HideSelection = True
		Me.txtPVAuths.ReadOnly = False
		Me.txtPVAuths.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPVAuths.MultiLine = False
		Me.txtPVAuths.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPVAuths.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPVAuths.TabStop = True
		Me.txtPVAuths.Visible = True
		Me.txtPVAuths.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPVAuths.Name = "txtPVAuths"
		Me.txtPVPendingDays.AutoSize = False
		Me.txtPVPendingDays.Enabled = False
		Me.txtPVPendingDays.Size = New System.Drawing.Size(47, 27)
		Me.txtPVPendingDays.Location = New System.Drawing.Point(423, 43)
		Me.txtPVPendingDays.ReadOnly = True
		Me.txtPVPendingDays.Maxlength = 4
		Me.txtPVPendingDays.TabIndex = 5
		Me.txtPVPendingDays.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPVPendingDays.AcceptsReturn = True
		Me.txtPVPendingDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPVPendingDays.BackColor = System.Drawing.SystemColors.Window
		Me.txtPVPendingDays.CausesValidation = True
		Me.txtPVPendingDays.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPVPendingDays.HideSelection = True
		Me.txtPVPendingDays.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPVPendingDays.MultiLine = False
		Me.txtPVPendingDays.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPVPendingDays.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPVPendingDays.TabStop = True
		Me.txtPVPendingDays.Visible = True
		Me.txtPVPendingDays.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPVPendingDays.Name = "txtPVPendingDays"
		Me.txtPVGraceDays.AutoSize = False
		Me.txtPVGraceDays.Size = New System.Drawing.Size(47, 27)
		Me.txtPVGraceDays.Location = New System.Drawing.Point(424, 8)
		Me.txtPVGraceDays.Maxlength = 2
		Me.txtPVGraceDays.TabIndex = 4
		Me.txtPVGraceDays.Tag = "1"
		Me.txtPVGraceDays.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPVGraceDays.AcceptsReturn = True
		Me.txtPVGraceDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPVGraceDays.BackColor = System.Drawing.SystemColors.Window
		Me.txtPVGraceDays.CausesValidation = True
		Me.txtPVGraceDays.Enabled = True
		Me.txtPVGraceDays.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPVGraceDays.HideSelection = True
		Me.txtPVGraceDays.ReadOnly = False
		Me.txtPVGraceDays.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPVGraceDays.MultiLine = False
		Me.txtPVGraceDays.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPVGraceDays.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPVGraceDays.TabStop = True
		Me.txtPVGraceDays.Visible = True
		Me.txtPVGraceDays.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPVGraceDays.Name = "txtPVGraceDays"
		Me.txtPVSaleDays.AutoSize = False
		Me.txtPVSaleDays.Size = New System.Drawing.Size(42, 27)
		Me.txtPVSaleDays.Location = New System.Drawing.Point(270, 40)
		Me.txtPVSaleDays.Maxlength = 4
		Me.txtPVSaleDays.TabIndex = 3
		Me.txtPVSaleDays.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPVSaleDays.AcceptsReturn = True
		Me.txtPVSaleDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPVSaleDays.BackColor = System.Drawing.SystemColors.Window
		Me.txtPVSaleDays.CausesValidation = True
		Me.txtPVSaleDays.Enabled = True
		Me.txtPVSaleDays.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPVSaleDays.HideSelection = True
		Me.txtPVSaleDays.ReadOnly = False
		Me.txtPVSaleDays.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPVSaleDays.MultiLine = False
		Me.txtPVSaleDays.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPVSaleDays.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPVSaleDays.TabStop = True
		Me.txtPVSaleDays.Visible = True
		Me.txtPVSaleDays.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPVSaleDays.Name = "txtPVSaleDays"
		cboPVAuthStatus.OcxState = CType(resources.GetObject("cboPVAuthStatus.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cboPVAuthStatus.Size = New System.Drawing.Size(155, 27)
		Me.cboPVAuthStatus.Location = New System.Drawing.Point(68, 132)
		Me.cboPVAuthStatus.TabIndex = 9
		Me.cboPVAuthStatus.Name = "cboPVAuthStatus"
		cboPVShipStatus.OcxState = CType(resources.GetObject("cboPVShipStatus.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cboPVShipStatus.Size = New System.Drawing.Size(155, 27)
		Me.cboPVShipStatus.Location = New System.Drawing.Point(68, 104)
		Me.cboPVShipStatus.TabIndex = 10
		Me.cboPVShipStatus.Name = "cboPVShipStatus"
		mskPVShipDate.OcxState = CType(resources.GetObject("mskPVShipDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskPVShipDate.Size = New System.Drawing.Size(87, 27)
		Me.mskPVShipDate.Location = New System.Drawing.Point(225, 104)
		Me.mskPVShipDate.TabIndex = 11
		Me.mskPVShipDate.Name = "mskPVShipDate"
		mskPVAuthDate.OcxState = CType(resources.GetObject("mskPVAuthDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskPVAuthDate.Size = New System.Drawing.Size(87, 27)
		Me.mskPVAuthDate.Location = New System.Drawing.Point(225, 130)
		Me.mskPVAuthDate.TabIndex = 12
		Me.mskPVAuthDate.Name = "mskPVAuthDate"
		cmdPVShipDate.OcxState = CType(resources.GetObject("cmdPVShipDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdPVShipDate.Size = New System.Drawing.Size(24, 27)
		Me.cmdPVShipDate.Location = New System.Drawing.Point(315, 104)
		Me.cmdPVShipDate.TabIndex = 13
		Me.cmdPVShipDate.Name = "cmdPVShipDate"
		cmdPVAuthDate.OcxState = CType(resources.GetObject("cmdPVAuthDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdPVAuthDate.Size = New System.Drawing.Size(24, 27)
		Me.cmdPVAuthDate.Location = New System.Drawing.Point(315, 132)
		Me.cmdPVAuthDate.TabIndex = 14
		Me.cmdPVAuthDate.Visible = False
		Me.cmdPVAuthDate.Name = "cmdPVAuthDate"
		cboPVDownloadStatus.OcxState = CType(resources.GetObject("cboPVDownloadStatus.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cboPVDownloadStatus.Size = New System.Drawing.Size(155, 27)
		Me.cboPVDownloadStatus.Location = New System.Drawing.Point(68, 75)
		Me.cboPVDownloadStatus.TabIndex = 15
		Me.cboPVDownloadStatus.Name = "cboPVDownloadStatus"
		mskPVDownloadDate.OcxState = CType(resources.GetObject("mskPVDownloadDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskPVDownloadDate.Size = New System.Drawing.Size(87, 27)
		Me.mskPVDownloadDate.Location = New System.Drawing.Point(225, 75)
		Me.mskPVDownloadDate.TabIndex = 16
		Me.mskPVDownloadDate.Name = "mskPVDownloadDate"
		cmdPVDownloadDate.OcxState = CType(resources.GetObject("cmdPVDownloadDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdPVDownloadDate.Size = New System.Drawing.Size(24, 27)
		Me.cmdPVDownloadDate.Location = New System.Drawing.Point(315, 75)
		Me.cmdPVDownloadDate.TabIndex = 17
		Me.cmdPVDownloadDate.Name = "cmdPVDownloadDate"
		mskPVSaleDate.OcxState = CType(resources.GetObject("mskPVSaleDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.mskPVSaleDate.Size = New System.Drawing.Size(87, 27)
		Me.mskPVSaleDate.Location = New System.Drawing.Point(74, 39)
		Me.mskPVSaleDate.TabIndex = 18
		Me.mskPVSaleDate.Name = "mskPVSaleDate"
		cmdPVSalesDate.OcxState = CType(resources.GetObject("cmdPVSalesDate.OcxState"), System.Windows.Forms.AxHost.State)
		Me.cmdPVSalesDate.Size = New System.Drawing.Size(24, 27)
		Me.cmdPVSalesDate.Location = New System.Drawing.Point(164, 39)
		Me.cmdPVSalesDate.TabIndex = 19
		Me.cmdPVSalesDate.Name = "cmdPVSalesDate"
		Me._Label10_5.Text = "Download:"
		Me._Label10_5.Size = New System.Drawing.Size(64, 24)
		Me._Label10_5.Location = New System.Drawing.Point(4, 78)
		Me._Label10_5.TabIndex = 31
		Me._Label10_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label10_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label10_5.BackColor = System.Drawing.Color.Transparent
		Me._Label10_5.Enabled = True
		Me._Label10_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label10_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label10_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label10_5.UseMnemonic = True
		Me._Label10_5.Visible = True
		Me._Label10_5.AutoSize = False
		Me._Label10_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label10_5.Name = "_Label10_5"
		Me._Label10_1.Text = "Shipping:"
		Me._Label10_1.Size = New System.Drawing.Size(64, 24)
		Me._Label10_1.Location = New System.Drawing.Point(4, 107)
		Me._Label10_1.TabIndex = 30
		Me._Label10_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label10_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label10_1.BackColor = System.Drawing.Color.Transparent
		Me._Label10_1.Enabled = True
		Me._Label10_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label10_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label10_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label10_1.UseMnemonic = True
		Me._Label10_1.Visible = True
		Me._Label10_1.AutoSize = False
		Me._Label10_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label10_1.Name = "_Label10_1"
		Me._Label4_6.Text = "Ver.:"
		Me._Label4_6.Size = New System.Drawing.Size(54, 24)
		Me._Label4_6.Location = New System.Drawing.Point(353, 107)
		Me._Label4_6.TabIndex = 29
		Me._Label4_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label4_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label4_6.BackColor = System.Drawing.Color.Transparent
		Me._Label4_6.Enabled = True
		Me._Label4_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label4_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label4_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label4_6.UseMnemonic = True
		Me._Label4_6.Visible = True
		Me._Label4_6.AutoSize = False
		Me._Label4_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label4_6.Name = "_Label4_6"
		Me._Label10_2.Text = "Auth:"
		Me._Label10_2.Size = New System.Drawing.Size(84, 24)
		Me._Label10_2.Location = New System.Drawing.Point(4, 134)
		Me._Label10_2.TabIndex = 28
		Me._Label10_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label10_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label10_2.BackColor = System.Drawing.Color.Transparent
		Me._Label10_2.Enabled = True
		Me._Label10_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label10_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label10_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label10_2.UseMnemonic = True
		Me._Label10_2.Visible = True
		Me._Label10_2.AutoSize = False
		Me._Label10_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label10_2.Name = "_Label10_2"
		Me.lblPVAuthRemaining.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblPVAuthRemaining.Size = New System.Drawing.Size(57, 27)
		Me.lblPVAuthRemaining.Location = New System.Drawing.Point(410, 132)
		Me.lblPVAuthRemaining.TabIndex = 27
		Me.lblPVAuthRemaining.Tag = "1"
		Me.lblPVAuthRemaining.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPVAuthRemaining.BackColor = System.Drawing.SystemColors.Control
		Me.lblPVAuthRemaining.Enabled = True
		Me.lblPVAuthRemaining.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPVAuthRemaining.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPVAuthRemaining.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPVAuthRemaining.UseMnemonic = True
		Me.lblPVAuthRemaining.Visible = True
		Me.lblPVAuthRemaining.AutoSize = False
		Me.lblPVAuthRemaining.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblPVAuthRemaining.Name = "lblPVAuthRemaining"
		Me.lblPVExpires.Text = "Expires:"
		Me.lblPVExpires.Size = New System.Drawing.Size(133, 24)
		Me.lblPVExpires.Location = New System.Drawing.Point(352, 79)
		Me.lblPVExpires.TabIndex = 26
		Me.lblPVExpires.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPVExpires.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPVExpires.BackColor = System.Drawing.Color.Transparent
		Me.lblPVExpires.Enabled = True
		Me.lblPVExpires.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPVExpires.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPVExpires.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPVExpires.UseMnemonic = True
		Me.lblPVExpires.Visible = True
		Me.lblPVExpires.AutoSize = False
		Me.lblPVExpires.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPVExpires.Name = "lblPVExpires"
		Me.Label11.Text = "PowerClaim XML"
		Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label11.Size = New System.Drawing.Size(152, 22)
		Me.Label11.Location = New System.Drawing.Point(4, 8)
		Me.Label11.TabIndex = 25
		Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label11.BackColor = System.Drawing.Color.Transparent
		Me.Label11.Enabled = True
		Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label11.UseMnemonic = True
		Me.Label11.Visible = True
		Me.Label11.AutoSize = False
		Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label11.Name = "Label11"
		Me.Label17.Text = "Available Online Auths:"
		Me.Label17.Size = New System.Drawing.Size(142, 22)
		Me.Label17.Location = New System.Drawing.Point(140, 10)
		Me.Label17.TabIndex = 24
		Me.Label17.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label17.BackColor = System.Drawing.Color.Transparent
		Me.Label17.Enabled = True
		Me.Label17.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label17.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label17.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label17.UseMnemonic = True
		Me.Label17.Visible = True
		Me.Label17.AutoSize = False
		Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label17.Name = "Label17"
		Me.Label14.Text = "Sale Date:"
		Me.Label14.Size = New System.Drawing.Size(82, 24)
		Me.Label14.Location = New System.Drawing.Point(4, 42)
		Me.Label14.TabIndex = 23
		Me.Label14.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label14.BackColor = System.Drawing.Color.Transparent
		Me.Label14.Enabled = True
		Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label14.UseMnemonic = True
		Me.Label14.Visible = True
		Me.Label14.AutoSize = False
		Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label14.Name = "Label14"
		Me._Label10_8.Text = "Pending Days:"
		Me._Label10_8.Size = New System.Drawing.Size(87, 24)
		Me._Label10_8.Location = New System.Drawing.Point(330, 44)
		Me._Label10_8.TabIndex = 22
		Me._Label10_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label10_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label10_8.BackColor = System.Drawing.Color.Transparent
		Me._Label10_8.Enabled = True
		Me._Label10_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label10_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label10_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label10_8.UseMnemonic = True
		Me._Label10_8.Visible = True
		Me._Label10_8.AutoSize = False
		Me._Label10_8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label10_8.Name = "_Label10_8"
		Me._Label10_9.Text = "Grace Period:"
		Me._Label10_9.Size = New System.Drawing.Size(84, 24)
		Me._Label10_9.Location = New System.Drawing.Point(340, 10)
		Me._Label10_9.TabIndex = 21
		Me._Label10_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label10_9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label10_9.BackColor = System.Drawing.Color.Transparent
		Me._Label10_9.Enabled = True
		Me._Label10_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label10_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label10_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label10_9.UseMnemonic = True
		Me._Label10_9.Visible = True
		Me._Label10_9.AutoSize = False
		Me._Label10_9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label10_9.Name = "_Label10_9"
		Me._Label10_12.Text = "Sale Days:"
		Me._Label10_12.Size = New System.Drawing.Size(87, 24)
		Me._Label10_12.Location = New System.Drawing.Point(200, 44)
		Me._Label10_12.TabIndex = 20
		Me._Label10_12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label10_12.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label10_12.BackColor = System.Drawing.Color.Transparent
		Me._Label10_12.Enabled = True
		Me._Label10_12.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label10_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label10_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label10_12.UseMnemonic = True
		Me._Label10_12.Visible = True
		Me._Label10_12.AutoSize = False
		Me._Label10_12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label10_12.Name = "_Label10_12"
		Me.Timer4.Interval = 650
		Me.Timer4.Enabled = True
		Me.Timer3.Enabled = False
		Me.Timer3.Interval = 100
		Me.Timer2.Enabled = False
		Me.Timer2.Interval = 500
		Me.Timer1.Interval = 80
		Me.Timer1.Enabled = True
		Me.Label2.BackColor = System.Drawing.SystemColors.ActiveCaptionText
		Me.Label2.Text = " Over"
		Me.Label2.Font = New System.Drawing.Font("Arial", 24!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(132, 62)
		Me.Label2.Location = New System.Drawing.Point(290, 20)
		Me.Label2.TabIndex = 1
		Me.Label2.Visible = False
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
		Me.Label1.Text = "Game "
		Me.Label1.Font = New System.Drawing.Font("Arial", 24!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(142, 62)
		Me.Label1.Location = New System.Drawing.Point(140, 20)
		Me.Label1.TabIndex = 0
		Me.Label1.Visible = False
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Image9.Size = New System.Drawing.Size(35, 42)
		Me.Image9.Location = New System.Drawing.Point(100, 130)
		Me.Image9.Image = CType(resources.GetObject("Image9.Image"), System.Drawing.Image)
		Me.Image9.Visible = False
		Me.Image9.Enabled = True
		Me.Image9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image9.Name = "Image9"
		Me.Image8.Size = New System.Drawing.Size(35, 42)
		Me.Image8.Location = New System.Drawing.Point(170, 160)
		Me.Image8.Image = CType(resources.GetObject("Image8.Image"), System.Drawing.Image)
		Me.Image8.Visible = False
		Me.Image8.Enabled = True
		Me.Image8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image8.Name = "Image8"
		Me.Image7.Size = New System.Drawing.Size(35, 42)
		Me.Image7.Location = New System.Drawing.Point(10, 100)
		Me.Image7.Image = CType(resources.GetObject("Image7.Image"), System.Drawing.Image)
		Me.Image7.Visible = False
		Me.Image7.Enabled = True
		Me.Image7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image7.Name = "Image7"
		Me.Image6.Size = New System.Drawing.Size(35, 42)
		Me.Image6.Location = New System.Drawing.Point(60, 80)
		Me.Image6.Image = CType(resources.GetObject("Image6.Image"), System.Drawing.Image)
		Me.Image6.Visible = False
		Me.Image6.Enabled = True
		Me.Image6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image6.Name = "Image6"
		Me.Image5.Size = New System.Drawing.Size(35, 42)
		Me.Image5.Location = New System.Drawing.Point(10, 160)
		Me.Image5.Image = CType(resources.GetObject("Image5.Image"), System.Drawing.Image)
		Me.Image5.Visible = False
		Me.Image5.Enabled = True
		Me.Image5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image5.Name = "Image5"
		Me.Image4.Size = New System.Drawing.Size(47, 42)
		Me.Image4.Location = New System.Drawing.Point(470, 60)
		Me.Image4.Image = CType(resources.GetObject("Image4.Image"), System.Drawing.Image)
		Me.Image4.Visible = False
		Me.Image4.Enabled = True
		Me.Image4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image4.Name = "Image4"
		Me.Image3.Size = New System.Drawing.Size(47, 42)
		Me.Image3.Location = New System.Drawing.Point(530, 60)
		Me.Image3.Image = CType(resources.GetObject("Image3.Image"), System.Drawing.Image)
		Me.Image3.Enabled = True
		Me.Image3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image3.Visible = True
		Me.Image3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image3.Name = "Image3"
		Me.Image2.Size = New System.Drawing.Size(47, 42)
		Me.Image2.Location = New System.Drawing.Point(160, 250)
		Me.Image2.Image = CType(resources.GetObject("Image2.Image"), System.Drawing.Image)
		Me.Image2.Visible = False
		Me.Image2.Enabled = True
		Me.Image2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image2.Name = "Image2"
		Me.Image1.Size = New System.Drawing.Size(35, 42)
		Me.Image1.Location = New System.Drawing.Point(20, 250)
		Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
		Me.Image1.Enabled = True
		Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Image1.Visible = True
		Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Image1.Name = "Image1"
		Me.Shape8.BackColor = System.Drawing.Color.Transparent
		Me.Shape8.Size = New System.Drawing.Size(592, 302)
		Me.Shape8.Location = New System.Drawing.Point(0, 0)
		Me.Shape8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape8.Visible = True
		Me.Shape8.Name = "Shape8"
		Me.Shape7.BackColor = System.Drawing.Color.Black
		Me.Shape7.Size = New System.Drawing.Size(152, 42)
		Me.Shape7.Location = New System.Drawing.Point(370, 200)
		Me.Shape7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape7.Visible = True
		Me.Shape7.Name = "Shape7"
		Me.Shape6.Size = New System.Drawing.Size(42, 112)
		Me.Shape6.Location = New System.Drawing.Point(480, 130)
		Me.Shape6.BackColor = System.Drawing.Color.Black
		Me.Shape6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape6.Visible = True
		Me.Shape6.Name = "Shape6"
		Me.Shape5.BackColor = System.Drawing.Color.Black
		Me.Shape5.Size = New System.Drawing.Size(152, 42)
		Me.Shape5.Location = New System.Drawing.Point(260, 110)
		Me.Shape5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape5.Visible = True
		Me.Shape5.Name = "Shape5"
		Me.Shape4.Size = New System.Drawing.Size(42, 112)
		Me.Shape4.Location = New System.Drawing.Point(260, 130)
		Me.Shape4.BackColor = System.Drawing.Color.Black
		Me.Shape4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape4.Visible = True
		Me.Shape4.Name = "Shape4"
		Me.Shape3.BackColor = System.Drawing.Color.Black
		Me.Shape3.Size = New System.Drawing.Size(152, 42)
		Me.Shape3.Location = New System.Drawing.Point(150, 110)
		Me.Shape3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape3.Visible = True
		Me.Shape3.Name = "Shape3"
		Me.Shape2.BackColor = System.Drawing.Color.Black
		Me.Shape2.Size = New System.Drawing.Size(152, 42)
		Me.Shape2.Location = New System.Drawing.Point(50, 200)
		Me.Shape2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape2.Visible = True
		Me.Shape2.Name = "Shape2"
		Me.Shape1.Size = New System.Drawing.Size(42, 112)
		Me.Shape1.Location = New System.Drawing.Point(50, 130)
		Me.Shape1.BackColor = System.Drawing.Color.Black
		Me.Shape1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape1.Visible = True
		Me.Shape1.Name = "Shape1"
		Me.Controls.Add(fmePVAuthStatus)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Image9)
		Me.Controls.Add(Image8)
		Me.Controls.Add(Image7)
		Me.Controls.Add(Image6)
		Me.Controls.Add(Image5)
		Me.Controls.Add(Image4)
		Me.Controls.Add(Image3)
		Me.Controls.Add(Image2)
		Me.Controls.Add(Image1)
		Me.Controls.Add(Shape8)
		Me.Controls.Add(Shape7)
		Me.Controls.Add(Shape6)
		Me.Controls.Add(Shape5)
		Me.Controls.Add(Shape4)
		Me.Controls.Add(Shape3)
		Me.Controls.Add(Shape2)
		Me.Controls.Add(Shape1)
		Me.fmePVAuthStatus.Controls.Add(txtPVAuthDays)
		Me.fmePVAuthStatus.Controls.Add(txtPVVersionShipped)
		Me.fmePVAuthStatus.Controls.Add(txtPVAuths)
		Me.fmePVAuthStatus.Controls.Add(txtPVPendingDays)
		Me.fmePVAuthStatus.Controls.Add(txtPVGraceDays)
		Me.fmePVAuthStatus.Controls.Add(txtPVSaleDays)
		Me.fmePVAuthStatus.Controls.Add(cboPVAuthStatus)
		Me.fmePVAuthStatus.Controls.Add(cboPVShipStatus)
		Me.fmePVAuthStatus.Controls.Add(mskPVShipDate)
		Me.fmePVAuthStatus.Controls.Add(mskPVAuthDate)
		Me.fmePVAuthStatus.Controls.Add(cmdPVShipDate)
		Me.fmePVAuthStatus.Controls.Add(cmdPVAuthDate)
		Me.fmePVAuthStatus.Controls.Add(cboPVDownloadStatus)
		Me.fmePVAuthStatus.Controls.Add(mskPVDownloadDate)
		Me.fmePVAuthStatus.Controls.Add(cmdPVDownloadDate)
		Me.fmePVAuthStatus.Controls.Add(mskPVSaleDate)
		Me.fmePVAuthStatus.Controls.Add(cmdPVSalesDate)
		Me.fmePVAuthStatus.Controls.Add(_Label10_5)
		Me.fmePVAuthStatus.Controls.Add(_Label10_1)
		Me.fmePVAuthStatus.Controls.Add(_Label4_6)
		Me.fmePVAuthStatus.Controls.Add(_Label10_2)
		Me.fmePVAuthStatus.Controls.Add(lblPVAuthRemaining)
		Me.fmePVAuthStatus.Controls.Add(lblPVExpires)
		Me.fmePVAuthStatus.Controls.Add(Label11)
		Me.fmePVAuthStatus.Controls.Add(Label17)
		Me.fmePVAuthStatus.Controls.Add(Label14)
		Me.fmePVAuthStatus.Controls.Add(_Label10_8)
		Me.fmePVAuthStatus.Controls.Add(_Label10_9)
		Me.fmePVAuthStatus.Controls.Add(_Label10_12)
		Me.Label10.SetIndex(_Label10_5, CType(5, Short))
		Me.Label10.SetIndex(_Label10_1, CType(1, Short))
		Me.Label10.SetIndex(_Label10_2, CType(2, Short))
		Me.Label10.SetIndex(_Label10_8, CType(8, Short))
		Me.Label10.SetIndex(_Label10_9, CType(9, Short))
		Me.Label10.SetIndex(_Label10_12, CType(12, Short))
		Me.Label4.SetIndex(_Label4_6, CType(6, Short))
		CType(Me.Label4, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label10, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdPVSalesDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskPVSaleDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdPVDownloadDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskPVDownloadDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboPVDownloadStatus, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdPVAuthDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdPVShipDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskPVAuthDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.mskPVShipDate, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboPVShipStatus, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cboPVAuthStatus, System.ComponentModel.ISupportInitialize).EndInit()
		Me.fmePVAuthStatus.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class