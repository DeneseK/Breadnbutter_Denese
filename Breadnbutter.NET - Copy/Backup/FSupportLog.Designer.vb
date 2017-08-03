<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FSupportLog
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
	Public WithEvents cmdFirst As System.Windows.Forms.Button
	Public WithEvents cmdLast As System.Windows.Forms.Button
	Public WithEvents cboShow As System.Windows.Forms.ComboBox
	Public WithEvents grdSupportLog As AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
	Public WithEvents lblShow As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FSupportLog))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdFirst = New System.Windows.Forms.Button
		Me.cmdLast = New System.Windows.Forms.Button
		Me.cboShow = New System.Windows.Forms.ComboBox
		Me.grdSupportLog = New AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
		Me.lblShow = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdSupportLog, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Support Log"
		Me.ClientSize = New System.Drawing.Size(793, 450)
		Me.Location = New System.Drawing.Point(318, 269)
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
		Me.Name = "FSupportLog"
		Me.cmdFirst.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdFirst.Text = "Go to First Record"
		Me.cmdFirst.Size = New System.Drawing.Size(139, 27)
		Me.cmdFirst.Location = New System.Drawing.Point(233, 418)
		Me.cmdFirst.TabIndex = 4
		Me.cmdFirst.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdFirst.BackColor = System.Drawing.SystemColors.Control
		Me.cmdFirst.CausesValidation = True
		Me.cmdFirst.Enabled = True
		Me.cmdFirst.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdFirst.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdFirst.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdFirst.TabStop = True
		Me.cmdFirst.Name = "cmdFirst"
		Me.cmdLast.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdLast.Text = "Go to Last Record"
		Me.cmdLast.Size = New System.Drawing.Size(139, 27)
		Me.cmdLast.Location = New System.Drawing.Point(375, 418)
		Me.cmdLast.TabIndex = 3
		Me.cmdLast.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdLast.BackColor = System.Drawing.SystemColors.Control
		Me.cmdLast.CausesValidation = True
		Me.cmdLast.Enabled = True
		Me.cmdLast.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdLast.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdLast.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdLast.TabStop = True
		Me.cmdLast.Name = "cmdLast"
		Me.cboShow.Size = New System.Drawing.Size(167, 27)
		Me.cboShow.Location = New System.Drawing.Point(60, 418)
		Me.cboShow.Items.AddRange(New Object(){"Today", "Previous 7 days", "All"})
		Me.cboShow.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboShow.TabIndex = 1
		Me.cboShow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboShow.BackColor = System.Drawing.SystemColors.Window
		Me.cboShow.CausesValidation = True
		Me.cboShow.Enabled = True
		Me.cboShow.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboShow.IntegralHeight = True
		Me.cboShow.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboShow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboShow.Sorted = False
		Me.cboShow.TabStop = True
		Me.cboShow.Visible = True
		Me.cboShow.Name = "cboShow"
		grdSupportLog.OcxState = CType(resources.GetObject("grdSupportLog.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdSupportLog.Size = New System.Drawing.Size(762, 414)
		Me.grdSupportLog.Location = New System.Drawing.Point(0, 0)
		Me.grdSupportLog.TabIndex = 0
		Me.grdSupportLog.Name = "grdSupportLog"
		Me.lblShow.Text = "Show:"
		Me.lblShow.Size = New System.Drawing.Size(54, 22)
		Me.lblShow.Location = New System.Drawing.Point(3, 420)
		Me.lblShow.TabIndex = 2
		Me.lblShow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblShow.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblShow.BackColor = System.Drawing.SystemColors.Control
		Me.lblShow.Enabled = True
		Me.lblShow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblShow.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblShow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblShow.UseMnemonic = True
		Me.lblShow.Visible = True
		Me.lblShow.AutoSize = False
		Me.lblShow.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblShow.Name = "lblShow"
		CType(Me.grdSupportLog, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(cmdFirst)
		Me.Controls.Add(cmdLast)
		Me.Controls.Add(cboShow)
		Me.Controls.Add(grdSupportLog)
		Me.Controls.Add(lblShow)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class