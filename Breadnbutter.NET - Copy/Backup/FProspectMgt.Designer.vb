<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FProspectMgt
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
	Public WithEvents chkTtls As System.Windows.Forms.CheckBox
	Public WithEvents _optSort_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optSort_0 As System.Windows.Forms.RadioButton
	Public WithEvents _chkFilter_1 As System.Windows.Forms.CheckBox
	Public WithEvents _chkFilter_0 As System.Windows.Forms.CheckBox
	Public WithEvents grdProspectGroup As AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
	Public WithEvents lstGroups As System.Windows.Forms.ListBox
	Public WithEvents grdHistory As AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents chkFilter As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents optSort As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FProspectMgt))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkTtls = New System.Windows.Forms.CheckBox
		Me._optSort_1 = New System.Windows.Forms.RadioButton
		Me._optSort_0 = New System.Windows.Forms.RadioButton
		Me._chkFilter_1 = New System.Windows.Forms.CheckBox
		Me._chkFilter_0 = New System.Windows.Forms.CheckBox
		Me.grdProspectGroup = New AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
		Me.lstGroups = New System.Windows.Forms.ListBox
		Me.grdHistory = New AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.chkFilter = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.optSort = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdProspectGroup, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.grdHistory, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkFilter, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optSort, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.ClientSize = New System.Drawing.Size(970, 594)
		Me.Location = New System.Drawing.Point(108, 134)
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
		Me.Name = "FProspectMgt"
		Me.chkTtls.Text = "Show &Totals"
		Me.chkTtls.Size = New System.Drawing.Size(99, 19)
		Me.chkTtls.Location = New System.Drawing.Point(190, 383)
		Me.chkTtls.TabIndex = 8
		Me.chkTtls.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTtls.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTtls.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTtls.BackColor = System.Drawing.SystemColors.Control
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
		Me._optSort_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optSort_1.Text = "Sort by &Label"
		Me._optSort_1.Size = New System.Drawing.Size(109, 22)
		Me._optSort_1.Location = New System.Drawing.Point(148, 410)
		Me._optSort_1.TabIndex = 7
		Me._optSort_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optSort_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optSort_1.BackColor = System.Drawing.SystemColors.Control
		Me._optSort_1.CausesValidation = True
		Me._optSort_1.Enabled = True
		Me._optSort_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optSort_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optSort_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optSort_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optSort_1.TabStop = True
		Me._optSort_1.Checked = False
		Me._optSort_1.Visible = True
		Me._optSort_1.Name = "_optSort_1"
		Me._optSort_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optSort_0.Text = "Sort by &Group"
		Me._optSort_0.Size = New System.Drawing.Size(112, 22)
		Me._optSort_0.Location = New System.Drawing.Point(20, 410)
		Me._optSort_0.TabIndex = 6
		Me._optSort_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optSort_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optSort_0.BackColor = System.Drawing.SystemColors.Control
		Me._optSort_0.CausesValidation = True
		Me._optSort_0.Enabled = True
		Me._optSort_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optSort_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optSort_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optSort_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optSort_0.TabStop = True
		Me._optSort_0.Checked = False
		Me._optSort_0.Visible = True
		Me._optSort_0.Name = "_optSort_0"
		Me._chkFilter_1.Text = "AM &Best"
		Me._chkFilter_1.Size = New System.Drawing.Size(77, 19)
		Me._chkFilter_1.Location = New System.Drawing.Point(108, 383)
		Me._chkFilter_1.TabIndex = 5
		Me._chkFilter_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkFilter_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkFilter_1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkFilter_1.BackColor = System.Drawing.SystemColors.Control
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
		Me._chkFilter_0.Text = "&Standard"
		Me._chkFilter_0.Size = New System.Drawing.Size(82, 19)
		Me._chkFilter_0.Location = New System.Drawing.Point(20, 383)
		Me._chkFilter_0.TabIndex = 4
		Me._chkFilter_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkFilter_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkFilter_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkFilter_0.BackColor = System.Drawing.SystemColors.Control
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
		grdProspectGroup.OcxState = CType(resources.GetObject("grdProspectGroup.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdProspectGroup.Size = New System.Drawing.Size(664, 394)
		Me.grdProspectGroup.Location = New System.Drawing.Point(303, 38)
		Me.grdProspectGroup.TabIndex = 3
		Me.grdProspectGroup.Name = "grdProspectGroup"
		Me.lstGroups.Size = New System.Drawing.Size(272, 334)
		Me.lstGroups.Location = New System.Drawing.Point(18, 38)
		Me.lstGroups.TabIndex = 0
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
		grdHistory.OcxState = CType(resources.GetObject("grdHistory.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdHistory.Size = New System.Drawing.Size(949, 144)
		Me.grdHistory.Location = New System.Drawing.Point(18, 443)
		Me.grdHistory.TabIndex = 9
		Me.grdHistory.Name = "grdHistory"
		Me._Label1_1.Text = "Current Group Members"
		Me._Label1_1.Size = New System.Drawing.Size(254, 22)
		Me._Label1_1.Location = New System.Drawing.Point(305, 13)
		Me._Label1_1.TabIndex = 2
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_0.Text = "Prospecting Groups"
		Me._Label1_0.Size = New System.Drawing.Size(247, 22)
		Me._Label1_0.Location = New System.Drawing.Point(20, 13)
		Me._Label1_0.TabIndex = 1
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.chkFilter.SetIndex(_chkFilter_1, CType(1, Short))
		Me.chkFilter.SetIndex(_chkFilter_0, CType(0, Short))
		Me.optSort.SetIndex(_optSort_1, CType(1, Short))
		Me.optSort.SetIndex(_optSort_0, CType(0, Short))
		CType(Me.optSort, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chkFilter, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdHistory, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdProspectGroup, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(chkTtls)
		Me.Controls.Add(_optSort_1)
		Me.Controls.Add(_optSort_0)
		Me.Controls.Add(_chkFilter_1)
		Me.Controls.Add(_chkFilter_0)
		Me.Controls.Add(grdProspectGroup)
		Me.Controls.Add(lstGroups)
		Me.Controls.Add(grdHistory)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_Label1_0)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class