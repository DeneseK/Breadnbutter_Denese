<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FSelect
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
	Public WithEvents _cmdSelect_1 As System.Windows.Forms.Button
	Public WithEvents _cmdSelect_0 As System.Windows.Forms.Button
	Public WithEvents Data1 As System.Windows.Forms.Label
	Public WithEvents grdSelect As AxSSDataWidgets_B.AxSSDBGrid
	Public WithEvents cmdSelect As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FSelect))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me._cmdSelect_1 = New System.Windows.Forms.Button
		Me._cmdSelect_0 = New System.Windows.Forms.Button
		Me.Data1 = New System.Windows.Forms.Label
		Me.grdSelect = New AxSSDataWidgets_B.AxSSDBGrid
		Me.cmdSelect = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdSelect, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdSelect, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Select Customers for Processing"
		Me.ClientSize = New System.Drawing.Size(753, 437)
		Me.Location = New System.Drawing.Point(85, 203)
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
		Me.Name = "FSelect"
		Me._cmdSelect_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdSelect_1.Text = "Deselect All"
		Me._cmdSelect_1.Size = New System.Drawing.Size(104, 27)
		Me._cmdSelect_1.Location = New System.Drawing.Point(633, 23)
		Me._cmdSelect_1.TabIndex = 2
		Me._cmdSelect_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdSelect_1.BackColor = System.Drawing.SystemColors.Control
		Me._cmdSelect_1.CausesValidation = True
		Me._cmdSelect_1.Enabled = True
		Me._cmdSelect_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdSelect_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdSelect_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdSelect_1.TabStop = True
		Me._cmdSelect_1.Name = "_cmdSelect_1"
		Me._cmdSelect_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdSelect_0.Text = "Select All"
		Me._cmdSelect_0.Size = New System.Drawing.Size(107, 27)
		Me._cmdSelect_0.Location = New System.Drawing.Point(518, 23)
		Me._cmdSelect_0.TabIndex = 1
		Me._cmdSelect_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdSelect_0.BackColor = System.Drawing.SystemColors.Control
		Me._cmdSelect_0.CausesValidation = True
		Me._cmdSelect_0.Enabled = True
		Me._cmdSelect_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdSelect_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdSelect_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdSelect_0.TabStop = True
		Me._cmdSelect_0.Name = "_cmdSelect_0"
		Me.Data1.Text = "Data1"
		Me.Data1.Enabled = False
		Me.Data1.Size = New System.Drawing.Size(157, 29)
		Me.Data1.Location = New System.Drawing.Point(540, 408)
		Me.Data1.Visible = False
		Me.Data1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Data1.BackColor = System.Drawing.Color.Red
		Me.Data1.ForeColor = System.Drawing.Color.Black
		Me.Data1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Data1.Text = "Data1"
		Me.Data1.Name = "Data1"
		grdSelect.OcxState = CType(resources.GetObject("grdSelect.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdSelect.Size = New System.Drawing.Size(722, 362)
		Me.grdSelect.Location = New System.Drawing.Point(13, 53)
		Me.grdSelect.TabIndex = 0
		Me.grdSelect.Name = "grdSelect"
		Me.cmdSelect.SetIndex(_cmdSelect_1, CType(1, Short))
		Me.cmdSelect.SetIndex(_cmdSelect_0, CType(0, Short))
		CType(Me.cmdSelect, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grdSelect, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(_cmdSelect_1)
		Me.Controls.Add(_cmdSelect_0)
		Me.Controls.Add(Data1)
		Me.Controls.Add(grdSelect)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class