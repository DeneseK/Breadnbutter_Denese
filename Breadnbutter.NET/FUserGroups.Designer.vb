<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FUserGroups
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
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents chkOperator As System.Windows.Forms.CheckBox
	Public WithEvents chkSupport As System.Windows.Forms.CheckBox
	Public WithEvents chkSales As System.Windows.Forms.CheckBox
	Public WithEvents chkAuthorizations As System.Windows.Forms.CheckBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FUserGroups))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdSave = New System.Windows.Forms.Button
		Me.chkOperator = New System.Windows.Forms.CheckBox
		Me.chkSupport = New System.Windows.Forms.CheckBox
		Me.chkSales = New System.Windows.Forms.CheckBox
		Me.chkAuthorizations = New System.Windows.Forms.CheckBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Select your groups"
		Me.ClientSize = New System.Drawing.Size(248, 212)
		Me.Location = New System.Drawing.Point(482, 364)
		Me.ControlBox = False
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FUserGroups"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(142, 32)
		Me.cmdCancel.Location = New System.Drawing.Point(50, 170)
		Me.cmdCancel.TabIndex = 5
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save Changes"
		Me.cmdSave.Size = New System.Drawing.Size(142, 32)
		Me.cmdSave.Location = New System.Drawing.Point(50, 130)
		Me.cmdSave.TabIndex = 4
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.chkOperator.Text = "Operator"
		Me.chkOperator.Size = New System.Drawing.Size(182, 22)
		Me.chkOperator.Location = New System.Drawing.Point(30, 100)
		Me.chkOperator.TabIndex = 3
		Me.chkOperator.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkOperator.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkOperator.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkOperator.BackColor = System.Drawing.SystemColors.Control
		Me.chkOperator.CausesValidation = True
		Me.chkOperator.Enabled = True
		Me.chkOperator.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkOperator.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkOperator.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkOperator.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkOperator.TabStop = True
		Me.chkOperator.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkOperator.Visible = True
		Me.chkOperator.Name = "chkOperator"
		Me.chkSupport.Text = "Support"
		Me.chkSupport.Size = New System.Drawing.Size(182, 22)
		Me.chkSupport.Location = New System.Drawing.Point(30, 70)
		Me.chkSupport.TabIndex = 2
		Me.chkSupport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSupport.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSupport.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSupport.BackColor = System.Drawing.SystemColors.Control
		Me.chkSupport.CausesValidation = True
		Me.chkSupport.Enabled = True
		Me.chkSupport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSupport.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSupport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSupport.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSupport.TabStop = True
		Me.chkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSupport.Visible = True
		Me.chkSupport.Name = "chkSupport"
		Me.chkSales.Text = "Sales"
		Me.chkSales.Size = New System.Drawing.Size(182, 22)
		Me.chkSales.Location = New System.Drawing.Point(30, 40)
		Me.chkSales.TabIndex = 1
		Me.chkSales.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSales.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSales.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSales.BackColor = System.Drawing.SystemColors.Control
		Me.chkSales.CausesValidation = True
		Me.chkSales.Enabled = True
		Me.chkSales.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSales.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSales.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSales.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSales.TabStop = True
		Me.chkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSales.Visible = True
		Me.chkSales.Name = "chkSales"
		Me.chkAuthorizations.Text = "Authorizations"
		Me.chkAuthorizations.Size = New System.Drawing.Size(182, 22)
		Me.chkAuthorizations.Location = New System.Drawing.Point(30, 10)
		Me.chkAuthorizations.TabIndex = 0
		Me.chkAuthorizations.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAuthorizations.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAuthorizations.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAuthorizations.BackColor = System.Drawing.SystemColors.Control
		Me.chkAuthorizations.CausesValidation = True
		Me.chkAuthorizations.Enabled = True
		Me.chkAuthorizations.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAuthorizations.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAuthorizations.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAuthorizations.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAuthorizations.TabStop = True
		Me.chkAuthorizations.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkAuthorizations.Visible = True
		Me.chkAuthorizations.Name = "chkAuthorizations"
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(chkOperator)
		Me.Controls.Add(chkSupport)
		Me.Controls.Add(chkSales)
		Me.Controls.Add(chkAuthorizations)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class