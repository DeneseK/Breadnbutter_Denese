<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FResult
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
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents TextResult As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FResult))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdClose = New System.Windows.Forms.Button
		Me.TextResult = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Result"
		Me.ClientSize = New System.Drawing.Size(520, 219)
		Me.Location = New System.Drawing.Point(4, 28)
		Me.Icon = CType(resources.GetObject("FResult.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FResult"
		Me.cmdClose.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdClose.Text = "Close"
		Me.cmdClose.Size = New System.Drawing.Size(169, 29)
		Me.cmdClose.Location = New System.Drawing.Point(173, 185)
		Me.cmdClose.TabIndex = 1
		Me.cmdClose.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClose.CausesValidation = True
		Me.cmdClose.Enabled = True
		Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClose.TabStop = True
		Me.cmdClose.Name = "cmdClose"
		Me.TextResult.AutoSize = False
		Me.TextResult.Size = New System.Drawing.Size(512, 177)
		Me.TextResult.Location = New System.Drawing.Point(5, 3)
		Me.TextResult.ReadOnly = True
		Me.TextResult.MultiLine = True
		Me.TextResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.TextResult.TabIndex = 0
		Me.TextResult.Text = "Text1" & Chr(13) & Chr(10)
		Me.TextResult.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextResult.AcceptsReturn = True
		Me.TextResult.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextResult.BackColor = System.Drawing.SystemColors.Window
		Me.TextResult.CausesValidation = True
		Me.TextResult.Enabled = True
		Me.TextResult.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextResult.HideSelection = True
		Me.TextResult.Maxlength = 0
		Me.TextResult.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextResult.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextResult.TabStop = True
		Me.TextResult.Visible = True
		Me.TextResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextResult.Name = "TextResult"
		Me.Controls.Add(cmdClose)
		Me.Controls.Add(TextResult)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class