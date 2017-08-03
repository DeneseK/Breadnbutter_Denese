<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FUtility
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
	Public WithEvents cmdScanForLicenses As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FUtility))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdScanForLicenses = New System.Windows.Forms.Button
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(612, 343)
		Me.Location = New System.Drawing.Point(5, 29)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FUtility"
		Me.cmdScanForLicenses.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdScanForLicenses.Text = "Scan for Licenses"
		Me.cmdScanForLicenses.Size = New System.Drawing.Size(147, 27)
		Me.cmdScanForLicenses.Location = New System.Drawing.Point(58, 55)
		Me.cmdScanForLicenses.TabIndex = 0
		Me.cmdScanForLicenses.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdScanForLicenses.BackColor = System.Drawing.SystemColors.Control
		Me.cmdScanForLicenses.CausesValidation = True
		Me.cmdScanForLicenses.Enabled = True
		Me.cmdScanForLicenses.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdScanForLicenses.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdScanForLicenses.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdScanForLicenses.TabStop = True
		Me.cmdScanForLicenses.Name = "cmdScanForLicenses"
		Me.Controls.Add(cmdScanForLicenses)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class