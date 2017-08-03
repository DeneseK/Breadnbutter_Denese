<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FDailyReports
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
	Public WithEvents PrintButton As System.Windows.Forms.Button
	Public WithEvents RefreshData As System.Windows.Forms.Button
	Public WithEvents ExitButton As System.Windows.Forms.Button
	Public WithEvents Preview As System.Windows.Forms.Button
	Public WithEvents ListView1 As System.Windows.Forms.ListView
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FDailyReports))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.PrintButton = New System.Windows.Forms.Button
		Me.RefreshData = New System.Windows.Forms.Button
		Me.ExitButton = New System.Windows.Forms.Button
		Me.Preview = New System.Windows.Forms.Button
		Me.ListView1 = New System.Windows.Forms.ListView
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.Text = "Daily Report"
		Me.ClientSize = New System.Drawing.Size(823, 497)
		Me.Location = New System.Drawing.Point(5, 29)
		Me.Icon = CType(resources.GetObject("FDailyReports.Icon"), System.Drawing.Icon)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
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
		Me.Name = "FDailyReports"
		Me.PrintButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.PrintButton.Text = "Print Report"
		Me.PrintButton.Size = New System.Drawing.Size(169, 27)
		Me.PrintButton.Location = New System.Drawing.Point(343, 233)
		Me.PrintButton.TabIndex = 3
		Me.PrintButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.PrintButton.BackColor = System.Drawing.SystemColors.Control
		Me.PrintButton.CausesValidation = True
		Me.PrintButton.Enabled = True
		Me.PrintButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.PrintButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.PrintButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.PrintButton.TabStop = True
		Me.PrintButton.Name = "PrintButton"
		Me.RefreshData.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.RefreshData.Text = "Refresh Data"
		Me.RefreshData.Size = New System.Drawing.Size(172, 27)
		Me.RefreshData.Location = New System.Drawing.Point(0, 233)
		Me.RefreshData.TabIndex = 1
		Me.RefreshData.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.RefreshData.BackColor = System.Drawing.SystemColors.Control
		Me.RefreshData.CausesValidation = True
		Me.RefreshData.Enabled = True
		Me.RefreshData.ForeColor = System.Drawing.SystemColors.ControlText
		Me.RefreshData.Cursor = System.Windows.Forms.Cursors.Default
		Me.RefreshData.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.RefreshData.TabStop = True
		Me.RefreshData.Name = "RefreshData"
		Me.ExitButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.ExitButton.Text = "Exit"
		Me.ExitButton.Size = New System.Drawing.Size(169, 27)
		Me.ExitButton.Location = New System.Drawing.Point(510, 233)
		Me.ExitButton.TabIndex = 4
		Me.ExitButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ExitButton.BackColor = System.Drawing.SystemColors.Control
		Me.ExitButton.CausesValidation = True
		Me.ExitButton.Enabled = True
		Me.ExitButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ExitButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.ExitButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ExitButton.TabStop = True
		Me.ExitButton.Name = "ExitButton"
		Me.Preview.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Preview.Text = "Preview Report"
		Me.Preview.Size = New System.Drawing.Size(169, 27)
		Me.Preview.Location = New System.Drawing.Point(173, 233)
		Me.Preview.TabIndex = 2
		Me.Preview.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Preview.BackColor = System.Drawing.SystemColors.Control
		Me.Preview.CausesValidation = True
		Me.Preview.Enabled = True
		Me.Preview.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Preview.Cursor = System.Windows.Forms.Cursors.Default
		Me.Preview.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Preview.TabStop = True
		Me.Preview.Name = "Preview"
		Me.ListView1.Size = New System.Drawing.Size(492, 172)
		Me.ListView1.Location = New System.Drawing.Point(0, 5)
		Me.ListView1.TabIndex = 0
		Me.ListView1.TabStop = 0
		Me.ListView1.View = System.Windows.Forms.View.Details
		Me.ListView1.LabelWrap = True
		Me.ListView1.HideSelection = True
		Me.ListView1.FullRowSelect = True
		Me.ListView1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ListView1.BackColor = System.Drawing.SystemColors.Window
		Me.ListView1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ListView1.LabelEdit = True
		Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ListView1.Name = "ListView1"
		Me.Controls.Add(PrintButton)
		Me.Controls.Add(RefreshData)
		Me.Controls.Add(ExitButton)
		Me.Controls.Add(Preview)
		Me.Controls.Add(ListView1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class