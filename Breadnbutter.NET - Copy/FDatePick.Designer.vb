<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FDatePick
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
	Public WithEvents cmdSet As System.Windows.Forms.Button
	Public WithEvents Calendar1 As AxMSComCtl2.AxMonthView
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FDatePick))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdSet = New System.Windows.Forms.Button
		Me.Calendar1 = New AxMSComCtl2.AxMonthView
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Calendar1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Text = "Double Click to Set"
		Me.ClientSize = New System.Drawing.Size(225, 230)
		Me.Location = New System.Drawing.Point(542, 395)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "FDatePick"
		Me.cmdSet.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSet.Text = "Set"
		Me.cmdSet.Size = New System.Drawing.Size(225, 29)
		Me.cmdSet.Location = New System.Drawing.Point(0, 200)
		Me.cmdSet.TabIndex = 1
		Me.cmdSet.TabStop = False
		Me.cmdSet.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSet.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSet.CausesValidation = True
		Me.cmdSet.Enabled = True
		Me.cmdSet.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSet.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSet.Name = "cmdSet"
		Calendar1.OcxState = CType(resources.GetObject("Calendar1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.Calendar1.Size = New System.Drawing.Size(225, 198)
		Me.Calendar1.Location = New System.Drawing.Point(0, 0)
		Me.Calendar1.TabIndex = 0
		Me.Calendar1.Name = "Calendar1"
		Me.Controls.Add(cmdSet)
		Me.Controls.Add(Calendar1)
		CType(Me.Calendar1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class