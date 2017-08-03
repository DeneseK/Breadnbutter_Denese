<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FSendTo
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
	Public WithEvents cboSendTo As System.Windows.Forms.ComboBox
	Public WithEvents lblEmailAddress As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdSend As System.Windows.Forms.Button
	Public WithEvents Shape1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FSendTo))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cboSendTo = New System.Windows.Forms.ComboBox
		Me.lblEmailAddress = New System.Windows.Forms.Label
		Me.cmdSend = New System.Windows.Forms.Button
		Me.Shape1 = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Forward Message To:"
		Me.ClientSize = New System.Drawing.Size(408, 158)
		Me.Location = New System.Drawing.Point(5, 29)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
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
		Me.Name = "FSendTo"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(174, 42)
		Me.cmdCancel.Location = New System.Drawing.Point(210, 110)
		Me.cmdCancel.TabIndex = 4
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.Frame1.BackColor = System.Drawing.Color.White
		Me.Frame1.Text = "Send To:"
		Me.Frame1.Size = New System.Drawing.Size(382, 92)
		Me.Frame1.Location = New System.Drawing.Point(10, 10)
		Me.Frame1.TabIndex = 1
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me.cboSendTo.Size = New System.Drawing.Size(232, 27)
		Me.cboSendTo.Location = New System.Drawing.Point(80, 20)
		Me.cboSendTo.TabIndex = 2
		Me.cboSendTo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboSendTo.BackColor = System.Drawing.SystemColors.Window
		Me.cboSendTo.CausesValidation = True
		Me.cboSendTo.Enabled = True
		Me.cboSendTo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboSendTo.IntegralHeight = True
		Me.cboSendTo.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboSendTo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboSendTo.Sorted = False
		Me.cboSendTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboSendTo.TabStop = True
		Me.cboSendTo.Visible = True
		Me.cboSendTo.Name = "cboSendTo"
		Me.lblEmailAddress.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblEmailAddress.BackColor = System.Drawing.Color.White
		Me.lblEmailAddress.Size = New System.Drawing.Size(232, 32)
		Me.lblEmailAddress.Location = New System.Drawing.Point(80, 50)
		Me.lblEmailAddress.TabIndex = 3
		Me.lblEmailAddress.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblEmailAddress.Enabled = True
		Me.lblEmailAddress.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblEmailAddress.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblEmailAddress.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEmailAddress.UseMnemonic = True
		Me.lblEmailAddress.Visible = True
		Me.lblEmailAddress.AutoSize = False
		Me.lblEmailAddress.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblEmailAddress.Name = "lblEmailAddress"
		Me.cmdSend.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSend.Text = "&Send"
		Me.cmdSend.Size = New System.Drawing.Size(172, 42)
		Me.cmdSend.Location = New System.Drawing.Point(10, 110)
		Me.cmdSend.TabIndex = 0
		Me.cmdSend.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSend.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSend.CausesValidation = True
		Me.cmdSend.Enabled = True
		Me.cmdSend.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSend.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSend.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSend.TabStop = True
		Me.cmdSend.Name = "cmdSend"
		Me.Shape1.Size = New System.Drawing.Size(407, 107)
		Me.Shape1.Location = New System.Drawing.Point(0, 0)
		Me.Shape1.BackColor = System.Drawing.Color.White
		Me.Shape1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Shape1.Visible = True
		Me.Shape1.Name = "Shape1"
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdSend)
		Me.Controls.Add(Shape1)
		Me.Frame1.Controls.Add(cboSendTo)
		Me.Frame1.Controls.Add(lblEmailAddress)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class