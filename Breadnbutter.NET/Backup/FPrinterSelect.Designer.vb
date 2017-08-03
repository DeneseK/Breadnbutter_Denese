<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FPrinterSelect
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
	Public WithEvents cboPrinter As System.Windows.Forms.ComboBox
	Public WithEvents cmdPrint As System.Windows.Forms.Button
	Public WithEvents txtCopies As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FPrinterSelect))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cboPrinter = New System.Windows.Forms.ComboBox
		Me.cmdPrint = New System.Windows.Forms.Button
		Me.txtCopies = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Select Printer"
		Me.ClientSize = New System.Drawing.Size(447, 102)
		Me.Location = New System.Drawing.Point(428, 324)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
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
		Me.Name = "FPrinterSelect"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(92, 32)
		Me.cmdCancel.Location = New System.Drawing.Point(240, 60)
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
		Me.cboPrinter.Size = New System.Drawing.Size(372, 27)
		Me.cboPrinter.Location = New System.Drawing.Point(10, 20)
		Me.cboPrinter.TabIndex = 3
		Me.cboPrinter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboPrinter.BackColor = System.Drawing.SystemColors.Window
		Me.cboPrinter.CausesValidation = True
		Me.cboPrinter.Enabled = True
		Me.cboPrinter.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboPrinter.IntegralHeight = True
		Me.cboPrinter.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboPrinter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboPrinter.Sorted = False
		Me.cboPrinter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboPrinter.TabStop = True
		Me.cboPrinter.Visible = True
		Me.cboPrinter.Name = "cboPrinter"
		Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPrint.Text = "Print"
		Me.cmdPrint.Size = New System.Drawing.Size(92, 32)
		Me.cmdPrint.Location = New System.Drawing.Point(110, 60)
		Me.cmdPrint.TabIndex = 1
		Me.cmdPrint.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPrint.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPrint.CausesValidation = True
		Me.cmdPrint.Enabled = True
		Me.cmdPrint.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPrint.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPrint.TabStop = True
		Me.cmdPrint.Name = "cmdPrint"
		Me.txtCopies.AutoSize = False
		Me.txtCopies.Size = New System.Drawing.Size(32, 24)
		Me.txtCopies.Location = New System.Drawing.Point(400, 20)
		Me.txtCopies.Maxlength = 2
		Me.txtCopies.TabIndex = 0
		Me.txtCopies.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCopies.AcceptsReturn = True
		Me.txtCopies.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCopies.BackColor = System.Drawing.SystemColors.Window
		Me.txtCopies.CausesValidation = True
		Me.txtCopies.Enabled = True
		Me.txtCopies.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtCopies.HideSelection = True
		Me.txtCopies.ReadOnly = False
		Me.txtCopies.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCopies.MultiLine = False
		Me.txtCopies.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCopies.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCopies.TabStop = True
		Me.txtCopies.Visible = True
		Me.txtCopies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCopies.Name = "txtCopies"
		Me.Label2.Text = "Copies"
		Me.Label2.Size = New System.Drawing.Size(40, 17)
		Me.Label2.Location = New System.Drawing.Point(400, 0)
		Me.Label2.TabIndex = 4
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = True
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "Select Printer"
		Me.Label1.Size = New System.Drawing.Size(79, 17)
		Me.Label1.Location = New System.Drawing.Point(10, 0)
		Me.Label1.TabIndex = 2
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = True
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cboPrinter)
		Me.Controls.Add(cmdPrint)
		Me.Controls.Add(txtCopies)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class