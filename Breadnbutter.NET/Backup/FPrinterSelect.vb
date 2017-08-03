Option Strict Off
Option Explicit On
Friend Class FPrinterSelect
	Inherits System.Windows.Forms.Form
	Public bPrintCancel As Boolean
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		bPrintCancel = True
		Me.Hide()
	End Sub
	
	Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
		If cboPrinter.Text <> "" Then
			sPrinterName = cboPrinter.Text
			iNumofCopies = Val(txtCopies.Text)
		Else
			'UPGRADE_ISSUE: Printer property Printer.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			sPrinterName = Printer.DeviceName
			iNumofCopies = Val(txtCopies.Text)
		End If
		bPrintCancel = False
		Me.Hide()
	End Sub
	
	'UPGRADE_WARNING: Form event FPrinterSelect.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FPrinterSelect_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		txtCopies.Text = CStr(1)
	End Sub
	
	Private Sub FPrinterSelect_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim prt As Printer
		Dim X As Short
		X = 0
		'UPGRADE_ISSUE: Printers object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		For	Each prt In Printers
			'UPGRADE_ISSUE: Printer property prt.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			cboPrinter.Items.Insert(X, prt.DeviceName)
			X = X + 1
		Next prt
		cboPrinter.SelectedIndex = 0
		txtCopies.Text = CStr(1)
	End Sub
End Class