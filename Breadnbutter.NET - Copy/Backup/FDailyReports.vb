Option Strict Off
Option Explicit On
Friend Class FDailyReports
	Inherits System.Windows.Forms.Form
	Private Report As New CReport
	'
	Private Sub ExitButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ExitButton.Click
		On Error GoTo EH
		'
		Me.Close()
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FDailyReports.Form_Load.")
	End Sub
	'
	Private Sub FDailyReports_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo EH
		'
		'Set Report = New CReport
		Report.Rtype = CReport.ReportType.daily
		RefreshList()
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FDailyReports.Form_Load.")
	End Sub
	'
	Private Sub RefreshList()
		On Error GoTo EH
		'
		ListView1.Items.Clear()
		ListView1.Columns.Clear()
		Report.FillList((Report.rsReport), ListView1)
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FDailyReports.RefreshList.")
	End Sub
	'
	Private Sub SetupList(ByRef list As System.Windows.Forms.ListView, ByRef rs As ADODB.Recordset)
		On Error GoTo EH
		'
		list.Items.Clear()
		list.Columns.Clear()
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FDailyReports.SetupList.")
	End Sub
	'
	'UPGRADE_WARNING: Event FDailyReports.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FDailyReports_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error GoTo EH
		'
		ListView1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 100)
		'
		If VB6.PixelsToTwipsY(Me.Height) > 1000 Then
			'
			ListView1.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 750)
			Me.RefreshData.SetBounds(0, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 700), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			Me.Preview.SetBounds(VB6.TwipsToPixelsX(2030), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 700), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			Me.ExitButton.SetBounds(VB6.TwipsToPixelsX(6090), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 700), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			Me.PrintButton.SetBounds(VB6.TwipsToPixelsX(4060), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 700), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			'
		End If
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FDailyReports.Form_Resize.")
	End Sub
	'
	Private Sub Preview_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Preview.Click
		On Error GoTo EH
		'
		Report.PreviewReport("Daily")
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FDailyReports.Preview_Click.")
	End Sub
	'
	Private Sub PrintButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PrintButton.Click
		On Error GoTo EH
		'
		Report.PrintReport("Daily")
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FDailyReports.PrintButton_Click.")
	End Sub
	'
	Private Sub RefreshData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles RefreshData.Click
		On Error GoTo EH
		'
		Report.Rtype = CReport.ReportType.daily
		RefreshList()
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FDailyReports.RefreshData_Click.")
	End Sub
End Class