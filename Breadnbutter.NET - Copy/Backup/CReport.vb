Option Strict Off
Option Explicit On
Friend Class CReport
	
	Public Enum ReportType
		daily
		Customer
		Eval
		WalkThrough
		Tech
		DaysNotAuth
		NotesSearch
		DaysRemaining
		SimpleContact
		SimpleContact2
		History
		Frontier
		FromList
		NoContactIn
		None
	End Enum
	'
	Public rsReport As New ADODB.Recordset
	'
	Private RRtype As ReportType
	'
	'These should be properties
	'
	Public ActionType As String
	Public SortOrder As String
	Public Company As String
	Public Branch As String
	Public Zip As String
	Public City As String
	Public LastName As String
	Public FirstName As String
	Public Status As String
	Public Notes As String
	Public DaysMax As Short
	Public DaysMin As Short
	Public ID As Short
	Public State As String
	Public Address1 As String
	Public Address2 As String
	Public Phone1 As String
	Public Phone2 As String
	Public Fax As String
	Public Email As String
	Public AuthStatus As String
	Public AuthDate As Date
	Public LastUpdate As Date
	Public ShipStatus As Date
	Public ShipDate As Date
	Public Results As String
	Public Subject As String
	Public User As String
	Public ResultsType As String
	Public Time As Date
	Public ResultsDateMin As Date
	Public ResultsDateMax As Date
	Public RecID As Integer
	Public CustRecID As Integer
	Public ContactType As Short
	Public ProductID As Integer
	Public Source As String
	Public ListQuery As String
	Public SortDirection As Short
	Public SortField As String
	Public RecLimit As Short
	'
	Private RQueryString As String
	'
	'
	Public Property Rtype() As ReportType
		Get
			On Error GoTo EH
			'
			Rtype = RRtype
			'
			Exit Function
EH: 
			MsgBox(Err.Description & " in Get Rtype.")
		End Get
		Set(ByVal Value As ReportType)
			On Error GoTo EH
			'
			Validate()
			RRtype = Value
			'
			Select Case RRtype
				Case ReportType.daily
					rsReport = CreateDaily(-1).rs
				Case ReportType.SimpleContact
					rsReport = CreateSimpleContact
				Case ReportType.History
					rsReport = CreateHistory
				Case ReportType.DaysRemaining
					rsReport = CreateDaysRemaining
				Case ReportType.NotesSearch
					rsReport = CreateNotesSearch
				Case ReportType.DaysNotAuth
					rsReport = CreateDaysNotAuth
				Case ReportType.Frontier
					rsReport = CreateFrontier
				Case ReportType.FromList
					rsReport = CreateFromList
				Case ReportType.SimpleContact2
					rsReport = CreateSimpleContact2
				Case ReportType.NoContactIn
					rsReport = CreateNoContactIn
				Case Else
					MsgBox("Huh?")
			End Select
			'
			Exit Property
EH: 
			MsgBox(Err.Description & " in Let Rtype.")
		End Set
	End Property
	'
	'
	Private Property QueryString() As String
		Get
			On Error GoTo EH
			'
			QueryString = RQueryString
			'
			Exit Property
EH: 
			MsgBox(Err.Description & " in Get Query String.")
		End Get
		Set(ByVal Value As String)
			On Error GoTo EH
			'
			RQueryString = Value
			'
			Exit Property
EH: 
			MsgBox(Err.Description & " in Let QueryString.")
		End Set
	End Property
	'
	Public Sub PreviewReport(ByRef t As String)
		Dim RDaysLeft As Object
		Dim RBranch As Object
		Dim RHistory As Object
		Dim RSimpleContact As Object
		On Error GoTo EH
		'
		'UPGRADE_ISSUE: RDailyBlank object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim rpt As New RDailyBlank
		Select Case t
			Case "Daily"
				rpt = CreateDailyReport
				'UPGRADE_WARNING: Couldn't resolve default property of object rpt.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rpt.Show()
				'UPGRADE_NOTE: Object rpt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rpt = Nothing
			Case "Simple Contact"
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.adc.Connection = cnMain.ConnectionString
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.adc.Source = QueryString
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.Show()
			Case "History"
				'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RHistory.adc.Connection = cnMain.ConnectionString
				'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RHistory.adc.Source = QueryString
				'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RHistory.Show()
			Case "Frontier"
				'UPGRADE_WARNING: Couldn't resolve default property of object RBranch.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RBranch.adc.Recordset = rsReport.Fields.Item.Value
				'UPGRADE_WARNING: Couldn't resolve default property of object RBranch.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RBranch.Show()
			Case "DaysLeft"
				'UPGRADE_WARNING: Couldn't resolve default property of object RDaysLeft.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RDaysLeft.adc.Recordset = rsReport.Fields.Item.Value
				'UPGRADE_WARNING: Couldn't resolve default property of object RDaysLeft.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RDaysLeft.Show()
			Case Else
				MsgBox("No report Type found")
		End Select
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in PreviewReport.")
	End Sub
	Public Sub PrintReport(ByRef t As String)
		Dim RDaysLeft As Object
		Dim RBranch As Object
		Dim RHistory As Object
		Dim RSimpleContact As Object
		On Error GoTo EH
		'
		'UPGRADE_ISSUE: RDailyBlank object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim rpt As New RDailyBlank
		Select Case t
			Case "Daily"
				rpt = CreateDailyReport
				'UPGRADE_WARNING: Couldn't resolve default property of object rpt.PrintReport. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rpt.PrintReport(True)
				'UPGRADE_NOTE: Object rpt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rpt = Nothing
			Case "Simple Contact"
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.adc.Connection = cnMain.ConnectionString
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.adc.Source = QueryString
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.PrintReport. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.PrintReport(True)
				'UPGRADE_ISSUE: Unload RSimpleContact was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
				Unload(RSimpleContact)
			Case "History"
				'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RHistory.adc.Connection = cnMain.ConnectionString
				'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RHistory.adc.Source = QueryString
				'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.PrintReport. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RHistory.PrintReport(True)
				'UPGRADE_ISSUE: Unload RHistory was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
				Unload(RHistory)
			Case "Frontier"
				'UPGRADE_WARNING: Couldn't resolve default property of object RBranch.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RBranch.adc.Recordset = rsReport.Fields.Item.Value
				'UPGRADE_WARNING: Couldn't resolve default property of object RBranch.PrintReport. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RBranch.PrintReport(True)
				'UPGRADE_ISSUE: Unload RBranch was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
				Unload(RBranch)
			Case "DaysLeft"
				'UPGRADE_WARNING: Couldn't resolve default property of object RDaysLeft.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RDaysLeft.adc.Recordset = rsReport.Fields.Item.Value
				'UPGRADE_WARNING: Couldn't resolve default property of object RDaysLeft.PrintReport. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RDaysLeft.PrintReport(True)
				'UPGRADE_ISSUE: Unload RDaysLeft was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
				Unload(RDaysLeft)
			Case "FromList"
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.adc.Connection = cnMain.ConnectionString
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.adc.Source = QueryString
				'UPGRADE_WARNING: Couldn't resolve default property of object RSimpleContact.PrintReport. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RSimpleContact.PrintReport(True)
				'UPGRADE_ISSUE: Unload RSimpleContact was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
				Unload(RSimpleContact)
			Case Else
				MsgBox("No report Type found")
		End Select
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in PrintReport.")
	End Sub
	'
	Private Function Validate() As Boolean
		'
		On Error GoTo EH
		'
		Validate = True
		'Results = "%" + Results + "%"
		'
		If Notes <> "" Then
			Notes = "%" & Notes & "%')"
		Else
			' Notes = "UNLIKELY@#$%^&*"
		End If
		'
		Company = "%" & Company & "%"
		Branch = "%" & Branch & "%')"
		'
		If DaysMin > 10000 Or DaysMin < -10000 Then DaysMin = -10000
		If DaysMax > 10000 Or DaysMax < -10000 Then DaysMax = -10000
		If FirstName = "" Then FirstName = "%"
		If LastName = "" Then LastName = "%"
		If State = "" Then State = "%"
		If State = "All" Then State = "%"
		If City = "" Then City = "%"
		If Zip = "" Then Zip = "%"
		If Company = "" Then Company = "%"
		If Branch = "" Then Branch = "%' OR TBranch.Name IS NULL)"
		' If Notes = "" Then Notes = "QQQQQQQ"
		If Status = "All Users" Then Status = "%"
		If Status = "Everyone" Then Status = "%"
		'
		If State = "" Then State = "%"
		'If Results = "" Then Results = "%"
		If User = "All Users" Then User = "%"
		If User = "" Then User = "%"
		If Notes = "" Then Notes = "%"
		If Source = "" Then Source = "%"
		If ResultsType = "All Categories" Then ResultsType = "%"
		If ResultsType = "" Then ResultsType = "%"
		'
		'If Results = "%%" Then Results = "%"
		
		If Notes = "%%" Then Notes = "%"
		If Company = "%%" Then Company = "%"
		If Branch = "%%')" Then Branch = "%' OR TBranch.Name IS NULL)"
		'
		If Status = "Everyone" Then Status = "%"
		'If ContactType = "0" Then ContactType = "%"
		If ContactType = CDbl("1") Then ContactType = CShort("%")
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in Validate.")
	End Function
	'
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'
		RRtype = ReportType.None
		'
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	'
	Public Function CreateDailyReport() As RDailyBlank
		'
		On Error GoTo EH
		'
		'UPGRADE_ISSUE: RDailyBlank object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim rpt As New RDailyBlank
		Dim ctl As Object
		Dim ctl2 As Object
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.Height. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.Height = 6375
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.PrintWidth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.PrintWidth = 10170
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.PageLeftMargin. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.PageLeftMargin = 500
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.PageRightMargin. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.PageRightMargin = 500
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.PageBottomMargin. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.PageBottomMargin = 700
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.PageTopMargin. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.PageTopMargin = 700
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.PageHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.PageHeader.Height = 1125
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.PageFooter. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.PageFooter.Height = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.Width. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.Width = 11520
		'
		Dim rs As ADODB.Recordset
		Dim iTop As Short
		Dim i As Short
		Dim DailySubItem As CDailySubItem
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.DateLabel. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rpt.DateLabel = Today
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object rpt.Sections. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With rpt.Sections("Detail").Controls
			'
			iTop = 0
			'
			For i = 1 To 10
				'
				DailySubItem = CreateDaily(i)
				'
				If DailySubItem.Name <> "" Then
					'
					'UPGRADE_WARNING: Couldn't resolve default property of object rpt.Sections. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl = .Add("DDActiveReports.Label")
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Caption. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Caption = DailySubItem.Name
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Width. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Width = 9900
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Height. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Height = 360
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Left. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Left = 90
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Top. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Top = iTop
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Font. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Font.Size = 14
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Height. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iTop = iTop + ctl.Height + 90
					'
					'UPGRADE_WARNING: Couldn't resolve default property of object rpt.Sections. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl = .Add("DDActiveReports.SubReport")
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Left. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Left = 270
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Height. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Height = 1080
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Left. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Left = 270
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Top. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Top = iTop
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Width. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.Width = 9900
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.object. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.object = New RDailyUnit
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.object. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ctl.object.adc.Recordset = DailySubItem.rs.Fields.Item.Value
					'UPGRADE_WARNING: Couldn't resolve default property of object ctl.Height. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iTop = iTop + ctl.Height + 100
					'
				End If
				'
			Next i
			'
		End With
		'
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
		'UPGRADE_NOTE: Object ctl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ctl = Nothing
		'UPGRADE_NOTE: Object DailySubItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		DailySubItem = Nothing
		CreateDailyReport = rpt
		'UPGRADE_NOTE: Object rpt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rpt = Nothing
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in CReport.CreateDailyReport.")
	End Function
	'
	Public Function CreateDaily(ByRef ItemNum As Short) As CDailySubItem 'As adodb.Recordset
		On Error GoTo EH
		'
		Dim iCount As Short
		'UPGRADE_WARNING: Arrays can't be declared with New. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC9D3AE5-6B95-4B43-91C7-28276302A5E8"'
		Dim DailyItemsSettings(10) As New CDailyItemSettings
		Dim tempDailySubItem As CDailySubItem
		Dim FinalDailySubItem As New CDailySubItem
		Dim DailySubItem As New CDailyItemSettings
		Dim rsAll As New ADODB.Recordset
		Dim rs As New ADODB.Recordset
		Dim rsOne As New ADODB.Recordset
		Dim CN As New ADODB.Connection
		Dim i As Short
		'
		rsAll.Fields.Append("FirstName", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		rsAll.Fields.Append("LastName", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		rsAll.Fields.Append("Company", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		rsAll.Fields.Append("Phone1", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		rsAll.Fields.Append("Status", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		rsAll.Fields.Append("Type", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		rsAll.Fields.Append("DaysSinceEvent", ADODB.DataTypeEnum.adInteger, 20, ADODB.FieldAttributeEnum.adFldIsNullable)
		rsAll.Open()
		rs.Open("SELECT * FROM TDailyReportsSettings", "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & My.Application.Info.DirectoryPath & "\" & "settings.mdb")
		'
		While (rs.eof = False)
			'
			iCount = iCount + 1
			rs.MoveNext()
			'
		End While
		'
		rs.MoveFirst()
		'
		If ItemNum = -1 Then
			'
			i = 1
			While (rs.eof = False) And i <= 10
				'
				DailyItemsSettings(i).StatusType = rs.Fields("StatusType").Value & vbNullString
				DailyItemsSettings(i).SupportType = rs.Fields("SupportType").Value & vbNullString
				DailyItemsSettings(i).InitialDays = CShort(rs.Fields("InitialDays").Value & vbNullString)
				DailyItemsSettings(i).RegularDays = CShort(rs.Fields("RegularDays").Value & vbNullString)
				DailyItemsSettings(i).DaysDivider = CShort(rs.Fields("DaysDivider").Value & vbNullString)
				DailyItemsSettings(i).DaysOperator = rs.Fields("DaysOperator").Value & vbNullString
				DailyItemsSettings(i).Name = rs.Fields("Name").Value & vbNullString
				DailyItemsSettings(i).DaysToSearch = CShort(rs.Fields("DaysToSearch").Value & vbNullString)
				'
				rsOne = CreateDailyItem(DailyItemsSettings(i), Today).rs
				CombineRS(rsAll, rsOne)
				i = i + 1
				rs.MoveNext()
				'
			End While
			FinalDailySubItem.Name = "Complete Report"
			FinalDailySubItem.rs = rsAll
		Else
			'
			If ItemNum <= iCount Then
				'
				rs.MoveFirst()
				rs.Move((ItemNum - 1))
				DailyItemsSettings(ItemNum).StatusType = rs.Fields("StatusType").Value & vbNullString
				DailyItemsSettings(ItemNum).SupportType = rs.Fields("SupportType").Value & vbNullString
				DailyItemsSettings(ItemNum).InitialDays = CShort(rs.Fields("InitialDays").Value & vbNullString)
				DailyItemsSettings(ItemNum).RegularDays = CShort(rs.Fields("RegularDays").Value & vbNullString)
				DailyItemsSettings(ItemNum).DaysDivider = CShort(rs.Fields("DaysDivider").Value & vbNullString)
				DailyItemsSettings(ItemNum).DaysOperator = rs.Fields("DaysOperator").Value & vbNullString
				DailyItemsSettings(ItemNum).Name = rs.Fields("Name").Value & vbNullString
				DailyItemsSettings(ItemNum).DaysToSearch = CShort(rs.Fields("DaysToSearch").Value & vbNullString)
				tempDailySubItem = CreateDailyItem(DailyItemsSettings(ItemNum), Today)
				rsOne = tempDailySubItem.rs
				FinalDailySubItem.Name = tempDailySubItem.Name
				FinalDailySubItem.rs = rsOne
				'
			End If
			'
		End If
		'
		CreateDaily = FinalDailySubItem
		'
		For i = 1 To 10
			'
			'UPGRADE_NOTE: Object DailyItemsSettings() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			DailyItemsSettings(i) = Nothing
			'
		Next i
		'
		'UPGRADE_NOTE: Object FinalDailySubItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FinalDailySubItem = Nothing
		'UPGRADE_NOTE: Object rsAll may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsAll = Nothing
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
		'UPGRADE_NOTE: Object rsOne may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsOne = Nothing
		'UPGRADE_NOTE: Object CN may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		CN = Nothing
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateDaily.")
	End Function
	Public Function CreateDailyItem(ByRef Item As CDailyItemSettings, ByRef ReportDate As Date) As CDailySubItem
		On Error GoTo EH
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		Dim DailySubItem As New CDailySubItem
		Dim QString As String
		'" WHERE ((([AuthDays]-DateDiff(d,[AuthDate]," & ReportDate & "))>0) AND" &
		QString = "SELECT [TCompany].[Name] as Company," & " [TContact].[FirstName]," & " [TContact].[LastName]," & " [TContact].[Phone1]," & " [TContact].[Status]," & " [AuthDays]-DateDiff(d,[AuthDate],GetDate()) AS Days," & " DateDiff(d,TSUPPORTACTS.DATE,GETDATE()) AS DaysSinceEvent," & " [TSupportActs].[Type]," & " [TSupportActs].[Date]," & " [TSupportActs].[Subject]," & " [TContact].[ID]" & " FROM (TCompany RIGHT JOIN TContact ON [TCompany].[ID]=[TContact].[CompanyID])" & " RIGHT JOIN TSupportActs ON [TContact].[ID]=[TSupportActs].[CustRecID]" & " WHERE ((([AuthDays]-DateDiff(d,[AuthDate],getdate()))>0) AND" & " (DateDiff(d,TSUPPORTACTS.DATE,GETDATE())<" & Item.InitialDays + Item.DaysToSearch & ") AND" & " (DateDiff(d,TSUPPORTACTS.DATE,GETDATE())>" & 0 & ") AND" & " (([TContact].[Status]) Like '" & Item.StatusType & "') AND" & " (([TSupportActs].[Type]) Like '" & Item.SupportType & "'))" & " ORDER BY DateDiff(d,TSUPPORTACTS.DATE,GETDATE());"
		'MsgBox (QString)
		Dim bigRS As ADODB.Recordset
		bigRS = New ADODB.Recordset
		bigRS.let_Source(QString)
		bigRS.Open( , cnMain)
		Dim smallerRS As New ADODB.Recordset
		smallerRS.Fields.Append("FirstName", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("LastName", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("Company", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("Phone1", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("Status", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("Type", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("Days", ADODB.DataTypeEnum.adInteger, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("DaysSinceEvent", ADODB.DataTypeEnum.adInteger, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("Subject", ADODB.DataTypeEnum.adInteger, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Fields.Append("Date", ADODB.DataTypeEnum.adDate, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		smallerRS.Open()
		Dim SingleChar As String
		Dim Complete As String
		Dim pos As Short
		Dim NumFlag As Boolean
		Dim ItemFlag As Boolean
		'
		Do Until bigRS.eof
			ItemFlag = True
			'If Not IsNull(bigRS("Subject")) Or (Item.RegularDays = 0) Then 'ItemFlag = False
			If Item.DaysOperator = "<" And Val(bigRS.Fields("Subject").Value & vbNullString) >= Item.DaysDivider Then ItemFlag = False
			If Item.DaysOperator = ">" And Val(bigRS.Fields("Subject").Value & vbNullString) <= Item.DaysDivider Then ItemFlag = False
			'If IsNull(bigRS("Subject")) Then ItemFlag = False
			If ItemFlag Then '1
				
				If (bigRS.Fields("Days").Value + bigRS.Fields("DaysSinceEvent").Value = Val(bigRS.Fields("Subject").Value & vbNullString)) Or (Item.RegularDays = 0) Then '3
					If CallOrNot(bigRS.Fields("Date").Value, (Item.InitialDays)) Then '4
						smallerRS.AddNew()
						smallerRS.Fields("FirstName").Value = bigRS.Fields("FirstName").Value & vbNullString
						smallerRS.Fields("LastName").Value = bigRS.Fields("LastName").Value & vbNullString
						smallerRS.Fields("Company").Value = bigRS.Fields("Company").Value & vbNullString
						smallerRS.Fields("Phone1").Value = bigRS.Fields("Phone1").Value & vbNullString
						smallerRS.Fields("Status").Value = bigRS.Fields("Status").Value & vbNullString
						smallerRS.Fields("Days").Value = bigRS.Fields("Days").Value & vbNullString
						smallerRS.Fields("DaysSinceEvent").Value = bigRS.Fields("DaysSinceEvent").Value & vbNullString
						smallerRS.Fields("Type").Value = bigRS.Fields("Type").Value & vbNullString
						smallerRS.Fields("Date").Value = bigRS.Fields("Date").Value & vbNullString
						smallerRS.Fields("Subject").Value = Val(bigRS.Fields("Subject").Value & vbNullString)
					Else '4
						If (Item.RegularDays <> 0) Then
							'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							If ((bigRS.Fields("DaysSinceEvent").Value - Item.InitialDays) Mod Item.RegularDays) = 0 Then '5
								If CallOrNot(bigRS.Fields("Date").Value, bigRS.Fields("DaysSinceEvent").Value) Then '6
									smallerRS.AddNew()
									smallerRS.Fields("FirstName").Value = bigRS.Fields("FirstName").Value & vbNullString
									smallerRS.Fields("LastName").Value = bigRS.Fields("LastName").Value & vbNullString
									smallerRS.Fields("Company").Value = bigRS.Fields("Company").Value & vbNullString
									smallerRS.Fields("Phone1").Value = bigRS.Fields("Phone1").Value & vbNullString
									smallerRS.Fields("Status").Value = bigRS.Fields("Status").Value & vbNullString
									smallerRS.Fields("Days").Value = bigRS.Fields("Days").Value & vbNullString
									smallerRS.Fields("DaysSinceEvent").Value = bigRS.Fields("DaysSinceEvent").Value & vbNullString
									smallerRS.Fields("Type").Value = bigRS.Fields("Type").Value & vbNullString
									smallerRS.Fields("Date").Value = bigRS.Fields("Date").Value & vbNullString
									smallerRS.Fields("Subject").Value = Val(bigRS.Fields("Subject").Value)
								End If '6
							End If '5
						End If
					End If '4
				End If '3
			End If '1
			'End If
			bigRS.MoveNext()
			
		Loop 
		If smallerRS.RecordCount > 0 Then
			smallerRS.MoveFirst()
		End If
		DailySubItem.Name = Item.Name
		DailySubItem.rs = smallerRS
		CreateDailyItem = DailySubItem
		'UPGRADE_NOTE: Object DailySubItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		DailySubItem = Nothing
		'UPGRADE_NOTE: Object bigRS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		bigRS = Nothing
		'UPGRADE_NOTE: Object smallerRS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		smallerRS = Nothing
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateDailyItem.")
	End Function
	'
	Public Function CreateNoContactIn() As ADODB.Recordset
		On Error GoTo EH
		Dim rsTemp As New ADODB.Recordset
		QueryString = "SELECT TCompany.ID AS CompanyID, TContact.ID AS ContactID" & ",TCompany.Name AS Company" & ", TContact.FirstName" & ", TContact.LastName" & ", MAX(TSupportActs.[Date]) AS [last called]" & ", TContact.Phone1" & ", TContact.Status" & ", TContact.AuthDays" & ", TSupportActs.Type" & " FROM TSupportActs LEFT OUTER JOIN" & " TContact ON TSupportActs.CustRecID = TContact.ID LEFT OUTER JOIN" & " TCompany ON TContact.CompanyID = TCompany.ID" & " GROUP BY TContact.AuthDays, TSupportActs.Type, TContact.ID" & ", TContact.Status, TContact.Phone1, TContact.LastName" & ", TCompany.Name, TContact.FirstName, TCompany.ID" & " HAVING(TCompany.Name LIKE N'" & Company & "') AND" & " (TContact.Status LIKE N'" & Status & "') AND" & " (MAX(TSupportActs.[Date]) <= GETDATE() - " & DaysMin & ") AND" & " (TSupportActs.Type LIKE N'" & ActionType & "')" & " ORDER BY MAX(TSupportActs.[Date]) DESC"
		
		'If SortField <> vbNullString Then
		'  QueryString = QueryString & " ORDER BY TContact." & SortField
		'  If SortDirection = 1 Then QueryString = QueryString & " DESC"
		'Else
		'  QueryString = QueryString & " ORDER BY TContact.LastName"
		'End If
		'
		rsTemp.let_Source(QueryString)
		'MsgBox (QueryString)
		rsTemp.Open( , cnMain)
		If rsTemp.RecordCount > 0 Then
			rsTemp.MoveFirst()
		End If
		CreateNoContactIn = rsTemp
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateNoContactIn.")
	End Function
	'
	Public Function CreateSimpleContact() As ADODB.Recordset
		On Error GoTo EH
		Dim rsTemp As New ADODB.Recordset
		QueryString = "SELECT TCompany.ID AS CompanyID, TContact.ID AS ContactID" & ",TCompany.Name AS Company" & ",TContact.FirstName" & ",TContact.LastName" & ",TContact.Phone1" & ",TContact.Status" & ",TContact.Source" & " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & " AND (TContact.LastName LIKE '" & LastName & "' )" & " AND (TCompany.Name LIKE '" & Company & "' )" & " AND (TContact.Status LIKE '" & Status & "' )" & " AND (TContact.City LIKE '" & City & "' )" & " AND (TContact.Zip LIKE '" & Zip & "' )" & " AND (TContact.State LIKE '" & State & "' )" & " AND (TContact.Source LIKE '" & Source & "' )" '& |                     '" AND (TContact.Notes LIKE '" & Notes & "')" '& |                     '" ORDER BY TContact.LastName"
		'
		If SortField <> vbNullString Then
			QueryString = QueryString & " ORDER BY TContact." & SortField
			If SortDirection = 1 Then QueryString = QueryString & " DESC"
		Else
			QueryString = QueryString & " ORDER BY TContact.LastName"
		End If
		'
		rsTemp.let_Source(QueryString)
		'MsgBox (QueryString)
		rsTemp.Open( , cnMain)
		If rsTemp.RecordCount > 0 Then
			rsTemp.MoveFirst()
		End If
		CreateSimpleContact = rsTemp
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateSimpleContact.")
	End Function
	'
	Public Function CreateSimpleContact2() As ADODB.Recordset
		'For main contact screen
		On Error GoTo EH
		Dim rsTemp As New ADODB.Recordset
		'
		QueryString = "SELECT TContact.ID, TContact.ContactType, TContact.Status" & ",TContact.FirstName" & ",TContact.LastName" & ",TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & ",TCompany.Name AS Company" & ",TContact.Phone1" & ",TContact.Source" & " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & " AND (TContact.LastName LIKE '" & LastName & "' )" & " AND (TCompany.Name LIKE '" & Company & "' )" & " AND (TContact.Status LIKE '" & Status & "' )" & " AND (TContact.City LIKE '" & City & "' )" & " AND (TContact.Zip LIKE '" & Zip & "' )" & " AND (TContact.State LIKE '" & State & "' )" & " AND (TContact.Source LIKE '" & Source & "' )" & " AND (TContact.Notes LIKE '" & Notes & "')" '& |                     '" ORDER BY TContact.LastName"
		'
		'",TContact.Notes" &
		If SortField <> vbNullString Then
			QueryString = QueryString & " ORDER BY TContact." & SortField
			If SortDirection = 1 Then QueryString = QueryString & " DESC"
		Else
			QueryString = QueryString & " ORDER BY TContact.LastName"
		End If
		'
		rsTemp.let_Source(QueryString)
		'MsgBox (QueryString)
		rsTemp.Open( , cnMain)
		If rsTemp.RecordCount > 0 Then
			rsTemp.MoveFirst()
		End If
		CreateSimpleContact2 = rsTemp
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateSimpleContact2.")
	End Function
	'
	Public Function CreateHistory() As ADODB.Recordset
		On Error GoTo EH
		Dim rsTemp As New ADODB.Recordset
		Dim sResultsQuery As String
		Dim sResultsTemp As String
		Dim sResultsChar As String
		Dim sProductID As String
		Dim sSortOrder As String
		Dim i As Short
		Select Case SortOrder
			Case "Company/Branch"
				sSortOrder = " ORDER by TCompany.Name DESC, TBranch.Name DESC, Date DESC, Time DESC"
			Case "User"
				sSortOrder = " ORDER by TSupportActs.[User] DESC, Date DESC, Time DESC"
			Case "Category"
				sSortOrder = " ORDER by TSupportActs.Type DESC, Date DESC, Time DESC"
			Case Else
				sSortOrder = " ORDER by Date DESC, Time DESC"
		End Select
		'
		If Results <> "" Then
			sResultsChar = Mid(Results, 1, 1)
			If sResultsChar = Chr(34) Then
				sResultsQuery = " (TSupportActs.Results LIKE '%" & Mid(Results, 2, Len(Results) - 2) & "%') AND"
			Else
				For i = 1 To Len(Results)
					sResultsChar = Mid(Results, i, 1)
					If sResultsChar <> " " Then
						sResultsTemp = sResultsTemp & sResultsChar
					Else
						If sResultsTemp <> "" Then
							sResultsQuery = sResultsQuery & " (TSupportActs.Results LIKE '%" & sResultsTemp & "%') AND"
							sResultsTemp = ""
						End If
					End If
				Next i
			End If
			'
			If sResultsTemp <> "" Then
				sResultsQuery = sResultsQuery & " (TSupportActs.Results LIKE '%" & sResultsTemp & "%') AND"
			End If
		Else
			sResultsQuery = " (TSupportActs.Results LIKE '%') AND"
		End If
		'
		Dim sRecLimit As String
		If RecLimit > 0 Then
			sRecLimit = " TOP " & CStr(RecLimit)
		Else
			sRecLimit = ""
		End If
		'
		Dim sContactType As String
		If ContactType <> 0 Then sContactType = "(TContact.ContactType = '" & ContactType & "') AND"
		'
		If ProductID = 1 Then sProductID = " (TSupportActs.ProductID = 1 or TSupportActs.ProductID is null) AND"
		If ProductID > 1 Then sProductID = " (TSupportActs.ProductID = " & ProductID & " ) AND"
		'
		QueryString = "SELECT" & sRecLimit & " TCompany.Name AS Company, TSupportActs.Type," & " TSupportActs.Date, TContact.FirstName," & " TContact.LastName, TSupportActs.Results," & " TSupportActs.[User], TContact.Status," & " TContact.State, TSupportActs.RecID," & " TSupportActs.CustRecID, TSupportActs.Subject," & " TSupportActs.Results, TSupportActs.Time," & " TContact.Phone1, TBranch.Name AS Branch" & " FROM TBranch RIGHT OUTER JOIN" & " TContact ON TBranch.BranchID = TContact.BranchID RIGHT OUTER JOIN" & " TSupportActs ON TContact.ID = TSupportActs.CustRecID LEFT OUTER JOIN" & " TCompany ON TContact.CompanyID = TCompany.ID" & " WHERE (TSupportActs.Type LIKE '" & ResultsType & "') AND" & " (TSupportActs.Date >= '" & ResultsDateMin & "') AND" & " (TSupportActs.Date <= '" & ResultsDateMax & "') AND" & " (TContact.Status LIKE '" & Status & "') AND" & " (TSupportActs.[User] LIKE '" & User & "') AND" & " (TContact.FirstName LIKE '" & FirstName & "') AND" & " (TContact.LastName LIKE '" & LastName & "') AND" & " (TCompany.Name LIKE '" & Company & "') AND" & " (TBranch.Name LIKE '" & Branch & " AND" & sContactType & sResultsQuery & sProductID & " (TContact.State LIKE '" & State & "')" & sSortOrder
		'" ORDER by Date DESC, Time DESC"
		'MsgBox (QueryString)
		'" (TSupportActs.Results LIKE '" & Results & "') AND" & _
		''If rsTemp.State <> 0 Then rsTemp.Close
		rsTemp.let_Source(QueryString)
		'MsgBox (QueryString)
		rsTemp.Open( , cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		If rsTemp.RecordCount > 0 Then
			rsTemp.MoveFirst()
		End If
		CreateHistory = rsTemp
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateHistory.")
	End Function
	'
	Public Function CreateDaysNotAuth() As ADODB.Recordset
		On Error GoTo EH
		Dim rsTemp As New ADODB.Recordset
		Dim sShipDateFieldName As String
		'
		If ProductID = 2 Then
			sShipDateFieldName = "TContact.PVShipDate"
		Else
			sShipDateFieldName = "TContact.ShipDate"
		End If
		
		QueryString = "SELECT TCompany.ID as CompanyID, TContact.ID as ContactID" & ",TCompany.Name AS Company" & ",TContact.FirstName" & ",TContact.LastName" & ",TContact.Phone1" & ",TContact.Status" & ",TContact.Source" & ",DATEDIFF(d, " & sShipDateFieldName & ", GetDate()) AS Days" & " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & " AND (TContact.LastName LIKE '" & LastName & "' )" & " AND (TCompany.Name LIKE '" & Company & "' )" & " AND (TContact.Status LIKE '" & Status & "' )" & " AND (TContact.City LIKE '" & City & "' )" & " AND (TContact.Zip LIKE '" & Zip & "' )" & " AND (TContact.State LIKE '" & State & "' )" & " AND (TContact.Source LIKE '" & Source & "' )" & " AND(DATEDIFF(d, " & sShipDateFieldName & ", GetDate()) <= '" & DaysMax & "')" & " AND(DATEDIFF(d, " & sShipDateFieldName & ", GetDate()) >= '" & DaysMin & "')" & " AND(TContact.AuthStatus LIKE 'Not Authorized')" '& |                     '" ORDER BY (DATEDIFF(d, dbo.TContact.ShipDate, GetDate()))"
		If SortField <> vbNullString Then
			QueryString = QueryString & " ORDER BY TContact." & SortField
			If SortDirection = 1 Then QueryString = QueryString & " DESC"
		Else
			QueryString = QueryString & " ORDER BY (DATEDIFF(d, dbo.TContact.ShipDate, GetDate()))"
		End If
		'
		rsTemp.let_Source(QueryString)
		'MsgBox (QueryString)
		rsTemp.Open( , cnMain)
		If rsTemp.RecordCount > 0 Then
			rsTemp.MoveFirst()
		End If
		CreateDaysNotAuth = rsTemp
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateDaysNotAuth.")
	End Function
	'
	Public Function CreateDaysRemaining() As ADODB.Recordset
		On Error GoTo EH
		Dim rsTemp As New ADODB.Recordset
		Dim sAuthDaysFieldName As String
		Dim sAuthDateFieldName As String
		Dim sStatus As String
		'
		If ProductID = 2 Then
			sAuthDaysFieldName = "TContact.PVAuthDays"
			sAuthDateFieldName = "TContact.PVAuthDate"
			'sStatus = "TContact.PVStatus"
		Else
			sAuthDaysFieldName = "TContact.AuthDays"
			sAuthDateFieldName = "TContact.AuthDate"
			'sStatus = "TContact.PVStatus"
		End If
		'
		QueryString = "SELECT TCompany.ID as CompanyID, TContact.ID as ContactID" & ",TCompany.Name AS Company" & ",TContact.FirstName" & ",TContact.LastName" & ",TContact.Phone1" & ",TContact.Status" & ",TContact.Source" & "," & sAuthDateFieldName & " + TContact.AuthDays AS [ExpirationDate]" & "," & sAuthDaysFieldName & "-DateDiff(d," & sAuthDateFieldName & ", GETDATE()) as Days" & " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & " AND (TContact.LastName LIKE '" & LastName & "' )" & " AND (TCompany.Name LIKE '" & Company & "' )" & " AND (TContact.Status LIKE '" & Status & "' )" & " AND (TContact.City LIKE '" & City & "' )" & " AND (TContact.Zip LIKE '" & Zip & "' )" & " AND (TContact.State LIKE '" & State & "' )" & " AND (TContact.Source LIKE '" & Source & "' )" & " AND (" & sAuthDaysFieldName & "- DateDiff(d, " & sAuthDateFieldName & ", GetDate())>='" & DaysMin & "' )" & " AND (" & sAuthDaysFieldName & "- DateDiff(d, " & sAuthDateFieldName & ", GetDate())<='" & DaysMax & " ' )" ' & |                   '" ORDER BY (" & sAuthDaysFieldName & "- DateDiff(d, " & sAuthDateFieldName & ", GetDate()))"
		'
		If SortField <> vbNullString Then
			QueryString = QueryString & " ORDER BY TContact." & SortField
			If SortDirection = 1 Then QueryString = QueryString & " DESC"
		Else
			QueryString = QueryString & " ORDER BY (" & sAuthDaysFieldName & "- DateDiff(d, " & sAuthDateFieldName & ", GetDate()))"
		End If
		'
		rsTemp.let_Source(QueryString)
		'MsgBox (QueryString)
		rsTemp.Open( , cnMain)
		If rsTemp.RecordCount > 0 Then
			rsTemp.MoveFirst()
		End If
		CreateDaysRemaining = rsTemp
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateDaysRemaining.")
	End Function
	'
	Public Function CreateNotesSearch() As ADODB.Recordset
		On Error GoTo EH
		Dim rsTemp As New ADODB.Recordset
		QueryString = "SELECT TCompany.ID as CompanyID, TContact.ID as ContactID" & ",TCompany.Name AS Company" & ",TContact.FirstName" & ",TContact.LastName" & ",TContact.Phone1" & ",TContact.Status" & ",TContact.Source" & ",TContact.Notes" & " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & " AND (TContact.LastName LIKE '" & LastName & "' )" & " AND (TCompany.Name LIKE '" & Company & "' )" & " AND (TContact.Status LIKE '" & Status & "' )" & " AND (TContact.City LIKE '" & City & "' )" & " AND (TContact.Zip LIKE '" & Zip & "' )" & " AND (TContact.State LIKE '" & State & "' )" & " AND (TContact.Notes LIKE '" & Notes & "')" & " AND (TContact.Source LIKE '" & Source & "' )" ' & |                   '" ORDER BY TContact.LastName"
		'
		If SortField <> vbNullString Then
			QueryString = QueryString & " ORDER BY TContact." & SortField
			If SortDirection = 1 Then QueryString = QueryString & " DESC"
		Else
			QueryString = QueryString & " ORDER BY TContact.LastName"
		End If
		'
		rsTemp.let_Source(QueryString)
		'MsgBox (QueryString)
		'If Notes <> "%" Then
		rsTemp.Open( , cnMain)
		'End If
		If rsTemp.RecordCount > 0 Then
			rsTemp.MoveFirst()
		End If
		CreateNotesSearch = rsTemp
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateNotesSearch.")
	End Function
	'
	Public Function CreateFromList() As ADODB.Recordset
		On Error GoTo EH
		'
		Dim rsTemp As New ADODB.Recordset
		'
		QueryString = ListQuery
		'
		If SortField <> vbNullString Then
			QueryString = QueryString & " ORDER BY TContact." & SortField
			If SortDirection = 1 Then QueryString = QueryString & " DESC"
		Else
			QueryString = QueryString & " ORDER BY TContact.State"
		End If
		rsTemp.let_Source(QueryString)
		'
		rsTemp.Open( , cnMain)
		'
		If rsTemp.RecordCount > 0 Then
			rsTemp.MoveFirst()
		End If
		'
		CreateFromList = rsTemp
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateFromlist.")
	End Function
	'
	Public Function CreateFrontier() As ADODB.Recordset
		On Error GoTo EH
		Dim rsTemp As New ADODB.Recordset
		Notes = "B#"
		ResultsType = "Sale"
		Validate()
		QueryString = "SELECT TCompany.ID as CompanyID" & ", TContact.ID as ContactID,TCompany.Name AS Company" & ",TContact.FirstName,TContact.LastName" & ",TContact.Phone1,TContact.Status" & ",TContact.Notes,TSupportActs.Date" & " FROM TSupportActs LEFT OUTER JOIN" & " TContact ON" & " TSupportActs.CustRecID = dbo.TContact.ID LEFT OUTER JOIN" & " TCompany ON" & " TContact.CompanyID = TCompany.ID" & " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & " AND (TContact.LastName LIKE '" & LastName & "' )" & " AND (TCompany.Name LIKE '" & Company & "' )" & " AND (TContact.Status LIKE '" & Status & "' )" & " AND (TContact.City LIKE '" & City & "' )" & " AND (TContact.Zip LIKE '" & Zip & "' )" & " AND (TContact.State LIKE '" & State & "' )" & " AND (TContact.Notes LIKE '" & Notes & "')" & " AND (TContact.Source LIKE '" & Source & "' )" & " AND (TSupportActs.Date >= '" & ResultsDateMin & "') " & " AND (TSupportActs.Date <= '" & ResultsDateMax & "') " & " AND (TSupportActs.Type LIKE '" & ResultsType & "') " '& |                   '" ORDER BY TContact.LastName"
		'
		If SortField <> vbNullString Then
			QueryString = QueryString & " ORDER BY TContact." & SortField
			If SortDirection = 1 Then QueryString = QueryString & " DESC"
		Else
			QueryString = QueryString & " ORDER BY TContact.LastName"
		End If
		rsTemp.let_Source(QueryString)
		'
		rsTemp.let_Source(QueryString)
		'MsgBox (QueryString)
		rsTemp.Open( , cnMain)
		Dim trs As New ADODB.Recordset
		trs.Fields.Append("CompanyID", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		trs.Fields.Append("ContactID", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		trs.Fields.Append("FirstName", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		trs.Fields.Append("LastName", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		trs.Fields.Append("Branch", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		trs.Fields.Append("Date", ADODB.DataTypeEnum.adDate, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		trs.Fields.Append("Status", ADODB.DataTypeEnum.adChar, 300, ADODB.FieldAttributeEnum.adFldIsNullable)
		trs.Open()
		Dim SingleChar As String
		Dim Complete As String
		Dim pos As Short
		Dim NumFlag As Boolean
		'
		Do Until rsTemp.eof
			trs.AddNew()
			trs.Fields("CompanyID").Value = rsTemp.Fields("CompanyID").Value
			trs.Fields("ContactID").Value = rsTemp.Fields("ContactID").Value
			trs.Fields("FirstName").Value = rsTemp.Fields("FirstName").Value
			trs.Fields("LastName").Value = rsTemp.Fields("LastName").Value
			trs.Fields("Status").Value = rsTemp.Fields("Status").Value
			SingleChar = Mid(rsTemp.Fields("Notes").Value, 3, 1)
			NumFlag = isNumber(SingleChar)
			Complete = ""
			pos = 1
			Do Until NumFlag = False
				Complete = Complete & SingleChar
				SingleChar = Mid(rsTemp.Fields("Notes").Value, 3 + pos, 1)
				NumFlag = isNumber(SingleChar)
				pos = pos + 1
			Loop 
			trs.Fields("Branch").Value = Complete
			trs.Fields("Date").Value = rsTemp.Fields("Date").Value
			rsTemp.MoveNext()
		Loop 
		If trs.RecordCount > 0 Then
			trs.MoveFirst()
		End If
		CreateFrontier = trs
		Exit Function
EH: 
		MsgBox(Err.Description & " in CreateFrontier.")
	End Function
	'
	Private Function CallOrNot(ByRef AccusedDate As Date, ByRef DaysUntilCall As Short) As Boolean
		On Error GoTo EH
		'
		Dim Difference As Short
		'UPGRADE_NOTE: Today was upgraded to Today_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Today_Renamed As Date
		'
		Today_Renamed = Today
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		Difference = DateDiff(Microsoft.VisualBasic.DateInterval.Day, AccusedDate, Today_Renamed)
		CallOrNot = False
		'
		If Difference = DaysUntilCall Then CallOrNot = True
		'
		If Difference = DaysUntilCall + 2 And WeekDay(Today_Renamed) = FirstDayOfWeek.Monday Then
			'
			CallOrNot = True
			'
		End If
		'
		If Difference = DaysUntilCall + 1 And WeekDay(Today_Renamed) = FirstDayOfWeek.Monday Then
			'
			CallOrNot = True
			'
		End If
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in CallOrNot.")
	End Function
	'
	Public Sub FillList(ByRef rs As ADODB.Recordset, ByRef list As System.Windows.Forms.ListView)
		'On Error GoTo EH
		'
		Dim iSortOrder As Short
		Dim lSortkey As Integer
		'
		Dim LineCount As Integer
		Dim FieldPos As Integer
		Dim TotalCharacters As Integer
		Dim sKey As String
		Dim iRecCount As Short
		'
		TotalCharacters = 0
		LineCount = 0
		FieldPos = 2
		'
		iSortOrder = list.Sorting
		'UPGRADE_ISSUE: MSComctlLib.ListView property list.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		lSortkey = list.SortKey
		'UPGRADE_ISSUE: MSComctlLib.ListView property list.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		list.SortKey = 0
		'UPGRADE_ISSUE: MSComctlLib.ListView property list.Sorted was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		list.Sorted = False
		With rs
			'
			If .RecordCount > 0 Then
				'
				Do 
					'
					iRecCount = 0
					TotalCharacters = 0
					.MoveFirst()
					Do 
						'
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If Not IsDbNull(.Fields(FieldPos).Value) Then
							'
							'TotalCharacters = Len(CStr(.Fields(FieldPos))) + TotalCharacters
							If Len(CStr(.Fields(FieldPos).Value)) > 0 Then
								TotalCharacters = Len(CStr(.Fields(FieldPos).Value)) + TotalCharacters
								iRecCount = iRecCount + 1
							End If
							'
						End If
						'
						.MoveNext()
						'
					Loop Until .eof
					'
					If TotalCharacters = 0 Then
						list.Columns.Add("w1" & FieldPos, .Fields(FieldPos).Name, CInt(VB6.TwipsToPixelsX(Len(CStr(.Fields(FieldPos).Name)) * 100)))
					Else
						list.Columns.Add("w1" & FieldPos, .Fields(FieldPos).Name, CInt(VB6.TwipsToPixelsX(400 + ((TotalCharacters / iRecCount) * 100))))
					End If
					'list.ColumnHeaders.Add , .Fields, .Fields(FieldPos).Name, 400 + ((TotalCharacters / .RecordCount) * 100)
					FieldPos = FieldPos + 1
					'
				Loop Until FieldPos = .Fields.Count
				'
				.MoveFirst()
				'
				FieldPos = 2
				LineCount = 0
				Do Until .eof
					'
					'list.ListItems.Add , "r1" & LineCount, .Fields(FieldPos)
					sKey = .Fields("CompanyID").Value & "A" & .Fields("ContactID").Value
					list.Items.Add(sKey, .Fields(FieldPos).Value, "")
					FieldPos = FieldPos + 1
					'
					Do 
						'        '
						' If Not IsNull(.Fields(FieldPos)) Then
						'         '
						' ListItems.Item(sKey).SubItems(FieldPos - 2) = .Fields(FieldPos)
						list.Items.Item(sKey).SubItems.Add(.Fields(FieldPos).Value & vbNullString) '.ForeColor = lColor
						'
						'
						'  End If
						'         '
						FieldPos = FieldPos + 1
						'         '
					Loop Until FieldPos = .Fields.Count
					'
					FieldPos = 2
					.MoveNext()
					LineCount = LineCount + 1
					'
				Loop 
				'
			End If
			'
		End With
		'
		list.Sorting = iSortOrder
		'UPGRADE_ISSUE: MSComctlLib.ListView property list.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		list.SortKey = lSortkey
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FillList.")
	End Sub
	
	Private Function isNumber(ByRef Value As String) As Boolean
		On Error GoTo EH
		'
		isNumber = False
		'
		If Value = "0" Then isNumber = True
		If Value = "1" Then isNumber = True
		If Value = "2" Then isNumber = True
		If Value = "3" Then isNumber = True
		If Value = "4" Then isNumber = True
		If Value = "5" Then isNumber = True
		If Value = "6" Then isNumber = True
		If Value = "7" Then isNumber = True
		If Value = "8" Then isNumber = True
		If Value = "9" Then isNumber = True
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in isNumber.")
	End Function
	'
	Private Sub CombineRS(ByRef RSFirst As ADODB.Recordset, ByRef RSSecond As ADODB.Recordset)
		'
		On Error GoTo EH
		If RSSecond.RecordCount > 0 Then
			'   RSSecond.MoveFirst
		End If
		'
		Do Until RSSecond.eof
			'
			RSFirst.AddNew()
			RSFirst.Fields("FirstName").Value = RSSecond.Fields("FirstName").Value
			RSFirst.Fields("LastName").Value = RSSecond.Fields("LastName").Value
			RSFirst.Fields("Company").Value = RSSecond.Fields("Company").Value
			RSFirst.Fields("Phone1").Value = RSSecond.Fields("Phone1").Value
			'   RSFirst!Status = RSSecond!Status
			RSFirst.Fields("Type").Value = RSSecond.Fields("Type").Value
			RSFirst.Fields("DaysSinceEvent").Value = RSSecond.Fields("DaysSinceEvent").Value
			RSFirst.Update()
			RSSecond.MoveNext()
			'
		Loop 
		'
		If RSSecond.RecordCount > 0 Then
			RSSecond.MoveFirst()
		End If
		'
		If RSFirst.RecordCount > 0 Then
			RSFirst.MoveFirst()
		End If
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in CombineRS.")
	End Sub
End Class