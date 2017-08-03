Option Strict Off
Option Explicit On
Friend Class FCallStats
	Inherits System.Windows.Forms.Form
	
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	' global in other prog
	Private strGroupType, strSQL, strDateType, strDirection As String
	Private DTStart, DTEnd As Date
	Private intEXT, intWorkgroup As Short
	'Public chrtArray()
	Private lGreatestValue As Integer
	Private iExt() As Short
	Private sName() As String
	Private iNumofCalls() As Short
	Private iBNBNotes() As Short
	Private lCallTime() As Integer
	Private sCallTime() As String
	Private iFollowups() As Short
	Private iWalkThroughs() As Short
	Private iSales() As Short
	Private lAvgCallTime() As Integer
	Private sAvgCallTime() As String
	Private iNumofUsers As Short
	Private strSQL2 As String
	Private iMins As Short
	
	'UPGRADE_NOTE: Form_Initialize was upgraded to Form_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Form_Initialize_Renamed()
		On Error GoTo ErrCall
		'
		FormControl = New CFormControl
		'
		FormControl.MinHeight = VB6.PixelsToTwipsY(Me.Height)
		FormControl.MinWidth = VB6.PixelsToTwipsX(Me.Width)
		FormControl.DataForm = False
		'
		cboMins.SelectedIndex = 0
		'
		optAll.Checked = True
		DTPicker1.Value = System.Date.FromOADate(Today.ToOADate - 7)
		DTPicker2.Value = Today
		optBNB.Checked = True
		'
		grdCallData.Visible = False
		ListView1.Visible = False
		'
		LoadExtList()
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSelect.Form_Initialize.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	
	Private Sub FCallStats_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		ListView1.View = System.Windows.Forms.View.Details
		ListView1.Sort()
		
		With ListView1.Columns
			.Add("Name")
			.Add("# Calls")
			.Add("# BNB Notes")
			.Add("Calls Vs. Notes")
			.Add("Total Call Time")
			.Add("# Follow-ups")
			.Add("# Walk Throughs")
			.Add("# Sales")
			.Add("Avg Call Time")
			.Add("") '-- Dummy column. No text in case user pulls out to view.
			
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(1).Width = VB6.TwipsToPixelsX(1500.09)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(2).Width = VB6.TwipsToPixelsX(700.15)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(3).Width = VB6.TwipsToPixelsX(1120.25)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(4).Width = VB6.TwipsToPixelsX(1239.87)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(5).Width = VB6.TwipsToPixelsX(1230.23)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(6).Width = VB6.TwipsToPixelsX(1080)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(7).Width = VB6.TwipsToPixelsX(1429.79)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(8).Width = VB6.TwipsToPixelsX(739.84)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(9).Width = VB6.TwipsToPixelsX(1149.73)
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.Item(10).Width = 0 '-- Dummy column.
		End With
		
	End Sub
	
	'UPGRADE_WARNING: Event FCallStats.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FCallStats_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		grdCallData.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 1000)
		ListView1.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 1000)
		If VB6.PixelsToTwipsY(Me.Height) > 0 Then
			grdCallData.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 2900)
			ListView1.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 2900)
		End If
		Label3.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdCallData.Top) + VB6.PixelsToTwipsY(grdCallData.Height) + 300)
		Label4.Top = Label3.Top
		txtTotal.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Label3.Top) + 200)
		txtAvg.Top = txtTotal.Top
	End Sub
	
	
	'UPGRADE_WARNING: Event cboGroup.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboGroup_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboGroup.SelectedIndexChanged
		'
		'Disable Print button
		'
		cmdPrintReport.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event cboCallDir.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCallDir_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCallDir.SelectedIndexChanged
		'
		'Disable Print button
		'
		cmdPrintReport.Enabled = False
	End Sub
	
	Private Sub cmdPrintReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrintReport.Click
		Dim RCallLog As Object
		Dim RCallsAndNotes As Object
		Dim sDate1 As String
		Dim sDate2 As String
		'Dim old_width As Integer
		If optBNB.Checked = True Then
			'
			sDate1 = "( " & DTPicker1._Value & " )"
			sDate2 = "( " & DTPicker2._Value & " )"
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object RCallsAndNotes.GetData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RCallsAndNotes.GetData(sName, iNumofCalls, iBNBNotes, sCallTime, iFollowups, iWalkThroughs, iSales, sAvgCallTime, iNumofUsers, sDate1, sDate2)
			'
		Else
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object RCallLog.GetData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RCallLog.GetData(DTStart, DTEnd, intWorkgroup, strSQL, strDirection)
			'
		End If
		'
		
		'Dim objPrintLV As clsPrintLV
		'    Set objPrintLV = New clsPrintLV
		'    objPrintLV.PrintListView ListView1, 0.1, 8, "Sample ListView Report", landscape, True
		'    Set objPrintLV = Nothing
		
	End Sub
	
	Private Sub DTPicker1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTPicker1.Change
		'
		'Disable Print button
		'
		cmdPrintReport.Enabled = False
	End Sub
	
	Private Sub DTPicker2_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTPicker2.Change
		'
		'Disable Print button
		'
		cmdPrintReport.Enabled = False
	End Sub
	
	Private Sub cmdGenReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGenReport.Click
		If optBNB.Checked = True Then
			BNBReport()
		Else
			DurationReport()
		End If
	End Sub
	
	Private Sub grdCallData_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdCallData.ClickEvent
		
	End Sub
	
	Private Sub ListView1_ColumnClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
		Dim ColumnHeader As System.Windows.Forms.ColumnHeader = ListView1.Columns(eventArgs.Column)
		Call SortListView(ListView1, ColumnHeader)
	End Sub
	
	'UPGRADE_WARNING: Event optBNB.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optBNB_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optBNB.CheckedChanged
		If eventSender.Checked Then
			fraGroup.Visible = False
			fraGroup.Visible = False
			Label3.Visible = False
			Label4.Visible = False
			txtTotal.Visible = False
			txtAvg.Visible = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optDuration.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDuration_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDuration.CheckedChanged
		If eventSender.Checked Then
			fraGroup.Visible = True
			Label3.Visible = True
			Label4.Visible = True
			txtTotal.Visible = True
			txtAvg.Visible = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optExt.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optExt_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optExt.CheckedChanged
		If eventSender.Checked Then
			'
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optAll.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAll.CheckedChanged
		If eventSender.Checked Then
			'
		End If
	End Sub
	
	Private Sub LoadExtList()
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		Dim X As Short
		'
		cmdCategories.CommandTimeout = 300
		'
		cmdCategories.ActiveConnection = cnMain
		cmdCategories.CommandText = "SELECT     EmployeeFirst + N' ' + EmployeeLast AS Name, EmployeeExt From tblEmployees GROUP BY EmployeeFirst + N' ' + EmployeeLast, EmployeeExt HAVING      (NOT (EmployeeFirst + N' ' + EmployeeLast IS NULL)) AND (NOT (EmployeeExt IS NULL)) ORDER BY EmployeeExt"
		rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		iNumofUsers = rstCategories.RecordCount
		'UPGRADE_WARNING: Lower bound of array iExt was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim iExt(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array sName was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim sName(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array iNumofCalls was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim iNumofCalls(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array iBNBNotes was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim iBNBNotes(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array lCallTime was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim lCallTime(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array sCallTime was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim sCallTime(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array iFollowups was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim iFollowups(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array iWalkThroughs was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim iWalkThroughs(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array iSales was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim iSales(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array lAvgCallTime was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim lAvgCallTime(iNumofUsers)
		'UPGRADE_WARNING: Lower bound of array sAvgCallTime was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim sAvgCallTime(iNumofUsers)
		X = 0
		'
		Do While Not rstCategories.eof
			If rstCategories.Fields("EmployeeExt").Value <> "" Then
				X = X + 1
				iExt(X) = rstCategories.Fields("EmployeeExt").Value
				cboGroup.Items.Add(rstCategories.Fields("Name").Value & " (" & rstCategories.Fields("EmployeeExt").Value & ")")
				sName(X) = rstCategories.Fields("Name").Value
			End If
			rstCategories.MoveNext()
		Loop 
		'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdCategories = Nothing
		rstCategories.Close()
		'
		cboGroup.SelectedIndex = 0
	End Sub
	
	Private Sub LoadNumOfCalls()
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		Dim X As Short
		'
		cmdCategories.CommandTimeout = 300
		'
		For X = 1 To iNumofUsers
			cmdCategories.ActiveConnection = cnMain
			'
			cmdCategories.CommandText = "SELECT COUNT(DISTINCT SESSID) AS Calls,  SUM(DISTINCT CALLDUR) AS Duration, AVG(CALLDUR) AS AvgTime From ICC_CDR WHERE (P1NO LIKE N'" & iExt(X) & "') AND (DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1970-01-01 00:00:00', 102)) BETWEEN CONVERT(DATETIME,'" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) AND (NOT (LEFT(TKRMNO, 3) LIKE N'270'))"
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			If Not rstCategories.eof Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rstCategories.Fields("Calls").Value) Then
					iNumofCalls(X) = rstCategories.Fields("Calls").Value
				End If
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rstCategories.Fields("Duration").Value) Then
					sCallTime(X) = ConvertTime(rstCategories.Fields("Duration").Value)
					'lCallTime(x) = rstCategories!Duration
				Else
					sCallTime(X) = "No Data"
				End If
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(rstCategories.Fields("AvgTime").Value) Then
					sAvgCallTime(X) = ConvertTime(rstCategories.Fields("AvgTime").Value)
					'lAvgCallTime(x) = rstCategories!AvgTime
				Else
					sAvgCallTime(X) = "No Data"
				End If
			End If
			'rstCategories.MoveNext
			'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cmdCategories = Nothing
			rstCategories.Close()
		Next 
		'
		
	End Sub
	
	Private Sub LoadBNBData()
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		Dim X As Short
		'
		cmdCategories.CommandTimeout = 300
		'
		For X = 1 To iNumofUsers
			cmdCategories.ActiveConnection = cnMain
			'
			cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS BNBNotes FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) AND (tblEmployees.EmployeeExt = " & iExt(X) & ") GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast HAVING (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL))"
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			If Not rstCategories.eof Then
				iBNBNotes(X) = rstCategories.Fields("BNBNotes").Value
			End If
			'rstCategories.MoveNext
			'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cmdCategories = Nothing
			rstCategories.Close()
		Next 
		'
		For X = 1 To iNumofUsers
			cmdCategories.ActiveConnection = cnMain
			'
			cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS Followups FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) AND (tblEmployees.EmployeeExt = " & iExt(X) & ") GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, TSupportActs.Type HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) AND (TSupportActs.Type LIKE N'Follow-up call')"
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			If Not rstCategories.eof Then
				iFollowups(X) = rstCategories.Fields("Followups").Value
			End If
			'rstCategories.MoveNext
			'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cmdCategories = Nothing
			rstCategories.Close()
		Next 
		'
		For X = 1 To iNumofUsers
			cmdCategories.ActiveConnection = cnMain
			'
			cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS WalkThroughs FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) AND (tblEmployees.EmployeeExt = " & iExt(X) & ") GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, TSupportActs.Type HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) AND (TSupportActs.Type LIKE N'Walk Through')"
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			If Not rstCategories.eof Then
				iWalkThroughs(X) = rstCategories.Fields("WalkThroughs").Value
			End If
			'rstCategories.MoveNext
			'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cmdCategories = Nothing
			rstCategories.Close()
		Next 
		'
		For X = 1 To iNumofUsers
			cmdCategories.ActiveConnection = cnMain
			'
			cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS Sales FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) AND (tblEmployees.EmployeeExt = " & iExt(X) & ") GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, TSupportActs.Type HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) AND (TSupportActs.Type LIKE N'Sale')"
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			If Not rstCategories.eof Then
				iSales(X) = rstCategories.Fields("Sales").Value
			End If
			'rstCategories.MoveNext
			'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cmdCategories = Nothing
			rstCategories.Close()
		Next 
		'
	End Sub
	
	Private Function ConvertTime(ByRef Seconds As Integer) As String
		Dim lHrs As Integer
		Dim lMinutes As Integer
		Dim lSeconds As Integer
		
		lSeconds = Seconds
		'
		If lSeconds >= 3600 Then
			'get hours which is equal to seconds divided by 3600
			lHrs = lSeconds / 3600
			
			'set the seconds to the numbers after the decimal sign
			'thats what mod does
			lSeconds = lSeconds Mod 3600
		Else
			'if not greater than 3600, just set it to 0
			lHrs = 0
		End If
		
		If lSeconds >= 60 Then
			'greater than or equal to 60
			'set the minutes equal to the value of (seconds divided by 60).
			'and get the remaining numbers after the decimal
			'which will be the seconds
			'using the mod sign
			
			lMinutes = lSeconds \ 60
			lSeconds = lSeconds Mod 60
		Else
			'if not set to 0
			lMinutes = 0
		End If
		'
		If lHrs > 0 Then
			ConvertTime = VB6.Format(CStr(lHrs), "#####0") & ":" & VB6.Format(CStr(lMinutes), "00") & "." & VB6.Format(CStr(lSeconds), "00")
		Else
			ConvertTime = VB6.Format(CStr(lMinutes), "#0") & "." & VB6.Format(CStr(lSeconds), "00")
		End If
		'
	End Function
	
	Private Sub LoadBNBData2()
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		Dim X As Short
		'
		cmdCategories.CommandTimeout = 300
		'
		cmdCategories.ActiveConnection = cnMain
		'
		cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS BNBNotes, tblEmployees.EmployeeExt FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, tblEmployees.EmployeeExt HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) ORDER BY tblEmployees.EmployeeExt"
		rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		Do Until rstCategories.eof
			For X = 1 To iNumofUsers
				If iExt(X) = rstCategories.Fields("EmployeeExt").Value Then
					iBNBNotes(X) = rstCategories.Fields("BNBNotes").Value
				End If
			Next 
			rstCategories.MoveNext()
		Loop 
		'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdCategories = Nothing
		rstCategories.Close()
		'
		cmdCategories.ActiveConnection = cnMain
		'
		cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS Followups, tblEmployees.EmployeeExt FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) AND (TSupportActs.Type LIKE N'Follow-up call') GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, tblEmployees.EmployeeExt HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) ORDER BY tblEmployees.EmployeeExt"
		rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		Do Until rstCategories.eof
			For X = 1 To iNumofUsers
				If iExt(X) = rstCategories.Fields("EmployeeExt").Value Then
					iFollowups(X) = rstCategories.Fields("Followups").Value
				End If
			Next 
			rstCategories.MoveNext()
		Loop 
		'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdCategories = Nothing
		rstCategories.Close()
		'
		cmdCategories.ActiveConnection = cnMain
		'
		cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS WalkThroughs, tblEmployees.EmployeeExt FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) AND (TSupportActs.Type LIKE N'Walk Through') GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, tblEmployees.EmployeeExt HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) ORDER BY tblEmployees.EmployeeExt"
		rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		Do Until rstCategories.eof
			For X = 1 To iNumofUsers
				If iExt(X) = rstCategories.Fields("EmployeeExt").Value Then
					iWalkThroughs(X) = rstCategories.Fields("WalkThroughs").Value
				End If
			Next 
			rstCategories.MoveNext()
		Loop 
		'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdCategories = Nothing
		rstCategories.Close()
		'
		cmdCategories.ActiveConnection = cnMain
		'
		cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS Sales, tblEmployees.EmployeeExt FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) AND (TSupportActs.Type LIKE N'Sale') GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, tblEmployees.EmployeeExt HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) ORDER BY tblEmployees.EmployeeExt"
		rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		Do Until rstCategories.eof
			For X = 1 To iNumofUsers
				If iExt(X) = rstCategories.Fields("EmployeeExt").Value Then
					iSales(X) = rstCategories.Fields("Sales").Value
				End If
			Next 
			rstCategories.MoveNext()
		Loop 
		'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdCategories = Nothing
		rstCategories.Close()
		'
	End Sub
	
	Private Sub LoadNumOfCalls2()
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		Dim X As Short
		'
		cmdCategories.CommandTimeout = 300
		'
		cmdCategories.ActiveConnection = cnMain
		'
		cmdCategories.CommandText = "SELECT COUNT(DISTINCT ICC_CDR.SESSID) AS Calls, SUM(DISTINCT ICC_CDR.CALLDUR) AS Duration, AVG(ICC_CDR.CALLDUR) AS AvgTime, ICC_CDR.P1NO AS EmployeeExt FROM ICC_CDR RIGHT OUTER JOIN tblEmployees ON ICC_CDR.P1NO = tblEmployees.EmployeeExt WHERE (DATEADD(ss, ICC_CDR.STARTTIME, CONVERT(DATETIME, '1970-01-01 00:00:00', 102)) BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value + 1 & "', 102)) GROUP BY ICC_CDR.P1NO ORDER BY ICC_CDR.P1NO"
		rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		Do Until rstCategories.eof
			For X = 1 To iNumofUsers
				If iExt(X) = rstCategories.Fields("EmployeeExt").Value Then
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(rstCategories.Fields("Calls").Value) Then
						iNumofCalls(X) = rstCategories.Fields("Calls").Value
					End If
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(rstCategories.Fields("Duration").Value) Then
						sCallTime(X) = ConvertTime(rstCategories.Fields("Duration").Value)
						'lCallTime(x) = rstCategories!Duration
					Else
						sCallTime(X) = "No Data"
					End If
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(rstCategories.Fields("AvgTime").Value) Then
						sAvgCallTime(X) = ConvertTime(rstCategories.Fields("AvgTime").Value)
						'lAvgCallTime(x) = rstCategories!AvgTime
					Else
						sAvgCallTime(X) = "No Data"
					End If
				End If
			Next 
			rstCategories.MoveNext()
		Loop 
		'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		cmdCategories = Nothing
		rstCategories.Close()
		'
	End Sub
	
	Private Sub BNBReport()
		Dim X As Short
		Dim sKey As String
		Dim iHolder As Short
		Dim sHolder As String
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If DTPicker1.Value > Today Or DTPicker2.Value < DTPicker1.Value Then
			MsgBox("Incorrect Date Values!")
			Exit Sub
		End If
		grdCallData.Visible = False
		ListView1.Visible = True
		'
		For X = 1 To iNumofUsers
			'iExt(x) = 0
			'sName(x) = 0
			iNumofCalls(X) = 0
			iBNBNotes(X) = 0
			lCallTime(X) = 0
			sCallTime(X) = "0"
			iFollowups(X) = 0
			iWalkThroughs(X) = 0
			iSales(X) = 0
			lAvgCallTime(X) = 0
			sAvgCallTime(X) = "0"
		Next 
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'
		LoadNumOfCalls2()
		LoadBNBData2()
		'
		SetupListView()
		''
		''Enable Print button
		''
		cmdPrintReport.Enabled = True
		
		
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub DurationReport()
		
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		Dim sCallDir As String
		Dim X As Short
		Dim iAvgTime As Integer
		'
		cmdCategories.CommandTimeout = 300
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'
		intWorkgroup = iExt(cboGroup.SelectedIndex + 1)
		'
		grdCallData.Clear()
		grdCallData.set_Cols( , 4)
		grdCallData.set_ColHeader(0, MSHierarchicalFlexGridLib.ColHeaderSettings.flexColHeaderOn)
		grdCallData.set_ColHeaderCaption(0, 0, "Caller's #")
		grdCallData.set_ColHeaderCaption(0, 1, "Date")
		grdCallData.set_ColHeaderCaption(0, 2, "Direction")
		grdCallData.set_ColHeaderCaption(0, 3, "Duration")
		'
		grdCallData.set_ColWidth(0, 0, 1500)
		grdCallData.set_ColWidth(1, 0, 1700)
		grdCallData.set_ColWidth(2, 0, 1000)
		grdCallData.set_ColWidth(3, 0, 1000)
		grdCallData.set_ColWidth(4, 0, 0)
		'
		'
		'Set the date values
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If DTPicker1.Value > DTPicker2.Value Then
			MsgBox("Incorrect Date Values!")
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Exit Sub
		Else
			If DTPicker1.Value > Today Or DTPicker2.Value > Today Then
				MsgBox("Incorrect Date Values!")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				Exit Sub
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DTStart = DTPicker1.Value
				'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DTEnd = DTPicker2.Value
			End If
		End If
		'
		grdCallData.Visible = True
		ListView1.Visible = False
		'
		'Set the sCallDir value
		'
		Select Case cboCallDir.Text
			Case "Incoming"
				sCallDir = " (TKDIR = 2) AND "
				strDirection = "Incoming "
			Case "Outgoing"
				sCallDir = " (TKDIR = 4) AND "
				strDirection = "Outgoing "
			Case "Both"
				sCallDir = " "
				strDirection = "All "
		End Select
		'
		'Set the Ext or Workgroup type and number
		'
		'If optExt.value = True Then
		strGroupType = "P1NO"
		'Else
		'strGroupType = "P1WGNO"
		'End If
		'
		'Create the SQL statement
		'
		strSQL = "SELECT TKRMNO as Phone, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)) AS Calls, TKDIR as Direction, CALLDUR as Duration FROM         ICC_CDR WHERE " & sCallDir & " (" & strGroupType & " LIKE '" & intWorkgroup & "')AND (DATEADD(ss, STARTTIME, '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, STARTTIME, '1969-12-31 19:00:00') < ('" & System.Date.FromOADate(DTEnd.ToOADate + 1) & "')) GROUP BY CALLDUR, TKRMNO, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)), TKDIR HAVING      (TKRMNO IS NOT NULL) ORDER BY DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))"
		strSQL2 = "SELECT TKRMNO AS Phone, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)) AS Calls, Direction = CASE TKDIR WHEN '2' THEN 'Incoming' WHEN '4' THEN 'Outgoing' end,  convert( char(8), dateadd( ss, CALLDUR, '00:00:00' ), 108 ) AS Duration, CALLDUR  From ICC_CDR WHERE " & sCallDir & " (" & strGroupType & " LIKE '" & intWorkgroup & "')AND (DATEADD(ss, STARTTIME, '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, STARTTIME, '1969-12-31 19:00:00') < ('" & System.Date.FromOADate(DTEnd.ToOADate + 1) & "')) GROUP BY CALLDUR, TKRMNO, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)), TKDIR HAVING      (TKRMNO IS NOT NULL) ORDER BY DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))"
		'
		'Select revelant data
		'
		cmdCategories.ActiveConnection = cnMain
		cmdCategories.CommandText = strSQL2
		rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		If rstCategories.RecordCount = 0 Then
			MsgBox("No revelant data!")
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Exit Sub
		End If
		'
		iAvgTime = 0
		rstCategories.MoveFirst()
		For X = 1 To rstCategories.RecordCount
			If Not rstCategories.eof Then
				iAvgTime = iAvgTime + Val(rstCategories.Fields("CALLDUR").Value)
				rstCategories.MoveNext()
			End If
		Next 
		txtTotal.Text = ConvertTime(iAvgTime)
		txtAvg.Text = ConvertTime(iAvgTime / rstCategories.RecordCount)
		'
		grdCallData.Recordset = rstCategories
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		grdCallData.CtlRefresh()
		'
		'Enable Print button
		'
		cmdPrintReport.Enabled = True
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		If chkUpdate.CheckState = System.Windows.Forms.CheckState.Checked Then
			If iMins = Val(cboMins.Text) Then
				cmdGenReport_Click(cmdGenReport, New System.EventArgs())
				iMins = 0
			End If
			iMins = iMins + 1
		End If
	End Sub
	
	Public Sub SortListView(ByRef oListView As System.Windows.Forms.ListView, ByRef oColumnHeader As System.Windows.Forms.ColumnHeader)
		'-- Sorts all list items correctly according to data type.
		'-- Requirements:
		'--     Any items without tag data will be sorted alphabetically.
		'--     When creating the list, add a dummy column to the end, width = 0.
		'--     Must be the last column in the list.
		'--     Create the dummy column subitems as you fill the loop.
		'--     Set .Sorted property = True.
		
		Dim oListItem As System.Windows.Forms.ListViewItem
		Dim i As Short
		Dim iTempColIndex As Short
		Dim bNoTagInColumn As Boolean
		
		With oListView
			
			'-- If 0 or 1 items or -1(uninitialized), then don't try to sort.
			If .Items.Count < 2 Then GoTo Exit_Point
			
			iTempColIndex = .Columns.Count - 1
			
			
			'-- Add the tag data from the clicked-on column to the dummy column.
			If oColumnHeader.Index = 1 Then
				'-- First column gets special treatment.
				For i = 1 To .Items.Count
					'UPGRADE_WARNING: Lower bound of collection oListView.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					oListItem = .Items.Item(i)
					'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oListItem.Tag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oListItem.ListSubItems(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oListItem.SubItems.Item(iTempColIndex) = oListItem.Tag
				Next 
				If Len(Trim(oListItem.Tag)) = 0 Then bNoTagInColumn = True
			Else
				'-- Subcolumns.
				For i = 1 To .Items.Count
					'UPGRADE_WARNING: Lower bound of collection oListView.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					oListItem = .Items.Item(i)
					'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oListItem.ListSubItems().Tag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oListItem.ListSubItems(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oListItem.SubItems.Item(iTempColIndex) = oListItem.SubItems.Item(oColumnHeader.Index - 1).Tag
				Next 
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				If Len(Trim(oListItem.SubItems.Item(iTempColIndex).Text)) = 0 Then bNoTagInColumn = True
			End If
			
			
			If bNoTagInColumn Then
				'-- If the tag is blank, sort by default - alphabetically.
				'UPGRADE_ISSUE: MSComctlLib.ListView property oListView.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				.SortKey = oColumnHeader.Index - 1
			Else
				'-- Otherwise sort by the dummy column.
				'UPGRADE_ISSUE: MSComctlLib.ListView property oListView.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				.SortKey = iTempColIndex
			End If
			
			'-- Sort.
			If .Sorting = System.Windows.Forms.SortOrder.Ascending Then
				.Sorting = System.Windows.Forms.SortOrder.Descending
			Else
				.Sorting = System.Windows.Forms.SortOrder.Ascending
			End If
			
			
			'-- Remove the data so no peeking.
			For i = 1 To .Items.Count
				'UPGRADE_WARNING: Lower bound of collection oListView.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem = .Items.Item(i)
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oListItem.ListSubItems(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oListItem.SubItems.Item(iTempColIndex) = ""
			Next 
			
		End With
		
Exit_Point: 
		'UPGRADE_NOTE: Object oListItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oListItem = Nothing
	End Sub
	
	Private Sub SetupListView()
		
		Dim oListItem As System.Windows.Forms.ListViewItem
		Dim dblDate As Double
		Dim X As Short
		Dim sKey As String
		Dim iHolder As Short
		Dim sHolder As String
		
		
		ListView1.Items.Clear()
		'
		'  For X = 1 To iNumofUsers
		'  '
		'  If iNumofCalls(X) = 0 Or iBNBNotes(X) = 0 Then
		'    If iBNBNotes(X) < 0 Then
		'        iHolder = 100
		'      Else
		'        iHolder = 0
		'      End If
		'    Else
		'      iHolder = (iBNBNotes(X) / iNumofCalls(X)) * 100
		'    End If
		'    sHolder = iHolder & "%"
		'    '
		'    sKey = "A" & CStr(X)
		'    Set oListItem = ListView1.ListItems.Add(, sKey, sName(X))
		'
		'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iNumofCalls(X), , iNumofCalls(X)).ForeColor = vbBlack
		'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iBNBNotes(X), , iBNBNotes(X)).ForeColor = vbBlack
		'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , sHolder, , sHolder).ForeColor = vbBlack
		'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , sCallTime(X), , sCallTime(X)).ForeColor = vbBlack
		'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iFollowups(X), , iFollowups(X)).ForeColor = vbBlack
		'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iWalkThroughs(X), , iWalkThroughs(X)).ForeColor = vbBlack
		'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iSales(X), , iSales(X)).ForeColor = vbBlack
		'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , sAvgCallTime(X), , sAvgCallTime(X)).ForeColor = vbBlack
		'  Next
		
		'-- Put some data in the listview.
		With ListView1.Items
			For X = 1 To iNumofUsers
				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				If iNumofCalls(X) = 0 Or iBNBNotes(X) = 0 Then
					If iBNBNotes(X) < 0 Then
						iHolder = 100
					Else
						iHolder = 0
					End If
				Else
					iHolder = (iBNBNotes(X) / iNumofCalls(X)) * 100
				End If
				sHolder = iHolder & "%"
				'
				sKey = "A" & CStr(X)
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				oListItem = .Add(sKey, sName(X), "")
				'oListItem.Tag = Format(X, "00000")
				
				oListItem.SubItems.Add(CStr(iNumofCalls(X)))
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem.SubItems.Item(1).Tag = VB6.Format(iNumofCalls(X), "0000000")
				
				oListItem.SubItems.Add(CStr(iBNBNotes(X)))
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem.SubItems.Item(2).Tag = VB6.Format(iBNBNotes(X), "0000000")
				
				oListItem.SubItems.Add(sHolder)
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem.SubItems.Item(3).Tag = VB6.Format(sHolder, "0000000000%")
				
				oListItem.SubItems.Add(sCallTime(X))
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem.SubItems.Item(4).Tag = VB6.Format(Val(sCallTime(X)), "yyyymmddHHMMSS")
				
				oListItem.SubItems.Add(CStr(iFollowups(X)))
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem.SubItems.Item(5).Tag = VB6.Format(iFollowups(X), "0000000")
				
				oListItem.SubItems.Add(CStr(iWalkThroughs(X)))
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem.SubItems.Item(6).Tag = VB6.Format(iWalkThroughs(X), "0000000")
				
				oListItem.SubItems.Add(CStr(iSales(X)))
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem.SubItems.Item(7).Tag = VB6.Format(iSales(X), "0000000")
				
				oListItem.SubItems.Add(sAvgCallTime(X))
				'UPGRADE_WARNING: Lower bound of collection oListItem.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				oListItem.SubItems.Item(8).Tag = VB6.Format(Val(sAvgCallTime(X)), "yyyymmddHHMMSS")
				
				'-- Dummy ListSubItems column.
				'UPGRADE_ISSUE: MSComctlLib.ListSubItems method ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				oListItem.SubItems.Add()
				
			Next 
			
		End With
	End Sub
End Class