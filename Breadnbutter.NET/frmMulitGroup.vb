Option Strict Off
Option Explicit On
Friend Class frmMultiChart
	Inherits System.Windows.Forms.Form 'Private rslabels As ADODB.Recordset
	
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	Public bChartData As Boolean
	'
	
	Public strGroupType, strSQL, strDateType, strDirection As String
	Public DTStart, DTEnd As Date
	Public intEXT, intWorkgroup As Short
	'Public chrtArray()
	Public lGreatestValue As Integer
	
	
	'UPGRADE_WARNING: Event cboCallDir.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboCallDir_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCallDir.SelectedIndexChanged
		'
		'Disable Print button
		'
		cmdPrintChart.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event cboChartType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboChartType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboChartType.SelectedIndexChanged
		Select Case cboChartType.Text
			Case "2D Bar"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType2dBar
			Case "3D Bar"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType3dBar
			Case "2D Line"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType2dLine
			Case "3D Line"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType3dLine
			Case "2D Area"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType2dArea
			Case "3D Area"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType3dArea
			Case "2D Step"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType2dStep
			Case "3D Step"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType3dStep
			Case "2D Combo"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType2dCombination
			Case "3D Combo"
				MSChart1.chartType = MSChart20Lib.VtChChartType.VtChChartType3dCombination
		End Select
		
	End Sub
	
	'UPGRADE_WARNING: Event cboDateNum.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboDateNum_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDateNum.SelectedIndexChanged
		Dim x As Short
		For x = 1 To 9
			DTPicker1(x).Enabled = False
		Next x
		For x = 2 To Val(cboDateNum.Text)
			DTPicker1(x - 1).Enabled = True
		Next x
		'
		'Disable Print button
		'
		cmdPrintChart.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event cboDateType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboDateType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDateType.SelectedIndexChanged
		Dim sTemp As String
		Dim x As Short
		'
		Select Case cboDateType.Text
			Case "Hour"
				'        cboYear.Enabled = False
				'        cboMonth.Enabled = False
				'        cboSunday.Enabled = False
				'        DTPicker1(0).Enabled = True
				'        optMultiDates.Enabled = True
				DTPicker2.value = DTPicker1(0).value + 1
			Case "Day"
				'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MonthView1.value = DTPicker1(0).value
				MonthView1.DayOfWeek = MSComCtl2.DayConstants.mvwSunday
				DTPicker1(0).value = MonthView1.value
				'        cboYear.Enabled = True
				'        cboMonth.Enabled = True
				'        cboSunday.Enabled = True
				'        DTPicker1(0).Enabled = False
				'        If optMultiDates.value = True Then
				'          optNone.value = True
				'        End If
				'        optMultiDates.Enabled = False
				'        GetSundays
				DTPicker2.value = DTPicker1(0).value + 7
			Case Else
				'        cboYear.Enabled = True
				'        cboMonth.Enabled = False
				'        cboSunday.Enabled = False
				'        DTPicker1(0).Enabled = False
				'        optMultiDates.Enabled = False
				'        If optMultiDates.value = True Then
				'          optNone.value = True
				'        End If
				'        optMultiDates.Enabled = False
				'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().year. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sTemp = DTPicker1(0).year
				DTPicker1(0).value = "1/1/" & sTemp
				If DTPicker1(0).year = 2004 Then
					DTPicker2.value = DTPicker1(0).value + 366
				Else
					DTPicker2.value = DTPicker1(0).value + 365
				End If
		End Select
		For x = 1 To 9
			DTPicker1_Change(DTPicker1.Item(x), New System.EventArgs())
		Next 
		'
		'Disable Print button
		'
		cmdPrintChart.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event cboGroup.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboGroup_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboGroup.SelectedIndexChanged
		Dim Index As Short = cboGroup.GetIndex(eventSender)
		'
		'Disable Print button
		'
		cmdPrintChart.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event cboMonth.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboMonth_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboMonth.SelectedIndexChanged
		GetSundays()
		MonthView1.month = CShort(cboMonth.Text)
		MonthView1.year = CShort(cboYear.Text)
		cboSunday.Text = CStr(MonthView1.day)
		'MonthView1.DayOfWeek = 1
		DTPicker1(0).value = MonthView1.value
		GetSundays()
	End Sub
	
	'UPGRADE_WARNING: Event cboNum.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboNum_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboNum.SelectedIndexChanged
		Dim x As Short
		For x = 1 To 9
			cboGroup(x).Enabled = False
		Next x
		For x = 2 To Val(cboNum.Text)
			cboGroup(x - 1).Enabled = True
		Next x
		'
		'Disable Print button
		'
		cmdPrintChart.Enabled = False
	End Sub
	
	'UPGRADE_WARNING: Event cboSunday.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboSunday_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSunday.SelectedIndexChanged
		MonthView1.year = CShort(cboYear.Text)
		MonthView1.day = CShort(cboSunday.Text)
		'MonthView1.DayOfWeek = 1
		cboMonth.Text = CStr(MonthView1.month)
		DTPicker1(0).value = MonthView1.value
	End Sub
	
	'UPGRADE_WARNING: Event cboYear.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboYear_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboYear.SelectedIndexChanged
		GetSundays()
		If cboDateType.Text = "Day" Then
			MonthView1.year = CShort(cboYear.Text)
			MonthView1.day = CShort(cboSunday.Text)
			'MonthView1.DayOfWeek = 1
			cboMonth.Text = CStr(MonthView1.month)
		Else
			MonthView1.value = CDate("1/1/" & cboYear.Text)
			cboMonth.Text = CStr(MonthView1.month)
			cboSunday.Text = CStr(MonthView1.day)
		End If
		DTPicker1(0).value = MonthView1.value
	End Sub
	
	'Private Sub chkVoiceMail_Click()
	''
	''Disable Print button
	''
	'  cmdPrintChart.Enabled = False
	''
	''If checked Disable Call Direction and set to Incoming
	''
	'  If chkVoiceMail.Value = 1 Then
	'    cboCallDir.Text = "Incoming"
	'    cboCallDir.Enabled = False
	'  Else
	'    If optExt.Value = True Then
	'      cboCallDir.Enabled = True
	'    End If
	'  End If
	'End Sub
	
	Private Sub cmdAvg_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAvg.Click
		Dim sVoiceMail As String
		Dim sCallDir As String
		Dim x As Short
		'UPGRADE_WARNING: Lower bound of array iInterval was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim iInterval(100) As Short
		Dim IAvg As Short
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		'
		cmdCategories.CommandTimeout = 300
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'
		intWorkgroup = CShort(cboGroup(0).Text)
		'
		' VoiceMail
		'
		If optVoiceMail.Checked = True Then
			sVoiceMail = "VMSTARTTM"
			strDirection = "VoiceMail "
		Else
			sVoiceMail = "STARTTIME"
			
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
					strDirection = ""
			End Select
			'
			If optCalls.Checked = True Then
				sCallDir = sCallDir & "VMSTARTTM = 0 AND "
			Else
				strDirection = "Total " & strDirection
			End If
			'
		End If
		
		'
		'Set the date values
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker4.value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker3.value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If DTPicker3.value > DTPicker4.value Then
			MsgBox("Incorrect Date Values!")
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Exit Sub
		Else
			If DTPicker3.value > Today Or DTPicker4.value > Today Then
				MsgBox("Incorrect Date Values!")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				Exit Sub
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker3.value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DTStart = DTPicker3.value
				'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker4.value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DTEnd = DTPicker4.value
			End If
		End If
		'
		'Set the DATEPART value
		'
		Select Case cboDateType.Text
			Case "Hour"
				strDateType = "hh"
				lGreatestValue = 24
				For x = 1 To 24
					iInterval(x) = (DTEnd.ToOADate - DTStart.ToOADate)
				Next 
			Case "Day"
				strDateType = "w"
				lGreatestValue = 7
				'
				For x = 0 To System.Date.FromOADate(DTEnd.ToOADate - DTStart.ToOADate).ToOADate
					IAvg = DatePart(Microsoft.VisualBasic.DateInterval.WeekDay, System.Date.FromOADate(DTStart.ToOADate + x), FirstDayOfWeek.Sunday, FirstWeekOfYear.System)
					iInterval(IAvg) = iInterval(IAvg) + 1
				Next 
			Case "Week"
				strDateType = "ww"
				lGreatestValue = 52
			Case "Month"
				strDateType = "m"
				lGreatestValue = 12
			Case "Year"
				'strDateType = "yyyy"
				Exit Sub
		End Select
		'
		'Set the Ext or Workgroup type and number
		'
		If optExt.Checked = True Then
			strGroupType = "P1NO"
		Else
			strGroupType = "P1WGNO"
		End If
		
		
		
		
		'
		'Create the SQL statement
		'
		strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & intWorkgroup & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & DTEnd & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
		'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
		'Select revelant data
		
		cmdCategories.ActiveConnection = cnMain
		cmdCategories.CommandText = strSQL
		rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		
		If rstCategories.RecordCount = 0 Then
			MsgBox("No revelant data!")
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			'Exit Sub
		End If
		bChartData = True 'set flag to true
		'                    C H A R T
		'
		' Dynamic 2-dimensional array to store series
		' The first index (x) is the total number of series
		' x-axis value in the 1st slot (i.e. chrtArray(x,1)
		' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
		'
		'ReDim chrtArray(1 To lGreatestValue, 1 To 2)
		
		MSChart1.ShowLegend = True
		'MSChart1.chartType = VtChChartType2dLine
		'
		'Chart Title centered on top
		'
		MSChart1.Title.Text = strDirection & "Calls between " & DTPicker3.value & " and " & DTPicker4.value
		'
		'Chart X and Y axis titles
		'
		MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = cboDateType.Text
		MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = ""
		'
		'Find the minimum and maximum date value in the record
		'
		Dim iMinVal, iMaxVal As Short
		rstCategories.MoveFirst()
		iMinVal = rstCategories.Fields("datetype").Value
		rstCategories.MoveLast()
		iMaxVal = rstCategories.Fields("datetype").Value
		rstCategories.MoveFirst()
		If iMinVal = 0 Then
			iMinVal = 1
			iMaxVal = 24
		End If
		'
		'UPGRADE_WARNING: Lower bound of array chrtArray was changed from iMinVal,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim chrtArray(iMaxVal, 2) As Object
		'
		'Load the array with 0s with correct # of rows
		'
		For x = iMinVal To iMaxVal
			If x = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(24, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chrtArray(24, 1) = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(24, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chrtArray(24, 2) = 24
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(x, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chrtArray(x, 1) = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(x, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chrtArray(x, 2) = x
			End If
		Next x
		'
		'Load the array with data
		'
		For x = 1 To rstCategories.RecordCount
			If strDateType = "hh" And rstCategories.Fields("datetype").Value = 0 Then
				rstCategories.Fields("datetype").Value = 24
			End If
			If iInterval(x) = 0 Then iInterval(x) = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(rstCategories!datetype, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			chrtArray(rstCategories.Fields("datetype").Value, 1) = rstCategories.Fields("Calls").Value / iInterval(x)
			rstCategories.MoveNext()
		Next x
		'
		'Attach the array of data to MS-CHART
		'
		With MSChart1
			.ChartData = VB6.CopyArray(chrtArray)
			.ColumnCount = 1
			.ColumnLabelCount = 1
			.Column = 1
			.ColumnLabel = "Calls for " & cboGroup(0).Text
			For x = 1 To iMaxVal - iMinVal + 1
				.Row = x
				Select Case strDateType
					Case "hh"
						Select Case chrtArray(iMinVal - 1 + x, 2)
							Case 1
								.RowLabel = "1 AM"
							Case 2
								.RowLabel = "2 AM"
							Case 3
								.RowLabel = "3 AM"
							Case 4
								.RowLabel = "4 AM"
							Case 5
								.RowLabel = "5 AM"
							Case 6
								.RowLabel = "6 AM"
							Case 7
								.RowLabel = "7 AM"
							Case 8
								.RowLabel = "8 AM"
							Case 9
								.RowLabel = "9 AM"
							Case 10
								.RowLabel = "10 AM"
							Case 11
								.RowLabel = "11 AM"
							Case 12
								.RowLabel = "12 AM"
							Case 13
								.RowLabel = "1 PM"
							Case 14
								.RowLabel = "2 PM"
							Case 15
								.RowLabel = "3 PM"
							Case 16
								.RowLabel = "4 PM"
							Case 17
								.RowLabel = "5 PM"
							Case 18
								.RowLabel = "6 PM"
							Case 19
								.RowLabel = "7 PM"
							Case 20
								.RowLabel = "8 PM"
							Case 21
								.RowLabel = "9 PM"
							Case 22
								.RowLabel = "10 PM"
							Case 23
								.RowLabel = "11 PM"
							Case 24
								.RowLabel = "12 PM"
						End Select
					Case "w"
						Select Case chrtArray(iMinVal - 1 + x, 2)
							Case 1
								.RowLabel = "Sun"
							Case 2
								.RowLabel = "Mon"
							Case 3
								.RowLabel = "Tue"
							Case 4
								.RowLabel = "Wed"
							Case 5
								.RowLabel = "Thur"
							Case 6
								.RowLabel = "Fri"
							Case 7
								.RowLabel = "Sat"
						End Select
					Case "m"
						Select Case chrtArray(iMinVal - 1 + x, 2)
							Case 1
								.RowLabel = "Jan"
							Case 2
								.RowLabel = "Feb"
							Case 3
								.RowLabel = "Mar"
							Case 4
								.RowLabel = "Apr"
							Case 5
								.RowLabel = "May"
							Case 6
								.RowLabel = "Jun"
							Case 7
								.RowLabel = "Jul"
							Case 8
								.RowLabel = "Aug"
							Case 9
								.RowLabel = "Sep"
							Case 10
								.RowLabel = "Oct"
							Case 11
								.RowLabel = "Nov"
							Case 12
								.RowLabel = "Dec"
						End Select
					Case Else
						'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(iMinVal - 1 + x, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.RowLabel = chrtArray(iMinVal - 1 + x, 2)
				End Select
			Next x
		End With
		If chkLablePoints.CheckState = System.Windows.Forms.CheckState.Checked Then LablePoints()
		'
		'Enable Print button
		'
		cmdPrintChart.Enabled = True
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		'
	End Sub
	
	Private Sub cmdChart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdChart.Click
		
		Dim sVoiceMail As String
		Dim sCallDir As String
		Dim x As Short
		Dim y As Short
		Dim iLines As Short
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		'
		cmdCategories.CommandTimeout = 300
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'
		'intWorkgroup = cboGroup.Text
		If optMultiDates.Checked = True And cboDateNum.Text = "" Or optMultiGroups.Checked = True And cboNum.Text = "" Then
			optNone.Checked = True
		End If
		'
		'  If DTPicker1(0).value > DTPicker2.value Then
		'      MsgBox "Incorrect Date Values!"
		'      Exit Sub
		'    Else
		'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
		'        MsgBox "Incorrect Date Values!"
		'        Exit Sub
		'      Else
		'        DTStart = DTPicker1(0).value
		'        DTEnd = DTPicker2.value
		'      End If
		'    End If
		'
		Select Case cboDateType.Text
			Case "Hour"
				strDateType = "hh"
				lGreatestValue = 24
			Case "Day"
				strDateType = "w"
				lGreatestValue = 7
			Case "Week"
				strDateType = "ww"
				lGreatestValue = 53
			Case "Month"
				strDateType = "m"
				lGreatestValue = 12
			Case "Year"
				'strDateType = "yyyy"
				Exit Sub
		End Select
		'
		If optMultiGroups.Checked = True Then
			'
			' VoiceMail
			'
			If optVoiceMail.Checked = True Then
				sVoiceMail = "VMSTARTTM"
				strDirection = "VoiceMail "
			Else
				sVoiceMail = "STARTTIME"
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
						strDirection = ""
				End Select
				'
				If optCalls.Checked = True Then
					sCallDir = sCallDir & "VMSTARTTM = 0 AND "
				Else
					strDirection = "Total " & strDirection
				End If
				'
			End If
			'
			'Set the date values
			'
			'    If DTPicker1(0).value > DTPicker2.value Then
			'      MsgBox "Incorrect Date Values!"
			'      Exit Sub
			'    Else
			'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
			'        MsgBox "Incorrect Date Values!"
			'        Exit Sub
			'      Else
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DTStart = DTPicker1(0).value
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2.value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DTEnd = DTPicker2.value
			'      End If
			'    End If
			'
			'Set the DATEPART value
			'
			'    Select Case cboDateType
			'      Case "Hour"
			'        strDateType = "hh"
			'        lGreatestValue = 24
			'      Case "Day"
			'        strDateType = "w"
			'        lGreatestValue = 7
			'      Case "Week"
			'        strDateType = "ww"
			'        lGreatestValue = 53
			'      Case "Month"
			'        strDateType = "m"
			'        lGreatestValue = 12
			'      Case "Year"
			'        'strDateType = "yyyy"
			'        Exit Sub
			'    End Select
			'
			'Set the Ext or Workgroup type and number
			'
			If optExt.Checked = True Then
				strGroupType = "P1NO"
			Else
				strGroupType = "P1WGNO"
			End If
			'                    C H A R T
			'
			' Dynamic 2-dimensional array to store series
			' The first index (x) is the total number of series
			' x-axis value in the 1st slot (i.e. chrtArray(x,1)
			' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
			'
			If cboNum.Text = "" Then cboNum.Text = CStr(1)
			'UPGRADE_WARNING: Lower bound of array chrtArray was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			Dim chrtArray(lGreatestValue, Val(cboNum.Text)) As Object
			MSChart1.ShowLegend = True
			'MSChart1.chartType = VtChChartType2dLine
			'
			'Chart Title centered on top
			'
			MSChart1.Title.Text = strDirection & "Calls between " & DTPicker1(0).value & " and " & DTPicker2.value
			'
			'Chart X and Y axis titles
			'
			MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = cboDateType.Text
			MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = ""
			'
			'Chart Foot note
			'
			'MSChart1.FootnoteText = "footnote"
			'
			'Load the array with 0s with correct # of rows
			'
			For x = 1 To lGreatestValue
				For y = 1 To Val(cboNum.Text)
					'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(x, y). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chrtArray(x, y) = 0
				Next y
			Next x
			For x = 0 To Val(cboNum.Text) - 1
				'
				'Create the SQL statement
				'
				strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & cboGroup(x).Text & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & System.Date.FromOADate(DTEnd.ToOADate + 1) & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
				'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
				'
				'Select revelant data
				'
				cmdCategories.ActiveConnection = cnMain
				cmdCategories.CommandText = strSQL
				rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
				rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
				If rstCategories.RecordCount = 0 Then
					MsgBox("No revelant data for " & cboGroup(x).Text & "!")
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
					'Exit Sub
				End If
				bChartData = True 'set flag to true
				'
				'Load the array with data
				'
				For y = 1 To rstCategories.RecordCount
					If strDateType = "hh" And rstCategories.Fields("datetype").Value = 0 Then 'this is for the 0 hour problem
						rstCategories.Fields("datetype").Value = 24
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(rstCategories!datetype, x + 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chrtArray(rstCategories.Fields("datetype").Value, x + 1) = rstCategories.Fields("Calls").Value
					rstCategories.MoveNext()
				Next y
				rstCategories.Close()
			Next x
			'
			'Attach the array of data to MS-CHART
			'
			With MSChart1
				.ChartData = VB6.CopyArray(chrtArray)
				.ColumnCount = Val(cboNum.Text)
				.ColumnLabelCount = Val(cboNum.Text) - 1
				For x = 0 To Val(cboNum.Text) - 1
					.Column = x + 1
					.ColumnLabel = "Calls for " & cboGroup(x).Text
				Next x
				Select Case strDateType
					Case "hh"
						.Row = 1
						.RowLabel = "1 AM"
						.Row = 2
						.RowLabel = "2 AM"
						.Row = 3
						.RowLabel = "3 AM"
						.Row = 4
						.RowLabel = "4 AM"
						.Row = 5
						.RowLabel = "5 AM"
						.Row = 6
						.RowLabel = "6 AM"
						.Row = 7
						.RowLabel = "7 AM"
						.Row = 8
						.RowLabel = "8 AM"
						.Row = 9
						.RowLabel = "9 AM"
						.Row = 10
						.RowLabel = "10 AM"
						.Row = 11
						.RowLabel = "11 AM"
						.Row = 12
						.RowLabel = "12 AM"
						.Row = 13
						.RowLabel = "1 PM"
						.Row = 14
						.RowLabel = "2 PM"
						.Row = 15
						.RowLabel = "3 PM"
						.Row = 16
						.RowLabel = "4 PM"
						.Row = 17
						.RowLabel = "5 PM"
						.Row = 18
						.RowLabel = "6 PM"
						.Row = 19
						.RowLabel = "7 PM"
						.Row = 20
						.RowLabel = "8 PM"
						.Row = 21
						.RowLabel = "9 PM"
						.Row = 22
						.RowLabel = "10 PM"
						.Row = 23
						.RowLabel = "11 PM"
						.Row = 24
						.RowLabel = "12 PM"
					Case "w"
						.Row = 1
						.RowLabel = "Sun"
						.Row = 2
						.RowLabel = "Mon"
						.Row = 3
						.RowLabel = "Tue"
						.Row = 4
						.RowLabel = "Wed"
						.Row = 5
						.RowLabel = "Thur"
						.Row = 6
						.RowLabel = "Fri"
						.Row = 7
						.RowLabel = "Sat"
					Case "m"
						.Row = 1
						.RowLabel = "Jan"
						.Row = 2
						.RowLabel = "Feb"
						.Row = 3
						.RowLabel = "Mar"
						.Row = 4
						.RowLabel = "Apr"
						.Row = 5
						.RowLabel = "May"
						.Row = 6
						.RowLabel = "Jun"
						.Row = 7
						.RowLabel = "Jul"
						.Row = 8
						.RowLabel = "Aug"
						.Row = 9
						.RowLabel = "Sep"
						.Row = 10
						.RowLabel = "Oct"
						.Row = 11
						.RowLabel = "Nov"
						.Row = 12
						.RowLabel = "Dec"
					Case Else
						For y = 1 To lGreatestValue
							.Row = y
							'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(y, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.RowLabel = chrtArray(y, 2)
						Next y
						'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						.CtlRefresh()
				End Select
			End With
		End If
		'
		'
		'
		'
		'
		'
		If optMultiDates.Checked = True Then
			'
			' VoiceMail
			'
			If optVoiceMail.Checked = True Then
				sVoiceMail = "VMSTARTTM"
				strDirection = "VoiceMail "
			Else
				sVoiceMail = "STARTTIME"
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
						strDirection = ""
				End Select
				'
				If optCalls.Checked = True Then
					sCallDir = sCallDir & "VMSTARTTM = 0 AND "
				Else
					strDirection = "Total " & strDirection
				End If
				'
			End If
			'
			'Set the DATEPART value
			'
			'    Select Case cboDateType
			'      Case "Hour"
			'        strDateType = "hh"
			'        lGreatestValue = 24
			'      Case "Day"
			'        strDateType = "w"
			'        lGreatestValue = 7
			'      Case "Week"
			'        strDateType = "ww"
			'        lGreatestValue = 53
			'      Case "Month"
			'        strDateType = "m"
			'        lGreatestValue = 12
			'      Case "Year"
			'        'strDateType = "yyyy"
			'        Exit Sub
			'    End Select
			'
			'Set the Ext or Workgroup type and number
			'
			If optExt.Checked = True Then
				strGroupType = "P1NO"
			Else
				strGroupType = "P1WGNO"
			End If
			'                    C H A R T
			'
			' Dynamic 2-dimensional array to store series
			' The first index (x) is the total number of series
			' x-axis value in the 1st slot (i.e. chrtArray(x,1)
			' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
			'
			If Val(cboDateNum.Text) = 0 Then
				cboDateNum.Text = "1"
			End If
			'UPGRADE_WARNING: Lower bound of array chrtArray was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim chrtArray(lGreatestValue, Val(cboDateNum.Text))
			MSChart1.ShowLegend = True
			'MSChart1.chartType = VtChChartType2dLine
			'
			'Chart X and Y axis titles
			'
			MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = cboDateType.Text
			MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = ""
			'
			'Set the date values
			'
			'    If DTPicker1(0).value > DTPicker2.value Then
			'      MsgBox "Incorrect Date Values!"
			'      Exit Sub
			'    Else
			'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
			'        MsgBox "Incorrect Date Values!"
			'        Exit Sub
			'      Else
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DTStart = DTPicker1(0).value
			DTEnd = DTPicker2.value - DTPicker1(0).value
			'      End If
			'    End If
			'
			'Chart Title centered on top
			'
			MSChart1.Title.Text = strDirection & "Calls for " & cboGroup(0).Text '& ", " & DTEnd & " " & cboDateType.Text & "s of data per line"
			'
			'Load the array with 0s with correct # of rows
			'
			For x = 1 To lGreatestValue
				For y = 1 To Val(cboDateNum.Text)
					'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(x, y). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chrtArray(x, y) = 0
				Next y
			Next x
			For x = 0 To Val(cboDateNum.Text) - 1
				'
				'Create the SQL statement
				'
				strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & cboGroup(0).Text & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTPicker1(x).value & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & DTPicker1(x).value + DTEnd + 1 & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
				'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
				'
				'Select revelant data
				'
				cmdCategories.ActiveConnection = cnMain
				cmdCategories.CommandText = strSQL
				rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
				rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
				If rstCategories.RecordCount = 0 Then
					MsgBox("No revelant data for " & DTPicker1(x).value & "!")
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
					'Exit Sub
				End If
				bChartData = True 'set flag to true
				'
				'Load the array with data
				'
				For y = 1 To rstCategories.RecordCount
					If strDateType = "hh" And rstCategories.Fields("datetype").Value = 0 Then
						rstCategories.Fields("datetype").Value = 24
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(rstCategories!datetype, x + 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chrtArray(rstCategories.Fields("datetype").Value, x + 1) = rstCategories.Fields("Calls").Value
					rstCategories.MoveNext()
				Next y
				rstCategories.Close()
			Next x
			'
			'Attach the array of data to MS-CHART
			'
			With MSChart1
				.ChartData = VB6.CopyArray(chrtArray)
				.ColumnCount = Val(cboDateNum.Text)
				.ColumnLabelCount = Val(cboDateNum.Text)
				For x = 0 To Val(cboDateNum.Text) - 1
					.Column = x + 1
					.ColumnLabel = cboGroup(0).Text & " on " & DTPicker1(x).value
				Next x
				Select Case strDateType
					Case "hh"
						.Row = 1
						.RowLabel = "1 AM"
						.Row = 2
						.RowLabel = "2 AM"
						.Row = 3
						.RowLabel = "3 AM"
						.Row = 4
						.RowLabel = "4 AM"
						.Row = 5
						.RowLabel = "5 AM"
						.Row = 6
						.RowLabel = "6 AM"
						.Row = 7
						.RowLabel = "7 AM"
						.Row = 8
						.RowLabel = "8 AM"
						.Row = 9
						.RowLabel = "9 AM"
						.Row = 10
						.RowLabel = "10 AM"
						.Row = 11
						.RowLabel = "11 AM"
						.Row = 12
						.RowLabel = "12 AM"
						.Row = 13
						.RowLabel = "1 PM"
						.Row = 14
						.RowLabel = "2 PM"
						.Row = 15
						.RowLabel = "3 PM"
						.Row = 16
						.RowLabel = "4 PM"
						.Row = 17
						.RowLabel = "5 PM"
						.Row = 18
						.RowLabel = "6 PM"
						.Row = 19
						.RowLabel = "7 PM"
						.Row = 20
						.RowLabel = "8 PM"
						.Row = 21
						.RowLabel = "9 PM"
						.Row = 22
						.RowLabel = "10 PM"
						.Row = 23
						.RowLabel = "11 PM"
						.Row = 24
						.RowLabel = "12 PM"
					Case "w"
						.Row = 1
						.RowLabel = "Sun"
						.Row = 2
						.RowLabel = "Mon"
						.Row = 3
						.RowLabel = "Tue"
						.Row = 4
						.RowLabel = "Wed"
						.Row = 5
						.RowLabel = "Thur"
						.Row = 6
						.RowLabel = "Fri"
						.Row = 7
						.RowLabel = "Sat"
					Case "m"
						.Row = 1
						.RowLabel = "Jan"
						.Row = 2
						.RowLabel = "Feb"
						.Row = 3
						.RowLabel = "Mar"
						.Row = 4
						.RowLabel = "Apr"
						.Row = 5
						.RowLabel = "May"
						.Row = 6
						.RowLabel = "Jun"
						.Row = 7
						.RowLabel = "Jul"
						.Row = 8
						.RowLabel = "Aug"
						.Row = 9
						.RowLabel = "Sep"
						.Row = 10
						.RowLabel = "Oct"
						.Row = 11
						.RowLabel = "Nov"
						.Row = 12
						.RowLabel = "Dec"
					Case Else
						For y = 1 To lGreatestValue
							.Row = y
							'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(y, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.RowLabel = chrtArray(y, 2)
						Next y
						'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						.CtlRefresh()
				End Select
			End With
		End If
		'
		'
		'
		'
		'
		'
		'
		If optDirection.Checked = True Then
			'
			' VoiceMail
			'
			If optVoiceMail.Checked = True Then
				sVoiceMail = "VMSTARTTM"
				strDirection = "VoiceMail "
			Else
				sVoiceMail = "STARTTIME"
				'
				'Set the sCallDir value
				'
				'      Select Case cboCallDir
				'        Case "Incoming"
				'          sCallDir = " (TKDIR = 2) AND "
				'          strDirection = "Incoming "
				'        Case "Outgoing"
				'          sCallDir = " (TKDIR = 4) AND "
				'          strDirection = "Outgoing "
				'        Case "Both"
				'          sCallDir = " "
				'          strDirection = ""
				'      End Select
				'
				If optCalls.Checked = True Then
					sCallDir = sCallDir & "VMSTARTTM = 0 AND "
				Else
					strDirection = "Total " & strDirection
				End If
				'
			End If
			'
			'Set the date values
			'
			'    If DTPicker1(0).value > DTPicker2.value Then
			'      MsgBox "Incorrect Date Values!"
			'      Exit Sub
			'    Else
			'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
			'        MsgBox "Incorrect Date Values!"
			'        Exit Sub
			'      Else
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DTStart = DTPicker1(0).value
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2.value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DTEnd = DTPicker2.value
			'      End If
			'    End If
			'
			'Set the DATEPART value
			'
			'    Select Case cboDateType
			'      Case "Hour"
			'        strDateType = "hh"
			'        lGreatestValue = 24
			'      Case "Day"
			'        strDateType = "w"
			'        lGreatestValue = 7
			'      Case "Week"
			'        strDateType = "ww"
			'        lGreatestValue = 53
			'      Case "Month"
			'        strDateType = "m"
			'        lGreatestValue = 12
			'      Case "Year"
			'        'strDateType = "yyyy"
			'        Exit Sub
			'    End Select
			'
			'Set the Ext or Workgroup type and number
			'
			If optExt.Checked = True Then
				strGroupType = "P1NO"
			Else
				strGroupType = "P1WGNO"
			End If
			'                    C H A R T
			'
			' Dynamic 2-dimensional array to store series
			' The first index (x) is the total number of series
			' x-axis value in the 1st slot (i.e. chrtArray(x,1)
			' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
			'
			
			'UPGRADE_WARNING: Lower bound of array chrtArray was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim chrtArray(lGreatestValue, 3)
			MSChart1.ShowLegend = True
			'MSChart1.chartType = VtChChartType2dLine
			'
			'Chart Title centered on top
			'
			MSChart1.Title.Text = "Calls between " & DTPicker1(0).value & " and " & DTPicker2.value
			'
			'Chart X and Y axis titles
			'
			MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = cboDateType.Text
			MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = ""
			'
			'Chart Foot note
			'
			'MSChart1.FootnoteText = "footnote"
			'
			'Load the array with 0s with correct # of rows
			'
			For x = 1 To lGreatestValue
				For y = 1 To 3
					'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(x, y). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chrtArray(x, y) = 0
				Next y
			Next x
			For x = 1 To 3
				If x = 1 Then sCallDir = " (TKDIR = 2) AND "
				If x = 2 Then sCallDir = " (TKDIR = 4) AND "
				If x = 3 Then sCallDir = " "
				'
				'Create the SQL statement
				'
				strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & cboGroup(0).Text & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & System.Date.FromOADate(DTEnd.ToOADate + 1) & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
				'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
				'
				'Select revelant data
				'
				cmdCategories.ActiveConnection = cnMain
				cmdCategories.CommandText = strSQL
				rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
				rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
				If rstCategories.RecordCount = 0 Then
					MsgBox("No revelant data for " & cboGroup(0).Text & "!")
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
					'Exit Sub
				End If
				bChartData = True 'set flag to true
				'
				'Load the array with data
				'
				For y = 1 To rstCategories.RecordCount
					If strDateType = "hh" And rstCategories.Fields("datetype").Value = 0 Then
						rstCategories.Fields("datetype").Value = 24
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(rstCategories!datetype, x). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					chrtArray(rstCategories.Fields("datetype").Value, x) = rstCategories.Fields("Calls").Value
					rstCategories.MoveNext()
				Next y
				rstCategories.Close()
			Next x
			'
			'Attach the array of data to MS-CHART
			'
			With MSChart1
				.ChartData = VB6.CopyArray(chrtArray)
				.ColumnCount = 3
				.ColumnLabelCount = 3
				.Column = 1
				.ColumnLabel = "Incoming"
				.Column = 2
				.ColumnLabel = "Outgoing"
				.Column = 3
				.ColumnLabel = "Both"
				Select Case strDateType
					Case "hh"
						.Row = 1
						.RowLabel = "1 AM"
						.Row = 2
						.RowLabel = "2 AM"
						.Row = 3
						.RowLabel = "3 AM"
						.Row = 4
						.RowLabel = "4 AM"
						.Row = 5
						.RowLabel = "5 AM"
						.Row = 6
						.RowLabel = "6 AM"
						.Row = 7
						.RowLabel = "7 AM"
						.Row = 8
						.RowLabel = "8 AM"
						.Row = 9
						.RowLabel = "9 AM"
						.Row = 10
						.RowLabel = "10 AM"
						.Row = 11
						.RowLabel = "11 AM"
						.Row = 12
						.RowLabel = "12 AM"
						.Row = 13
						.RowLabel = "1 PM"
						.Row = 14
						.RowLabel = "2 PM"
						.Row = 15
						.RowLabel = "3 PM"
						.Row = 16
						.RowLabel = "4 PM"
						.Row = 17
						.RowLabel = "5 PM"
						.Row = 18
						.RowLabel = "6 PM"
						.Row = 19
						.RowLabel = "7 PM"
						.Row = 20
						.RowLabel = "8 PM"
						.Row = 21
						.RowLabel = "9 PM"
						.Row = 22
						.RowLabel = "10 PM"
						.Row = 23
						.RowLabel = "11 PM"
						.Row = 24
						.RowLabel = "12 PM"
					Case "w"
						.Row = 1
						.RowLabel = "Sun"
						.Row = 2
						.RowLabel = "Mon"
						.Row = 3
						.RowLabel = "Tue"
						.Row = 4
						.RowLabel = "Wed"
						.Row = 5
						.RowLabel = "Thur"
						.Row = 6
						.RowLabel = "Fri"
						.Row = 7
						.RowLabel = "Sat"
					Case "m"
						.Row = 1
						.RowLabel = "Jan"
						.Row = 2
						.RowLabel = "Feb"
						.Row = 3
						.RowLabel = "Mar"
						.Row = 4
						.RowLabel = "Apr"
						.Row = 5
						.RowLabel = "May"
						.Row = 6
						.RowLabel = "Jun"
						.Row = 7
						.RowLabel = "Jul"
						.Row = 8
						.RowLabel = "Aug"
						.Row = 9
						.RowLabel = "Sep"
						.Row = 10
						.RowLabel = "Oct"
						.Row = 11
						.RowLabel = "Nov"
						.Row = 12
						.RowLabel = "Dec"
					Case Else
						For y = 1 To lGreatestValue
							.Row = y
							'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(y, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.RowLabel = chrtArray(y, 2)
						Next y
						'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						.CtlRefresh()
				End Select
			End With
		End If
		'
		'
		'
		'
		
		If optNone.Checked = True Then
			
			intWorkgroup = CShort(cboGroup(0).Text)
			'
			' VoiceMail
			'
			If optVoiceMail.Checked = True Then
				sVoiceMail = "VMSTARTTM"
				strDirection = "VoiceMail "
			Else
				sVoiceMail = "STARTTIME"
				
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
						strDirection = ""
				End Select
				'
				If optCalls.Checked = True Then
					sCallDir = sCallDir & "VMSTARTTM = 0 AND "
				Else
					strDirection = "Total " & strDirection
				End If
				'
			End If
			
			'
			'Set the date values
			'
			'    If DTPicker1(0).value > DTPicker2.value Then
			'        MsgBox "Incorrect Date Values!"
			'        Exit Sub
			'    Else
			'        If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
			'            MsgBox "Incorrect Date Values!"
			'            Exit Sub
			'        Else
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DTStart = DTPicker1(0).value
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2.value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DTEnd = DTPicker2.value
			'        End If
			'    End If
			'
			'Set the DATEPART value
			'
			'    Select Case cboDateType
			'        Case "Hour"
			'            strDateType = "hh"
			'            lGreatestValue = 24
			'        Case "Day"
			'            strDateType = "w"
			'            lGreatestValue = 7
			'        Case "Week"
			'            strDateType = "ww"
			'            lGreatestValue = 53
			'        Case "Month"
			'            strDateType = "m"
			'            lGreatestValue = 12
			'        Case "Year"
			'            'strDateType = "yyyy"
			'            Exit Sub
			'    End Select
			'
			'Set the Ext or Workgroup type and number
			'
			If optExt.Checked = True Then
				strGroupType = "P1NO"
			Else
				strGroupType = "P1WGNO"
			End If
			
			
			
			
			'
			'Create the SQL statement
			'
			strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & intWorkgroup & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & System.Date.FromOADate(DTEnd.ToOADate + 1) & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
			'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
			'Select revelant data
			
			cmdCategories.ActiveConnection = cnMain
			cmdCategories.CommandText = strSQL
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			
			If rstCategories.RecordCount = 0 Then
				MsgBox("No revelant data!")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				'Exit Sub
			End If
			bChartData = True 'set flag to true
			'                    C H A R T
			'
			' Dynamic 2-dimensional array to store series
			' The first index (x) is the total number of series
			' x-axis value in the 1st slot (i.e. chrtArray(x,1)
			' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
			'
			'UPGRADE_WARNING: Lower bound of array chrtArray was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim chrtArray(lGreatestValue, 2)
			MSChart1.ShowLegend = True
			'MSChart1.chartType = VtChChartType2dLine
			'
			'Chart Title centered on top
			'
			MSChart1.Title.Text = strDirection & "Calls between " & DTPicker1(0).value & " and " & DTPicker2.value
			'
			'Chart X and Y axis titles
			'
			MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = cboDateType.Text
			MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = ""
			'
			'Chart Foot note
			'
			'MSChart1.FootnoteText = "footnote"
			'
			'Load the array with 0s with correct # of rows
			'
			For x = 1 To lGreatestValue
				'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(x, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chrtArray(x, 1) = 0
				'chrtArray(X, 2) = X
			Next x
			'
			'Load the array with data
			'
			For x = 1 To rstCategories.RecordCount
				If strDateType = "hh" And rstCategories.Fields("datetype").Value = 0 Then
					rstCategories.Fields("datetype").Value = 24
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(rstCategories!datetype, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chrtArray(rstCategories.Fields("datetype").Value, 1) = rstCategories.Fields("Calls").Value
				rstCategories.MoveNext()
			Next x
			'
			'Attach the array of data to MS-CHART
			'
			With MSChart1
				.ChartData = VB6.CopyArray(chrtArray)
				.ColumnCount = 1
				.ColumnLabelCount = 1
				.Column = 1
				.ColumnLabel = "Calls for " & cboGroup(0).Text
				Select Case strDateType
					Case "hh"
						.Row = 1
						.RowLabel = "1 AM"
						.Row = 2
						.RowLabel = "2 AM"
						.Row = 3
						.RowLabel = "3 AM"
						.Row = 4
						.RowLabel = "4 AM"
						.Row = 5
						.RowLabel = "5 AM"
						.Row = 6
						.RowLabel = "6 AM"
						.Row = 7
						.RowLabel = "7 AM"
						.Row = 8
						.RowLabel = "8 AM"
						.Row = 9
						.RowLabel = "9 AM"
						.Row = 10
						.RowLabel = "10 AM"
						.Row = 11
						.RowLabel = "11 AM"
						.Row = 12
						.RowLabel = "12 AM"
						.Row = 13
						.RowLabel = "1 PM"
						.Row = 14
						.RowLabel = "2 PM"
						.Row = 15
						.RowLabel = "3 PM"
						.Row = 16
						.RowLabel = "4 PM"
						.Row = 17
						.RowLabel = "5 PM"
						.Row = 18
						.RowLabel = "6 PM"
						.Row = 19
						.RowLabel = "7 PM"
						.Row = 20
						.RowLabel = "8 PM"
						.Row = 21
						.RowLabel = "9 PM"
						.Row = 22
						.RowLabel = "10 PM"
						.Row = 23
						.RowLabel = "11 PM"
						.Row = 24
						.RowLabel = "12 PM"
					Case "w"
						.Row = 1
						.RowLabel = "Sun"
						.Row = 2
						.RowLabel = "Mon"
						.Row = 3
						.RowLabel = "Tue"
						.Row = 4
						.RowLabel = "Wed"
						.Row = 5
						.RowLabel = "Thur"
						.Row = 6
						.RowLabel = "Fri"
						.Row = 7
						.RowLabel = "Sat"
					Case "m"
						.Row = 1
						.RowLabel = "Jan"
						.Row = 2
						.RowLabel = "Feb"
						.Row = 3
						.RowLabel = "Mar"
						.Row = 4
						.RowLabel = "Apr"
						.Row = 5
						.RowLabel = "May"
						.Row = 6
						.RowLabel = "Jun"
						.Row = 7
						.RowLabel = "Jul"
						.Row = 8
						.RowLabel = "Aug"
						.Row = 9
						.RowLabel = "Sep"
						.Row = 10
						.RowLabel = "Oct"
						.Row = 11
						.RowLabel = "Nov"
						.Row = 12
						.RowLabel = "Dec"
					Case Else
						For x = 1 To lGreatestValue
							.Row = x
							.RowLabel = CStr(x) '"Week " & X
							'.RowLabel = chrtArray(X, 2)
						Next x
						'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						.CtlRefresh()
				End Select
			End With
		End If
		
		If optCallVsVoice.Checked = True Then
			CallVsVoice()
		End If
		'
		If chkLablePoints.CheckState = System.Windows.Forms.CheckState.Checked Then LablePoints()
		'
		'
		'Enable Print button
		'
		cmdPrintChart.Enabled = True
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		'
	End Sub
	
	Private Sub cmdPrintChart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrintChart.Click
		On Error GoTo PrintErrHandler
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		dlgCommon.CancelError = True
		dlgCommonPrint.ShowDialog()
		'UPGRADE_ISSUE: Constant vbPRPSLetter was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Printer property Printer.PaperSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.PaperSize = vbPRPSLetter
		'UPGRADE_ISSUE: Constant vbPRORLandscape was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Printer property Printer.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.Orientation = vbPRORLandscape
		'UPGRADE_ISSUE: Printer property Printer.Copies was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.Copies = dlgCommonPrint.PrinterSettings.Copies
		MSChart1.EditCopy()
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		'UPGRADE_ISSUE: Printer method Printer.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.Print(" ")
		'Printer.PaintPicture Clipboard.GetData(), 0, 0
		'UPGRADE_ISSUE: Printer method Printer.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.PaintPicture(My.Computer.Clipboard.GetImage(), 150, 0, 15000, 12000)
		'UPGRADE_ISSUE: Printer method Printer.EndDoc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.EndDoc()
		Exit Sub
		
PrintErrHandler: 
		Select Case Err.Number
			Case 32755
				'MsgBox "Print cancelled."
		End Select
		Exit Sub
	End Sub
	
	Private Sub DTPicker1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTPicker1.Change
		Dim Index As Short = DTPicker1.GetIndex(eventSender)
		'
		If cboDateType.Text = "Day" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MonthView1.value = DTPicker1(Index).value
			MonthView1.DayOfWeek = MSComCtl2.DayConstants.mvwSunday
			DTPicker1(Index).value = MonthView1.value
		End If
		'
		If Index = 0 Then
			cboDateType_SelectedIndexChanged(cboDateType, New System.EventArgs())
		Else
			SetDate(Index)
		End If
		'Disable Print button
		'
		cmdPrintChart.Enabled = False
	End Sub
	
	Private Sub DTPicker2_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTPicker2.Change
		'
		'Disable Print button
		'
		cmdPrintChart.Enabled = False
	End Sub
	
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
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSelect.Form_Initialize.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub frmMultiChart_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim x As Short
		Dim y As Short
		'frmMultiChart.Top = (Screen.Height - frmMultiChart.Height) / 2
		'frmMultiChart.Left = (Screen.Width - frmMultiChart.Width) / 2
		optExt.Checked = True
		optNone.Checked = True
		optBoth.Checked = True
		For x = 0 To 9
			DTPicker1(x).value = Today
			DTPicker2.value = Today
			DTPicker3.value = Today
			DTPicker4.value = Today
		Next x
		cboDateType_SelectedIndexChanged(cboDateType, New System.EventArgs())
		optTotal.Checked = True
		'
		'  Y = 0
		'  For X = 2003 To 2010
		'    cboYear.AddItem X, Y
		'    Y = Y + 1
		'  Next
		'  '
		''  Y = 0
		''  For x = 1 To 53
		''    cboSunday.AddItem x, Y
		''    Y = Y + 1
		''  Next
		'  '
		'  Y = 0
		'  For X = 1 To 12
		'    cboMonth.AddItem X, Y
		'    Y = Y + 1
		'  Next
		'  cboYear.ListIndex = 0
		'  cboMonth.ListIndex = 0
		'  GetSundays
		'  '
		'  bChartData = False 'set flag to false
		'  cboDateType_Click
	End Sub
	
	
	Private Sub MSChart1_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MSChart1.DblClick
		'  If bChartData = True Then
		'    frmBigChart.Show
		'    frmBigChart.MSChart2.ChartData = MSChart1.ChartData
		'    frmBigChart.MSChart2.chartType = MSChart1.chartType
		'    frmBigChart.MSChart2.Title = MSChart1.Title
		'  End If
	End Sub
	
	
	'UPGRADE_WARNING: Event optAvg.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optAvg_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAvg.CheckedChanged
		If eventSender.Checked Then
			DTPicker1(0).Visible = False
			DTPicker2.Visible = False
			DTPicker3.Visible = True
			DTPicker4.Visible = True
			fraMultiLines.Enabled = False
			optNone.Enabled = False
			optMultiGroups.Enabled = False
			optMultiDates.Enabled = False
			optDirection.Enabled = False
			optCallVsVoice.Enabled = False
			cmdAvg.Visible = True
			cmdChart.Visible = False
			fraGroups.Visible = False
			fraMultiDates.Visible = False
			cmdPrintChart.Enabled = False
			optNone.Checked = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optBoth.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optBoth_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optBoth.CheckedChanged
		If eventSender.Checked Then
			If optExt.Checked = True Then
				cboCallDir.Enabled = True
			End If
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optCalls.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optCalls_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCalls.CheckedChanged
		If eventSender.Checked Then
			If optExt.Checked = True Then
				cboCallDir.Enabled = True
			End If
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optCallVsVoice.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optCallVsVoice_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCallVsVoice.CheckedChanged
		If eventSender.Checked Then
			fraGroups.Visible = False
			fraMultiDates.Visible = False
			fraCallType.Enabled = False
			optBoth.Enabled = False
			optCalls.Enabled = False
			optVoiceMail.Enabled = False
			'
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optExt.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optExt_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optExt.CheckedChanged
		If eventSender.Checked Then
			Dim cmdCategories As New ADODB.Command
			Dim rstCategories As New ADODB.Recordset
			Dim x As Short
			'
			cmdCategories.CommandTimeout = 300
			'
			'Clear the combo box
			'
			For x = 0 To 9
				cboGroup(x).Items.Clear()
			Next x
			'
			'Select all Ext#s and put in combo box
			'
			cmdCategories.ActiveConnection = cnMain
			cmdCategories.CommandText = "SELECT P1NO FROM ICC_CDR GROUP BY P1NO ORDER BY P1NO"
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			'Add Ext# to drop down cbo
			'
			Do While Not rstCategories.eof
				If rstCategories.Fields("P1NO").Value <> "" Then
					For x = 0 To 9
						cboGroup(x).Items.Add(rstCategories.Fields("P1NO").Value)
					Next x
				End If
				rstCategories.MoveNext()
			Loop 
			'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cmdCategories = Nothing
			rstCategories.Close()
			'
			'Set the first Ext# as default
			'
			For x = 0 To 9
				cboGroup(x).SelectedIndex = 0
			Next x
			'
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
			'
			'Enable call direction
			'
			If optVoiceMail.Checked = False Then
				cboCallDir.Enabled = True
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optDirection.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDirection_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDirection.CheckedChanged
		If eventSender.Checked Then
			DTPicker1(0).Visible = True
			DTPicker2.Visible = True
			DTPicker3.Visible = False
			DTPicker4.Visible = False
			cmdAvg.Visible = False
			cmdChart.Visible = True
			fraGroups.Visible = False
			fraMultiDates.Visible = False
			'frmMultiChart.Width = 10440
			fraCallType.Enabled = True
			optBoth.Enabled = True
			optCalls.Enabled = True
			optVoiceMail.Enabled = True
			'
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optMultiDates.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMultiDates_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMultiDates.CheckedChanged
		If eventSender.Checked Then
			DTPicker1(0).Visible = True
			DTPicker2.Visible = True
			DTPicker3.Visible = False
			DTPicker4.Visible = False
			cmdAvg.Visible = False
			cmdChart.Visible = True
			Dim x As Short
			For x = 1 To 9
				DTPicker1_Change(DTPicker1.Item(x), New System.EventArgs())
			Next 
			fraGroups.Visible = False
			fraMultiDates.Visible = True
			'frmMultiChart.Width = 12705
			fraCallType.Enabled = True
			optBoth.Enabled = True
			optCalls.Enabled = True
			optVoiceMail.Enabled = True
			'
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optMultiGroups.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMultiGroups_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMultiGroups.CheckedChanged
		If eventSender.Checked Then
			DTPicker1(0).Visible = True
			DTPicker2.Visible = True
			DTPicker3.Visible = False
			DTPicker4.Visible = False
			cmdAvg.Visible = False
			cmdChart.Visible = True
			fraGroups.Visible = True
			fraMultiDates.Visible = False
			'frmMultiChart.Width = 12705
			fraCallType.Enabled = True
			optBoth.Enabled = True
			optCalls.Enabled = True
			optVoiceMail.Enabled = True
			'
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optNone.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optNone_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNone.CheckedChanged
		If eventSender.Checked Then
			fraGroups.Visible = False
			fraMultiDates.Visible = False
			'frmMultiChart.Width = 10440
			fraCallType.Enabled = True
			optBoth.Enabled = True
			optCalls.Enabled = True
			optVoiceMail.Enabled = True
			'
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optTotal.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optTotal_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTotal.CheckedChanged
		If eventSender.Checked Then
			DTPicker1(0).Visible = True
			DTPicker2.Visible = True
			DTPicker3.Visible = False
			DTPicker4.Visible = False
			cmdAvg.Visible = False
			cmdChart.Visible = True
			fraMultiLines.Enabled = True
			optNone.Enabled = True
			optMultiGroups.Enabled = True
			optMultiDates.Enabled = True
			optDirection.Enabled = True
			optCallVsVoice.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optVoiceMail.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optVoiceMail_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optVoiceMail.CheckedChanged
		If eventSender.Checked Then
			'
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
			'
			cboCallDir.Text = "Incoming"
			cboCallDir.Enabled = False
			'
		End If
	End Sub
	
	'Private Sub optMean_Click()
	'  MSChart1.Plot.SeriesCollection(1).StatLine.Flag = VtChStatsMean
	'End Sub
	'
	'Private Sub optMinMax_Click()
	'  MSChart1.Plot.SeriesCollection(1).StatLine.Flag = VtChStatsMinimum Or VtChStatsMaximum
	'End Sub
	'
	'Private Sub optRegression_Click()
	'  MSChart1.Plot.SeriesCollection(1).StatLine.Flag = VtChStatsRegression
	'End Sub
	'
	'Private Sub optStdDev_Click()
	'  MSChart1.Plot.SeriesCollection(1).StatLine.Flag = VtChStatsStddev
	'End Sub
	
	'UPGRADE_WARNING: Event optWorkgroup.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optWorkgroup_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optWorkgroup.CheckedChanged
		If eventSender.Checked Then
			Dim cmdCategories As New ADODB.Command
			Dim rstCategories As New ADODB.Recordset
			Dim x As Short
			'
			'Clear the combo box
			'
			For x = 0 To 9
				cboGroup(x).Items.Clear()
			Next x
			'
			'Select all Workgroup#s and put in combo box
			'
			cmdCategories.ActiveConnection = cnMain
			cmdCategories.CommandText = "SELECT P1WGNO FROM ICC_CDR GROUP BY P1WGNO ORDER BY P1WGNO"
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			'Add Workgroup# to drop down cbo
			'
			Do While Not rstCategories.eof
				If rstCategories.Fields("P1WGNO").Value <> "" Then
					For x = 0 To 9
						cboGroup(x).Items.Add(rstCategories.Fields("P1WGNO").Value)
					Next x
				End If
				rstCategories.MoveNext()
			Loop 
			'UPGRADE_NOTE: Object cmdCategories may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cmdCategories = Nothing
			rstCategories.Close()
			'
			'Set the first Workgroup# as default id
			'
			For x = 0 To 9
				cboGroup(x).SelectedIndex = 0
			Next x
			'
			'Disable Print button
			'
			cmdPrintChart.Enabled = False
			'
			'Disable call direction and set to incoming. There are no outgoing calls from a workgroup only Ext
			'
			cboCallDir.Text = "Incoming"
			cboCallDir.Enabled = False
		End If
	End Sub
	
	Private Sub GetSundays()
		Dim x As Short
		Dim y As Short
		Dim Z As Short
		'
		cboSunday.Items.Clear()
		MonthView1.year = CShort(cboYear.Text)
		MonthView1.month = CShort(cboMonth.Text)
		'
		Select Case cboMonth.Text
			Case CStr(4), CStr(6), CStr(9), CStr(11)
				y = 30
			Case CStr(1), CStr(3), CStr(5), CStr(7), CStr(8), CStr(10), CStr(12)
				y = 31
			Case Else
				If CDbl(cboYear.Text) = 2004 Then
					y = 29
				Else
					y = 28
				End If
		End Select
		'
		Z = 0
		For x = 1 To y
			MonthView1.day = x
			If MonthView1.DayOfWeek = 1 Then
				cboSunday.Items.Insert(Z, CStr(x))
				Z = Z + 1
			End If
		Next 
		'
		cboSunday.SelectedIndex = 0
		MonthView1.day = CShort(cboSunday.Text)
		'
	End Sub
	
	Private Function SetDate(ByRef iControl As Short) As Object
		Dim sTemp As String
		'
		Select Case cboDateType.Text
			Case "Hour"
				'DTPicker1_Change iControl
				'DTPicker2.value = DTPicker1(iControl).value + 1
			Case "Day"
				'DTPicker1_Change iControl
				'DTPicker2.value = DTPicker1(iControl).value + 7
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().year. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sTemp = DTPicker1(iControl).year
				DTPicker1(iControl).value = "1/1/" & sTemp
		End Select
		'
	End Function
	
	Private Sub CallVsVoice()
		'
		Dim sVoiceMail As String
		Dim sCallDir As String
		Dim x As Short
		Dim y As Short
		'UPGRADE_WARNING: Lower bound of array iInterval was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim iInterval(100) As Short
		Dim IAvg As Short
		Dim cmdCategories As New ADODB.Command
		Dim rstCategories As New ADODB.Recordset
		'
		' VoiceMail
		'
		'  If optVoiceMail.Value = True Then
		'    sVoiceMail = "VMSTARTTM"
		'    strDirection = "VoiceMail "
		'  Else
		'    sVoiceMail = "STARTTIME"
		'    '
		'    'Set the sCallDir value
		'    '
		'      Select Case cboCallDir
		'        Case "Incoming"
		'          sCallDir = " (TKDIR = 2) AND "
		'          strDirection = "Incoming "
		'        Case "Outgoing"
		'          sCallDir = " (TKDIR = 4) AND "
		'          strDirection = "Outgoing "
		'        Case "Both"
		'          sCallDir = " "
		'          strDirection = ""
		'      End Select
		'  End If
		'
		'Set the date values
		'
		'    If DTPicker1(0).value > DTPicker2.value Then
		'      MsgBox "Incorrect Date Values!"
		'      Exit Sub
		'    Else
		'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
		'        MsgBox "Incorrect Date Values!"
		'        Exit Sub
		'      Else
		'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1().value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DTStart = DTPicker1(0).value
		'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2.value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DTEnd = DTPicker2.value
		'      End If
		'    End If
		'
		'Set the DATEPART value
		'
		Select Case cboDateType.Text
			Case "Hour"
				strDateType = "hh"
				lGreatestValue = 24
			Case "Day"
				strDateType = "w"
				lGreatestValue = 7
			Case "Week"
				strDateType = "ww"
				lGreatestValue = 53
			Case "Month"
				strDateType = "m"
				lGreatestValue = 12
			Case "Year"
				'strDateType = "yyyy"
				Exit Sub
		End Select
		'
		'Set the Ext or Workgroup type and number
		'
		If optExt.Checked = True Then
			strGroupType = "P1NO"
		Else
			strGroupType = "P1WGNO"
		End If
		'                    C H A R T
		'
		' Dynamic 2-dimensional array to store series
		' The first index (x) is the total number of series
		' x-axis value in the 1st slot (i.e. chrtArray(x,1)
		' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
		'
		
		'UPGRADE_WARNING: Lower bound of array chrtArray was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim chrtArray(lGreatestValue, 2) As Object
		MSChart1.ShowLegend = True
		'MSChart1.chartType = VtChChartType2dLine
		'
		'Chart Title centered on top
		'
		MSChart1.Title.Text = "Calls between " & DTPicker1(0).value & " and " & DTPicker2.value
		'
		'Chart X and Y axis titles
		'
		MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = cboDateType.Text
		MSChart1.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = ""
		'
		'Chart Foot note
		'
		'MSChart1.FootnoteText = "footnote"
		'
		'Load the array with 0s with correct # of rows
		'
		For x = 1 To lGreatestValue
			For y = 1 To 2
				'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(x, y). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chrtArray(x, y) = 0
			Next y
		Next x
		For x = 1 To 2
			If x = 1 Then
				sVoiceMail = "STARTTIME"
				strDirection = ""
				sCallDir = sCallDir & "VMSTARTTM = 0 AND "
			End If
			If x = 2 Then
				sVoiceMail = "VMSTARTTM"
				strDirection = "VoiceMail "
			End If
			'sCallDir = sCallDir + " "
			'
			'Create the SQL statement
			'
			strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & cboGroup(0).Text & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & System.Date.FromOADate(DTEnd.ToOADate + 1) & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
			sCallDir = " "
			'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
			'
			'Select revelant data
			'
			cmdCategories.ActiveConnection = cnMain
			cmdCategories.CommandText = strSQL
			rstCategories.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			rstCategories.Open(cmdCategories,  , ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			If rstCategories.RecordCount = 0 Then
				MsgBox("No revelant data for " & cboGroup(0).Text & "!")
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				'Exit Sub
			End If
			bChartData = True 'set flag to true
			'
			'Load the array with data
			'
			For y = 1 To rstCategories.RecordCount
				If strDateType = "hh" And rstCategories.Fields("datetype").Value = 0 Then
					rstCategories.Fields("datetype").Value = 24
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(rstCategories!datetype, x). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				chrtArray(rstCategories.Fields("datetype").Value, x) = rstCategories.Fields("Calls").Value
				rstCategories.MoveNext()
			Next y
			rstCategories.Close()
		Next x
		'
		'Attach the array of data to MS-CHART
		'
		With MSChart1
			.ChartData = VB6.CopyArray(chrtArray)
			.ColumnCount = 2
			.ColumnLabelCount = 2
			.Column = 1
			.ColumnLabel = "Calls"
			.Column = 2
			.ColumnLabel = "VoiceMails"
			Select Case strDateType
				Case "hh"
					.Row = 1
					.RowLabel = "1 AM"
					.Row = 2
					.RowLabel = "2 AM"
					.Row = 3
					.RowLabel = "3 AM"
					.Row = 4
					.RowLabel = "4 AM"
					.Row = 5
					.RowLabel = "5 AM"
					.Row = 6
					.RowLabel = "6 AM"
					.Row = 7
					.RowLabel = "7 AM"
					.Row = 8
					.RowLabel = "8 AM"
					.Row = 9
					.RowLabel = "9 AM"
					.Row = 10
					.RowLabel = "10 AM"
					.Row = 11
					.RowLabel = "11 AM"
					.Row = 12
					.RowLabel = "12 AM"
					.Row = 13
					.RowLabel = "1 PM"
					.Row = 14
					.RowLabel = "2 PM"
					.Row = 15
					.RowLabel = "3 PM"
					.Row = 16
					.RowLabel = "4 PM"
					.Row = 17
					.RowLabel = "5 PM"
					.Row = 18
					.RowLabel = "6 PM"
					.Row = 19
					.RowLabel = "7 PM"
					.Row = 20
					.RowLabel = "8 PM"
					.Row = 21
					.RowLabel = "9 PM"
					.Row = 22
					.RowLabel = "10 PM"
					.Row = 23
					.RowLabel = "11 PM"
					.Row = 24
					.RowLabel = "12 PM"
				Case "w"
					.Row = 1
					.RowLabel = "Sun"
					.Row = 2
					.RowLabel = "Mon"
					.Row = 3
					.RowLabel = "Tue"
					.Row = 4
					.RowLabel = "Wed"
					.Row = 5
					.RowLabel = "Thur"
					.Row = 6
					.RowLabel = "Fri"
					.Row = 7
					.RowLabel = "Sat"
				Case "m"
					.Row = 1
					.RowLabel = "Jan"
					.Row = 2
					.RowLabel = "Feb"
					.Row = 3
					.RowLabel = "Mar"
					.Row = 4
					.RowLabel = "Apr"
					.Row = 5
					.RowLabel = "May"
					.Row = 6
					.RowLabel = "Jun"
					.Row = 7
					.RowLabel = "Jul"
					.Row = 8
					.RowLabel = "Aug"
					.Row = 9
					.RowLabel = "Sep"
					.Row = 10
					.RowLabel = "Oct"
					.Row = 11
					.RowLabel = "Nov"
					.Row = 12
					.RowLabel = "Dec"
				Case Else
					For y = 1 To lGreatestValue
						.Row = y
						'UPGRADE_WARNING: Couldn't resolve default property of object chrtArray(y, 2). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.RowLabel = chrtArray(y, 2)
					Next y
					'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
					.CtlRefresh()
			End Select
		End With
		'If chkLablePoints.Value = vbChecked Then LablePoints
	End Sub
	
	Private Sub LablePoints()
		Dim y As Short
		With MSChart1
			For y = 1 To .Plot.SeriesCollection.Count
				With .Plot.SeriesCollection(y).DataPoints(-1).DataPointLabel
					.LocationType = MSChart20Lib.VtChLabelLocationType.VtChLabelLocationTypeAbovePoint
					.Component = MSChart20Lib.VtChLabelComponent.VtChLabelComponentValue
					'.PercentFormat = "0%"
					.VtFont.Style = 1
					.VtFont.Size = 10
				End With
			Next y
		End With
	End Sub
End Class