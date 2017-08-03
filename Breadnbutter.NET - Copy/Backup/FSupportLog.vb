Option Strict Off
Option Explicit On
Friend Class FSupportLog
	Inherits System.Windows.Forms.Form
	
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	
	Private rsSupportAct As ADODB.Recordset
	Private lSupportActRecID As Integer
	
	'UPGRADE_WARNING: Event cboShow.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboShow_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboShow.SelectedIndexChanged
		SetRecordset()
	End Sub
	
	Private Sub cmdFirst_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFirst.Click
		On Error GoTo EH
		'
		rsSupportAct.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object rsSupportAct.Bookmark. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.grdSupportLog.Bookmark = rsSupportAct.Bookmark
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in Command First Click.")
	End Sub
	
	Private Sub cmdLast_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLast.Click
		On Error GoTo EH
		'
		rsSupportAct.MoveLast()
		'UPGRADE_WARNING: Couldn't resolve default property of object rsSupportAct.Bookmark. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Me.grdSupportLog.Bookmark = rsSupportAct.Bookmark
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in Command Last Click.")
	End Sub
	
	'UPGRADE_NOTE: Form_Initialize was upgraded to Form_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Form_Initialize_Renamed()
		On Error GoTo ErrCall
		'
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.MinHeight = 1965
		FormControl.MinWidth = VB6.PixelsToTwipsX(Me.Width)
		FormControl.DataForm = True
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSupportLog.Form_Initialize.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub FSupportLog_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo EH
		'
		rsSupportAct = New ADODB.Recordset
		'
		cboShow.SelectedIndex = 0
		SetRecordset()
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FSupportLog:Load")
	End Sub
	
	'UPGRADE_WARNING: Event FSupportLog.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FSupportLog_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		grdSupportLog.Redraw = False
		grdSupportLog.SetBounds(0, 0, Me.ClientRectangle.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 435))
		lblShow.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdSupportLog.Height) + 60)
		cboShow.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(grdSupportLog.Height) + 45)
		cmdLast.Top = cboShow.Top
		cmdFirst.Top = cboShow.Top
		'
		grdSupportLog.Columns(6).Width = VB6.PixelsToTwipsX(grdSupportLog.Width) - grdSupportLog.Columns(6).Left - 255
		grdSupportLog.Redraw = True
	End Sub
	
	Private Sub SetRecordset()
		On Error GoTo EH
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		System.Windows.Forms.Application.DoEvents()
		'
		If rsSupportAct.State = ADODB.ObjectStateEnum.adStateOpen Then
			rsSupportAct.Close()
		End If
		'
		Dim sSQL As String
		'
		Dim cmdRecords As New ADODB.Command
		Dim dtOldestDate As Date
		If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
			Select Case cboShow.SelectedIndex
				Case 0 'Today
					sSQL = "SELECT * FROM QSupportActs WHERE Date = #" & VB6.Format(Now, "Short Date") & "# ORDER BY Date DESC, Time DESC"
				Case 1 'Previous 7 Days
					sSQL = "SELECT * FROM QSupportActs where Date > #" & VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -7, Now), "Short Date") & "# ORDER BY Date DESC, Time DESC"
				Case Else 'All
					sSQL = "SELECT * FROM QSupportActs ORDER BY Date DESC, Time DESC"
			End Select
			'
			rsSupportAct.Open(sSQL, cnMain, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		Else
			'
			With cmdRecords
				.ActiveConnection = cnMain
				.CommandText = "dbo.UpParmSelSupportActsByDateRange"
				.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
				'
				'
				Select Case cboShow.SelectedIndex
					Case 0 'Today
						dtOldestDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, Now)
					Case 1 'Previous 7 Days
						dtOldestDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -8, Now)
					Case Else 'All
						dtOldestDate = CDate("01/01/1980")
				End Select
				'
				.Parameters.Append(.CreateParameter("Date", ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput,  , dtOldestDate))
				rsSupportAct = .Execute
			End With
			'
			'UPGRADE_NOTE: Object cmdRecords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			cmdRecords = Nothing
		End If
		'
		Me.grdSupportLog.ReBind()
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		System.Windows.Forms.Application.DoEvents()
		'
		Exit Sub
EH: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		System.Windows.Forms.Application.DoEvents()
		MsgBox(Err.Description & " in FSupportLog: Set Recordset.")
	End Sub
	
	Private Sub grdSupportLog_UnboundReadData(ByVal eventSender As System.Object, ByVal eventArgs As AxSSDataWidgets_B_OLEDB._DSSDBGridEvents_UnboundReadDataEvent) Handles grdSupportLog.UnboundReadData
		On Error GoTo EH
		'
		Dim iRBRow As Short
		Dim iGridRows As Short
		Dim iFld As Short
		'
		iGridRows = 0
		'
		'This code initializes the procedure by declaring the variables
		'that will be used to move data into the Grid.
		'iRBRow will be used to count the number of rows of data
		'requested by the ssRowBuffer object (RowBuf) that is passed to the event.
		'iGridRows will count how many rows of data should be supplied to the Grid
		'from the data source.
		'iFld will be used as a generic counter when pulling data from the recordset.
		'Setting iGridRows to 0 indicates that, at the start of the event,
		'no rows have been read from the recordset.
		'
		With rsSupportAct
			If Not (.BOF And .eof) Then
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If IsDbNull(eventArgs.StartLocation) Then 'If the Grid is empty
					If eventArgs.ReadPriorRows Then 'If ReadPriorRows is True
						.MoveLast() 'then the Grid is being
					Else 'scrolled up towards the top
						.MoveFirst()
					End If
				Else 'If Grid contains data
					'UPGRADE_WARNING: Couldn't resolve default property of object StartLocation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Bookmark = eventArgs.StartLocation
					If eventArgs.ReadPriorRows Then
						.MovePrevious()
					Else
						.MoveNext()
					End If
				End If
				'
				For iRBRow = 0 To eventArgs.RowBuf.RowCount - 1
					If .BOF Or .eof Then Exit For
					'
					Select Case eventArgs.RowBuf.ReadType
						Case SSDataWidgets_B.Constants_ReadType.ssReadTypeAllData 'All data must be read
							'UPGRADE_WARNING: Couldn't resolve default property of object rsSupportAct.Bookmark. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.Bookmark(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.Bookmark(iRBRow) = .Bookmark
							'
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.value(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.value(iRBRow, 0) = .Fields("RecID")
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.value(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.value(iRBRow, 1) = .Fields("CustRecID")
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.value(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.value(iRBRow, 2) = .Fields("CompanyContact")
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.value(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.value(iRBRow, 3) = .Fields("Date")
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.value(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.value(iRBRow, 4) = .Fields("Type")
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.value(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.value(iRBRow, 5) = .Fields("User")
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.value(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.value(iRBRow, 6) = .Fields("Results")
						Case SSDataWidgets_B.Constants_ReadType.ssReadTypeBookmarkOnly 'Only bookmarks must be read
							'UPGRADE_WARNING: Couldn't resolve default property of object rsSupportAct.Bookmark. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object RowBuf.Bookmark(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							eventArgs.RowBuf.Bookmark(iRBRow) = .Bookmark
					End Select 'Cases 2 and 3 are not used by DBGrid
					'
					If eventArgs.ReadPriorRows Then
						.MovePrevious()
					Else
						.MoveNext()
					End If
					'
					iGridRows = iGridRows + 1
				Next iRBRow
				'
				eventArgs.RowBuf.RowCount = iGridRows
			End If
		End With 'rsSupportAct
		'
		Exit Sub
EH: 
		MsgBox(Err.Description)
	End Sub
End Class