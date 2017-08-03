Option Strict Off
Option Explicit On
Friend Class FHours
	Inherits System.Windows.Forms.Form
	
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	'
	Private rsLog As ADODB.Recordset
	Private rsEmployee As ADODB.Recordset
	Private rsHours As ADODB.Recordset
	
	Private Sub cmbEmployee_InitColumnProps(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbEmployee.InitColumnProps
		On Error GoTo ErrCall
		'
		rsEmployee = New ADODB.Recordset
		'
		If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
			rsEmployee.Open("SELECT EmployeeFirst & ' ' & EmployeeLast AS EmployeeName FROM tblEmployees ORDER BY EmployeeFirst, EmployeeLast", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		Else
			rsEmployee.Open("SELECT EmployeeFirst + ' ' + EmployeeLast AS EmployeeName FROM tblEmployees ORDER BY EmployeeFirst, EmployeeLast", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		End If
		'
		With rsEmployee
			Do Until .eof
				cmbEmployee.AddItem(.Fields("EmployeeName").Value)
				.MoveNext()
			Loop 
		End With
		'
		Exit Sub
		'
ErrCall: 
		MsgBox(Err.Description)
	End Sub
	
	Private Sub cmdActual_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdActual.Click
		'  Me.grdHours.Columns("ActualIn").Visible = Not (Me.grdHours.Columns("ActualIn").Visible)
		'  Me.grdHours.Columns("ActualOut").Visible = Not (Me.grdHours.Columns("ActualOut").Visible)
	End Sub
	
	Private Sub cmdCalcHours_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCalcHours.Click
		Dim X As Short
		'
		GetEmployeeLog()
		'
		If chkShow.CheckState = 1 Then
			For X = 0 To 5
				FlexGridHours.set_ColWidth(X,  , ((VB6.PixelsToTwipsX(FlexGridHours.Width) - 1180) / 6) - 8)
			Next 
			FlexGridHours.set_ColWidth(6,  , 1080)
		Else
			For X = 0 To 3
				FlexGridHours.set_ColWidth(X,  , ((VB6.PixelsToTwipsX(FlexGridHours.Width) - 1180) / 4) - 15)
			Next 
			FlexGridHours.set_ColWidth(4,  , 1080)
		End If
	End Sub
	
	Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
		Dim printDlg As VBPrnDlgLib.PrinterDlg
		printDlg = New VBPrnDlgLib.PrinterDlg
		' Set the starting information for the dialog box based on the current
		' printer settings.
		'UPGRADE_ISSUE: Printer property Printer.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		printDlg.PrinterName = Printer.DeviceName
		'UPGRADE_ISSUE: Printer property Printer.DriverName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		printDlg.DriverName = Printer.DriverName
		'UPGRADE_ISSUE: Printer property Printer.Port was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		printDlg.Port = Printer.Port
		
		' Set the default PaperBin so that a valid value is returned even
		' in the Cancel case.
		'UPGRADE_ISSUE: Printer property Printer.PaperBin was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		printDlg.PaperBin = Printer.PaperBin
		
		' Set the flags for the PrinterDlg object using the same flags as in the
		' common dialog control. The structure starts with VBPrinterConstants.
		printDlg.FLAGS = VBPrnDlgLib.VBPrinterConstants.cdlPDNoSelection Or VBPrnDlgLib.VBPrinterConstants.cdlPDNoPageNums Or VBPrnDlgLib.VBPrinterConstants.cdlPDReturnDC
		'UPGRADE_ISSUE: Printer property Printer.TrackDefault was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.TrackDefault = False
		
		' When CancelError is set to True the ShowPrinterDlg will return error
		' 32755. You can handle the error to know when the Cancel button was
		' clicked. Enable this by uncommenting the lines prefixed with "'**".
		'**printDlg.CancelError = True
		
		' Add error handling for Cancel.
		'**On Error GoTo Cancel
		If Not printDlg.ShowPrinter(Me.Handle.ToInt32) Then
			Debug.Print("Cancel Selected")
			Exit Sub
		End If
		
		'Turn off Error Handling for Cancel.
		'**On Error GoTo 0
		Dim NewPrinterName As String
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim objPrinter As Printer
		Dim strsetting As String
		
		' Locate the printer that the user selected in the Printers collection.
		NewPrinterName = UCase(printDlg.PrinterName)
		'UPGRADE_ISSUE: Printer property Printer.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If Printer.DeviceName <> NewPrinterName Then
			'UPGRADE_ISSUE: Printers object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
			For	Each objPrinter In Printers
				'UPGRADE_ISSUE: Printer property objPrinter.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				If UCase(objPrinter.DeviceName) = NewPrinterName Then
					'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
					Printer = objPrinter
				End If
			Next objPrinter
		End If
		
		' Copy user input from the dialog box to the properties of the selected printer.
		'UPGRADE_ISSUE: Printer property Printer.Copies was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.Copies = printDlg.Copies
		'UPGRADE_ISSUE: Printer property Printer.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.Orientation = printDlg.Orientation
		'UPGRADE_ISSUE: Printer property Printer.ColorMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.ColorMode = printDlg.ColorMode
		'UPGRADE_ISSUE: Printer property Printer.Duplex was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.Duplex = printDlg.Duplex
		'UPGRADE_ISSUE: Printer property Printer.PaperBin was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.PaperBin = printDlg.PaperBin
		'UPGRADE_ISSUE: Printer property Printer.PaperSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.PaperSize = printDlg.PaperSize
		'UPGRADE_ISSUE: Printer property Printer.PrintQuality was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.PrintQuality = printDlg.PrintQuality
		
		' Display the results in the immediate (Debug) window.
		' NOTE: Supported values for PaperBin and Size are printer specific. Some
		' common defaults are defined in the Win32 SDK in MSDN and in Visual Basic.
		' Print quality is the number of dots per inch.
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		With Printer
			'UPGRADE_ISSUE: Printer property Printer.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Debug.Print(.DeviceName)
			'UPGRADE_ISSUE: Printer property Printer.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			If .Orientation = 1 Then
				strsetting = "Portrait. "
			Else
				strsetting = "Landscape. "
			End If
			'UPGRADE_ISSUE: Printer property Printer.Copies was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Debug.Print(VB6.TabLayout("Copies = " & .Copies, "Orientation = " & strsetting))
			'UPGRADE_ISSUE: Printer property Printer.ColorMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			If .ColorMode = 1 Then
				strsetting = "Black and White. "
			Else
				strsetting = "Color. "
			End If
			Debug.Print("ColorMode = " & strsetting)
			'UPGRADE_ISSUE: Printer property Printer.Duplex was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			If .Duplex = 1 Then
				strsetting = "None. "
				'UPGRADE_ISSUE: Printer property Printer.Duplex was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			ElseIf .Duplex = 2 Then 
				strsetting = "Horizontal/Long Edge. "
				'UPGRADE_ISSUE: Printer property Printer.Duplex was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			ElseIf .Duplex = 3 Then 
				strsetting = "Vertical/Short Edge. "
			Else
				strsetting = "Unknown. "
			End If
			Debug.Print("Duplex = " & strsetting)
			'UPGRADE_ISSUE: Printer property Printer.PaperBin was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Debug.Print("PaperBin = " & .PaperBin)
			'UPGRADE_ISSUE: Printer property Printer.PaperSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Debug.Print("PaperSize = " & .PaperSize)
			'UPGRADE_ISSUE: Printer property Printer.PrintQuality was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			Debug.Print("PrintQuality = " & .PrintQuality)
			If (printDlg.FLAGS And VBPrnDlgLib.VBPrinterConstants.cdlPDPrintToFile) = VBPrnDlgLib.VBPrinterConstants.cdlPDPrintToFile Then
				Debug.Print("Print to File Selected")
			Else
				Debug.Print("Print to File Not Selected")
			End If
			Debug.Print("hDC = " & printDlg.hDC)
		End With
		'
		Dim old_width As Short
		'
		old_width = VB6.PixelsToTwipsX(FlexGridHours.Width)
		'UPGRADE_ISSUE: Printer property Printer.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		FlexGridHours.Width = VB6.TwipsToPixelsX(Printer.Width)
		'UPGRADE_ISSUE: Printer method Printer.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.PaintPicture(FlexGridHours.Picture, 0, 0)
		'UPGRADE_ISSUE: Printer method Printer.EndDoc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		Printer.EndDoc()
		FlexGridHours.Width = VB6.TwipsToPixelsX(old_width)
		'
		Exit Sub
		'**Cancel:
		'**If Err.Number = 32755 Then
		'**    Debug.Print "Cancel Selected"
		'**Else
		'**    Debug.Print "A nonCancel Error Occured - "; Err.Number
		'**End If
	End Sub
	
	Private Sub FHours_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'On Error GoTo ErrCall
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.Setup(Me, False,  ,  , "Hours")
		'
		Dim rsEmployee2 As ADODB.Recordset
		'
		rsEmployee2 = New ADODB.Recordset
		'
		rsHours = New ADODB.Recordset
		'
		rsEmployee2.Open("SELECT Password FROM tblEmployees WHERE (EmployeeFirst + ' ' + EmployeeLast = '" & User.Name & "')", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		rsLog = New ADODB.Recordset
		rsLog.Open("SELECT * FROM tblHours WHERE RecID = NULL", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		dcHours.Recordset = rsLog
		dcHours.Password = Rot39(DecryptStr(rsEmployee2.Fields("Password").Value & ""))
		'
		mskBeginDate.value = Today
		mskEndDate.value = Today
		Exit Sub
		'
		'UPGRADE_NOTE: Object rsEmployee2 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsEmployee2 = Nothing
ErrCall: 
		MsgBox(Err.Description)
	End Sub
	
	Private Sub cmdBeginDate_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBeginDate.ClickEvent
		mskBeginDate.value = FDatePick.DateText(mskBeginDate.value)
	End Sub
	
	Private Sub cmdEndDate_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEndDate.ClickEvent
		mskEndDate.value = FDatePick.DateText(mskEndDate.value)
	End Sub
	
	Private Sub GetEmployeeLog()
		'On Error GoTo ErrCall
		'
		Dim TotalHours As Single
		With rsEmployee
			If Not .BOF Then .MoveFirst()
			.Find("EmployeeName = '" & cmbEmployee.Text & "'",  , ADODB.SearchDirectionEnum.adSearchForward)
			'
			If Not .eof Then
				'grdHours.Redraw = False
				'
				If chkShow.CheckState = 1 Then
					If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
						'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						rsHours.Open("SELECT  Employee, LogDate, InTime, ActualIn, OutTime, ActualOut, Hours FROM tblHours WHERE (Employee = '" & cmbEmployee.Text & "') AND (LogDate >= #" & Trim(mskBeginDate.CtlText) & "# AND LogDate <= #" & Trim(mskEndDate.CtlText) & "#)", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
					Else 'SQL Server
						'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						rsHours.Open("SELECT  Employee, LogDate, InTime, ActualIn, OutTime, ActualOut, Hours FROM tblHours WHERE (Employee = '" & cmbEmployee.Text & "') AND (LogDate >= '" & Trim(mskBeginDate.CtlText) & "' AND LogDate <= '" & Trim(mskEndDate.CtlText) & "')", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
					End If
				Else
					If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
						'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						rsHours.Open("SELECT  Employee, LogDate, InTime, OutTime, Hours FROM tblHours WHERE (Employee = '" & cmbEmployee.Text & "') AND (LogDate >= #" & Trim(mskBeginDate.CtlText) & "# AND LogDate <= #" & Trim(mskEndDate.CtlText) & "#)", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
					Else 'SQL Server
						'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						rsHours.Open("SELECT  Employee, LogDate, InTime, OutTime, Hours FROM tblHours WHERE (Employee = '" & cmbEmployee.Text & "') AND (LogDate >= '" & Trim(mskBeginDate.CtlText) & "' AND LogDate <= '" & Trim(mskEndDate.CtlText) & "')", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
					End If
				End If
				'
				'dcHours.Refresh
				'
				If Not (rsHours.eof And rsHours.BOF) Then rsHours.MoveFirst()
				If Not rsHours.RecordCount = 0 Then
					'
					rsHours.MoveFirst()
					Do While Not rsHours.eof
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDbNull(rsHours.Fields("outtime").Value) Or IsDbNull(rsHours.Fields("intime").Value) Then
							MsgBox("An entry does not contain both an in-time and out-time. Please correct this before calculating final hours.")
						Else
							rsHours.Fields("hours").Value = ConvertToHours(CDate(0 & rsHours.Fields("outtime").Value).ToOADate - CDate(0 & rsHours.Fields("intime").Value).ToOADate)
							'dcHours.Recordset!hours = 6 'Format(dcHours.Recordset!hours, "h:mm")
							rsHours.Update()
							TotalHours = TotalHours + rsHours.Fields("hours").Value
						End If
						'dcHours.Recordset!LogDate = DatePart("yyyy", dcHours.Recordset!LogDate) & "/" & DatePart("m", dcHours.Recordset!LogDate) & "/" & DatePart("d", dcHours.Recordset!LogDate)
						'
						rsHours.MoveNext()
					Loop 
					'
					'dcHours.Refresh
					'
					txtHours.Text = CStr(TotalHours)
				End If
				'
				'grdHours.Redraw = True
			Else
				MsgBox("Employee not found.")
			End If
		End With
		'
		rsHours.MoveFirst()
		PopulateFlexGrid(FlexGridHours, rsHours)
		'
		rsHours.Close()
		Exit Sub
		'
		'ErrCall:
		'grdHours.Redraw = True
		'MsgBox Err.Description
	End Sub
	
	'UPGRADE_WARNING: Event FHours.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FHours_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		Frame1.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(Frame1.Width)) / 2), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
	End Sub
	
	'UPGRADE_WARNING: Event txtHours.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtHours_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHours.TextChanged
		On Error GoTo ErrCall
		'
		txtTotalHours.Text = CStr(Val(txtHours.Text) - Val(txtLunch.Text))
		'
		Exit Sub
		'
ErrCall: 
		MsgBox(Err.Description)
	End Sub
	
	'UPGRADE_WARNING: Event txtLunch.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtLunch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLunch.TextChanged
		On Error GoTo ErrCall
		'
		txtTotalHours.Text = CStr(Val(txtHours.Text) - Val(txtLunch.Text))
		'
		Exit Sub
		'
ErrCall: 
		MsgBox(Err.Description)
	End Sub
	
	Private Function ConvertToHours(ByRef HoursIn As Double) As Single
		On Error GoTo ErrCall
		'
		Dim intDays As Short
		Dim sngHours As Single
		
		Dim intHours As Short
		Dim intMinutes As Short
		
		intDays = Int(HoursIn)
		sngHours = HoursIn - intDays
		
		intHours = Hour(System.Date.FromOADate(sngHours))
		intMinutes = Minute(System.Date.FromOADate(sngHours))
		
		ConvertToHours = CSng(VB6.Format((intDays * 24) + intHours + (intMinutes / 60), "fixed"))
		'
		Exit Function
		'
ErrCall: 
		MsgBox(Err.Description)
	End Function
	
	Public Function PopulateFlexGrid(ByRef FlexGrid As Object, ByRef rs As Object) As Boolean
		'*******************************************************
		'PURPOSE: Populate MSFlexGrid with data from an
		'         ADO Recordset
		'PARAMETERS:  FlexGrid: MsFlexGrid to Populate
		'             RS: Open ADO Recordset
		'RETURNS:     True if successful, false otherwise
		'REQUIRES:    -- Reference to Microsoft Active Data Objects
		'             -- Recordset should be open with cursor set at
		'                first row when passed and must
		'                support recordcount property
		'             -- FlexGrid should be empty when passed
		'EXAMPLE:
		'Dim conn As New ADODB.Connection
		'Dim rs As New ADODB.Recordset
		'Dim sConnString As String
		'
		'sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDatabase.mdb"
		'conn.Open sConnString
		'rs.Open " SELECT * FROM MyTable", oConn, adOpenKeyset, adLockOptimistic
		'PopulateFlexGrid MSFlexGrid1, rs
		'
		'rs.Close
		'conn.Close
		'***********************************************************
		On Error GoTo ErrorHandler
		'
		If Not TypeOf FlexGrid Is AxMSHierarchicalFlexGridLib.AxMSHFlexGrid Then Exit Function
		If Not TypeOf rs Is ADODB.Recordset Then Exit Function
		'
		Dim i As Short
		Dim J As Short
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object FlexGrid.FixedRows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FlexGrid.FixedRows = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object FlexGrid.FixedCols. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FlexGrid.FixedCols = 0
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object rs.eof. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not rs.eof Then
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object FlexGrid.Rows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object rs.RecordCount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FlexGrid.Rows = rs.RecordCount + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object FlexGrid.Cols. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object rs.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FlexGrid.Cols = rs.Fields.Count
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object rs.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For i = 0 To rs.Fields.Count - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object FlexGrid.TextMatrix. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object rs.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FlexGrid.TextMatrix(0, i) = rs.Fields(i).Name
			Next 
			
			i = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object rs.eof. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Do While Not rs.eof
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object rs.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For J = 0 To rs.Fields.Count - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object rs.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDbNull(rs.Fields(J).value) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object FlexGrid.TextMatrix. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object rs.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						FlexGrid.TextMatrix(i, J) = rs.Fields(J).value
					End If
				Next 
				'
				i = i + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object rs.MoveNext. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rs.MoveNext()
			Loop 
		End If
		'
		PopulateFlexGrid = True
		'
		Exit Function
ErrorHandler: 
		Exit Function
	End Function
End Class