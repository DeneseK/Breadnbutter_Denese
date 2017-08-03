Option Strict Off
Option Explicit On
Friend Class FHistory
	Inherits System.Windows.Forms.Form
	Private Report As New CReport
	'
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	
	'UPGRADE_WARNING: Event chkLimit.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkLimit_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLimit.CheckStateChanged
		If chkLimit.CheckState = System.Windows.Forms.CheckState.Checked Then
			txtLimit.Enabled = True
		Else
			txtLimit.Enabled = False
		End If
	End Sub
	
	'
	Private Sub cmdDateSet1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDateSet1.Click
		On Error GoTo EH
		Me.lblDate1.Text = FDatePick.DateText(lblDate1.Text)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.cmdDateSet1_Click.")
	End Sub
	'
	Private Sub cmdDateSet2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDateSet2.Click
		On Error GoTo EH
		Me.lblDate2.Text = FDatePick.DateText(lblDate2.Text)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.cmdDateSet2_Click.")
	End Sub
	
	Private Sub cmdCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopy.Click
		My.Computer.Clipboard.Clear()
		'
		Dim Count As Integer
		Dim sText As String
		'
		On Error GoTo EH
		SendData()
		Me.grdHistory.RemoveAll()
		With Report.rsReport
			.MoveLast()
			Do While Not .BOF
				Count = Count + 1
				sText = sText & .Fields("FirstName").Value & " " & .Fields("LastName").Value & ", "
				sText = sText & .Fields("Company").Value & vbCrLf
				sText = sText & .Fields("Date").Value & ", "
				sText = sText & .Fields("Time").Value & ", "
				sText = sText & .Fields("User").Value & vbCrLf
				sText = sText & .Fields("Type").Value & ", "
				'
				If .Fields("Subject").Value & vbNullString <> "" Then
					sText = sText & .Fields("Subject").Value & ", "
				End If
				'
				sText = sText & .Fields("Results").Value & vbCrLf & vbCrLf
				'grdHistory.AddItem !RecID & vbTab & !CustRecID & vbTab & !Date & vbTab _
				''& !Time & vbTab & !Type & vbTab & !User & vbTab & !Subject & vbTab _
				''& "Company: " & !Company & vbTab & "Contact: " & !FirstName & " " & !LastName _
				''& vbTab & !Results
				.MovePrevious()
			Loop 
		End With
		'
		LblCount.Text = CStr(Count)
		'
		My.Computer.Clipboard.SetText(sText)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.cmdShowResults_Click.")
	End Sub
	
	Private Sub Command1_Click()
		'   Dim SupportAct As New CSupportACT
		'   Dim SupportActData As New CSupportActData
		'   '
		'   SupportActData.CustRecID = 21815
		'   SupportActData.ActDate = Date
		'   SupportActData.Subject = "TEST SUBJECT"
		'   SupportActData.Results = "TEST RESULTS"
		'   SupportActData.ActUser = "TEST USER"
		'   SupportActData.ActTime = "3:00"
		'   SupportActData.ActType = "TEST TYPE"
		'   SupportActData.ProductID = 1
		'   SupportActData.ClosedTime = Now
		'   SupportActData.OpenCall = False
		'   SupportAct.Save SupportActData, True
		'
		''  'Password = Awesome
		''  'User = Hurray
		''  '
		''  'Create Security Login
		''  cnMain.Execute "EXEC sp_addlogin 'Hurray', 'awesome', 'BNB_DATA'"
		''  'Give access to DB. Current one I guess.
		''  cnMain.Execute "EXEC sp_grantdbaccess N'Hurray', N'Hurray'"
		''  'Assign "User" role.
		''  cnMain.Execute "EXEC sp_addrolemember N'User', N'Hurray'"
		''  'BONUS: Change Password
		''  cnMain.Execute "EXEC sp_password NULL, 'gnarly', 'Hurray'"
		
	End Sub
	
	'UPGRADE_WARNING: Form event FHistory.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FHistory_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error GoTo EH
		'
		Frame1.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(Frame1.Width)) / 2), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.Form_Activate.")
	End Sub
	
	'UPGRADE_WARNING: Event FHistory.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FHistory_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error GoTo EH
		'
		Frame1.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(Frame1.Width)) / 2), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.Form_Resize.")
	End Sub
	
	Private Sub FHistory_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo EH
		Dim rs As New ADODB.Recordset
		'
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.Setup(Me, False,  ,  , "History Reports")
		'
		lblDate1.Text = CStr(Today)
		lblDate2.Text = CStr(Today)
		'
		rs.Open("SELECT * FROM tblactivities ORDER BY Activity", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		cboCategory.Items.Add("All Categories")
		'
		Do While Not rs.eof
			cboCategory.Items.Add(rs.Fields("Activity").Value)
			rs.MoveNext()
		Loop 
		'
		rs.Close()
		'
		cboCategory.Items.Add("Note")
		cboCategory.Text = "All Categories"
		cboCategory.Refresh()
		'
		lstReport.Items.Add("Detail")
		lstReport.Items.Add("Simple")
		lstReport.Text = "Detail"
		'
		rs.Open("SELECT * FROM tblStatus ORDER BY Status", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		lstStatus.Items.Add("Everyone")
		'
		Do While Not rs.eof
			lstStatus.Items.Add(rs.Fields("Status").Value)
			rs.MoveNext()
		Loop 
		'
		rs.Close()
		'
		lstStatus.Text = "Everyone"
		'
		rs.Open("SELECT * FROM tblEmployees ORDER BY EmployeeLast", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		cboUser.Items.Add("All Users")
		'
		Do While Not rs.eof
			cboUser.Items.Add(rs.Fields("EmployeeFirst").Value & " " & rs.Fields("EmployeeLast").Value)
			rs.MoveNext()
		Loop 
		'
		rs.Close()
		'
		cboUser.Text = "All Users"
		'
		rs.Open("SELECT * FROM TType ORDER BY TypeID", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		cboType.Items.Add("All Types")
		'
		Do While Not rs.eof
			cboType.Items.Add(rs.Fields("Description").Value)
			rs.MoveNext()
		Loop 
		'
		rs.Close()
		'
		cboType.SelectedIndex = 0
		'
		FillProductBox()
		'
		cboState.Items.Add("All")
		cboState.Items.Add("AL")
		cboState.Items.Add("AK")
		cboState.Items.Add("AZ")
		cboState.Items.Add("AR")
		cboState.Items.Add("CA")
		cboState.Items.Add("CO")
		cboState.Items.Add("CT")
		cboState.Items.Add("DE")
		cboState.Items.Add("DC")
		cboState.Items.Add("FL")
		cboState.Items.Add("GA")
		cboState.Items.Add("HI")
		cboState.Items.Add("ID")
		cboState.Items.Add("IL")
		cboState.Items.Add("IN")
		cboState.Items.Add("IA")
		cboState.Items.Add("KS")
		cboState.Items.Add("KY")
		cboState.Items.Add("LA")
		cboState.Items.Add("ME")
		cboState.Items.Add("MD")
		cboState.Items.Add("MA")
		cboState.Items.Add("MI")
		cboState.Items.Add("MN")
		cboState.Items.Add("MS")
		cboState.Items.Add("MO")
		cboState.Items.Add("MT")
		cboState.Items.Add("NE")
		cboState.Items.Add("NV")
		cboState.Items.Add("NH")
		cboState.Items.Add("NJ")
		cboState.Items.Add("NM")
		cboState.Items.Add("NY")
		cboState.Items.Add("NC")
		cboState.Items.Add("ND")
		cboState.Items.Add("OH")
		cboState.Items.Add("OK")
		cboState.Items.Add("OR")
		cboState.Items.Add("PA")
		cboState.Items.Add("PR")
		cboState.Items.Add("RI")
		cboState.Items.Add("SC")
		cboState.Items.Add("SD")
		cboState.Items.Add("TN")
		cboState.Items.Add("TX")
		cboState.Items.Add("UT")
		cboState.Items.Add("VT")
		cboState.Items.Add("WA")
		cboState.Items.Add("WV")
		cboState.Items.Add("WI")
		cboState.Items.Add("WY")
		'
		cboOrder.Items.Add("Date/Time")
		cboOrder.Items.Add("Company/Branch")
		cboOrder.Items.Add("User")
		cboOrder.Items.Add("Category")
		cboOrder.SelectedIndex = 0
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.Form_Load.")
	End Sub
	
	
	Private Sub grdHistory_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdHistory.DblClick
		On Error GoTo EH
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(FResult)
		FResult.TextResult.Text = grdHistory.Columns(9).Value
		FResult.ShowDialog()
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.grdHistory_DblClick.")
	End Sub
	Private Sub SendData()
		On Error GoTo EH
		Report.FirstName = Me.txtFirstName.Text
		Report.LastName = Me.txtLastName.Text
		Report.State = Me.cboState.Text
		Report.Company = Me.txtCompany.Text
		Report.Branch = Me.txtBranch.Text
		Report.ResultsDateMin = CDate(Me.lblDate1.Text)
		Report.ResultsDateMax = CDate(Me.lblDate2.Text)
		Report.Results = Me.txtHistory.Text
		Report.ResultsType = Me.cboCategory.Text
		Report.User = Me.cboUser.Text
		Report.Status = Me.lstStatus.Text
		Report.ContactType = Me.cboType.SelectedIndex
		Report.ProductID = Product.GetProductID((cboProduct.Text))
		Report.SortOrder = Me.cboOrder.Text
		If Me.chkLimit.CheckState = System.Windows.Forms.CheckState.Checked Then
			Report.RecLimit = CShort(Me.txtLimit.Text)
		Else
			Report.RecLimit = 0
		End If
		'
		Report.Rtype = CReport.ReportType.History
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.SendData.")
	End Sub
	
	Private Sub cmdPreviewReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreviewReport.Click
		On Error GoTo EH
		SendData()
		If Me.lstReport.Text = "Detail" Then
			Report.PreviewReport(("History"))
		Else
			Report.PreviewReport(("Simple Contact"))
		End If
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.cmdPreviewReport_Click.")
	End Sub
	
	Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
		On Error GoTo EH
		SendData()
		If Me.lstReport.Text = "Detail" Then
			Report.PrintReport(("History"))
		Else
			Report.PrintReport(("Simple Contact"))
		End If
		Exit Sub
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FHistory.cmdPrint_Click.")
	End Sub
	
	Private Sub cmdShowResults_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowResults.Click
		'
		Dim sCompanyandBranch As String
		Dim Count As Integer
		'On Error GoTo EH
		SendData()
		Me.grdHistory.RemoveAll()
		With Report.rsReport
			Do While Not .eof
				Count = Count + 1
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				If Not IsDbNull(.Fields("Branch").Value) Then
					sCompanyandBranch = .Fields("Company").Value & "   Branch: " & .Fields("Branch").Value
				Else
					sCompanyandBranch = .Fields("Company").Value
				End If
				grdHistory.AddItem(.Fields("RecID").Value & vbTab & .Fields("CustRecID").Value & vbTab & .Fields("Date").Value & vbTab & .Fields("Time").Value & vbTab & .Fields("Type").Value & vbTab & .Fields("User").Value & vbTab & .Fields("Subject").Value & vbTab & "Company: " & sCompanyandBranch & vbTab & "Contact: " & .Fields("FirstName").Value & " " & .Fields("LastName").Value & vbTab & .Fields("Results").Value)
				.MoveNext()
			Loop 
		End With
		'
		LblCount.Text = CStr(Count)
		Exit Sub
		'EH:
		' MsgBox Err.Description & " in FHistory.cmdShowResults_Click."
	End Sub
	
	Private Sub FillProductBox()
		Dim Products As New CProducts
		Dim i As Short
		'
		cboProduct.Items.Clear()
		'
		Product.LoadCollection(Products)
		'
		cboProduct.Items.Add("All")
		'
		For i = 1 To Products.Count
			cboProduct.Items.Add(Products.Item(i).Product)
		Next i
		'
		cboProduct.SelectedIndex = 0
		'
		'UPGRADE_NOTE: Object Products may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Products = Nothing
	End Sub
	
	Private Sub txtHistory_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtHistory.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 13 Then cmdShowResults_Click(cmdShowResults, New System.EventArgs())
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtLimit_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLimit.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii <> 8 And KeyAscii <> 127 Then
			If KeyAscii < 48 Or KeyAscii > 57 Then
				KeyAscii = 0
			End If
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class