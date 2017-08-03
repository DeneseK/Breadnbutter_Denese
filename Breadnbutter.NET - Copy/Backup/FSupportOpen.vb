Option Strict Off
Option Explicit On
Friend Class FSupportOpen
	Inherits System.Windows.Forms.Form
	Private Report As New CReport
	'
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	Private sQueryString As String
	Private sText As String
	
	'UPGRADE_WARNING: Event chkDate.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkDate_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDate.CheckStateChanged
		If chkDate.CheckState = 1 Then
			Date_Renamed.Enabled = True
			DTPicker1.Enabled = True
			DTPicker2.Enabled = True
		Else
			Date_Renamed.Enabled = False
			DTPicker1.Enabled = False
			DTPicker2.Enabled = False
		End If
	End Sub
	
	Private Sub cmdCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopy.Click
		My.Computer.Clipboard.Clear()
		'
		On Error GoTo EH
		'
		My.Computer.Clipboard.SetText(sText)
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FSupportOpen.cmdShowResults_Click.")
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		CloseAllCalls()
	End Sub
	
	'UPGRADE_WARNING: Form event FSupportOpen.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FSupportOpen_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error GoTo EH
		'
		Frame1.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(Frame1.Width)) / 2), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FSupportOpen.Form_Activate.")
	End Sub
	
	'UPGRADE_WARNING: Event FSupportOpen.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FSupportOpen_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error GoTo EH
		'
		Frame1.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(Frame1.Width)) / 2), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FSupportOpen.Form_Resize.")
	End Sub
	
	Private Sub FSupportOpen_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo EH
		Dim rs As New ADODB.Recordset
		'
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.Setup(Me, False,  ,  , "Open Support Calls")
		'
		rs.Open("SELECT * FROM tblactivities ORDER BY Activity", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		Do While Not rs.eof
			lstCategory.Items.Add(rs.Fields("Activity").Value)
			rs.MoveNext()
		Loop 
		'
		rs.Close()
		'
		rs.Open("SELECT * FROM tblEmployees ORDER BY EmployeeLast", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		Do While Not rs.eof
			lstUsers.Items.Add(rs.Fields("EmployeeFirst").Value & " " & rs.Fields("EmployeeLast").Value)
			rs.MoveNext()
		Loop 
		'
		rs.Close()
		'
		lstGroup.Items.Add("Management")
		lstGroup.Items.Add("Sales")
		lstGroup.Items.Add("Support")
		lstGroup.Items.Add("Development")
		'
		optUser.Checked = True
		chkDate.CheckState = System.Windows.Forms.CheckState.Unchecked
		Date_Renamed.Enabled = False
		DTPicker1.Enabled = False
		DTPicker2.Enabled = False
		DTPicker1.Value = System.Date.FromOADate(Now.ToOADate - 1)
		DTPicker2.Value = Now
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FSupportOpen.Form_Load.")
	End Sub
	
	
	Private Sub grdHistory_BtnClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdHistory.BtnClick
		If MsgBox("Are you sure you want to close this call?", MsgBoxStyle.YesNo, "Close Call") = MsgBoxResult.Yes Then
			CloseCall((grdHistory.Columns(0).Value))
			LoadHistory()
		End If
	End Sub
	
	Private Sub grdHistory_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdHistory.DblClick
		On Error GoTo EH
		'grdHistory.Redraw = False
		'FEditDetail.ShowRecord grdHistory.Columns(0).Value, True
		'grdHistory.Redraw = True
		'LoadHistory
		'
		'Load FResult
		'FResult.TextResult.Text = grdHistory.Columns(9).Value
		'FResult.Show vbModal
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FSupportOpen.grdHistory_DblClick.")
	End Sub
	
	Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
		Dim RHistory As Object
		'On Error GoTo EH
		'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.Company. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RHistory.Company.Text = "Open Calls"
		'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RHistory.adc.Connection = cnMain.ConnectionString
		'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.adc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RHistory.adc.Source = sQueryString
		'UPGRADE_WARNING: Couldn't resolve default property of object RHistory.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RHistory.Show()
		'EH:
		' MsgBox Err.Description & " in FSupportOpen.cmdPrint_Click."
	End Sub
	
	Private Sub cmdShowResults_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowResults.Click
		LoadHistory()
		'  '
		'  Dim Count As Long
		'  On Error GoTo EH
		'  SendData
		'  Me.grdHistory.RemoveAll
		'  With Report.rsReport
		'      Do While Not .eof
		'         Count = Count + 1
		'         grdHistory.AddItem !RecID & vbTab & !CustRecID & vbTab & !Date & vbTab _
		''         & !Time & vbTab & !Type & vbTab & !User & vbTab & !Subject & vbTab _
		''         & "Company: " & !Company & vbTab & "Contact: " & !FirstName & " " & !LastName _
		''         & vbTab & !Results
		'         .MoveNext
		'      Loop
		'  End With
		'  '
		'  LblCount.Caption = Count
		'  Exit Sub
		'EH:
		' MsgBox Err.Description & " in FSupportOpen.cmdShowResults_Click."
	End Sub
	
	Private Sub CloseAllCalls()
		'
		Dim iWorkGroup As Short
		Dim sCategory As String
		Dim sUsers As String
		Dim sDate As String
		Dim iIndexNum As Short
		Dim iConjunction As Short
		Dim Employee As New CEmployee
		'
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		'
		sText = ""
		'
		grdHistory.Redraw = False
		'
		Me.grdHistory.RemoveAll()
		'
		'If lCustomerID <> -1 Then
		Dim rsHistory As ADODB.Recordset
		'
		rsHistory = New ADODB.Recordset
		'
		iConjunction = 0
		For iIndexNum = 0 To lstCategory.Items.Count - 1
			If lstCategory.GetItemChecked(iIndexNum) = True Then
				If iConjunction = 0 Then
					sCategory = "AND ("
					iConjunction = 1
				Else
					sCategory = sCategory & "OR"
				End If
				sCategory = sCategory & " (TSupportActs.Type = '" & VB6.GetItemString(lstCategory, iIndexNum) & "') "
			End If
		Next 
		If Len(sCategory) > 0 Then sCategory = sCategory & ")"
		'
		If optUser.Checked = True Then
			iConjunction = 0
			For iIndexNum = 0 To lstUsers.Items.Count - 1
				If lstUsers.GetItemChecked(iIndexNum) = True Then
					If iConjunction = 0 Then
						sUsers = "AND ("
						iConjunction = 1
					Else
						sUsers = sUsers & "OR"
					End If
					sUsers = sUsers & " (TSupportActs.[User] = '" & VB6.GetItemString(lstUsers, iIndexNum) & "') "
				End If
			Next 
			If Len(sUsers) > 0 Then sUsers = sUsers & ")"
		Else '##################################################################################
			iConjunction = 0
			For iWorkGroup = 0 To lstGroup.Items.Count - 1
				If lstGroup.GetItemChecked(iWorkGroup) = True Then
					For iIndexNum = 0 To lstUsers.Items.Count - 1
						If Employee.InGroup(VB6.GetItemString(lstUsers, iIndexNum), VB6.GetItemString(lstGroup, iWorkGroup)) = True Then
							If iConjunction = 0 Then
								sUsers = "AND ("
								iConjunction = 1
							Else
								sUsers = sUsers & "OR"
							End If
							sUsers = sUsers & " (TSupportActs.[User] = '" & VB6.GetItemString(lstUsers, iIndexNum) & "') "
						End If
					Next 
				End If
			Next 
			If Len(sUsers) > 0 Then sUsers = sUsers & ")"
		End If '###############################################################################
		'
		If chkDate.CheckState = 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2._Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1._Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If DTPicker1._Value < DTPicker2._Value Then
				sDate = "AND (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) "
			Else
				MsgBox("Incorrect Date Values", MsgBoxStyle.Exclamation, "Date Error")
				LblCount.Text = CStr(0)
				Exit Sub
			End If
		Else
			sDate = ""
		End If
		'
		sQueryString = "SELECT TSupportActs.*, TContact.FirstName, TContact.LastName, TCompany.Name AS CompanyName FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID RIGHT OUTER JOIN TSupportActs ON TContact.ID = TSupportActs.CustRecID Where (TSupportActs.OpenCall = 1) " & sCategory & sUsers & sDate & "ORDER BY TSupportActs.[Date] DESC, TSupportActs.[Time] DESC"
		
		rsHistory.Open(sQueryString, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		With rsHistory
			If Not (.eof And .BOF) Then
				Do While Not .eof
					CloseCall(.Fields("RecID").Value)
					.MoveNext()
				Loop 
			Else
				MsgBox("No Open Calls Found!", MsgBoxStyle.Exclamation, "Close Open Calls")
				LblCount.Text = CStr(0)
			End If
		End With 'rsHistory
		'
		rsHistory.Close()
		'UPGRADE_NOTE: Object rsHistory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsHistory = Nothing
		'UPGRADE_NOTE: Object Employee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employee = Nothing
		'  End If
		'
		grdHistory.Redraw = True
		'
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Form property FSupportOpen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		Me.Cursor = vbNormal
	End Sub
	
	Private Sub LoadHistory()
		'  On Error GoTo EH
		'
		Dim iWorkGroup As Short
		Dim sCategory As String
		Dim sUsers As String
		Dim sDate As String
		Dim iIndexNum As Short
		Dim iConjunction As Short
		Dim Employee As New CEmployee
		'
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		'
		sText = ""
		'
		grdHistory.Redraw = False
		'
		Me.grdHistory.RemoveAll()
		'
		'If lCustomerID <> -1 Then
		Dim rsHistory As ADODB.Recordset
		'
		rsHistory = New ADODB.Recordset
		'
		iConjunction = 0
		For iIndexNum = 0 To lstCategory.Items.Count - 1
			If lstCategory.GetItemChecked(iIndexNum) = True Then
				If iConjunction = 0 Then
					sCategory = "AND ("
					iConjunction = 1
				Else
					sCategory = sCategory & "OR"
				End If
				sCategory = sCategory & " (TSupportActs.Type = '" & VB6.GetItemString(lstCategory, iIndexNum) & "') "
			End If
		Next 
		If Len(sCategory) > 0 Then sCategory = sCategory & ")"
		'
		If optUser.Checked = True Then
			iConjunction = 0
			For iIndexNum = 0 To lstUsers.Items.Count - 1
				If lstUsers.GetItemChecked(iIndexNum) = True Then
					If iConjunction = 0 Then
						sUsers = "AND ("
						iConjunction = 1
					Else
						sUsers = sUsers & "OR"
					End If
					sUsers = sUsers & " (TSupportActs.[User] = '" & VB6.GetItemString(lstUsers, iIndexNum) & "') "
				End If
			Next 
			If Len(sUsers) > 0 Then sUsers = sUsers & ")"
		Else '##################################################################################
			iConjunction = 0
			For iWorkGroup = 0 To lstGroup.Items.Count - 1
				If lstGroup.GetItemChecked(iWorkGroup) = True Then
					For iIndexNum = 0 To lstUsers.Items.Count - 1
						If Employee.InGroup(VB6.GetItemString(lstUsers, iIndexNum), VB6.GetItemString(lstGroup, iWorkGroup)) = True Then
							If iConjunction = 0 Then
								sUsers = "AND ("
								iConjunction = 1
							Else
								sUsers = sUsers & "OR"
							End If
							sUsers = sUsers & " (TSupportActs.[User] = '" & VB6.GetItemString(lstUsers, iIndexNum) & "') "
						End If
					Next 
				End If
			Next 
			If Len(sUsers) > 0 Then sUsers = sUsers & ")"
		End If '###############################################################################
		'
		If chkDate.CheckState = 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker2._Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object DTPicker1._Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If DTPicker1._Value < DTPicker2._Value Then
				sDate = "AND (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1._Value & "', 102) AND CONVERT(DATETIME, '" & DTPicker2._Value & "', 102)) "
			Else
				MsgBox("Incorrect Date Values", MsgBoxStyle.Exclamation, "Date Error")
				LblCount.Text = CStr(0)
				Exit Sub
			End If
		Else
			sDate = ""
		End If
		'
		sQueryString = "SELECT TSupportActs.*, TContact.FirstName, TContact.LastName, TCompany.Name AS CompanyName FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID RIGHT OUTER JOIN TSupportActs ON TContact.ID = TSupportActs.CustRecID Where (TSupportActs.OpenCall = 1) " & sCategory & sUsers & sDate & "ORDER BY TSupportActs.[Date] DESC, TSupportActs.[Time] DESC"
		
		rsHistory.Open(sQueryString, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		With rsHistory
			If Not (.eof And .BOF) Then
				LblCount.Text = CStr(.RecordCount)
				Do While Not .eof
					'###################################################################################
					sText = sText & .Fields("FirstName").Value & " " & .Fields("LastName").Value & ", "
					sText = sText & .Fields("CompanyName").Value & vbCrLf
					sText = sText & .Fields("Date").Value & ", "
					sText = sText & .Fields("Time").Value & ", "
					sText = sText & .Fields("User").Value & vbCrLf
					sText = sText & .Fields("Type").Value & ", "
					
					If .Fields("Subject").Value & vbNullString <> "" Then
						sText = sText & .Fields("Subject").Value & ", "
					End If
					'
					sText = sText & .Fields("Results").Value & vbCrLf & vbCrLf
					'###################################################################################
					grdHistory.AddItem(.Fields("RecID").Value & vbTab & .Fields("CustRecID").Value & vbTab & .Fields("Date").Value & vbTab & .Fields("Time").Value & vbTab & .Fields("Type").Value & vbTab & .Fields("User").Value & vbTab & .Fields("Subject").Value & vbTab & "Company: " & .Fields("CompanyName").Value & vbTab & "Contact: " & .Fields("FirstName").Value & " " & .Fields("LastName").Value & vbTab & .Fields("Results").Value & vbTab & "Close")
					'
					.MoveNext()
				Loop 
			Else
				MsgBox("No Open Calls Found!", MsgBoxStyle.Exclamation, "Open Calls")
				LblCount.Text = CStr(0)
			End If
		End With 'rsHistory
		'
		rsHistory.Close()
		'UPGRADE_NOTE: Object rsHistory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsHistory = Nothing
		'UPGRADE_NOTE: Object Employee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employee = Nothing
		'  End If
		'
		grdHistory.Redraw = True
		'
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Form property FSupportOpen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		Me.Cursor = vbNormal
		'
		Exit Sub
		'EH:
		'  grdHistory.Redraw = True
		'  MsgBox Err.Description
	End Sub
	
	Private Sub CloseCall(ByRef lID As Integer)
		Dim sSQL As String
		'
		sSQL = "UPDATE TSupportActs SET " & "ClosedTime = '" & Replace(VB6.Format(Now, "m/d/yy h:mm AM/PM"), "'", "''") & "', OpenCall = '0' " & "WHERE RecID = " & lID
		'
		cnMain.Execute(sSQL)
		'
	End Sub
	
	'UPGRADE_WARNING: Event optGroup.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optGroup_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGroup.CheckedChanged
		If eventSender.Checked Then
			If optGroup.Checked = True Then
				User.Enabled = False
				lstUsers.Enabled = False
				Group.Enabled = True
				lstGroup.Enabled = True
			Else
				User.Enabled = True
				lstUsers.Enabled = True
				Group.Enabled = False
				lstGroup.Enabled = False
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optUser.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optUser_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUser.CheckedChanged
		If eventSender.Checked Then
			If optUser.Checked = True Then
				User.Enabled = True
				lstUsers.Enabled = True
				Group.Enabled = False
				lstGroup.Enabled = False
			Else
				User.Enabled = False
				lstUsers.Enabled = False
				Group.Enabled = True
				lstGroup.Enabled = True
			End If
		End If
	End Sub
	
	'Private Function InGroup(sName As String, sWorkGroup As String) As Boolean
	'  InGroup = False
	'  Dim rsList As New ADODB.Recordset
	'  Dim EmployeeData As CEmployeeData
	'  Dim iWorkGroupNum As Integer
	'  '
	'  rsList.Open "select * from tblEmployees", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
	'  With rsList
	'    Do While Not .eof
	'      If LCase(sName) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
	'        iWorkGroupNum = nnNum(!WorkGroups)
	'      End If
	'      .MoveNext
	'    Loop
	'  End With
	'    Select Case sWorkGroup
	'      Case "Management"
	'        If iWorkGroupNum > 7 Then InGroup = True
	'      Case "Sales"
	'        Select Case iWorkGroupNum
	'          Case 4, 5, 6, 7, 12, 13, 14, 15
	'            InGroup = True
	'        End Select
	'      Case "Support"
	'        Select Case iWorkGroupNum
	'          Case 2, 3, 6, 7, 10, 11, 14, 15
	'            InGroup = True
	'        End Select
	'      Case "Development"
	'        Select Case iWorkGroupNum
	'          Case 1, 3, 5, 7, 9, 11, 13, 15
	'            InGroup = True
	'        End Select
	'    End Select
	'    '
	'    rsList.Close
	'  Set rsList = Nothing
	'  Set EmployeeData = Nothing
	'  '
	'End Function
End Class