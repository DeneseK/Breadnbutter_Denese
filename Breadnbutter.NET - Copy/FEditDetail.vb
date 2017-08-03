Option Strict Off
Option Explicit On
Friend Class FEditDetail
	Inherits System.Windows.Forms.Form
	
	Private CboEvents As New CComboSearch
	Private lID As Integer
	Private OpenEdit As Boolean
	Private ClosedDate As Date
	
	Private Sub cboType_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.Enter
		CboEvents.Setup(cboType)
	End Sub
	
	Private Sub cboType_InitColumnProps(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.InitColumnProps
		On Error GoTo EH
		'
		Dim rsType As ADODB.Recordset
		'
		rsType = New ADODB.Recordset
		rsType.Open("SELECT * FROM tblActivities ORDER BY Activity", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		cboType.Redraw = False
		Do While Not rsType.eof
			cboType.AddItem(rsType.Fields("Activity").Value)
			rsType.MoveNext()
		Loop 
		cboType.Redraw = True
		'
		DBOps.ZapRS(rsType)
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FEditDetail: cboType_InitColumnProps.")
	End Sub
	
	'UPGRADE_WARNING: Event chkOpenCall.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkOpenCall_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOpenCall.CheckStateChanged
		'ClosedDate = Format(Now, "m/d/yy h:mm AM/PM")
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEdit.Click
		On Error Resume Next
		cboType.Enabled = True
		cboType.Focus()
		txtSubject.Enabled = True
		txtResults.Enabled = True
		cboCase.Enabled = True
		If OpenEdit = True Then
			chkOpenCall.Enabled = True
		End If
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		On Error GoTo EH
		'
		Dim sSQL As String
		Dim rsCase As New ADODB.Recordset
		Dim rsCaseLink As New ADODB.Recordset
		Dim iCaseNum As Short
		'
		If chkOpenCall.CheckState = False Then
			If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
				sSQL = "UPDATE tblSupportActs SET Type = '" & cboType.Text & "', " & "Results = '" & Replace(txtResults.Text, "'", "''") & "', Subject = '" & Replace(txtSubject.Text, "'", "''") & "', ClosedTime = '" & Replace(VB6.Format(Now, "m/d/yy h:mm AM/PM"), "'", "''") & "', OpenCall = '" & Replace(CStr(chkOpenCall.CheckState), "'", "''") & "' " & "WHERE RecID = " & lID
			Else
				sSQL = "UPDATE TSupportActs SET Type = '" & cboType.Text & "', " & "Results = '" & Replace(txtResults.Text, "'", "''") & "', Subject = '" & Replace(txtSubject.Text, "'", "''") & "', ClosedTime = '" & Replace(VB6.Format(Now, "m/d/yy h:mm AM/PM"), "'", "''") & "', OpenCall = '" & Replace(CStr(chkOpenCall.CheckState), "'", "''") & "' " & "WHERE RecID = " & lID
			End If
		Else
			If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
				sSQL = "UPDATE tblSupportActs SET Type = '" & cboType.Text & "', " & "Results = '" & Replace(txtResults.Text, "'", "''") & "', Subject = '" & Replace(txtSubject.Text, "'", "''") & "', OpenCall = '" & Replace(CStr(chkOpenCall.CheckState), "'", "''") & "' " & "WHERE RecID = " & lID
			Else
				sSQL = "UPDATE TSupportActs SET Type = '" & cboType.Text & "', " & "Results = '" & Replace(txtResults.Text, "'", "''") & "', Subject = '" & Replace(txtSubject.Text, "'", "''") & "', OpenCall = '" & Replace(CStr(chkOpenCall.CheckState), "'", "''") & "' " & "WHERE RecID = " & lID
			End If
		End If
		'
		cnMain.Execute(sSQL)
		'
		If Not cboCase.SelectedIndex = -1 Then
			rsCase.Open("Select [CaseName], [CaseID] from TCase Where [CaseName] = '" & cboCase.Text & "'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			With rsCase
				If Not .eof Then
					iCaseNum = .Fields("CaseID").Value
				End If
			End With
			'
			rsCaseLink.Open("Select * from TCaseSupportActLink where CaseID = '" & iCaseNum & "'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			With rsCaseLink
				If Not .eof Then
					While Not .eof
						If .Fields("SupportActID").Value = lID Then
							cboCase.Visible = False
							Label1.Visible = False
							Me.Close()
							Exit Sub
						End If
						.MoveNext()
					End While
				End If
				.Close()
			End With
			'
			rsCaseLink.Open("Select * from TCaseSupportActLink", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			With rsCaseLink
				.AddNew()
				.Fields("CaseID").Value = iCaseNum
				.Fields("SupportActID").Value = lID
				.UpdateBatch()
			End With
			rsCase.Close()
			rsCaseLink.Close()
		End If
		'
		Me.Close()
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FEditDetail: cmdSave_Click.")
	End Sub
	
	Public Sub ShowRecord(ByRef plID As Integer, Optional ByRef pOpenEdit As Boolean = False)
		'On Error GoTo EH
		'
		Dim rsDetail As ADODB.Recordset
		Dim rsListCases As New ADODB.Recordset
		Dim rsCaseLink As New ADODB.Recordset
		Dim i As Short
		Dim sCaseName As String
		Dim iCaseID As Short
		'
		lID = plID
		OpenEdit = pOpenEdit
		If OpenEdit = True Then
			chkOpenCall.Visible = True
		Else
			chkOpenCall.Visible = False
		End If
		'
		rsDetail = New ADODB.Recordset
		'
		If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
			rsDetail.Open("SELECT * FROM tblSupportActs WHERE RecID = " & lID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		Else
			rsDetail.Open("SELECT * FROM TSupportActs WHERE RecID = " & lID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		End If
		'
		With rsDetail
			If Not .eof Then
				Me.cboType.Text = .Fields("Type").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Me.mskDate.DateValue = nnNum(.Fields("Date"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Me.ttmTime.ValidateMode = nnNum(.Fields("Time"))
				Me.txtSubject.Text = .Fields("Subject").Value & vbNullString
				Me.txtResults.Text = .Fields("Results").Value & vbNullString
				Me.lblUser.Text = .Fields("User").Value & vbNullString
				If .Fields("OpenCall").Value = True Then
					Me.chkOpenCall.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					Me.chkOpenCall.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
			Else
				MsgBox("Record not found.", MsgBoxStyle.Information, "Edit Detail")
			End If
		End With
		'
		'Loads combo box with case names
		If Not bCases Then
			cboCase.Items.Clear()
			cboCase.Visible = True
			Label1.Visible = True
			rsListCases.Open("Select [CaseName] from TCase", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			With rsListCases
				If Not .eof Then
					While Not .eof
						cboCase.Items.Add(.Fields("CaseName").Value)
						.MoveNext()
					End While
				End If
				.Close()
			End With
		End If
		'
		'Checks to see if Message is connected to a case and displays it in combo box
		rsCaseLink.Open("Select * from TCaseSupportActLink where [SupportActID] = '" & lID & "'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		With rsCaseLink
			If Not .eof Then
				iCaseID = .Fields("CaseID").Value
			End If
		End With
		'
		If iCaseID > 0 Then
			rsListCases.Open("Select * from TCase where [CaseID] = '" & iCaseID & "'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
			With rsListCases
				If Not .eof Then
					sCaseName = .Fields("CaseName").Value
				End If
			End With
			'
			For i = 0 To cboCase.Items.Count
				If sCaseName = VB6.GetItemString(cboCase, i) Then
					cboCase.SelectedIndex = i
				End If
			Next i
		End If
		'
		VB6.ShowForm(Me, VB6.FormShowConstants.Modal, FMain)
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FEditDetail: ShowRecord.")
	End Sub
End Class