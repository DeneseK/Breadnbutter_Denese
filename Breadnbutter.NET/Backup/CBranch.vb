Option Strict Off
Option Explicit On
Friend Class CBranch
	
	
	Public Sub LoadCollection(ByVal plCompanyID As Integer, ByRef pBranchs As CBranchs)
		
		Dim rslist As New ADODB.Recordset
		Dim BranchData As CBranchData
		'
		pBranchs = New CBranchs
		'
		rslist.Open("SELECT * FROM TBranch WHERE CompanyID = " & plCompanyID & " ORDER BY Name", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		
		'Set rsList = dbPropertyValuation.OpenRecordset("SELECT * FROM TAttachedGarage" & _
		'" ORDER BY SquareFoot", dbOpenForwardOnly)
		'
		While Not rslist.eof
			With rslist
				BranchData = New CBranchData
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				BranchData.BranchID = nnNum(.Fields("BranchID"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				BranchData.CompanyID = nnNum(.Fields("CompanyID"))
				BranchData.Name = .Fields("Name").Value & vbNullString
				'        BranchData.Address1 = !Address1 & vbNullString
				'        BranchData.Address2 = !Address2 & vbNullString
				'        BranchData.Address3 = !Address3 & vbNullString
				'        BranchData.City = !City & vbNullString
				'        BranchData.Email = !Email & vbNullString
				'        BranchData.FaxNumber = !FaxNumber & vbNullString
				'        BranchData.ManagerFirstName = !ManagerFirstName & vbNullString
				'        BranchData.ManagerLastName = !ManagerLastName & vbNullString
				'        BranchData.Number = !Number & vbNullString
				'        BranchData.PhoneNumber = !PhoneNumber & vbNullString
				'        BranchData.State = !State & vbNullString
				'        BranchData.Zip = !Zip & vbNullString
				'
				pBranchs.Add(BranchData)
				'
				rslist.MoveNext()
			End With
		End While
		'
		rslist.Close()
		'
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object BranchData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		BranchData = Nothing
		'
	End Sub
	
	Public Function Load(ByRef pBranchData As CBranchData, ByRef plBranchID As Integer) As Object
		Dim rs As New ADODB.Recordset
		'
		pBranchData = New CBranchData
		'
		rs.Open("SELECT * FROM TBranch WHERE BranchID = " & plBranchID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		If Not rs.eof Then
			With rs
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pBranchData.BranchID = nnNum(.Fields("BranchID"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pBranchData.CompanyID = nnNum(.Fields("CompanyID"))
				pBranchData.Name = .Fields("Name").Value & vbNullString
				pBranchData.Address1 = .Fields("Address1").Value & vbNullString
				pBranchData.Address2 = .Fields("Address2").Value & vbNullString
				pBranchData.Address3 = .Fields("Address3").Value & vbNullString
				pBranchData.City = .Fields("City").Value & vbNullString
				pBranchData.Email = .Fields("Email").Value & vbNullString
				pBranchData.FaxNumber = .Fields("FaxNumber").Value & vbNullString
				pBranchData.ManagerFirstName = .Fields("ManagerFirstName").Value & vbNullString
				pBranchData.ManagerLastName = .Fields("ManagerLastName").Value & vbNullString
				pBranchData.Number = .Fields("Number").Value & vbNullString
				pBranchData.PhoneNumber = .Fields("PhoneNumber").Value & vbNullString
				pBranchData.State = .Fields("State").Value & vbNullString
				pBranchData.Zip = .Fields("Zip").Value & vbNullString
			End With
		End If
		'
		rs.Close()
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
		'
	End Function
	
	Public Function Save(ByRef pBranchData As CBranchData, ByRef pbNew As Boolean) As Boolean
		On Error GoTo EH
		'
		Dim rsBranch As New ADODB.Recordset
		'
		If pbNew Then
			pBranchData.BranchID = NextID("BranchID", "TBranch", cnMain)
			rsBranch.Open("SELECT * FROM TBranch WHERE BranchID = -1", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			rsBranch.AddNew()
			rsBranch.Fields("BranchID").Value = pBranchData.BranchID
			rsBranch.Fields("DateEntered").Value = CDate(Now)
		Else
			rsBranch.Open("SELECT * FROM TBranch WHERE BranchID = " & pBranchData.BranchID, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			'
			If rsBranch.eof Then
				Save = False
				Exit Function
			End If
		End If
		'
		If Not (rsBranch.BOF And rsBranch.eof) Then
			With rsBranch
				'
				.Fields("CompanyID").Value = pBranchData.CompanyID
				.Fields("Name").Value = pBranchData.Name
				.Fields("Address1").Value = pBranchData.Address1
				.Fields("Address2").Value = pBranchData.Address2
				.Fields("Address3").Value = pBranchData.Address3
				.Fields("City").Value = pBranchData.City
				.Fields("Email").Value = pBranchData.Email
				.Fields("FaxNumber").Value = pBranchData.FaxNumber
				.Fields("ManagerFirstName").Value = pBranchData.ManagerFirstName
				.Fields("ManagerLastName").Value = pBranchData.ManagerLastName
				.Fields("Number").Value = pBranchData.Number
				.Fields("PhoneNumber").Value = pBranchData.PhoneNumber
				.Fields("State").Value = pBranchData.State
				.Fields("Zip").Value = pBranchData.Zip
				'
				.UpdateBatch()
				'
			End With
			'
			rsBranch.Close()
			Save = True
			'
		End If
		'UPGRADE_NOTE: Object rsBranch may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsBranch = Nothing
		'
		Exit Function
EH: 
		If Not rsBranch Is Nothing Then
			If rsBranch.State = ADODB.ObjectStateEnum.adStateOpen Then
				rsBranch.CancelUpdate()
				rsBranch.Close()
			End If
		End If
		MsgBox(Err.Description)
	End Function
	
	'Public Function AddNew(ByRef pBranchData As CBranchData) As Boolean
	'  AddNew = False
	'  Dim rslist As New ADODB.Recordset
	'  Dim lNewID As Long
	'  Dim sTempPass As String
	'  '
	'  rslist.Open "SELECT * FROM TBranch WHERE BranchID = -1", cnMain, adOpenKeyset, adLockBatchOptimistic
	'  '
	'  If Not (rslist.BOF And rslist.eof) Then
	'    With rslist
	'      '.MoveLast
	'      lNewID = NextID("BranchID", "TBranch", cnMain) '(!BranchID + 1)
	'      .AddNew
	'      '
	'      !BranchID = lNewID
	'      !CompanyID = pBranchData.CompanyID
	'      !Name = pBranchData.Name
	'      !Address1 = pBranchData.Address1
	'      !Address2 = pBranchData.Address2
	'      !Address3 = pBranchData.Address3
	'      !City = pBranchData.City
	'      !Email = pBranchData.Email
	'      !FaxNumber = pBranchData.FaxNumber
	'      !ManagerFirstName = pBranchData.ManagerFirstName
	'      !ManagerLastName = pBranchData.ManagerLastName
	'      !Number = pBranchData.Number
	'      !PhoneNumber = pBranchData.PhoneNumber
	'      !State = pBranchData.State
	'      !Zip = pBranchData.Zip
	'      '
	'      rslist.UpdateBatch
	'      '
	'    End With
	'    '
	'    rslist.Close
	'    AddNew = True
	'    '
	'  End If
	'  Set rslist = Nothing
	'
	'End Function
	
	Public Function Delete(ByRef pID As Integer) As Boolean
		Delete = False
		'
		Dim rslist As New ADODB.Recordset
		Dim rsContactList As New ADODB.Recordset
		'
		rsContactList.Open("SELECT ID FROM TContact WHERE BranchID = " & pID, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		If rsContactList.RecordCount > 0 Then
			Delete = False
			'
			Exit Function
		Else
			'
			rslist.Open("SELECT * FROM TBranch WHERE BranchID = " & pID, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			If Not (rslist.BOF And rslist.eof) Then
				'
				rslist.Delete()
				rslist.UpdateBatch()
				'
				rslist.Close()
				'
				Delete = True
			End If
			
			'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rslist = Nothing
		End If
	End Function
End Class