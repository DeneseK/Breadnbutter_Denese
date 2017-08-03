Option Strict Off
Option Explicit On
Friend Class CEmployee
	
	Public Function LoadCollection(ByRef pEmployees As CEmployees) As Boolean
		LoadCollection = False
		Dim rslist As New ADODB.Recordset
		Dim EmployeeData As CEmployeeData
		'
		rslist.Open("SELECT * FROM tblEmployees ORDER BY EmployeeID", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		If Not (rslist.BOF And rslist.eof) Then
			While Not rslist.eof
				With rslist
					EmployeeData = New CEmployeeData
					'
					'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EmployeeData.EmployeeID = nnNum(.Fields("EmployeeID"))
					EmployeeData.EmployeeNumber = .Fields("EmployeeNumber").Value & vbNullString
					EmployeeData.EmployeeLast = .Fields("EmployeeLast").Value & vbNullString
					EmployeeData.EmployeeFirst = .Fields("EmployeeFirst").Value & vbNullString
					EmployeeData.EmployeeMiddle = .Fields("EmployeeMiddle").Value & vbNullString
					EmployeeData.Password = .Fields("Password").Value & vbNullString
					'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EmployeeData.Groups = nnNum(.Fields("Groups"))
					'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EmployeeData.SecurityLevel = nnNum(.Fields("SecurityLevel"))
					'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EmployeeData.EmployeeExt = nnNum(.Fields("EmployeeExt"))
					'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EmployeeData.WorkGroups = nnNum(.Fields("WorkGroups"))
					EmployeeData.EMailAddress = .Fields("EMailAddress").Value & vbNullString
					'
					pEmployees.Add(EmployeeData, EmployeeData.EmployeeID)
					'
					rslist.MoveNext()
				End With
			End While
			'
			rslist.Close()
			LoadCollection = True
		End If
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object EmployeeData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EmployeeData = Nothing
		'
	End Function
	
	Public Function Load(ByRef pEmployeeData As CEmployeeData, ByRef pID As Short) As Boolean
		Load = False
		Dim rslist As New ADODB.Recordset
		'
		rslist.Open("SELECT * FROM tblEmployees WHERE EmployeeID = " & pID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		If Not (rslist.BOF And rslist.eof) Then
			With rslist
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEmployeeData.EmployeeID = nnNum(.Fields("EmployeeID"))
				pEmployeeData.EmployeeNumber = .Fields("EmployeeNumber").Value & vbNullString
				pEmployeeData.EmployeeLast = .Fields("EmployeeLast").Value & vbNullString
				pEmployeeData.EmployeeFirst = .Fields("EmployeeFirst").Value & vbNullString
				pEmployeeData.EmployeeMiddle = .Fields("EmployeeMiddle").Value & vbNullString
				pEmployeeData.Password = .Fields("Password").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEmployeeData.Groups = nnNum(.Fields("Groups"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEmployeeData.SecurityLevel = nnNum(.Fields("SecurityLevel"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEmployeeData.EmployeeExt = nnNum(.Fields("EmployeeExt"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEmployeeData.WorkGroups = nnNum(.Fields("WorkGroups"))
				pEmployeeData.EMailAddress = .Fields("EMailAddress").Value & vbNullString
				'
			End With
			'
			rslist.Close()
			Load = True
		End If
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'
	End Function
	
	Public Function Save(ByRef pEmployeeData As CEmployeeData, ByRef pID As Short) As Boolean
		Save = False
		Dim rslist As New ADODB.Recordset
		'
		rslist.Open("SELECT * FROM tblEmployees WHERE EmployeeID = " & pID, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		If Not (rslist.BOF And rslist.eof) Then
			With rslist
				'
				'!EmployeeID = pEmployeeData.EmployeeID
				'!EmployeeNumber = pEmployeeData.EmployeeNumber
				.Fields("EmployeeLast").Value = pEmployeeData.EmployeeLast
				.Fields("EmployeeFirst").Value = pEmployeeData.EmployeeFirst
				.Fields("EmployeeMiddle").Value = pEmployeeData.EmployeeMiddle
				.Fields("Password").Value = EncryptStr((pEmployeeData.Password))
				.Fields("Groups").Value = pEmployeeData.Groups
				.Fields("SecurityLevel").Value = pEmployeeData.SecurityLevel
				.Fields("EmployeeExt").Value = pEmployeeData.EmployeeExt
				.Fields("WorkGroups").Value = pEmployeeData.WorkGroups
				.Fields("EMailAddress").Value = pEmployeeData.EMailAddress
				
				rslist.UpdateBatch()
				'
			End With
			'
			rslist.Close()
			Save = True
			'BONUS: Change Password
			cnMain.Execute("EXEC sp_password NULL, '" & Rot39(pEmployeeData.Password) & "', '" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "'")
		End If
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'
	End Function
	
	Public Function AddNew(ByRef pEmployeeData As CEmployeeData) As Boolean
		AddNew = False
		Dim rslist As New ADODB.Recordset
		Dim iNewID As Short
		Dim sTempPass As String
		'
		rslist.Open("SELECT * FROM tblEmployees ORDER BY EmployeeID", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		If Not (rslist.BOF And rslist.eof) Then
			With rslist
				.MoveLast()
				iNewID = (.Fields("EmployeeID").Value + 1)
				.AddNew()
				'
				.Fields("EmployeeID").Value = iNewID
				'!EmployeeNumber = pEmployeeData.EmployeeNumber
				.Fields("EmployeeLast").Value = pEmployeeData.EmployeeLast
				.Fields("EmployeeFirst").Value = pEmployeeData.EmployeeFirst
				.Fields("EmployeeMiddle").Value = pEmployeeData.EmployeeMiddle
				.Fields("Password").Value = EncryptStr((pEmployeeData.Password))
				.Fields("Groups").Value = pEmployeeData.Groups
				.Fields("SecurityLevel").Value = pEmployeeData.SecurityLevel
				.Fields("EmployeeExt").Value = pEmployeeData.EmployeeExt
				.Fields("WorkGroups").Value = pEmployeeData.WorkGroups
				.Fields("EMailAddress").Value = pEmployeeData.EMailAddress
				rslist.UpdateBatch()
				'
			End With
			'
			rslist.Close()
			AddNew = True
			'
			sTempPass = Rot39(pEmployeeData.Password)
			'
			'Create Security Login
			cnMain.Execute("EXEC sp_addlogin  '" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "', '" & sTempPass & "', 'BNB_DATA'")
			'Give access to DB. Current one I guess.
			cnMain.Execute("EXEC sp_grantdbaccess N'" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "', N'" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "'")
			'Assign "User" role.
			cnMain.Execute("EXEC sp_addrolemember N'User', N'" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "'")
			'BONUS: Change Password
			'cnMain.Execute "EXEC sp_password NULL, 'gnarly', ' & sTempPass & '"
			'
		End If
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		
	End Function
	
	Public Function Delete(ByRef pID As Short) As Boolean
		Delete = False
		Dim rslist As New ADODB.Recordset
		Dim sTempName As String
		'
		rslist.Open("SELECT * FROM tblEmployees WHERE EmployeeID = " & pID, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		If Not (rslist.BOF And rslist.eof) Then
			'
			sTempName = rslist.Fields("EmployeeFirst").Value & rslist.Fields("EmployeeLast").Value
			'
			rslist.Delete()
			rslist.UpdateBatch()
			'
			rslist.Close()
			'
			cnMain.Execute("EXEC sp_dropuser   '" & sTempName & "'")
			cnMain.Execute("EXEC sp_droplogin   '" & sTempName & "'")
			'
			Delete = True
		End If
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'
	End Function
	
	Public Function InGroup(ByRef sName As String, ByRef sWorkGroup As String) As Boolean
		InGroup = False
		Dim rslist As New ADODB.Recordset
		Dim EmployeeData As CEmployeeData
		Dim iWorkGroupNum As Short
		'
		rslist.Open("select * from tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		With rslist
			Do While Not .eof
				If LCase(sName) = LCase(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iWorkGroupNum = nnNum(.Fields("WorkGroups"))
				End If
				.MoveNext()
			Loop 
		End With
		Select Case sWorkGroup
			Case "Management"
				If iWorkGroupNum > 7 Then InGroup = True
			Case "Sales"
				Select Case iWorkGroupNum
					Case 4, 5, 6, 7, 12, 13, 14, 15
						InGroup = True
				End Select
			Case "Support"
				Select Case iWorkGroupNum
					Case 2, 3, 6, 7, 10, 11, 14, 15
						InGroup = True
				End Select
			Case "Development"
				Select Case iWorkGroupNum
					Case 1, 3, 5, 7, 9, 11, 13, 15
						InGroup = True
				End Select
		End Select
		'
		rslist.Close()
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object EmployeeData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EmployeeData = Nothing
		'
	End Function
	
	Public Function GetEmployeeID(ByRef psName As String) As Integer
		GetEmployeeID = 0
		'
		Dim rslist As New ADODB.Recordset
		Dim EmployeeData As CEmployeeData
		'
		rslist.Open("select * from tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		With rslist
			Do While Not .eof
				If LCase(psName) = LCase(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetEmployeeID = nnNum(.Fields("EmployeeID"))
				End If
				.MoveNext()
			Loop 
		End With
		rslist.Close()
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object EmployeeData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EmployeeData = Nothing
		'
		'
	End Function
End Class