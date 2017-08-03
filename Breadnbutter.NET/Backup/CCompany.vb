Option Strict Off
Option Explicit On
Friend Class CCompany
	
	Public Function Delete(ByVal plID As Integer) As Object
		
		If InputBox("Type DELETE if you sure you want to delete this company.", "Delete Contact") = "DELETE" Then
			
			'* TODO Delete Contacts
			
			cnMain.Execute("DELETE FROM TCompany WHERE ID = " & plID)
			'
		End If
		
	End Function
	
	Public Function Load(ByRef pCompanyData As CCompanyData, ByRef plID As Integer) As Boolean
		Dim rsCompany As New ADODB.Recordset
		'
		pCompanyData = New CCompanyData
		'
		rsCompany.Open("SELECT * FROM TCompany WHERE ID = " & plID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		If Not rsCompany.eof Then
			pCompanyData = New CCompanyData
			With rsCompany
				pCompanyData.ID = .Fields("ID").Value
				'pCompanyData.DateEntered = nnNum(!DateEntered)
				'pCompanyData.LastUpdate = nnNum(!LastUpdate)
				pCompanyData.Name = .Fields("Name").Value & vbNullString
				'pCompanyData.Division = !Division & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pCompanyData.Individual = nnNum(.Fields("Individual"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pCompanyData.DoNotContact = nnNum(.Fields("DoNotContact"))
				pCompanyData.Note = .Fields("Note").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pCompanyData.InterestRank = nnNum(.Fields("InterestRank"))
				'
			End With
			'
			Load = True
		Else
			Load = False
		End If
		'
		rsCompany.Close()
		'UPGRADE_NOTE: Object rsCompany may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCompany = Nothing
	End Function
	
	Public Function Save(ByRef pCompanyData As CCompanyData, ByRef pbNew As Boolean) As Boolean
		On Error GoTo EH
		'
		Dim rsCompany As ADODB.Recordset
		'
		rsCompany = New ADODB.Recordset
		'
		If pbNew Then
			pCompanyData.ID = NextID("ID", "TCompany", cnMain)
			rsCompany.Open("SELECT * FROM TCompany WHERE ID = -1", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
			rsCompany.AddNew()
			rsCompany.Fields("ID").Value = pCompanyData.ID
			rsCompany.Fields("DateEntered").Value = CDate(Now) 'pCompanyData.DateEntered
		Else
			'
			rsCompany.Open("SELECT * FROM TCompany WHERE ID =" & pCompanyData.ID, cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
			'
			If rsCompany.eof Then
				Save = False
				Exit Function
				'Else
				' rsCompany.Edit
				'End If
			End If
		End If
		'
		With rsCompany
			.Fields("Name").Value = pCompanyData.Name
			.Fields("LastUpdate").Value = Now
			'!Division = pCompanyData.Division
			.Fields("Individual").Value = pCompanyData.Individual
			.Fields("DoNotContact").Value = pCompanyData.DoNotContact
			.Fields("Note").Value = pCompanyData.Note
			.Fields("InterestRank").Value = pCompanyData.InterestRank
			.Update()
		End With
		'
		rsCompany.Close()
		'UPGRADE_NOTE: Object rsCompany may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCompany = Nothing
		'
		Save = True
		'fAdding = False
		'
		Exit Function
EH: 
		Save = False
		If Err.Number = 3022 Then 'Duplicate company name record
			MsgBox("Duplicate company name detected. You must enter a name and division that is unique.", MsgBoxStyle.Information, "Save Company")
		Else
			MsgBox(Err.Description)
		End If
	End Function
	
	'Public Sub LoadCompanyList()
	'
	'  Set rsCompanyList = New adodb.Recordset
	'  rsCompanyList.Open "SELECT ID FROM TCompany ORDER BY Name", cnMain, adOpenKeyset, adLockOptimistic, adCmdText
	'  '
	'  With rsCompanyList
	'  If Not .eof Then .MoveLast
	'  lCount = .RecordCount
	'  If Not .BOF Then .MoveFirst
	'  End With
	'
	'End Sub
	
	'Public Function ListID() As Long
	'
	'  If Not (rsCompanyList.BOF Or rsCompanyList.eof) Then
	'    ListID = rsCompanyList!ID
	'  Else
	'    ListID = 0
	'  End If
	'
	'End Function
	'
	'Public Function MoveFirst() As Boolean
	'
	'  With rsCompanyList
	'  If Not .BOF Then
	'    .MoveFirst
	'    If Not (.BOF Or .eof) Then
	'      MoveFirst = True
	'    End If
	'  End If
	'  End With
	'
	'End Function
	'
	'Public Function MovePrevious() As Long
	'
	'  With rsCompanyList
	'  If Not .BOF Then
	'    .MovePrevious
	'    If Not .BOF Then
	'      MovePrevious = !ID
	'    Else
	'      .MoveFirst
	'      MovePrevious = !ID
	'    End If
	'  Else
	'    MovePrevious = 0
	'  End If
	'  End With
	'
	'End Function
	'
	'Public Function MoveNext() As Boolean
	'
	'  With rsCompanyList
	'  If Not .eof Then
	'    .MoveNext
	'    If Not .eof Then
	'      MoveNext = True
	'    Else
	'      .MoveLast
	'    End If
	'  End If
	'  End With
	'
	'End Function
	'
	'Public Function MoveLast() As Long
	'
	'  With rsCompanyList
	'  If Not .eof Then
	'    .MoveLast
	'    If Not .eof Then
	'      MoveLast = !ID
	'    Else
	'      MoveLast = 0
	'    End If
	'  End If
	'  End With
	'
	'End Function
	
	
	'Public Property Get ID() As Long
	'  ID = CR.ID
	'End Property
	'
	'Public Property Let ID(ByVal NewValue As Long)
	'  CR.ID = NewValue
	'End Property
	'
	'Public Property Get DateEntered() As Date
	'  DateEntered = CR.DateEntered
	'End Property
	'
	'Public Property Get LastUpdate() As Date
	'  LastUpdate = CR.LastUpdate
	'End Property
	'
	'Public Property Get Name() As String
	'  Name = CR.Name
	'End Property
	'
	'Public Property Let Name(ByVal NewValue As String)
	'  CR.Name = NewValue
	'End Property
	'
	'Public Property Get Division() As String
	'  Division = CR.Division
	'End Property
	'
	'Public Property Let Office(ByVal NewValue As String)
	'  CR.Division = NewValue
	'End Property
	'
	'Public Property Get Individual() As Boolean
	'  Individual = CR.Individual
	'End Property
	'
	'Public Property Let Individual(ByVal NewValue As Boolean)
	'  CR.Individual = NewValue
	'End Property
	'
	'Public Property Get DoNotContact() As Boolean
	'  DoNotContact = CR.DoNotContact
	'End Property
	'
	'Public Property Let DoNotContact(ByVal NewValue As Boolean)
	'  CR.DoNotContact = NewValue
	'End Property
	'
	'Public Property Get Note() As String
	'  Note = CR.Note
	'End Property
	'
	'Public Property Let Note(ByVal NewValue As String)
	'  CR.Note = NewValue
	'End Property
	
	'Private Sub Class_Initialize()
	'  Set Contact = New CContact
	'End Sub
	
	'Private Sub Class_Terminate()
	'
	'  On Error Resume Next
	'  Set Contact = Nothing
	'  '
	'  rsCompanyList.Close
	'  Set rsCompanyList = Nothing
	'  rsCompany.Close
	'  Set rsCompany = Nothing
	'
	'End Sub
	
	'Public Property Get SearchID() As Long
	'  SearchID = lSearchID
	'End Property
	'
	'Public Property Let SearchID(ByVal NewValue As Long)
	'  lSearchID = NewValue
	'End Property
	
	'Public Function UsersCount(SearchCriteria As String) As Long
	'  On Error GoTo EH
	'  '
	'  Dim rs As New adodb.Recordset
	'  rs.Open "SELECT TContact.Status FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
	''          "WHERE (((TContact.Status) Like '" & SearchCriteria & "') AND ((TCompany.ID)=" & CR.ID & "));", cnMain
	'  '
	'  UsersCount = rs.RecordCount
	'   '
	'  DBOps.ZapRS rs
	'  Exit Function
	'EH:
	'  MsgBox Err.Description & " in UsersCount"
	'  UsersCount = -1
	'End Function
	
	'Public Function CustomersCount() As Long
	'  CustomersCount = UsersCount("Customer")
	'End Function
	'
	'Public Function ProspectsCount() As Long
	'  ProspectsCount = UsersCount("Prospect")
	'End Function
	'
	'Public Function FutureProspectsCount() As Long
	'  FutureProspectsCount = UsersCount("Future Prospect")
	'End Function
	
	'Public Function TotalUserCount() As Long
	'  TotalUserCount = UsersCount("%")
	'End Function
	
	'Public Function GetContactList(pCN As ADODB.Connection, plBranchID As Long) As Recordset
	'  '
	'  On Error GoTo EH
	'  '
	'  Dim sQuery As String
	'  '
	'  If Not GetContactList Is Nothing Then
	'    If GetContactList.State = adStateOpen Then GetContactList.Close
	'  Else
	'    Set GetContactList = New ADODB.Recordset
	'  End If
	'  '
	'  sQuery = "SELECT TContact.ID, TContact.ContactType, TContact.Status, TContact.FirstName, TContact.LastName, " & _
	''            "TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & _
	''            "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
	''            "WHERE (((TCOMPANY.ID)=" & Company.ID & "))" 'ORDER BY TContact.LastName, TContact.FirstName"
	'  '
	'  If plBranchID > 0 Then
	'    sQuery = sQuery & " AND (TCONTACT.BRANCHID = " & plBranchID & ") "
	'  End If
	'  '
	'  sQuery = sQuery & " ORDER BY TContact.LastName, TContact.FirstName"
	'
	'  '
	'  GetContactList.Open sQuery, cnMain, adOpenForwardOnly, adLockReadOnly
	'  '
	'  Exit Function
	'EH:
	'    MsgBox ("Error " & Err.Description & " in CCompany.GetContactList")
	'End Function
	
	'Public Function GetDetailContactList(pCN As ADODB.Connection) As Recordset
	'  On Error GoTo EH
	'  '
	'  Dim sQuery As String
	'  '
	'  If Not GetDetailContactList Is Nothing Then
	'    If GetDetailContactList.State = adStateOpen Then GetDetailContactList.Close
	'  Else
	'    Set GetDetailContactList = New ADODB.Recordset
	'  End If
	'  '
	'  If ConnType = SQL Then
	'    sQuery = "SELECT TContact.ID, TContact.ContactType, TContact.Status, TContact.FirstName, TContact.LastName, " & _
	''                "TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as [Days], " & _
	''                "TContact.AuthDate + TContact.AuthDays AS [ExpirationDate], " & _
	''                "TContact.Status, TContact.Title, TContact.ShipStatus, TContact.VersionShipped " & _
	''                "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
	''                "WHERE (((TCOMPANY.ID)=" & Company.ID & ")) ORDER BY TContact.LastName, TContact.FirstName"
	'  Else
	'     sQuery = "SELECT TContact.ID, TContact.ContactType, TContact.Status, TContact.FirstName, TContact.LastName, " & _
	''                "TContact.AuthDate + TContact.AuthDays AS [ExpirationDate], " & _
	''                "TContact.Status, TContact.Title, TContact.ShipStatus, TContact.VersionShipped " & _
	''                "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
	''                "WHERE (((TCOMPANY.ID)=" & Company.ID & ")) ORDER BY TContact.LastName, TContact.FirstName"
	'  End If
	'  '
	'  '
	'  GetDetailContactList.Open sQuery, cnMain, adOpenForwardOnly, adLockReadOnly
	'   Exit Function
	'EH:
	'    MsgBox ("Error in CCompany.GetDetailContactList")
	'End Function
	
	Public Sub LoadCollection(ByRef pCompanys As CCompanys)
		Dim rslist As New ADODB.Recordset
		'
		Dim CompanyData As CCompanyData
		'
		Dim sQuery As String
		'
		pCompanys = New CCompanys
		'
		sQuery = "SELECT TCompany.ID, TCompany.Name " & "FROM TCompany ORDER BY TCompany.Name"
		'
		rslist.Open(sQuery, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		'
		While Not rslist.eof
			With rslist
				CompanyData = New CCompanyData
				'
				CompanyData.ID = rslist.Fields("ID").Value
				CompanyData.Name = rslist.Fields("Name").Value & vbNullString
				'
				pCompanys.Add(CompanyData)
				'
				rslist.MoveNext()
			End With
		End While
		'
		rslist.Close()
		'
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object CompanyData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		CompanyData = Nothing
		'
	End Sub
End Class