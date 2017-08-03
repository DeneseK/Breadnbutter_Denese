Option Strict Off
Option Explicit On
Friend Class CContact
	
	Public Function Save(ByRef pContactData As CContactData, ByRef pbNew As Boolean) As Boolean
		On Error GoTo EH
		'
		Dim rsContact As New ADODB.Recordset
		'
		If pbNew Then
			pContactData.ID = NextID("ID", "TContact", cnMain)
			'pContactData.DateEntered = Now
			rsContact.Open("SELECT * FROM TContact WHERE ID = -1", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			rsContact.AddNew()
			rsContact.Fields("ID").Value = pContactData.ID
			rsContact.Fields("DateEntered").Value = CDate(Now) 'pContactData.DateEntered
		Else
			rsContact.Open("SELECT * FROM TContact WHERE ID = " & pContactData.ID, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			'
			If rsContact.eof Then
				Save = False
				Exit Function
				'Else
				'rsContact.Edit
			End If
		End If
		'
		With rsContact
			.Fields("CompanyID").Value = pContactData.CompanyID
			.Fields("BranchID").Value = pContactData.BranchID
			.Fields("FirstName").Value = pContactData.FirstName
			.Fields("LastName").Value = pContactData.LastName
			.Fields("Salutation").Value = pContactData.Salutation
			.Fields("Title").Value = pContactData.Title
			.Fields("Address1").Value = pContactData.Address1
			.Fields("Address2").Value = pContactData.Address2
			.Fields("City").Value = pContactData.City
			.Fields("State").Value = pContactData.State
			.Fields("Zip").Value = pContactData.Zip
			.Fields("PermMailAddress1").Value = pContactData.MailAddress1
			.Fields("PermMailAddress2").Value = pContactData.MailAddress2
			.Fields("PermMailCity").Value = pContactData.MailCity
			.Fields("PermMailState").Value = pContactData.MailState
			.Fields("PermMailZip").Value = pContactData.MailZip
			'
			.Fields("PCEmail").Value = pContactData.PCEmail
			.Fields("PCEmailPassword").Value = pContactData.PCEmailPassword
			'
			.Fields("Phone1").Value = pContactData.Phone1
			.Fields("Phone2").Value = pContactData.Phone2
			.Fields("Fax").Value = pContactData.Fax
			.Fields("Email").Value = pContactData.Email
			.Fields("Source").Value = pContactData.Source
			.Fields("betatester").Value = IIf(pContactData.Selected = 1, True, False)
			.Fields("PreferredAddress").Value = pContactData.PreferredAddress
			.Fields("Notes").Value = pContactData.Notes
			'
			.Fields("Status").Value = pContactData.Status
			.Fields("ShipStatus").Value = pContactData.ShipStatus
			.Fields("AuthStatus").Value = pContactData.AuthStatus
			'!DateEntered = pContactData.DateEntered
			.Fields("ShipDate").Value = pContactData.ShipDate
			.Fields("AuthDate").Value = pContactData.AuthDate
			.Fields("AuthDays").Value = pContactData.AuthDays
			.Fields("Copies").Value = pContactData.Copies
			.Fields("VersionShipped").Value = pContactData.VersionShipped
			' !AuthRemaining = pContactData.AuthRemaining
			'
			.Fields("PVShipStatus").Value = pContactData.PVShipStatus
			.Fields("PVAuthStatus").Value = pContactData.PVAuthStatus
			.Fields("PVDownloadStatus").Value = pContactData.PVDownloadStatus
			.Fields("DownloadStatus").Value = pContactData.DownloadStatus
			'!PVDateEntered = pContactData.PVDateEntered
			.Fields("PVShipDate").Value = pContactData.PVShipDate
			.Fields("PVDownloadDate").Value = pContactData.PVDownloadDate
			.Fields("DownloadDate").Value = pContactData.DownloadDate
			.Fields("PVAuthDate").Value = pContactData.PVAuthDate
			.Fields("PVAuthDays").Value = pContactData.PVAuthDays
			.Fields("PVCopies").Value = pContactData.PVCopies
			.Fields("PVVersionShipped").Value = pContactData.PVVersionShipped
			'!PVAuthRemaining = pContactData.PVAuthRemaining
			'
			.Fields("Rate").Value = pContactData.Rate_Renamed
			.Fields("ContactType").Value = pContactData.ContactType
			.Fields("RateExpDate").Value = pContactData.RateExpDate
			'
			.Fields("GraceDays").Value = pContactData.GraceDays
			.Fields("OnlineAuths").Value = pContactData.OnlineAuths
			.Fields("SaleDate").Value = pContactData.SaleDate
			.Fields("SaleDays").Value = pContactData.SaleDays
			'
			.Fields("PVGraceDays").Value = pContactData.PVGraceDays
			.Fields("PVOnlineAuths").Value = pContactData.PVOnlineAuths
			.Fields("PVSaleDate").Value = pContactData.PVSaleDate
			.Fields("PVSaleDays").Value = pContactData.PVSaleDays
			'
			.Fields("WebPassword").Value = pContactData.WebPassword
			.Fields("ContactByEmail").Value = pContactData.ContactByEmail
			.Fields("ChangedData").Value = pContactData.ChangedData
			'
			.Fields("AdjusterID").Value = pContactData.AdjusterID
			'
			.Fields("LastUpdate").Value = Now
			.Update()
		End With
		'
		rsContact.Close()
		'
		'UPGRADE_NOTE: Object rsContact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsContact = Nothing
		'
		Save = True
		'fAdding = False
		'
		Exit Function
EH: 
		If Not rsContact Is Nothing Then
			If rsContact.State = ADODB.ObjectStateEnum.adStateOpen Then
				rsContact.CancelUpdate()
				rsContact.Close()
			End If
		End If
		MsgBox(Err.Description)
	End Function
	
	Public Function Load(ByRef pContactData As CContactData, ByRef plID As Integer) As Boolean
		On Error GoTo EH
		'
		Dim rsContact As New ADODB.Recordset
		'
		pContactData = New CContactData
		'
		rsContact.Open("SELECT * FROM TContact WHERE ID = " & plID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		'
		If Not rsContact.eof Then
			With rsContact
				pContactData.ID = .Fields("ID").Value
				pContactData.CompanyID = .Fields("CompanyID").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.BranchID = nnNum(.Fields("BranchID"))
				'pContactData.LastUpdate = nnNum(!LastUpdate)
				pContactData.FirstName = .Fields("FirstName").Value & vbNullString
				pContactData.LastName = .Fields("LastName").Value & vbNullString
				pContactData.Salutation = .Fields("Salutation").Value & vbNullString
				pContactData.Title = .Fields("Title").Value & vbNullString
				pContactData.Address1 = .Fields("Address1").Value & vbNullString
				pContactData.Address2 = .Fields("Address2").Value & vbNullString
				pContactData.City = .Fields("City").Value & vbNullString
				pContactData.State = .Fields("State").Value & vbNullString
				pContactData.Zip = .Fields("Zip").Value & vbNullString
				pContactData.MailAddress1 = .Fields("PermMailAddress1").Value & vbNullString
				pContactData.MailAddress2 = .Fields("PermMailAddress2").Value & vbNullString
				pContactData.MailCity = .Fields("PermMailCity").Value & vbNullString
				pContactData.MailState = .Fields("PermMailState").Value & vbNullString
				pContactData.MailZip = .Fields("PermMailZip").Value & vbNullString
				'
				pContactData.PCEmail = .Fields("PCEmail").Value & vbNullString
				pContactData.PCEmailPassword = .Fields("PCEmailPassword").Value & vbNullString
				'
				pContactData.Phone1 = .Fields("Phone1").Value & vbNullString
				pContactData.Phone2 = .Fields("Phone2").Value & vbNullString
				pContactData.Fax = .Fields("Fax").Value & vbNullString
				pContactData.Email = .Fields("Email").Value & vbNullString
				pContactData.Source = .Fields("Source").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.Selected = IIf(nnNum(.Fields("betatester")), 1, 0)
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PreferredAddress = nnNum(.Fields("PreferredAddress"))
				pContactData.Notes = .Fields("Notes").Value & vbNullString
				pContactData.Status = .Fields("Status").Value & vbNullString
				'
				pContactData.ShipStatus = .Fields("ShipStatus").Value & vbNullString
				pContactData.AuthStatus = .Fields("AuthStatus").Value & vbNullString
				'pContactData.DateEntered = nnNum(!DateEntered)
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.ShipDate = nnNum(.Fields("ShipDate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.AuthDate = nnNum(.Fields("AuthDate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.AuthDays = nnNum(.Fields("AuthDays"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.Copies = nnNum(.Fields("Copies"))
				pContactData.VersionShipped = .Fields("VersionShipped").Value & vbNullString
				'pContactData.AuthRemaining = nnNum(!AuthRemaining)
				'
				pContactData.PVShipStatus = .Fields("PVShipStatus").Value & vbNullString
				pContactData.PVDownloadStatus = .Fields("PVDownloadStatus").Value & vbNullString
				pContactData.DownloadStatus = .Fields("DownloadStatus").Value & vbNullString
				pContactData.PVAuthStatus = .Fields("PVAuthStatus").Value & vbNullString
				'     pContactData.PVDateEntered = nnNum(!PVDateEntered)
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVShipDate = nnNum(.Fields("PVShipDate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVDownloadDate = nnNum(.Fields("PVDownloadDate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.DownloadDate = nnNum(.Fields("DownloadDate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVAuthDate = nnNum(.Fields("PVAuthDate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVAuthDays = nnNum(.Fields("PVAuthDays"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVCopies = nnNum(.Fields("PVCopies"))
				pContactData.PVVersionShipped = .Fields("PVVersionShipped").Value & vbNullString
				'pContactData.PVAuthRemaining = nnNum(!PVAuthRemaining)
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.Rate_Renamed = nnNum(.Fields("Rate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.ContactType = nnNum(.Fields("ContactType"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.RateExpDate = nnNum(.Fields("RateExpDate"))
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.GraceDays = nnNum(.Fields("GraceDays"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.OnlineAuths = nnNum(.Fields("OnlineAuths"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.SaleDate = nnNum(.Fields("SaleDate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.SaleDays = nnNum(.Fields("SaleDays"))
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVGraceDays = nnNum(.Fields("PVGraceDays"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVOnlineAuths = nnNum(.Fields("PVOnlineAuths"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVSaleDate = nnNum(.Fields("PVSaleDate"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.PVSaleDays = nnNum(.Fields("PVSaleDays"))
				'
				pContactData.WebPassword = .Fields("WebPassword").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.ContactByEmail = nnNum(.Fields("ContactByEmail"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pContactData.ChangedData = nnNum(.Fields("ChangedData"))
				'
				pContactData.AdjusterID = .Fields("AdjusterID").Value & vbNullString
			End With
			'
			If IsDate(pContactData.AuthDate) Then
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				pContactData.AuthRemaining = DateDiff(Microsoft.VisualBasic.DateInterval.Day, Now, DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(pContactData.AuthDays), pContactData.AuthDate))
			End If
			'
			If IsDate(pContactData.AuthDate) Then
				'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
				pContactData.PVAuthRemaining = DateDiff(Microsoft.VisualBasic.DateInterval.Day, Now, DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(pContactData.PVAuthDays), pContactData.PVAuthDate))
			End If
			'
			pContactData.DaysPending = CalculatePendingDays((pContactData.SaleDate), (pContactData.GraceDays), (pContactData.SaleDays))
			'
			pContactData.PVDaysPending = CalculatePendingDays((pContactData.PVSaleDate), (pContactData.PVGraceDays), (pContactData.PVSaleDays))
			'
			Load = True
		Else
			Load = False
		End If
		'
		rsContact.Close()
		'
		'UPGRADE_NOTE: Object rsContact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsContact = Nothing
		Exit Function
EH: 
		MsgBox(Err.Description & " in Class Contact: Load.")
	End Function
	
	'Private Function CalculatePendingDays(pdSaleDate As Date, plGraceDays As Long, plSaleDays As Long) As Long
	'  Dim lTempPending As Long
	'  Dim lDaysPassed As Long
	'  '
	'  lDaysPassed = Abs(DateDiff("y", pdSaleDate, Now))
	'  '
	'  If plGraceDays < 0 Then
	'    CalculatePendingDays = plSaleDays
	'  Else
	'    If lDaysPassed < plGraceDays Then
	'      CalculatePendingDays = plSaleDays
	'    Else
	'      lTempPending = plSaleDays - (lDaysPassed - plGraceDays)
	'      '
	'      If lTempPending >= 0 Then
	'        CalculatePendingDays = lTempPending
	'      Else
	'        CalculatePendingDays = 0
	'      End If
	'    End If
	'  End If
	'End Function
	
	Public Function Delete(ByVal plID As Integer) As Boolean
		
		If InputBox("Type DELETE and click OK if you really want to delete this contact.", "Delete Contact") = "DELETE" Then
			If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
				cnMain.Execute("DELETE FROM tblSupportActs WHERE CustRecID = " & plID,  , ADODB.CommandTypeEnum.adCmdText)
				cnMain.Execute("DELETE FROM TContact WHERE ID = " & plID,  , ADODB.CommandTypeEnum.adCmdText)
			Else
				cnMain.Execute("DELETE FROM TSupportActs WHERE CustRecID = " & plID,  , ADODB.CommandTypeEnum.adCmdText)
				cnMain.Execute("DELETE FROM TContact WHERE ID = " & plID,  , ADODB.CommandTypeEnum.adCmdText)
			End If
			'Me.Clear
			Delete = True
		Else
			Delete = False
		End If
		
	End Function
	
	Public Sub LoadCollection(ByVal plCompanyID As Integer, ByRef pContacts As CContacts, ByVal plBranchID As Integer)
		Dim rslist As New ADODB.Recordset
		Dim ContactData As CContactData
		Dim BranchData As CBranchData
		'
		Dim sQuery As String
		'
		pContacts = New CContacts
		'
		sQuery = "SELECT TContact.ID, TContact.ContactType, TContact.Status, TContact.FirstName, TContact.LastName, " & "TContact.AuthDate, TContact.AuthDays " & "FROM TContact " & "WHERE (((TContact.CompanyID)=" & plCompanyID & "))" 'ORDER BY TContact.LastName, TContact.FirstName"
		'"TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & _
		''
		If plBranchID > 0 Then
			sQuery = sQuery & " AND (TCONTACT.BRANCHID = " & plBranchID & ") "
		End If
		'
		sQuery = sQuery & " ORDER BY TContact.LastName, TContact.FirstName"
		'
		rslist.Open(sQuery, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		'
		' If rslist.RecordCount = 0 Then
		' Debug.Print rslist.RecordCount
		' End If
		'
		While Not rslist.eof
			With rslist
				ContactData = New CContactData
				'
				ContactData.ID = rslist.Fields("ID").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ContactData.AuthDays = nnNum(rslist.Fields("AuthDays"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ContactData.AuthDate = nnNum(rslist.Fields("AuthDate"))
				ContactData.Status = rslist.Fields("Status").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ContactData.ContactType = nnNum(rslist.Fields("ContactType"))
				ContactData.FirstName = rslist.Fields("FirstName").Value & vbNullString
				ContactData.LastName = rslist.Fields("LastName").Value & vbNullString
				'Load ContactData, rslist!ID
				'
				pContacts.Add(ContactData)
				'
				rslist.MoveNext()
			End With
		End While
		'
		rslist.Close()
		'
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object ContactData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ContactData = Nothing
		'
	End Sub
End Class