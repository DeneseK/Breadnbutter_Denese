Option Strict Off
Option Explicit On
Friend Class CGroupListLink
	
	Public Function Delete(ByVal plID As Integer) As Object
		
		If InputBox("Type DELETE if you sure you want to delete this Custom Group.", "Delete Contact") = "DELETE" Then
			
			'* TODO Delete Contacts
			
			cnMain.Execute("DELETE FROM TGroupListLink WHERE ID = " & plID)
			'
		End If
		
	End Function
	
	Public Function CheckContact(ByRef plListID As Integer, ByRef plContactID As Integer) As Boolean
		Dim rsGroupListLink As New ADODB.Recordset
		'
		rsGroupListLink.Open("SELECT * FROM TGroupListLink WHERE ContactID = " & plContactID & " AND ListID = " & plListID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		If Not rsGroupListLink.eof Then
			'
			CheckContact = True
		Else
			CheckContact = False
		End If
		'
		rsGroupListLink.Close()
		'UPGRADE_NOTE: Object rsGroupListLink may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsGroupListLink = Nothing
	End Function
	
	Public Function Load(ByRef pGroupListLinkData As CGroupListLinkData, ByRef plID As Integer) As Boolean
		Dim rsGroupListLink As New ADODB.Recordset
		'
		pGroupListLinkData = New CGroupListLinkData
		'
		rsGroupListLink.Open("SELECT * FROM TGroupListLink WHERE ID = " & plID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		If Not rsGroupListLink.eof Then
			pGroupListLinkData = New CGroupListLinkData
			With rsGroupListLink
				pGroupListLinkData.ID = .Fields("ID").Value
				pGroupListLinkData.ContactID = .Fields("ContactID").Value
				pGroupListLinkData.ListID = .Fields("ListID").Value
				'
			End With
			'
			Load = True
		Else
			Load = False
		End If
		'
		rsGroupListLink.Close()
		'UPGRADE_NOTE: Object rsGroupListLink may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsGroupListLink = Nothing
	End Function
	
	Public Function AddContact(ByVal plContact As Integer, ByRef plList As Integer) As Object
		On Error GoTo EH
		'
		Dim rsGroupListLink As ADODB.Recordset
		'
		rsGroupListLink = New ADODB.Recordset
		'
		rsGroupListLink.Open("SELECT * FROM TGroupListLink WHERE ID = -1", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		rsGroupListLink.AddNew()
		'
		With rsGroupListLink
			.Fields("ID").Value = NextID("ID", "TGroupListLink", cnMain)
			.Fields("ContactID").Value = plContact
			.Fields("ListID").Value = plList
			.Update()
		End With
		'
		rsGroupListLink.Close()
		'UPGRADE_NOTE: Object rsGroupListLink may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsGroupListLink = Nothing
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object AddContact. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddContact = True
		'
		Exit Function
EH: 
		'UPGRADE_WARNING: Couldn't resolve default property of object AddContact. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddContact = False
		MsgBox(Err.Description)
	End Function
	
	Public Function DelContact(ByVal plContact As Integer, ByRef plList As Integer) As Object
		On Error GoTo EH
		'
		Dim rsGroupListLink As ADODB.Recordset
		'
		rsGroupListLink = New ADODB.Recordset
		'
		rsGroupListLink.Open("SELECT * FROM TGroupListLink WHERE ContactID = " & plContact & " AND ListID = " & plList, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		If Not rsGroupListLink.eof Then
			'
			With rsGroupListLink
				cnMain.Execute("DELETE FROM TGroupListLink WHERE ID = " & .Fields("ID").Value)
			End With
			'
		End If
		rsGroupListLink.Close()
		'UPGRADE_NOTE: Object rsGroupListLink may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsGroupListLink = Nothing
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object DelContact. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DelContact = True
		'
		Exit Function
EH: 
		'UPGRADE_WARNING: Couldn't resolve default property of object DelContact. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DelContact = False
		MsgBox(Err.Description)
	End Function
	
	Public Function Save(ByRef pGroupListLinkData As CGroupListLinkData, ByRef pbNew As Boolean) As Boolean
		On Error GoTo EH
		'
		Dim rsGroupListLink As ADODB.Recordset
		'
		rsGroupListLink = New ADODB.Recordset
		'
		If pbNew Then
			pGroupListLinkData.ID = NextID("ID", "TGroupListLink", cnMain)
			rsGroupListLink.Open("SELECT * FROM TGroupListLink WHERE ID = -1", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
			rsGroupListLink.AddNew()
			rsGroupListLink.Fields("ID").Value = pGroupListLinkData.ID
		Else
			'
			rsGroupListLink.Open("SELECT * FROM TGroupListLink WHERE ID =" & pGroupListLinkData.ID, cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
			'
			If rsGroupListLink.eof Then
				Save = False
				Exit Function
			End If
		End If
		'
		With rsGroupListLink
			.Fields("ContactID").Value = pGroupListLinkData.ContactID
			.Fields("ListID").Value = pGroupListLinkData.ListID
			.Update()
		End With
		'
		rsGroupListLink.Close()
		'UPGRADE_NOTE: Object rsGroupListLink may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsGroupListLink = Nothing
		'
		Save = True
		'
		Exit Function
EH: 
		Save = False
		If Err.Number = 3022 Then 'Duplicate GroupListLink name record
			MsgBox("Duplicate Group name detected. You must enter a name that is unique.", MsgBoxStyle.Information, "Save Group")
		Else
			MsgBox(Err.Description)
		End If
	End Function
	
	'Public Sub LoadCollection(ByRef pGroupListLinks As CGroupListLinkDatas)
	'  Dim rslist As New ADODB.Recordset
	'  '
	'  Dim GroupListLinkData As CGroupListLinkData
	'  '
	'  Dim sQuery As String
	'  '
	'  Set pGroupListLinks = New CGroupListLinks
	'  '
	'  sQuery = "SELECT TGroupListLink.ID, TGroupListLink.ListName " & _
	''            "FROM TGroupListLink ORDER BY TGroupListLink.ListName"
	'  '
	'  rslist.Open sQuery, cnMain, adOpenForwardOnly, adLockReadOnly
	'  '
	'  While Not rslist.eof
	'    With rslist
	'        Set GroupListLinkData = New CGroupListLinkData
	'        '
	'        GroupListLinkData.ID = rslist!ID
	'        GroupListLinkData.ContactID = rslist!ContactID
	'        GroupListLinkData.ListID = rslist!ListID
	'        '
	'        pGroupListLinks.Add GroupListLinkData
	'        '
	'        rslist.MoveNext
	'      End With
	'    Wend
	'    '
	'    rslist.Close
	'  '
	'  Set rslist = Nothing
	'  Set GroupListLinkData = Nothing
	'  '
	'End Sub
End Class