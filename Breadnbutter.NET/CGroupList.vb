Option Strict Off
Option Explicit On
Friend Class CGroupList
	
	Public Function Clear(ByVal plID As Integer) As Object
		cnMain.Execute("DELETE FROM TGroupListLink WHERE ListID = " & plID)
	End Function
	
	Public Function Delete(ByVal plID As Integer) As Object
		
		If InputBox("Type DELETE if you sure you want to delete this Custom Group.", "Delete Contact") = "DELETE" Then
			Clear(plID)
			'
			cnMain.Execute("DELETE FROM TGroupList WHERE ID = " & plID)
			'
		End If
		
	End Function
	
	Public Function Load(ByRef pGroupListData As CGroupListData, ByRef plID As Integer) As Boolean
		Dim rsGroupList As New ADODB.Recordset
		'
		pGroupListData = New CGroupListData
		'
		rsGroupList.Open("SELECT * FROM TGroupList WHERE ID = " & plID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		If Not rsGroupList.eof Then
			pGroupListData = New CGroupListData
			With rsGroupList
				pGroupListData.ID = .Fields("ID").Value
				pGroupListData.ListName = .Fields("ListName").Value & vbNullString
				pGroupListData.EmployeeID = .Fields("EmployeeID").Value
				'
			End With
			'
			Load = True
		Else
			Load = False
		End If
		'
		rsGroupList.Close()
		'UPGRADE_NOTE: Object rsGroupList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsGroupList = Nothing
	End Function
	
	Public Function Save(ByRef pGroupListData As CGroupListData, ByRef pbNew As Boolean) As Boolean
		On Error GoTo EH
		'
		Dim rsGroupList As ADODB.Recordset
		'
		rsGroupList = New ADODB.Recordset
		'
		If pbNew Then
			pGroupListData.ID = NextID("ID", "TGroupList", cnMain)
			rsGroupList.Open("SELECT * FROM TGroupList WHERE ID = -1", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
			rsGroupList.AddNew()
			rsGroupList.Fields("ID").Value = pGroupListData.ID
		Else
			'
			rsGroupList.Open("SELECT * FROM TGroupList WHERE ID =" & pGroupListData.ID, cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
			'
			If rsGroupList.eof Then
				Save = False
				Exit Function
			End If
		End If
		'
		rsGroupList.Fields("ListName").Value = pGroupListData.ListName
		rsGroupList.Fields("EmployeeID").Value = pGroupListData.EmployeeID
		rsGroupList.Update()
		'
		rsGroupList.Close()
		'UPGRADE_NOTE: Object rsGroupList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsGroupList = Nothing
		'
		Save = True
		'
		Exit Function
EH: 
		Save = False
		If Err.Number = 3022 Then 'Duplicate GroupList name record
			MsgBox("Duplicate Group name detected. You must enter a name that is unique.", MsgBoxStyle.Information, "Save Group")
		Else
			MsgBox(Err.Description)
		End If
	End Function
	
	Public Sub LoadCollection(ByRef pGroupLists As CGroupListDatas)
		Dim rslist As New ADODB.Recordset
		'
		Dim GroupListData As CGroupListData
		'
		Dim sQuery As String
		'
		pGroupLists = New CGroupListDatas
		'
		sQuery = "SELECT * FROM TGroupList ORDER BY TGroupList.ListName"
		'
		rslist.Open(sQuery, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		'
		While Not rslist.eof
			With rslist
				GroupListData = New CGroupListData
				'
				GroupListData.ID = rslist.Fields("ID").Value
				GroupListData.ListName = rslist.Fields("ListName").Value & vbNullString
				GroupListData.EmployeeID = rslist.Fields("EmployeeID").Value
				'
				pGroupLists.Add(GroupListData)
				'
				rslist.MoveNext()
			End With
		End While
		'
		rslist.Close()
		'
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object GroupListData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		GroupListData = Nothing
		'
	End Sub
End Class