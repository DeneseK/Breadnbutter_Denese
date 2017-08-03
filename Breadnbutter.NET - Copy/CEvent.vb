Option Strict Off
Option Explicit On
Friend Class CEvent
	Public Function Save(ByRef pEventData As CEventData, ByRef pbNew As Boolean) As Boolean
		On Error GoTo EH
		'
		Dim rsEvent As New ADODB.Recordset
		'
		If pbNew Then
			pEventData.RecID = NextID("RecID", "TSupportActs", cnMain)
			'pEventData.DateEntered = Now
			rsEvent.Open("SELECT * FROM TSupportActs WHERE RecID = -1", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			rsEvent.AddNew()
			'rsEvent!RecID = pEventData.RecID
			rsEvent.Fields("Date").Value = CDate(Now)
			rsEvent.Fields("Time").Value = VB6.Format(Now, "hh:nn AM/PM")
		Else
			rsEvent.Open("SELECT * FROM TSupportActs WHERE RecID = " & pEventData.RecID, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
			'
			If rsEvent.eof Then
				Exit Function
				'Else
				'rsEvent.Edit
			End If
		End If
		'
		With rsEvent
			'!RecID = pEventData.CompanyID
			.Fields("Date").Value = VB6.Format(pEventData.EventDate, "Short Date") 'CDate(Now)
			.Fields("Time").Value = pEventData.EventTime 'Format(Now, "hh:nn AM/PM")
			.Fields("CustRecID").Value = pEventData.CustRecID
			.Fields("Type").Value = pEventData.EventType
			.Fields("Results").Value = pEventData.EventResults
			.Fields("User").Value = pEventData.EventUser
			.Fields("Subject").Value = pEventData.EventSubject
			.Fields("ProductID").Value = pEventData.ProductID
			.Fields("ClosedTime").Value = pEventData.ClosedTime
			.Fields("OpenCall").Value = pEventData.OpenCall
			.Fields("Sticky").Value = pEventData.Sticky
			'
			.Update()
			'
			pEventData.RecID = rsEvent.Fields("RecID").Value
		End With
		'
		rsEvent.Close()
		'
		'UPGRADE_NOTE: Object rsEvent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsEvent = Nothing
		'
		Save = True
		'
		Exit Function
EH: 
		If rsEvent.State = ADODB.ObjectStateEnum.adStateOpen Then
			rsEvent.CancelUpdate()
			rsEvent.Close()
		End If
		MsgBox(Err.Description)
	End Function
	
	Public Function Load(ByRef pEventData As CEventData, ByRef plID As Integer) As Boolean
		On Error GoTo EH
		'
		Dim rsEvent As New ADODB.Recordset
		'
		pEventData = New CEventData
		'
		rsEvent.Open("SELECT * FROM TSupportActs WHERE RecID = " & plID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		'
		If Not rsEvent.eof Then
			With rsEvent
				pEventData.RecID = .Fields("RecID").Value
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEventData.EventDate = nnNum(.Fields("Date"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEventData.EventTime = nnNum(.Fields("Time"))
				pEventData.EventType = .Fields("Type").Value & vbNullString
				pEventData.EventResults = .Fields("Results").Value & vbNullString
				pEventData.EventUser = .Fields("User").Value & vbNullString
				pEventData.EventSubject = .Fields("Subject").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEventData.ProductID = nnNum(.Fields("ProductID"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEventData.ClosedTime = nnNum(.Fields("ClosedTime"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEventData.OpenCall = nnNum(.Fields("OpenCall"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pEventData.Sticky = nnNum(.Fields("Sticky"))
			End With
			'
			Load = True
		Else
			Load = False
		End If
		'
		rsEvent.Close()
		'
		'UPGRADE_NOTE: Object rsEvent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsEvent = Nothing
		Exit Function
EH: 
		MsgBox(Err.Description & " in Class Event: Load.")
	End Function
	
	'Public Function Delete(ByVal plID As Long) As Boolean
	'
	'  If InputBox("Type DELETE and click OK if you really want to delete this contact.", "Delete Contact") = "DELETE" Then
	'    cnMain.Execute "DELETE FROM TSupportActs WHERE CustRecID = " & plID, , adCmdText
	'    cnMain.Execute "DELETE FROM TContact WHERE ID = " & plID, , adCmdText
	'    Delete = True
	'  Else
	'    Delete = False
	'  End If
	'
	'End Function
	
	Public Sub LoadCollection(ByVal plContactID As Integer, ByRef pEvents As CEvents)
		
		Dim rslist As New ADODB.Recordset
		Dim EventData As CEventData
		'
		Dim sQuery As String
		'
		sQuery = "SELECT * FROM TSupportActs WHERE (CustRecID=" & plContactID & ") ORDER BY Date DESC, Time DESC"
		'
		rslist.Open(sQuery, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		'
		While Not rslist.eof
			With rslist
				EventData = New CEventData
				'
				Load(EventData, rslist.Fields("RecID").Value)
				'
				pEvents.Add(EventData)
				'
				rslist.MoveNext()
			End With
		End While
		'
		rslist.Close()
		'
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object EventData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventData = Nothing
		'
	End Sub
End Class