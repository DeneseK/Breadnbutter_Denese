VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Save(ByRef pEventData As CEventData, pbNew As Boolean) As Boolean
  On Error GoTo EH
  '
  Dim rsEvent As New ADODB.Recordset
  '
  If pbNew Then
    pEventData.RecID = NextID("RecID", "TSupportActs", cnMain)
    'pEventData.DateEntered = Now
    rsEvent.Open "SELECT * FROM TSupportActs WHERE RecID = -1", cnMain, adOpenKeyset, adLockOptimistic
    rsEvent.AddNew
    'rsEvent!RecID = pEventData.RecID
    rsEvent!Date = CDate(Now)
    rsEvent!Time = Format(Now, "hh:nn AM/PM")
  Else
    rsEvent.Open "SELECT * FROM TSupportActs WHERE RecID = " & pEventData.RecID, cnMain, adOpenKeyset, adLockOptimistic
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
  !Date = Format(pEventData.EventDate, "Short Date")  'CDate(Now)
  !Time = pEventData.EventTime 'Format(Now, "hh:nn AM/PM")
  !CustRecID = pEventData.CustRecID
  !Type = pEventData.EventType
  !Results = pEventData.EventResults
  !User = pEventData.EventUser
  !Subject = pEventData.EventSubject
  !ProductID = pEventData.ProductID
  !ClosedTime = pEventData.ClosedTime
  !OpenCall = pEventData.OpenCall
  !Sticky = pEventData.Sticky
  '
  .Update
  '
  pEventData.RecID = rsEvent!RecID
  End With
  '
  rsEvent.Close
  '
  Set rsEvent = Nothing
  '
  Save = True
  '
  Exit Function
EH:
  If rsEvent.State = adStateOpen Then
    rsEvent.CancelUpdate
    rsEvent.Close
  End If
  MsgBox Err.Description
End Function

Public Function Load(ByRef pEventData As CEventData, plID As Long) As Boolean
  On Error GoTo EH
  '
  Dim rsEvent As New Recordset
  '
    Set pEventData = New CEventData
    '
      rsEvent.Open "SELECT * FROM TSupportActs WHERE RecID = " & plID, cnMain, adOpenForwardOnly, adLockReadOnly
      '
      If Not rsEvent.eof Then
        With rsEvent
          pEventData.RecID = !RecID
          pEventData.EventDate = nnNum(!Date)
          pEventData.EventTime = nnNum(!Time)
          pEventData.EventType = !Type & vbNullString
          pEventData.EventResults = !Results & vbNullString
          pEventData.EventUser = !User & vbNullString
          pEventData.EventSubject = !Subject & vbNullString
          pEventData.ProductID = nnNum(!ProductID)
          pEventData.ClosedTime = nnNum(!ClosedTime)
          pEventData.OpenCall = nnNum(!OpenCall)
          pEventData.Sticky = nnNum(!Sticky)
        End With
        '
        Load = True
      Else
        Load = False
      End If
      '
      rsEvent.Close
      '
      Set rsEvent = Nothing
  Exit Function
EH:
  MsgBox Err.Description & " in Class Event: Load."
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

Public Sub LoadCollection(ByVal plContactID As Long, ByRef pEvents As CEvents)

  Dim rslist As New ADODB.Recordset
  Dim EventData As CEventData
  '
  Dim sQuery As String
  '
  sQuery = "SELECT * FROM TSupportActs WHERE (CustRecID=" & plContactID & ") ORDER BY Date DESC, Time DESC"
  '
  rslist.Open sQuery, cnMain, adOpenForwardOnly, adLockReadOnly
  '
  While Not rslist.eof
    With rslist
        Set EventData = New CEventData
        '
        Load EventData, rslist!RecID
        '
        pEvents.Add EventData
        '
        rslist.MoveNext
      End With
    Wend
    '
    rslist.Close
  '
  Set rslist = Nothing
  Set EventData = Nothing
  '
End Sub



