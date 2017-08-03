VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroupListLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Delete(ByVal plID As Long)
  
  If InputBox("Type DELETE if you sure you want to delete this Custom Group.", "Delete Contact") = "DELETE" Then
    
    '* TODO Delete Contacts
    
    cnMain.Execute "DELETE FROM TGroupListLink WHERE ID = " & plID
    '
  End If
  
End Function

Public Function CheckContact(ByRef plListID As Long, plContactID As Long) As Boolean
    Dim rsGroupListLink As New Recordset
    '
    rsGroupListLink.Open "SELECT * FROM TGroupListLink WHERE ContactID = " & plContactID & " AND ListID = " & plListID, cnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsGroupListLink.eof Then
      '
      CheckContact = True
    Else
      CheckContact = False
    End If
    '
    rsGroupListLink.Close
    Set rsGroupListLink = Nothing
End Function

Public Function Load(ByRef pGroupListLinkData As CGroupListLinkData, plID As Long) As Boolean
    Dim rsGroupListLink As New Recordset
    '
    Set pGroupListLinkData = New CGroupListLinkData
    '
    rsGroupListLink.Open "SELECT * FROM TGroupListLink WHERE ID = " & plID, cnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsGroupListLink.eof Then
      Set pGroupListLinkData = New CGroupListLinkData
      With rsGroupListLink
      pGroupListLinkData.ID = !ID
      pGroupListLinkData.ContactID = !ContactID
      pGroupListLinkData.ListID = !ListID
      '
      End With
      '
      Load = True
    Else
      Load = False
    End If
    '
    rsGroupListLink.Close
    Set rsGroupListLink = Nothing
End Function

Public Function AddContact(ByVal plContact As Long, plList As Long)
  On Error GoTo EH
  '
  Dim rsGroupListLink As Recordset
  '
  Set rsGroupListLink = New ADODB.Recordset
  '
  rsGroupListLink.Open "SELECT * FROM TGroupListLink WHERE ID = -1", cnMain, adOpenDynamic, adLockOptimistic, adCmdText
  rsGroupListLink.AddNew
  '
  With rsGroupListLink
  !ID = NextID("ID", "TGroupListLink", cnMain)
  !ContactID = plContact
  !ListID = plList
  .Update
  End With
  '
  rsGroupListLink.Close
  Set rsGroupListLink = Nothing
  '
  AddContact = True
  '
  Exit Function
EH:
  AddContact = False
  MsgBox Err.Description
End Function

Public Function DelContact(ByVal plContact As Long, plList As Long)
  On Error GoTo EH
  '
  Dim rsGroupListLink As Recordset
  '
  Set rsGroupListLink = New ADODB.Recordset
  '
  rsGroupListLink.Open "SELECT * FROM TGroupListLink WHERE ContactID = " & plContact & " AND ListID = " & plList, cnMain, adOpenForwardOnly, adLockReadOnly
  If Not rsGroupListLink.eof Then
    '
    With rsGroupListLink
    cnMain.Execute "DELETE FROM TGroupListLink WHERE ID = " & !ID
    End With
    '
  End If
  rsGroupListLink.Close
  Set rsGroupListLink = Nothing
  '
  DelContact = True
  '
  Exit Function
EH:
  DelContact = False
  MsgBox Err.Description
End Function

Public Function Save(ByRef pGroupListLinkData As CGroupListLinkData, pbNew As Boolean) As Boolean
  On Error GoTo EH
  '
  Dim rsGroupListLink As Recordset
  '
  Set rsGroupListLink = New ADODB.Recordset
  '
  If pbNew Then
    pGroupListLinkData.ID = NextID("ID", "TGroupListLink", cnMain)
    rsGroupListLink.Open "SELECT * FROM TGroupListLink WHERE ID = -1", cnMain, adOpenDynamic, adLockOptimistic, adCmdText
    rsGroupListLink.AddNew
    rsGroupListLink!ID = pGroupListLinkData.ID
  Else
    '
    rsGroupListLink.Open "SELECT * FROM TGroupListLink WHERE ID =" & pGroupListLinkData.ID, cnMain, adOpenDynamic, adLockOptimistic, adCmdText
    '
    If rsGroupListLink.eof Then
      Save = False
      Exit Function
    End If
  End If
  '
  With rsGroupListLink
  !ContactID = pGroupListLinkData.ContactID
  !ListID = pGroupListLinkData.ListID
  .Update
  End With
  '
  rsGroupListLink.Close
  Set rsGroupListLink = Nothing
  '
  Save = True
  '
  Exit Function
EH:
  Save = False
  If Err.Number = 3022 Then 'Duplicate GroupListLink name record
    MsgBox "Duplicate Group name detected. You must enter a name that is unique.", vbInformation, "Save Group"
  Else
    MsgBox Err.Description
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
'            "FROM TGroupListLink ORDER BY TGroupListLink.ListName"
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


