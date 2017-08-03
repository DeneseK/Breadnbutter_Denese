VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroupList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Clear(ByVal plID As Long)
    cnMain.Execute "DELETE FROM TGroupListLink WHERE ListID = " & plID
End Function

Public Function Delete(ByVal plID As Long)
  
  If InputBox("Type DELETE if you sure you want to delete this Custom Group.", "Delete Contact") = "DELETE" Then
    Clear plID
    '
    cnMain.Execute "DELETE FROM TGroupList WHERE ID = " & plID
    '
  End If
  
End Function

Public Function Load(ByRef pGroupListData As CGroupListData, plID As Long) As Boolean
    Dim rsGroupList As New Recordset
    '
    Set pGroupListData = New CGroupListData
    '
    rsGroupList.Open "SELECT * FROM TGroupList WHERE ID = " & plID, cnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsGroupList.eof Then
      Set pGroupListData = New CGroupListData
      With rsGroupList
      pGroupListData.ID = !ID
      pGroupListData.ListName = !ListName & vbNullString
      pGroupListData.EmployeeID = !EmployeeID
      '
      End With
      '
      Load = True
    Else
      Load = False
    End If
    '
    rsGroupList.Close
    Set rsGroupList = Nothing
End Function

Public Function Save(ByRef pGroupListData As CGroupListData, pbNew As Boolean) As Boolean
  On Error GoTo EH
  '
  Dim rsGroupList As Recordset
  '
  Set rsGroupList = New ADODB.Recordset
  '
  If pbNew Then
    pGroupListData.ID = NextID("ID", "TGroupList", cnMain)
    rsGroupList.Open "SELECT * FROM TGroupList WHERE ID = -1", cnMain, adOpenDynamic, adLockOptimistic, adCmdText
    rsGroupList.AddNew
    rsGroupList!ID = pGroupListData.ID
  Else
    '
    rsGroupList.Open "SELECT * FROM TGroupList WHERE ID =" & pGroupListData.ID, cnMain, adOpenDynamic, adLockOptimistic, adCmdText
    '
    If rsGroupList.eof Then
      Save = False
      Exit Function
    End If
  End If
  '
  rsGroupList!ListName = pGroupListData.ListName
  rsGroupList!EmployeeID = pGroupListData.EmployeeID
  rsGroupList.Update
  '
  rsGroupList.Close
  Set rsGroupList = Nothing
  '
  Save = True
  '
  Exit Function
EH:
  Save = False
  If Err.Number = 3022 Then 'Duplicate GroupList name record
    MsgBox "Duplicate Group name detected. You must enter a name that is unique.", vbInformation, "Save Group"
  Else
    MsgBox Err.Description
  End If
End Function

Public Sub LoadCollection(ByRef pGroupLists As CGroupListDatas)
  Dim rslist As New ADODB.Recordset
  '
  Dim GroupListData As CGroupListData
  '
  Dim sQuery As String
  '
  Set pGroupLists = New CGroupListDatas
  '
  sQuery = "SELECT * FROM TGroupList ORDER BY TGroupList.ListName"
  '
  rslist.Open sQuery, cnMain, adOpenForwardOnly, adLockReadOnly
  '
  While Not rslist.eof
    With rslist
        Set GroupListData = New CGroupListData
        '
        GroupListData.ID = rslist!ID
        GroupListData.ListName = rslist!ListName & vbNullString
        GroupListData.EmployeeID = rslist!EmployeeID
        '
        pGroupLists.Add GroupListData
        '
        rslist.MoveNext
      End With
    Wend
    '
    rslist.Close
  '
  Set rslist = Nothing
  Set GroupListData = Nothing
  '
End Sub



