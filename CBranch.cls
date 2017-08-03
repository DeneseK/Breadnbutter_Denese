VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub LoadCollection(ByVal plCompanyID As Long, ByRef pBranchs As CBranchs)

  Dim rslist As New ADODB.Recordset
  Dim BranchData As CBranchData
  '
  Set pBranchs = New CBranchs
  '
  rslist.Open "SELECT * FROM TBranch WHERE CompanyID = " & plCompanyID & " ORDER BY Name", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText

  'Set rsList = dbPropertyValuation.OpenRecordset("SELECT * FROM TAttachedGarage" & _
    " ORDER BY SquareFoot", dbOpenForwardOnly)
  '
    While Not rslist.eof
      With rslist
        Set BranchData = New CBranchData
        '
        BranchData.BranchID = nnNum(!BranchID)
        BranchData.CompanyID = nnNum(!CompanyID)
        BranchData.Name = !Name & vbNullString
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
        pBranchs.Add BranchData
        '
        rslist.MoveNext
      End With
    Wend
    '
    rslist.Close
  '
  Set rslist = Nothing
  Set BranchData = Nothing
  '
End Sub

Public Function Load(ByRef pBranchData As CBranchData, plBranchID As Long)
  Dim rs As New ADODB.Recordset
  '
  Set pBranchData = New CBranchData
  '
  rs.Open "SELECT * FROM TBranch WHERE BranchID = " & plBranchID, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  If Not rs.eof Then
    With rs
        pBranchData.BranchID = nnNum(!BranchID)
        pBranchData.CompanyID = nnNum(!CompanyID)
        pBranchData.Name = !Name & vbNullString
        pBranchData.Address1 = !Address1 & vbNullString
        pBranchData.Address2 = !Address2 & vbNullString
        pBranchData.Address3 = !Address3 & vbNullString
        pBranchData.City = !City & vbNullString
        pBranchData.Email = !Email & vbNullString
        pBranchData.FaxNumber = !FaxNumber & vbNullString
        pBranchData.ManagerFirstName = !ManagerFirstName & vbNullString
        pBranchData.ManagerLastName = !ManagerLastName & vbNullString
        pBranchData.Number = !Number & vbNullString
        pBranchData.PhoneNumber = !PhoneNumber & vbNullString
        pBranchData.State = !State & vbNullString
        pBranchData.Zip = !Zip & vbNullString
    End With
  End If
  '
  rs.Close
  Set rs = Nothing
  '
End Function

Public Function Save(ByRef pBranchData As CBranchData, pbNew As Boolean) As Boolean
  On Error GoTo EH
  '
  Dim rsBranch As New ADODB.Recordset
  '
  If pbNew Then
    pBranchData.BranchID = NextID("BranchID", "TBranch", cnMain)
    rsBranch.Open "SELECT * FROM TBranch WHERE BranchID = -1", cnMain, adOpenKeyset, adLockOptimistic
    rsBranch.AddNew
    rsBranch!BranchID = pBranchData.BranchID
    rsBranch!DateEntered = CDate(Now)
  Else
    rsBranch.Open "SELECT * FROM TBranch WHERE BranchID = " & pBranchData.BranchID, cnMain, adOpenKeyset, adLockOptimistic
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
      !CompanyID = pBranchData.CompanyID
      !Name = pBranchData.Name
      !Address1 = pBranchData.Address1
      !Address2 = pBranchData.Address2
      !Address3 = pBranchData.Address3
      !City = pBranchData.City
      !Email = pBranchData.Email
      !FaxNumber = pBranchData.FaxNumber
      !ManagerFirstName = pBranchData.ManagerFirstName
      !ManagerLastName = pBranchData.ManagerLastName
      !Number = pBranchData.Number
      !PhoneNumber = pBranchData.PhoneNumber
      !State = pBranchData.State
      !Zip = pBranchData.Zip
      '
      .UpdateBatch
      '
    End With
    '
    rsBranch.Close
    Save = True
    '
    End If
  Set rsBranch = Nothing
  '
  Exit Function
EH:
  If Not rsBranch Is Nothing Then
    If rsBranch.State = adStateOpen Then
      rsBranch.CancelUpdate
      rsBranch.Close
    End If
  End If
  MsgBox Err.Description
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

Public Function Delete(pID As Long) As Boolean
  Delete = False
  '
  Dim rslist As New ADODB.Recordset
  Dim rsContactList As New Recordset
  '
  rsContactList.Open "SELECT ID FROM TContact WHERE BranchID = " & pID, cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  If rsContactList.RecordCount > 0 Then
    Delete = False
    '
    Exit Function
  Else
    '
    rslist.Open "SELECT * FROM TBranch WHERE BranchID = " & pID, cnMain, adOpenKeyset, adLockBatchOptimistic
    '
    If Not (rslist.BOF And rslist.eof) Then
      '
      rslist.Delete
      rslist.UpdateBatch
      '
      rslist.Close
      '
      Delete = True
    End If
    
    Set rslist = Nothing
  End If
End Function
