VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function LoadCollection(ByRef pEmployees As CEmployees) As Boolean
  LoadCollection = False
  Dim rslist As New ADODB.Recordset
  Dim EmployeeData As CEmployeeData
  '
  rslist.Open "SELECT * FROM tblEmployees ORDER BY EmployeeID", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  If Not (rslist.BOF And rslist.eof) Then
    While Not rslist.eof
      With rslist
        Set EmployeeData = New CEmployeeData
        '
        EmployeeData.EmployeeID = nnNum(!EmployeeID)
        EmployeeData.EmployeeNumber = !EmployeeNumber & vbNullString
        EmployeeData.EmployeeLast = !EmployeeLast & vbNullString
        EmployeeData.EmployeeFirst = !EmployeeFirst & vbNullString
        EmployeeData.EmployeeMiddle = !EmployeeMiddle & vbNullString
        EmployeeData.Password = !Password & vbNullString
        EmployeeData.Groups = nnNum(!Groups)
        EmployeeData.SecurityLevel = nnNum(!SecurityLevel)
        EmployeeData.EmployeeExt = nnNum(!EmployeeExt)
        EmployeeData.WorkGroups = nnNum(!WorkGroups)
        EmployeeData.EMailAddress = !EMailAddress & vbNullString
        '
        pEmployees.Add EmployeeData, EmployeeData.EmployeeID
        '
        rslist.MoveNext
      End With
    Wend
    '
    rslist.Close
    LoadCollection = True
  End If
  Set rslist = Nothing
  Set EmployeeData = Nothing
  '
End Function

Public Function Load(ByRef pEmployeeData As CEmployeeData, pID As Integer) As Boolean
  Load = False
  Dim rslist As New ADODB.Recordset
  '
  rslist.Open "SELECT * FROM tblEmployees WHERE EmployeeID = " & pID, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  If Not (rslist.BOF And rslist.eof) Then
    With rslist
      '
      pEmployeeData.EmployeeID = nnNum(!EmployeeID)
      pEmployeeData.EmployeeNumber = !EmployeeNumber & vbNullString
      pEmployeeData.EmployeeLast = !EmployeeLast & vbNullString
      pEmployeeData.EmployeeFirst = !EmployeeFirst & vbNullString
      pEmployeeData.EmployeeMiddle = !EmployeeMiddle & vbNullString
      pEmployeeData.Password = !Password & vbNullString
      pEmployeeData.Groups = nnNum(!Groups)
      pEmployeeData.SecurityLevel = nnNum(!SecurityLevel)
      pEmployeeData.EmployeeExt = nnNum(!EmployeeExt)
      pEmployeeData.WorkGroups = nnNum(!WorkGroups)
      pEmployeeData.EMailAddress = !EMailAddress & vbNullString
      '
    End With
    '
    rslist.Close
    Load = True
  End If
  Set rslist = Nothing
  '
End Function

Public Function Save(ByRef pEmployeeData As CEmployeeData, pID As Integer) As Boolean
  Save = False
  Dim rslist As New ADODB.Recordset
  '
  rslist.Open "SELECT * FROM tblEmployees WHERE EmployeeID = " & pID, cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  If Not (rslist.BOF And rslist.eof) Then
    With rslist
      '
      '!EmployeeID = pEmployeeData.EmployeeID
      '!EmployeeNumber = pEmployeeData.EmployeeNumber
      !EmployeeLast = pEmployeeData.EmployeeLast
      !EmployeeFirst = pEmployeeData.EmployeeFirst
      !EmployeeMiddle = pEmployeeData.EmployeeMiddle
      !Password = EncryptStr(pEmployeeData.Password)
      !Groups = pEmployeeData.Groups
      !SecurityLevel = pEmployeeData.SecurityLevel
      !EmployeeExt = pEmployeeData.EmployeeExt
      !WorkGroups = pEmployeeData.WorkGroups
      !EMailAddress = pEmployeeData.EMailAddress
      
      rslist.UpdateBatch
      '
    End With
    '
    rslist.Close
    Save = True
    'BONUS: Change Password
    cnMain.Execute "EXEC sp_password NULL, '" & Rot39(pEmployeeData.Password) & "', '" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "'"
  End If
  Set rslist = Nothing
  '
End Function

Public Function AddNew(ByRef pEmployeeData As CEmployeeData) As Boolean
  AddNew = False
  Dim rslist As New ADODB.Recordset
  Dim iNewID As Integer
  Dim sTempPass As String
  '
  rslist.Open "SELECT * FROM tblEmployees ORDER BY EmployeeID", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  If Not (rslist.BOF And rslist.eof) Then
    With rslist
      .MoveLast
      iNewID = (!EmployeeID + 1)
      .AddNew
      '
      !EmployeeID = iNewID
      '!EmployeeNumber = pEmployeeData.EmployeeNumber
      !EmployeeLast = pEmployeeData.EmployeeLast
      !EmployeeFirst = pEmployeeData.EmployeeFirst
      !EmployeeMiddle = pEmployeeData.EmployeeMiddle
      !Password = EncryptStr(pEmployeeData.Password)
      !Groups = pEmployeeData.Groups
      !SecurityLevel = pEmployeeData.SecurityLevel
      !EmployeeExt = pEmployeeData.EmployeeExt
      !WorkGroups = pEmployeeData.WorkGroups
      !EMailAddress = pEmployeeData.EMailAddress
      rslist.UpdateBatch
      '
    End With
    '
    rslist.Close
    AddNew = True
    '
    sTempPass = Rot39(pEmployeeData.Password)
    '
    'Create Security Login
    cnMain.Execute "EXEC sp_addlogin  '" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "', '" & sTempPass & "', 'BNB_DATA'"
    'Give access to DB. Current one I guess.
    cnMain.Execute "EXEC sp_grantdbaccess N'" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "', N'" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "'"
    'Assign "User" role.
    cnMain.Execute "EXEC sp_addrolemember N'User', N'" & pEmployeeData.EmployeeFirst & pEmployeeData.EmployeeLast & "'"
    'BONUS: Change Password
    'cnMain.Execute "EXEC sp_password NULL, 'gnarly', ' & sTempPass & '"
    '
  End If
  Set rslist = Nothing

End Function

Public Function Delete(pID As Integer) As Boolean
  Delete = False
  Dim rslist As New ADODB.Recordset
  Dim sTempName As String
  '
  rslist.Open "SELECT * FROM tblEmployees WHERE EmployeeID = " & pID, cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  If Not (rslist.BOF And rslist.eof) Then
    '
    sTempName = rslist!EmployeeFirst & rslist!EmployeeLast
    '
    rslist.Delete
    rslist.UpdateBatch
    '
    rslist.Close
    '
    cnMain.Execute "EXEC sp_dropuser   '" & sTempName & "'"
    cnMain.Execute "EXEC sp_droplogin   '" & sTempName & "'"
    '
    Delete = True
  End If
  Set rslist = Nothing
  '
End Function

Public Function InGroup(sName As String, sWorkGroup As String) As Boolean
  InGroup = False
  Dim rslist As New ADODB.Recordset
  Dim EmployeeData As CEmployeeData
  Dim iWorkGroupNum As Integer
  '
  rslist.Open "select * from tblEmployees", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rslist
    Do While Not .eof
      If LCase(sName) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
        iWorkGroupNum = nnNum(!WorkGroups)
      End If
      .MoveNext
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
    rslist.Close
  Set rslist = Nothing
  Set EmployeeData = Nothing
  '
End Function

Public Function GetEmployeeID(psName As String) As Long
  GetEmployeeID = 0
  '
  Dim rslist As New ADODB.Recordset
  Dim EmployeeData As CEmployeeData
  '
  rslist.Open "select * from tblEmployees", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  With rslist
    Do While Not .eof
      If LCase(psName) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
        GetEmployeeID = nnNum(!EmployeeID)
      End If
      .MoveNext
    Loop
  End With
    rslist.Close
  Set rslist = Nothing
  Set EmployeeData = Nothing
  '
  '
End Function

