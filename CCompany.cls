VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Delete(ByVal plID As Long)
  
  If InputBox("Type DELETE if you sure you want to delete this company.", "Delete Contact") = "DELETE" Then
    
    '* TODO Delete Contacts
    
    cnMain.Execute "DELETE FROM TCompany WHERE ID = " & plID
    '
  End If
  
End Function

Public Function Load(ByRef pCompanyData As CCompanyData, plID As Long) As Boolean
    Dim rsCompany As New Recordset
    '
    Set pCompanyData = New CCompanyData
    '
    rsCompany.Open "SELECT * FROM TCompany WHERE ID = " & plID, cnMain, adOpenForwardOnly, adLockReadOnly
    If Not rsCompany.eof Then
      Set pCompanyData = New CCompanyData
      With rsCompany
      pCompanyData.ID = !ID
      'pCompanyData.DateEntered = nnNum(!DateEntered)
      'pCompanyData.LastUpdate = nnNum(!LastUpdate)
      pCompanyData.Name = !Name & vbNullString
      'pCompanyData.Division = !Division & vbNullString
      pCompanyData.Individual = nnNum(!Individual)
      pCompanyData.DoNotContact = nnNum(!DoNotContact)
      pCompanyData.Note = !Note & vbNullString
      pCompanyData.InterestRank = nnNum(!InterestRank)
      '
      End With
      '
      Load = True
    Else
      Load = False
    End If
    '
    rsCompany.Close
    Set rsCompany = Nothing
End Function

Public Function Save(ByRef pCompanyData As CCompanyData, pbNew As Boolean) As Boolean
  On Error GoTo EH
  '
  Dim rsCompany As Recordset
  '
  Set rsCompany = New ADODB.Recordset
  '
  If pbNew Then
    pCompanyData.ID = NextID("ID", "TCompany", cnMain)
    rsCompany.Open "SELECT * FROM TCompany WHERE ID = -1", cnMain, adOpenDynamic, adLockOptimistic, adCmdText
    rsCompany.AddNew
    rsCompany!ID = pCompanyData.ID
    rsCompany!DateEntered = CDate(Now) 'pCompanyData.DateEntered
  Else
    '
    rsCompany.Open "SELECT * FROM TCompany WHERE ID =" & pCompanyData.ID, cnMain, adOpenDynamic, adLockOptimistic, adCmdText
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
  !Name = pCompanyData.Name
  !LastUpdate = Now()
  '!Division = pCompanyData.Division
  !Individual = pCompanyData.Individual
  !DoNotContact = pCompanyData.DoNotContact
  !Note = pCompanyData.Note
  !InterestRank = pCompanyData.InterestRank
  .Update
  End With
  '
  rsCompany.Close
  Set rsCompany = Nothing
  '
  Save = True
  'fAdding = False
  '
  Exit Function
EH:
  Save = False
  If Err.Number = 3022 Then 'Duplicate company name record
    MsgBox "Duplicate company name detected. You must enter a name and division that is unique.", vbInformation, "Save Company"
  Else
    MsgBox Err.Description
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
'          "WHERE (((TContact.Status) Like '" & SearchCriteria & "') AND ((TCompany.ID)=" & CR.ID & "));", cnMain
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
'            "TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & _
'            "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
'            "WHERE (((TCOMPANY.ID)=" & Company.ID & "))" 'ORDER BY TContact.LastName, TContact.FirstName"
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
'                "TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as [Days], " & _
'                "TContact.AuthDate + TContact.AuthDays AS [ExpirationDate], " & _
'                "TContact.Status, TContact.Title, TContact.ShipStatus, TContact.VersionShipped " & _
'                "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
'                "WHERE (((TCOMPANY.ID)=" & Company.ID & ")) ORDER BY TContact.LastName, TContact.FirstName"
'  Else
'     sQuery = "SELECT TContact.ID, TContact.ContactType, TContact.Status, TContact.FirstName, TContact.LastName, " & _
'                "TContact.AuthDate + TContact.AuthDays AS [ExpirationDate], " & _
'                "TContact.Status, TContact.Title, TContact.ShipStatus, TContact.VersionShipped " & _
'                "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
'                "WHERE (((TCOMPANY.ID)=" & Company.ID & ")) ORDER BY TContact.LastName, TContact.FirstName"
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
  Set pCompanys = New CCompanys
  '
  sQuery = "SELECT TCompany.ID, TCompany.Name " & _
            "FROM TCompany ORDER BY TCompany.Name"
  '
  rslist.Open sQuery, cnMain, adOpenForwardOnly, adLockReadOnly
  '
  While Not rslist.eof
    With rslist
        Set CompanyData = New CCompanyData
        '
        CompanyData.ID = rslist!ID
        CompanyData.Name = rslist!Name & vbNullString
        '
        pCompanys.Add CompanyData
        '
        rslist.MoveNext
      End With
    Wend
    '
    rslist.Close
  '
  Set rslist = Nothing
  Set CompanyData = Nothing
  '
End Sub

