VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Save(ByRef pContactData As CContactData, pbNew As Boolean) As Boolean
  On Error GoTo EH
  '
  Dim rsContact As New ADODB.Recordset
  '
  If pbNew Then
    pContactData.ID = NextID("ID", "TContact", cnMain)
    'pContactData.DateEntered = Now
    rsContact.Open "SELECT * FROM TContact WHERE ID = -1", cnMain, adOpenKeyset, adLockOptimistic
    rsContact.AddNew
    rsContact!ID = pContactData.ID
    rsContact!DateEntered = CDate(Now) 'pContactData.DateEntered
  Else
    rsContact.Open "SELECT * FROM TContact WHERE ID = " & pContactData.ID, cnMain, adOpenKeyset, adLockOptimistic
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
  !CompanyID = pContactData.CompanyID
  !BranchID = pContactData.BranchID
  !FirstName = pContactData.FirstName
  !LastName = pContactData.LastName
  !Salutation = pContactData.Salutation
  !Title = pContactData.Title
  !Address1 = pContactData.Address1
  !Address2 = pContactData.Address2
  !City = pContactData.City
  !State = pContactData.State
  !Zip = pContactData.Zip
  !PermMailAddress1 = pContactData.MailAddress1
  !PermMailAddress2 = pContactData.MailAddress2
  !PermMailCity = pContactData.MailCity
  !PermMailState = pContactData.MailState
  !PermMailZip = pContactData.MailZip
  '
  !PCEmail = pContactData.PCEmail
  !PCEmailPassword = pContactData.PCEmailPassword
  '
  !Phone1 = pContactData.Phone1
  !Phone2 = pContactData.Phone2
  !Fax = pContactData.Fax
  !Email = pContactData.Email
  !Source = pContactData.Source
  !betatester = IIf(pContactData.Selected = 1, True, False)
  !PreferredAddress = pContactData.PreferredAddress
  !Notes = pContactData.Notes
  '
  !Status = pContactData.Status
  !ShipStatus = pContactData.ShipStatus
  !AuthStatus = pContactData.AuthStatus
  '!DateEntered = pContactData.DateEntered
  !ShipDate = pContactData.ShipDate
  !AuthDate = pContactData.AuthDate
  !AuthDays = pContactData.AuthDays
  !Copies = pContactData.Copies
  !VersionShipped = pContactData.VersionShipped
 ' !AuthRemaining = pContactData.AuthRemaining
  '
  !PVShipStatus = pContactData.PVShipStatus
  !PVAuthStatus = pContactData.PVAuthStatus
  !PVDownloadStatus = pContactData.PVDownloadStatus
  !DownloadStatus = pContactData.DownloadStatus
  '!PVDateEntered = pContactData.PVDateEntered
  !PVShipDate = pContactData.PVShipDate
  !PVDownloadDate = pContactData.PVDownloadDate
  !DownloadDate = pContactData.DownloadDate
  !PVAuthDate = pContactData.PVAuthDate
  !PVAuthDays = pContactData.PVAuthDays
  !PVCopies = pContactData.PVCopies
  !PVVersionShipped = pContactData.PVVersionShipped
  '!PVAuthRemaining = pContactData.PVAuthRemaining
  '
  !Rate = pContactData.Rate
  !ContactType = pContactData.ContactType
  !RateExpDate = pContactData.RateExpDate
  '
  !GraceDays = pContactData.GraceDays
  !OnlineAuths = pContactData.OnlineAuths
  !SaleDate = pContactData.SaleDate
  !SaleDays = pContactData.SaleDays
  '
  !PVGraceDays = pContactData.PVGraceDays
  !PVOnlineAuths = pContactData.PVOnlineAuths
  !PVSaleDate = pContactData.PVSaleDate
  !PVSaleDays = pContactData.PVSaleDays
  '
  !WebPassword = pContactData.WebPassword
  !ContactByEmail = pContactData.ContactByEmail
  !ChangedData = pContactData.ChangedData
  '
  !AdjusterID = pContactData.AdjusterID
  '
  !LastUpdate = Now()
  .Update
  End With
  '
  rsContact.Close
  '
  Set rsContact = Nothing
  '
  Save = True
  'fAdding = False
  '
  Exit Function
EH:
  If Not rsContact Is Nothing Then
    If rsContact.State = adStateOpen Then
      rsContact.CancelUpdate
      rsContact.Close
    End If
  End If
  MsgBox Err.Description
End Function

Public Function Load(ByRef pContactData As CContactData, plID As Long) As Boolean
  On Error GoTo EH
  '
  Dim rsContact As New Recordset
  '
    Set pContactData = New CContactData
    '
      rsContact.Open "SELECT * FROM TContact WHERE ID = " & plID, cnMain, adOpenForwardOnly, adLockReadOnly
      '
      If Not rsContact.eof Then
        With rsContact
        pContactData.ID = !ID
        pContactData.CompanyID = !CompanyID
        pContactData.BranchID = nnNum(!BranchID)
        'pContactData.LastUpdate = nnNum(!LastUpdate)
        pContactData.FirstName = !FirstName & vbNullString
        pContactData.LastName = !LastName & vbNullString
        pContactData.Salutation = !Salutation & vbNullString
        pContactData.Title = !Title & vbNullString
        pContactData.Address1 = !Address1 & vbNullString
        pContactData.Address2 = !Address2 & vbNullString
        pContactData.City = !City & vbNullString
        pContactData.State = !State & vbNullString
        pContactData.Zip = !Zip & vbNullString
        pContactData.MailAddress1 = !PermMailAddress1 & vbNullString
        pContactData.MailAddress2 = !PermMailAddress2 & vbNullString
        pContactData.MailCity = !PermMailCity & vbNullString
        pContactData.MailState = !PermMailState & vbNullString
        pContactData.MailZip = !PermMailZip & vbNullString
        '
        pContactData.PCEmail = !PCEmail & vbNullString
        pContactData.PCEmailPassword = !PCEmailPassword & vbNullString
        '
        pContactData.Phone1 = !Phone1 & vbNullString
        pContactData.Phone2 = !Phone2 & vbNullString
        pContactData.Fax = !Fax & vbNullString
        pContactData.Email = !Email & vbNullString
        pContactData.Source = !Source & vbNullString
        pContactData.Selected = IIf(nnNum(!betatester), 1, 0)
        pContactData.PreferredAddress = nnNum(!PreferredAddress)
        pContactData.Notes = !Notes & vbNullString
        pContactData.Status = !Status & vbNullString
        '
        pContactData.ShipStatus = !ShipStatus & vbNullString
        pContactData.AuthStatus = !AuthStatus & vbNullString
        'pContactData.DateEntered = nnNum(!DateEntered)
        pContactData.ShipDate = nnNum(!ShipDate)
        pContactData.AuthDate = nnNum(!AuthDate)
        pContactData.AuthDays = nnNum(!AuthDays)
        pContactData.Copies = nnNum(!Copies)
        pContactData.VersionShipped = !VersionShipped & vbNullString
        'pContactData.AuthRemaining = nnNum(!AuthRemaining)
        '
        pContactData.PVShipStatus = !PVShipStatus & vbNullString
        pContactData.PVDownloadStatus = !PVDownloadStatus & vbNullString
        pContactData.DownloadStatus = !DownloadStatus & vbNullString
        pContactData.PVAuthStatus = !PVAuthStatus & vbNullString
   '     pContactData.PVDateEntered = nnNum(!PVDateEntered)
        pContactData.PVShipDate = nnNum(!PVShipDate)
        pContactData.PVDownloadDate = nnNum(!PVDownloadDate)
        pContactData.DownloadDate = nnNum(!DownloadDate)
        pContactData.PVAuthDate = nnNum(!PVAuthDate)
        pContactData.PVAuthDays = nnNum(!PVAuthDays)
        pContactData.PVCopies = nnNum(!PVCopies)
        pContactData.PVVersionShipped = !PVVersionShipped & vbNullString
        'pContactData.PVAuthRemaining = nnNum(!PVAuthRemaining)
        '
        pContactData.Rate = nnNum(!Rate)
        pContactData.ContactType = nnNum(!ContactType)
        pContactData.RateExpDate = nnNum(!RateExpDate)
        '
        pContactData.GraceDays = nnNum(!GraceDays)
        pContactData.OnlineAuths = nnNum(!OnlineAuths)
        pContactData.SaleDate = nnNum(!SaleDate)
        pContactData.SaleDays = nnNum(!SaleDays)
        '
        pContactData.PVGraceDays = nnNum(!PVGraceDays)
        pContactData.PVOnlineAuths = nnNum(!PVOnlineAuths)
        pContactData.PVSaleDate = nnNum(!PVSaleDate)
        pContactData.PVSaleDays = nnNum(!PVSaleDays)
        '
        pContactData.WebPassword = !WebPassword & vbNullString
        pContactData.ContactByEmail = nnNum(!ContactByEmail)
        pContactData.ChangedData = nnNum(!ChangedData)
        '
        pContactData.AdjusterID = !AdjusterID & vbNullString
        End With
        '
        If IsDate(pContactData.AuthDate) Then
          pContactData.AuthRemaining = DateDiff("d", Now, DateAdd("d", CDbl(pContactData.AuthDays), pContactData.AuthDate))
        End If
        '
        If IsDate(pContactData.AuthDate) Then
          pContactData.PVAuthRemaining = DateDiff("d", Now, DateAdd("d", CDbl(pContactData.PVAuthDays), pContactData.PVAuthDate))
        End If
        '
        pContactData.DaysPending = CalculatePendingDays(pContactData.SaleDate, pContactData.GraceDays, pContactData.SaleDays)
        '
        pContactData.PVDaysPending = CalculatePendingDays(pContactData.PVSaleDate, pContactData.PVGraceDays, pContactData.PVSaleDays)
        '
        Load = True
      Else
        Load = False
      End If
      '
      rsContact.Close
      '
      Set rsContact = Nothing
  Exit Function
EH:
  MsgBox Err.Description & " in Class Contact: Load."
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

Public Function Delete(ByVal plID As Long) As Boolean
  
  If InputBox("Type DELETE and click OK if you really want to delete this contact.", "Delete Contact") = "DELETE" Then
    If ConnType = Access Then
      cnMain.Execute "DELETE FROM tblSupportActs WHERE CustRecID = " & plID, , adCmdText
      cnMain.Execute "DELETE FROM TContact WHERE ID = " & plID, , adCmdText
    Else
      cnMain.Execute "DELETE FROM TSupportActs WHERE CustRecID = " & plID, , adCmdText
      cnMain.Execute "DELETE FROM TContact WHERE ID = " & plID, , adCmdText
    End If
    'Me.Clear
    Delete = True
  Else
    Delete = False
  End If
  
End Function

Public Sub LoadCollection(ByVal plCompanyID As Long, ByRef pContacts As CContacts, ByVal plBranchID As Long)
  Dim rslist As New ADODB.Recordset
  Dim ContactData As CContactData
  Dim BranchData As CBranchData
  '
  Dim sQuery As String
  '
  Set pContacts = New CContacts
  '
  sQuery = "SELECT TContact.ID, TContact.ContactType, TContact.Status, TContact.FirstName, TContact.LastName, " & _
            "TContact.AuthDate, TContact.AuthDays " & _
            "FROM TContact " & _
            "WHERE (((TContact.CompanyID)=" & plCompanyID & "))" 'ORDER BY TContact.LastName, TContact.FirstName"
  '"TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & _
  '
  If plBranchID > 0 Then
    sQuery = sQuery & " AND (TCONTACT.BRANCHID = " & plBranchID & ") "
  End If
  '
  sQuery = sQuery & " ORDER BY TContact.LastName, TContact.FirstName"
  '
  rslist.Open sQuery, cnMain, adOpenForwardOnly, adLockReadOnly
  '
 ' If rslist.RecordCount = 0 Then
   ' Debug.Print rslist.RecordCount
 ' End If
  '
  While Not rslist.eof
    With rslist
        Set ContactData = New CContactData
        '
        ContactData.ID = rslist!ID
        ContactData.AuthDays = nnNum(rslist!AuthDays)
        ContactData.AuthDate = nnNum(rslist!AuthDate)
        ContactData.Status = rslist!Status & vbNullString
        ContactData.ContactType = nnNum(rslist!ContactType)
        ContactData.FirstName = rslist!FirstName & vbNullString
        ContactData.LastName = rslist!LastName & vbNullString
        'Load ContactData, rslist!ID
        '
        pContacts.Add ContactData
        '
        rslist.MoveNext
      End With
    Wend
    '
    rslist.Close
  '
  Set rslist = Nothing
  Set ContactData = Nothing
  '
End Sub

