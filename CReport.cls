VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ReportType
  daily
  Customer
  Eval
  WalkThrough
  Tech
  DaysNotAuth
  NotesSearch
  DaysRemaining
  SimpleContact
  SimpleContact2
  History
  Frontier
  FromList
  NoContactIn
  None
End Enum
'
Public rsReport As New ADODB.Recordset
'
Private RRtype As ReportType
'
'These should be properties
'
Public ActionType As String
Public SortOrder As String
Public Company As String
Public Branch As String
Public Zip As String
Public City As String
Public LastName As String
Public FirstName As String
Public Status As String
Public Notes As String
Public DaysMax As Integer
Public DaysMin As Integer
Public ID As Integer
Public State As String
Public Address1 As String
Public Address2 As String
Public Phone1 As String
Public Phone2 As String
Public Fax As String
Public Email As String
Public AuthStatus As String
Public AuthDate As Date
Public LastUpdate As Date
Public ShipStatus As Date
Public ShipDate As Date
Public Results As String
Public Subject As String
Public User As String
Public ResultsType As String
Public Time As Date
Public ResultsDateMin As Date
Public ResultsDateMax As Date
Public RecID As Long
Public CustRecID As Long
Public ContactType As Integer
Public ProductID As Long
Public Source As String
Public ListQuery As String
Public SortDirection As Integer
Public SortField As String
Public RecLimit As Integer
'
Private RQueryString As String
'
Public Property Get Rtype() As ReportType
  On Error GoTo EH
  '
  Rtype = RRtype
  '
Exit Function
EH:
 MsgBox Err.Description & " in Get Rtype."
End Property
'
Public Property Let Rtype(ByVal Value As ReportType)
  On Error GoTo EH
  '
  Validate
  RRtype = Value
  '
  Select Case RRtype
  Case daily
    Set rsReport = CreateDaily(-1).rs
  Case SimpleContact
    Set rsReport = CreateSimpleContact
  Case History
    Set rsReport = CreateHistory
  Case DaysRemaining
    Set rsReport = CreateDaysRemaining
  Case NotesSearch
    Set rsReport = CreateNotesSearch
  Case DaysNotAuth
    Set rsReport = CreateDaysNotAuth
  Case Frontier
   Set rsReport = CreateFrontier
  Case FromList
    Set rsReport = CreateFromList
  Case SimpleContact2
    Set rsReport = CreateSimpleContact2
  Case NoContactIn
    Set rsReport = CreateNoContactIn
  Case Else
    MsgBox ("Huh?")
  End Select
  '
Exit Property
EH:
 MsgBox Err.Description & " in Let Rtype."
End Property
'
Property Get QueryString() As String
  On Error GoTo EH
  '
  QueryString = RQueryString
  '
Exit Property
EH:
 MsgBox Err.Description & " in Get Query String."
End Property
'
Private Property Let QueryString(ByVal Value As String)
  On Error GoTo EH
  '
  RQueryString = Value
  '
Exit Property
EH:
 MsgBox Err.Description & " in Let QueryString."
End Property
'
Public Sub PreviewReport(t As String)
  On Error GoTo EH
  '
  Select Case t
    Case "Daily"
      Dim rpt As New RDailyBlank
      Set rpt = CreateDailyReport
      rpt.Show
      Set rpt = Nothing
    Case "Simple Contact"
      RSimpleContact.adc.Connection = cnMain
      RSimpleContact.adc.Source = QueryString
      RSimpleContact.Show
    Case "History"
      RHistory.adc.Connection = cnMain
      RHistory.adc.Source = QueryString
      RHistory.Show
    Case "Frontier"
      RBranch.adc.Recordset = rsReport
      RBranch.Show
    Case "DaysLeft"
      RDaysLeft.adc.Recordset = rsReport
      RDaysLeft.Show
  Case Else
      MsgBox ("No report Type found")
  End Select
'
Exit Sub
EH:
 MsgBox Err.Description & " in PreviewReport."
End Sub
Public Sub PrintReport(t As String)
  On Error GoTo EH
  '
  Select Case t
    Case "Daily"
      Dim rpt As New RDailyBlank
      Set rpt = CreateDailyReport
      rpt.PrintReport True
      Set rpt = Nothing
    Case "Simple Contact"
      RSimpleContact.adc.Connection = cnMain
      RSimpleContact.adc.Source = QueryString
      RSimpleContact.PrintReport True
      Unload RSimpleContact
    Case "History"
      RHistory.adc.Connection = cnMain
      RHistory.adc.Source = QueryString
      RHistory.PrintReport True
      Unload RHistory
    Case "Frontier"
      RBranch.adc.Recordset = rsReport
      RBranch.PrintReport True
      Unload RBranch
    Case "DaysLeft"
      RDaysLeft.adc.Recordset = rsReport
      RDaysLeft.PrintReport True
      Unload RDaysLeft
    Case "FromList"
      RSimpleContact.adc.Connection = cnMain
      RSimpleContact.adc.Source = QueryString
      RSimpleContact.PrintReport True
      Unload RSimpleContact
  Case Else
      MsgBox ("No report Type found")
  End Select
'
Exit Sub
EH:
 MsgBox Err.Description & " in PrintReport."
End Sub
'
Private Function Validate() As Boolean
'
  On Error GoTo EH
  '
  Validate = True
  'Results = "%" + Results + "%"
  '
  If Notes <> "" Then
    Notes = "%" + Notes + "%')"
  Else
   ' Notes = "UNLIKELY@#$%^&*"
  End If
  '
  Company = "%" + Company + "%"
  Branch = "%" + Branch + "%')"
  '
  If DaysMin > 10000 Or DaysMin < -10000 Then DaysMin = -10000
  If DaysMax > 10000 Or DaysMax < -10000 Then DaysMax = -10000
  If FirstName = "" Then FirstName = "%"
  If LastName = "" Then LastName = "%"
  If State = "" Then State = "%"
  If State = "All" Then State = "%"
  If City = "" Then City = "%"
  If Zip = "" Then Zip = "%"
  If Company = "" Then Company = "%"
  If Branch = "" Then Branch = "%' OR TBranch.Name IS NULL)"
 ' If Notes = "" Then Notes = "QQQQQQQ"
  If Status = "All Users" Then Status = "%"
  If Status = "Everyone" Then Status = "%"
  '
  If State = "" Then State = "%"
  'If Results = "" Then Results = "%"
  If User = "All Users" Then User = "%"
  If User = "" Then User = "%"
  If Notes = "" Then Notes = "%"
  If Source = "" Then Source = "%"
  If ResultsType = "All Categories" Then ResultsType = "%"
  If ResultsType = "" Then ResultsType = "%"
  '
  'If Results = "%%" Then Results = "%"
  
  If Notes = "%%" Then Notes = "%"
  If Company = "%%" Then Company = "%"
  If Branch = "%%')" Then Branch = "%' OR TBranch.Name IS NULL)"
  '
  If Status = "Everyone" Then Status = "%"
  'If ContactType = "0" Then ContactType = "%"
  If ContactType = "1" Then ContactType = "%"
  '
Exit Function
EH:
 MsgBox Err.Description & " in Validate."
End Function
'
Private Sub Class_Initialize()
'
RRtype = None
'
End Sub
'
Public Function CreateDailyReport() As RDailyBlank
'
  On Error GoTo EH
  '
  Dim rpt As New RDailyBlank
  Dim ctl As Object
  Dim ctl2 As Object
  '
  rpt.Height = 6375
  rpt.PrintWidth = 10170
  rpt.PageLeftMargin = 500
  rpt.PageRightMargin = 500
  rpt.PageBottomMargin = 700
  rpt.PageTopMargin = 700
  rpt.PageHeader.Height = 1125
  rpt.PageFooter.Height = 0
  rpt.Width = 11520
  '
  Dim rs As Recordset
  Dim iTop As Integer
  Dim i As Integer
  Dim DailySubItem As CDailySubItem
  '
  rpt.DateLabel = Date
  '
  With rpt.Sections("Detail").Controls
    '
    iTop = 0
    '
    For i = 1 To 10
      '
      Set DailySubItem = CreateDaily(i)
       '
       If DailySubItem.Name <> "" Then
       '
        Set ctl = .Add("DDActiveReports.Label")
        ctl.Caption = DailySubItem.Name
        ctl.Width = 9900
        ctl.Height = 360
        ctl.Left = 90
        ctl.Top = iTop
        ctl.Font.Size = 14
        iTop = iTop + ctl.Height + 90
        '
        Set ctl = .Add("DDActiveReports.SubReport")
        ctl.Left = 270
        ctl.Height = 1080
        ctl.Left = 270
        ctl.Top = iTop
        ctl.Width = 9900
        Set ctl.object = New RDailyUnit
        ctl.object.adc.Recordset = DailySubItem.rs
        iTop = iTop + ctl.Height + 100
       '
       End If
    '
    Next i
  '
  End With
  '
  Set rs = Nothing
  Set ctl = Nothing
  Set DailySubItem = Nothing
  Set CreateDailyReport = rpt
  Set rpt = Nothing
  '
  Exit Function
EH:
 MsgBox Err.Description & " in CReport.CreateDailyReport."
End Function
'
Public Function CreateDaily(ItemNum As Integer) As CDailySubItem 'As adodb.Recordset
  On Error GoTo EH
  '
  Dim iCount As Integer
  Dim DailyItemsSettings(10) As New CDailyItemSettings
  Dim tempDailySubItem As CDailySubItem
  Dim FinalDailySubItem As New CDailySubItem
  Dim DailySubItem As New CDailyItemSettings
  Dim rsAll As New ADODB.Recordset
  Dim rs As New ADODB.Recordset
  Dim rsOne As New ADODB.Recordset
  Dim CN As New ADODB.Connection
  Dim i As Integer
  '
  rsAll.Fields.Append "FirstName", adChar, 300, adFldIsNullable
  rsAll.Fields.Append "LastName", adChar, 300, adFldIsNullable
  rsAll.Fields.Append "Company", adChar, 300, adFldIsNullable
  rsAll.Fields.Append "Phone1", adChar, 300, adFldIsNullable
  rsAll.Fields.Append "Status", adChar, 300, adFldIsNullable
  rsAll.Fields.Append "Type", adChar, 300, adFldIsNullable
  rsAll.Fields.Append "DaysSinceEvent", adInteger, 20, adFldIsNullable
  rsAll.Open
  rs.Open "SELECT * FROM TDailyReportsSettings", "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source=" & App.Path & "\" & "settings.mdb"
  '
  While (rs.eof = False)
  '
    iCount = iCount + 1
    rs.MoveNext
  '
  Wend
  '
  rs.MoveFirst
  '
  If ItemNum = -1 Then
    '
    i = 1
    While (rs.eof = False) And i <= 10
    '
      DailyItemsSettings(i).StatusType = rs!StatusType & vbNullString
      DailyItemsSettings(i).SupportType = rs!SupportType & vbNullString
      DailyItemsSettings(i).InitialDays = rs!InitialDays & vbNullString
      DailyItemsSettings(i).RegularDays = rs!RegularDays & vbNullString
      DailyItemsSettings(i).DaysDivider = rs!DaysDivider & vbNullString
      DailyItemsSettings(i).DaysOperator = rs!DaysOperator & vbNullString
      DailyItemsSettings(i).Name = rs!Name & vbNullString
      DailyItemsSettings(i).DaysToSearch = rs!DaysToSearch & vbNullString
      '
      Set rsOne = CreateDailyItem(DailyItemsSettings(i), Date).rs
      CombineRS rsAll, rsOne
        i = i + 1
        rs.MoveNext
    '
    Wend
     FinalDailySubItem.Name = "Complete Report"
    Set FinalDailySubItem.rs = rsAll
  Else
  '
    If ItemNum <= iCount Then
    '
      rs.MoveFirst
      rs.Move (ItemNum - 1)
      DailyItemsSettings(ItemNum).StatusType = rs!StatusType & vbNullString
      DailyItemsSettings(ItemNum).SupportType = rs!SupportType & vbNullString
      DailyItemsSettings(ItemNum).InitialDays = rs!InitialDays & vbNullString
      DailyItemsSettings(ItemNum).RegularDays = rs!RegularDays & vbNullString
      DailyItemsSettings(ItemNum).DaysDivider = rs!DaysDivider & vbNullString
      DailyItemsSettings(ItemNum).DaysOperator = rs!DaysOperator & vbNullString
      DailyItemsSettings(ItemNum).Name = rs!Name & vbNullString
      DailyItemsSettings(ItemNum).DaysToSearch = rs!DaysToSearch & vbNullString
      Set tempDailySubItem = CreateDailyItem(DailyItemsSettings(ItemNum), Date)
      Set rsOne = tempDailySubItem.rs
       FinalDailySubItem.Name = tempDailySubItem.Name
      Set FinalDailySubItem.rs = rsOne
    '
    End If
  '
  End If
  '
  Set CreateDaily = FinalDailySubItem
  '
  For i = 1 To 10
  '
    Set DailyItemsSettings(i) = Nothing
  '
  Next i
  '
  Set FinalDailySubItem = Nothing
  Set rsAll = Nothing
  Set rs = Nothing
  Set rsOne = Nothing
  Set CN = Nothing
  '
Exit Function
EH:
 MsgBox Err.Description & " in CreateDaily."
End Function
Public Function CreateDailyItem(Item As CDailyItemSettings, ReportDate As Date) As CDailySubItem
  On Error GoTo EH
  Screen.MousePointer = vbHourglass
  Dim DailySubItem As New CDailySubItem
  Dim QString As String
  '" WHERE ((([AuthDays]-DateDiff(d,[AuthDate]," & ReportDate & "))>0) AND" &
  QString = "SELECT [TCompany].[Name] as Company," & _
                    " [TContact].[FirstName]," & _
                    " [TContact].[LastName]," & _
                    " [TContact].[Phone1]," & _
                    " [TContact].[Status]," & _
                    " [AuthDays]-DateDiff(d,[AuthDate],GetDate()) AS Days," & _
                    " DateDiff(d,TSUPPORTACTS.DATE,GETDATE()) AS DaysSinceEvent," & _
                    " [TSupportActs].[Type]," & _
                    " [TSupportActs].[Date]," & _
                    " [TSupportActs].[Subject]," & _
                    " [TContact].[ID]" & _
                  " FROM (TCompany RIGHT JOIN TContact ON [TCompany].[ID]=[TContact].[CompanyID])" & _
                    " RIGHT JOIN TSupportActs ON [TContact].[ID]=[TSupportActs].[CustRecID]" & _
                  " WHERE ((([AuthDays]-DateDiff(d,[AuthDate],getdate()))>0) AND" & _
                    " (DateDiff(d,TSUPPORTACTS.DATE,GETDATE())<" & Item.InitialDays + Item.DaysToSearch & ") AND" & _
                    " (DateDiff(d,TSUPPORTACTS.DATE,GETDATE())>" & 0 & ") AND" & _
                    " (([TContact].[Status]) Like '" & Item.StatusType & "') AND" & _
                    " (([TSupportActs].[Type]) Like '" & Item.SupportType & "'))" & _
                    " ORDER BY DateDiff(d,TSUPPORTACTS.DATE,GETDATE());"
  'MsgBox (QString)
  Dim bigRS As ADODB.Recordset
  Set bigRS = New ADODB.Recordset
  bigRS.Source = QString
  bigRS.Open , cnMain
  Dim smallerRS As New ADODB.Recordset
  smallerRS.Fields.Append "FirstName", adChar, 300, adFldIsNullable
  smallerRS.Fields.Append "LastName", adChar, 300, adFldIsNullable
  smallerRS.Fields.Append "Company", adChar, 300, adFldIsNullable
  smallerRS.Fields.Append "Phone1", adChar, 300, adFldIsNullable
  smallerRS.Fields.Append "Status", adChar, 300, adFldIsNullable
  smallerRS.Fields.Append "Type", adChar, 300, adFldIsNullable
  smallerRS.Fields.Append "Days", adInteger, 300, adFldIsNullable
  smallerRS.Fields.Append "DaysSinceEvent", adInteger, 300, adFldIsNullable
  smallerRS.Fields.Append "Subject", adInteger, 300, adFldIsNullable
  smallerRS.Fields.Append "Date", adDate, 300, adFldIsNullable
  smallerRS.Open
  Dim SingleChar As String
  Dim Complete As String
  Dim pos As Integer
  Dim NumFlag As Boolean
  Dim ItemFlag As Boolean
  '
   Do Until bigRS.eof
      ItemFlag = True
    'If Not IsNull(bigRS("Subject")) Or (Item.RegularDays = 0) Then 'ItemFlag = False
      If Item.DaysOperator = "<" And Val(bigRS("Subject") & vbNullString) >= Item.DaysDivider Then ItemFlag = False
      If Item.DaysOperator = ">" And Val(bigRS("Subject") & vbNullString) <= Item.DaysDivider Then ItemFlag = False
      'If IsNull(bigRS("Subject")) Then ItemFlag = False
      If ItemFlag Then '1
       
            If (bigRS("Days") + bigRS("DaysSinceEvent") = Val(bigRS("Subject") & vbNullString)) Or (Item.RegularDays = 0) Then '3
            If CallOrNot(bigRS("Date"), Item.InitialDays) Then '4
              smallerRS.AddNew
              smallerRS("FirstName") = bigRS("FirstName") & vbNullString
              smallerRS("LastName") = bigRS("LastName") & vbNullString
              smallerRS("Company") = bigRS("Company") & vbNullString
              smallerRS("Phone1") = bigRS("Phone1") & vbNullString
              smallerRS("Status") = bigRS("Status") & vbNullString
              smallerRS("Days") = bigRS("Days") & vbNullString
              smallerRS("DaysSinceEvent") = bigRS("DaysSinceEvent") & vbNullString
              smallerRS("Type") = bigRS("Type") & vbNullString
              smallerRS("Date") = bigRS("Date") & vbNullString
              smallerRS("Subject") = Val(bigRS("Subject") & vbNullString)
            Else '4
             If (Item.RegularDays <> 0) Then
              If ((bigRS("DaysSinceEvent") - Item.InitialDays) Mod Item.RegularDays) = 0 Then  '5
                If CallOrNot(bigRS("Date"), bigRS("DaysSinceEvent")) Then '6
                  smallerRS.AddNew
                  smallerRS("FirstName") = bigRS("FirstName") & vbNullString
                  smallerRS("LastName") = bigRS("LastName") & vbNullString
                  smallerRS("Company") = bigRS("Company") & vbNullString
                  smallerRS("Phone1") = bigRS("Phone1") & vbNullString
                  smallerRS("Status") = bigRS("Status") & vbNullString
                  smallerRS("Days") = bigRS("Days") & vbNullString
                  smallerRS("DaysSinceEvent") = bigRS("DaysSinceEvent") & vbNullString
                  smallerRS("Type") = bigRS("Type") & vbNullString
                  smallerRS("Date") = bigRS("Date") & vbNullString
                  smallerRS("Subject") = Val(bigRS("Subject"))
                End If '6
              End If '5
             End If
            End If '4
          End If '3
         End If '1
       'End If
      bigRS.MoveNext
    
  Loop
  If smallerRS.RecordCount > 0 Then
    smallerRS.MoveFirst
  End If
  DailySubItem.Name = Item.Name
  Set DailySubItem.rs = smallerRS
  Set CreateDailyItem = DailySubItem
  Set DailySubItem = Nothing
  Set bigRS = Nothing
  Set smallerRS = Nothing
  Screen.MousePointer = vbDefault
Exit Function
EH:
 MsgBox Err.Description & " in CreateDailyItem."
End Function
'
Public Function CreateNoContactIn() As ADODB.Recordset
  On Error GoTo EH
  Dim rsTemp As New ADODB.Recordset
  QueryString = "SELECT TCompany.ID AS CompanyID, TContact.ID AS ContactID" & _
                          ",TCompany.Name AS Company" & _
                          ", TContact.FirstName" & _
                          ", TContact.LastName" & _
                          ", MAX(TSupportActs.[Date]) AS [last called]" & _
                          ", TContact.Phone1" & _
                          ", TContact.Status" & _
                          ", TContact.AuthDays" & _
                          ", TSupportActs.Type" & _
                          " FROM TSupportActs LEFT OUTER JOIN" & _
                          " TContact ON TSupportActs.CustRecID = TContact.ID LEFT OUTER JOIN" & _
                          " TCompany ON TContact.CompanyID = TCompany.ID" & _
                          " GROUP BY TContact.AuthDays, TSupportActs.Type, TContact.ID" & _
                          ", TContact.Status, TContact.Phone1, TContact.LastName" & _
                          ", TCompany.Name, TContact.FirstName, TCompany.ID" & _
                          " HAVING(TCompany.Name LIKE N'" & Company & "') AND" & _
                          " (TContact.Status LIKE N'" & Status & "') AND" & _
                          " (MAX(TSupportActs.[Date]) <= GETDATE() - " & DaysMin & ") AND" & _
                          " (TSupportActs.Type LIKE N'" & ActionType & "')" & _
                          " ORDER BY MAX(TSupportActs.[Date]) DESC"
                          
    'If SortField <> vbNullString Then
    '  QueryString = QueryString & " ORDER BY TContact." & SortField
    '  If SortDirection = 1 Then QueryString = QueryString & " DESC"
    'Else
    '  QueryString = QueryString & " ORDER BY TContact.LastName"
    'End If
    '
    rsTemp.Source = QueryString
    'MsgBox (QueryString)
    rsTemp.Open , cnMain
    If rsTemp.RecordCount > 0 Then
      rsTemp.MoveFirst
    End If
    Set CreateNoContactIn = rsTemp
Exit Function
EH:
 MsgBox Err.Description & " in CreateNoContactIn."
End Function
'
Public Function CreateSimpleContact() As ADODB.Recordset
  On Error GoTo EH
  Dim rsTemp As New ADODB.Recordset
  QueryString = "SELECT TCompany.ID AS CompanyID, TContact.ID AS ContactID" & _
                          ",TCompany.Name AS Company" & _
                          ",TContact.FirstName" & _
                          ",TContact.LastName" & _
                          ",TContact.Phone1" & _
                          ",TContact.Status" & _
                          ",TContact.Source" & _
                     " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & _
                     " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & _
                     " AND (TContact.LastName LIKE '" & LastName & "' )" & _
                     " AND (TCompany.Name LIKE '" & Company & "' )" & _
                     " AND (TContact.Status LIKE '" & Status & "' )" & _
                     " AND (TContact.City LIKE '" & City & "' )" & _
                     " AND (TContact.Zip LIKE '" & Zip & "' )" & _
                     " AND (TContact.State LIKE '" & State & "' )" & _
                     " AND (TContact.Source LIKE '" & Source & "' )" '& _
                     '" AND (TContact.Notes LIKE '" & Notes & "')" '& _
                     '" ORDER BY TContact.LastName"
    '
    If SortField <> vbNullString Then
      QueryString = QueryString & " ORDER BY TContact." & SortField
      If SortDirection = 1 Then QueryString = QueryString & " DESC"
    Else
      QueryString = QueryString & " ORDER BY TContact.LastName"
    End If
    '
    rsTemp.Source = QueryString
    'MsgBox (QueryString)
    rsTemp.Open , cnMain
    If rsTemp.RecordCount > 0 Then
      rsTemp.MoveFirst
    End If
    Set CreateSimpleContact = rsTemp
Exit Function
EH:
 MsgBox Err.Description & " in CreateSimpleContact."
End Function
'
Public Function CreateSimpleContact2() As ADODB.Recordset
'For main contact screen
  On Error GoTo EH
  Dim rsTemp As New ADODB.Recordset
  '
  QueryString = "SELECT TContact.ID, TContact.ContactType, TContact.Status" & _
                          ",TContact.FirstName" & _
                          ",TContact.LastName" & _
                          ",TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & _
                          ",TCompany.Name AS Company" & _
                          ",TContact.Phone1" & _
                          ",TContact.Source" & _
                     " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & _
                     " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & _
                     " AND (TContact.LastName LIKE '" & LastName & "' )" & _
                     " AND (TCompany.Name LIKE '" & Company & "' )" & _
                     " AND (TContact.Status LIKE '" & Status & "' )" & _
                     " AND (TContact.City LIKE '" & City & "' )" & _
                     " AND (TContact.Zip LIKE '" & Zip & "' )" & _
                     " AND (TContact.State LIKE '" & State & "' )" & _
                     " AND (TContact.Source LIKE '" & Source & "' )" & _
                     " AND (TContact.Notes LIKE '" & Notes & "')" '& _
                     '" ORDER BY TContact.LastName"
    '
    '",TContact.Notes" &
    If SortField <> vbNullString Then
      QueryString = QueryString & " ORDER BY TContact." & SortField
      If SortDirection = 1 Then QueryString = QueryString & " DESC"
    Else
      QueryString = QueryString & " ORDER BY TContact.LastName"
    End If
    '
    rsTemp.Source = QueryString
    'MsgBox (QueryString)
    rsTemp.Open , cnMain
    If rsTemp.RecordCount > 0 Then
      rsTemp.MoveFirst
    End If
    Set CreateSimpleContact2 = rsTemp
Exit Function
EH:
 MsgBox Err.Description & " in CreateSimpleContact2."
End Function
'
Public Function CreateHistory() As ADODB.Recordset
  On Error GoTo EH
  Dim rsTemp As New ADODB.Recordset
  Dim sResultsQuery As String
  Dim sResultsTemp As String
  Dim sResultsChar As String
  Dim sProductID As String
  Dim sSortOrder As String
  Dim i As Integer
  Select Case SortOrder
    Case "Company/Branch"
      sSortOrder = " ORDER by TCompany.Name DESC, TBranch.Name DESC, Date DESC, Time DESC"
    Case "User"
      sSortOrder = " ORDER by TSupportActs.[User] DESC, Date DESC, Time DESC"
    Case "Category"
      sSortOrder = " ORDER by TSupportActs.Type DESC, Date DESC, Time DESC"
    Case Else
      sSortOrder = " ORDER by Date DESC, Time DESC"
  End Select
  '
  If Results <> "" Then
    sResultsChar = Mid(Results, 1, 1)
    If sResultsChar = Chr(34) Then
      sResultsQuery = " (TSupportActs.Results LIKE '%" & Mid(Results, 2, Len(Results) - 2) & "%') AND"
    Else
      For i = 1 To Len(Results)
        sResultsChar = Mid(Results, i, 1)
        If sResultsChar <> " " Then
          sResultsTemp = sResultsTemp & sResultsChar
        Else
          If sResultsTemp <> "" Then
            sResultsQuery = sResultsQuery & " (TSupportActs.Results LIKE '%" & sResultsTemp & "%') AND"
            sResultsTemp = ""
          End If
        End If
      Next i
    End If
    '
    If sResultsTemp <> "" Then
       sResultsQuery = sResultsQuery & " (TSupportActs.Results LIKE '%" & sResultsTemp & "%') AND"
    End If
  Else
    sResultsQuery = " (TSupportActs.Results LIKE '%') AND"
  End If
  '
  Dim sRecLimit As String
  If RecLimit > 0 Then
    sRecLimit = " TOP " & CStr(RecLimit)
  Else
    sRecLimit = ""
  End If
  '
  Dim sContactType As String
  If ContactType <> 0 Then sContactType = "(TContact.ContactType = '" & ContactType & "') AND"
  '
  If ProductID = 1 Then sProductID = " (TSupportActs.ProductID = 1 or TSupportActs.ProductID is null) AND"
  If ProductID > 1 Then sProductID = " (TSupportActs.ProductID = " & ProductID & " ) AND"
  '
  QueryString = "SELECT" & sRecLimit & " TCompany.Name AS Company, TSupportActs.Type," & _
                  " TSupportActs.Date, TContact.FirstName," & _
                  " TContact.LastName, TSupportActs.Results," & _
                  " TSupportActs.[User], TContact.Status," & _
                  " TContact.State, TSupportActs.RecID," & _
                  " TSupportActs.CustRecID, TSupportActs.Subject," & _
                  " TSupportActs.Results, TSupportActs.Time," & _
                  " TContact.Phone1, TBranch.Name AS Branch" & _
              " FROM TBranch RIGHT OUTER JOIN" & _
                  " TContact ON TBranch.BranchID = TContact.BranchID RIGHT OUTER JOIN" & _
                  " TSupportActs ON TContact.ID = TSupportActs.CustRecID LEFT OUTER JOIN" & _
                  " TCompany ON TContact.CompanyID = TCompany.ID" & _
              " WHERE (TSupportActs.Type LIKE '" & ResultsType & "') AND" & _
                " (TSupportActs.Date >= '" & ResultsDateMin & "') AND" & _
                " (TSupportActs.Date <= '" & ResultsDateMax & "') AND" & _
                " (TContact.Status LIKE '" & Status & "') AND" & _
                " (TSupportActs.[User] LIKE '" & User & "') AND" & _
                " (TContact.FirstName LIKE '" & FirstName & "') AND" & _
                " (TContact.LastName LIKE '" & LastName & "') AND" & _
                " (TCompany.Name LIKE '" & Company & "') AND" & _
                " (TBranch.Name LIKE '" & Branch & " AND" & _
                sContactType & sResultsQuery & sProductID & _
                " (TContact.State LIKE '" & State & "')" & sSortOrder
              '" ORDER by Date DESC, Time DESC"
            'MsgBox (QueryString)
    '" (TSupportActs.Results LIKE '" & Results & "') AND" & _
    'If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Source = QueryString
    'MsgBox (QueryString)
    rsTemp.Open , cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsTemp.RecordCount > 0 Then
      rsTemp.MoveFirst
    End If
    Set CreateHistory = rsTemp
Exit Function
EH:
 MsgBox Err.Description & " in CreateHistory."
End Function
'
Public Function CreateDaysNotAuth() As ADODB.Recordset
  On Error GoTo EH
  Dim rsTemp As New ADODB.Recordset
  Dim sShipDateFieldName As String
  '
  If ProductID = 2 Then
    sShipDateFieldName = "TContact.PVShipDate"
  Else
    sShipDateFieldName = "TContact.ShipDate"
  End If
  
  QueryString = "SELECT TCompany.ID as CompanyID, TContact.ID as ContactID" & _
                          ",TCompany.Name AS Company" & _
                          ",TContact.FirstName" & _
                          ",TContact.LastName" & _
                          ",TContact.Phone1" & _
                          ",TContact.Status" & _
                          ",TContact.Source" & _
                          ",DATEDIFF(d, " & sShipDateFieldName & ", GetDate()) AS Days" & _
                     " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & _
                     " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & _
                     " AND (TContact.LastName LIKE '" & LastName & "' )" & _
                     " AND (TCompany.Name LIKE '" & Company & "' )" & _
                     " AND (TContact.Status LIKE '" & Status & "' )" & _
                     " AND (TContact.City LIKE '" & City & "' )" & _
                     " AND (TContact.Zip LIKE '" & Zip & "' )" & _
                     " AND (TContact.State LIKE '" & State & "' )" & _
                     " AND (TContact.Source LIKE '" & Source & "' )" & _
                     " AND(DATEDIFF(d, " & sShipDateFieldName & ", GetDate()) <= '" & DaysMax & "')" & _
                     " AND(DATEDIFF(d, " & sShipDateFieldName & ", GetDate()) >= '" & DaysMin & "')" & _
                     " AND(TContact.AuthStatus LIKE 'Not Authorized')" '& _
                     '" ORDER BY (DATEDIFF(d, dbo.TContact.ShipDate, GetDate()))"
  If SortField <> vbNullString Then
    QueryString = QueryString & " ORDER BY TContact." & SortField
    If SortDirection = 1 Then QueryString = QueryString & " DESC"
  Else
    QueryString = QueryString & " ORDER BY (DATEDIFF(d, dbo.TContact.ShipDate, GetDate()))"
  End If
  '
  rsTemp.Source = QueryString
  'MsgBox (QueryString)
  rsTemp.Open , cnMain
  If rsTemp.RecordCount > 0 Then
    rsTemp.MoveFirst
  End If
  Set CreateDaysNotAuth = rsTemp
Exit Function
EH:
 MsgBox Err.Description & " in CreateDaysNotAuth."
End Function
'
Public Function CreateDaysRemaining() As ADODB.Recordset
  On Error GoTo EH
  Dim rsTemp As New ADODB.Recordset
  Dim sAuthDaysFieldName As String
  Dim sAuthDateFieldName As String
  Dim sStatus As String
  '
  If ProductID = 2 Then
    sAuthDaysFieldName = "TContact.PVAuthDays"
    sAuthDateFieldName = "TContact.PVAuthDate"
    'sStatus = "TContact.PVStatus"
  Else
    sAuthDaysFieldName = "TContact.AuthDays"
    sAuthDateFieldName = "TContact.AuthDate"
    'sStatus = "TContact.PVStatus"
  End If
  '
  QueryString = "SELECT TCompany.ID as CompanyID, TContact.ID as ContactID" & _
                        ",TCompany.Name AS Company" & _
                        ",TContact.FirstName" & _
                        ",TContact.LastName" & _
                        ",TContact.Phone1" & _
                        ",TContact.Status" & _
                        ",TContact.Source" & _
                        "," & sAuthDateFieldName & " + TContact.AuthDays AS [ExpirationDate]" & _
                        "," & sAuthDaysFieldName & "-DateDiff(d," & sAuthDateFieldName & ", GETDATE()) as Days" & _
                   " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & _
                   " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & _
                   " AND (TContact.LastName LIKE '" & LastName & "' )" & _
                   " AND (TCompany.Name LIKE '" & Company & "' )" & _
                   " AND (TContact.Status LIKE '" & Status & "' )" & _
                   " AND (TContact.City LIKE '" & City & "' )" & _
                   " AND (TContact.Zip LIKE '" & Zip & "' )" & _
                   " AND (TContact.State LIKE '" & State & "' )" & _
                   " AND (TContact.Source LIKE '" & Source & "' )" & _
                   " AND (" & sAuthDaysFieldName & "- DateDiff(d, " & sAuthDateFieldName & ", GetDate())>='" & DaysMin & "' )" & _
                   " AND (" & sAuthDaysFieldName & "- DateDiff(d, " & sAuthDateFieldName & ", GetDate())<='" & DaysMax & " ' )" ' & _
                   '" ORDER BY (" & sAuthDaysFieldName & "- DateDiff(d, " & sAuthDateFieldName & ", GetDate()))"
  '
  If SortField <> vbNullString Then
    QueryString = QueryString & " ORDER BY TContact." & SortField
    If SortDirection = 1 Then QueryString = QueryString & " DESC"
  Else
    QueryString = QueryString & " ORDER BY (" & sAuthDaysFieldName & "- DateDiff(d, " & sAuthDateFieldName & ", GetDate()))"
  End If
  '
  rsTemp.Source = QueryString
  'MsgBox (QueryString)
  rsTemp.Open , cnMain
  If rsTemp.RecordCount > 0 Then
    rsTemp.MoveFirst
  End If
  Set CreateDaysRemaining = rsTemp
Exit Function
EH:
 MsgBox Err.Description & " in CreateDaysRemaining."
End Function
'
Public Function CreateNotesSearch() As ADODB.Recordset
  On Error GoTo EH
  Dim rsTemp As New ADODB.Recordset
  QueryString = "SELECT TCompany.ID as CompanyID, TContact.ID as ContactID" & _
                        ",TCompany.Name AS Company" & _
                        ",TContact.FirstName" & _
                        ",TContact.LastName" & _
                        ",TContact.Phone1" & _
                        ",TContact.Status" & _
                        ",TContact.Source" & _
                        ",TContact.Notes" & _
                   " FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID" & _
                   " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & _
                   " AND (TContact.LastName LIKE '" & LastName & "' )" & _
                   " AND (TCompany.Name LIKE '" & Company & "' )" & _
                   " AND (TContact.Status LIKE '" & Status & "' )" & _
                   " AND (TContact.City LIKE '" & City & "' )" & _
                   " AND (TContact.Zip LIKE '" & Zip & "' )" & _
                   " AND (TContact.State LIKE '" & State & "' )" & _
                   " AND (TContact.Notes LIKE '" & Notes & "')" & _
                   " AND (TContact.Source LIKE '" & Source & "' )" ' & _
                   '" ORDER BY TContact.LastName"
  '
  If SortField <> vbNullString Then
    QueryString = QueryString & " ORDER BY TContact." & SortField
    If SortDirection = 1 Then QueryString = QueryString & " DESC"
  Else
    QueryString = QueryString & " ORDER BY TContact.LastName"
  End If
  '
  rsTemp.Source = QueryString
  'MsgBox (QueryString)
  'If Notes <> "%" Then
    rsTemp.Open , cnMain
  'End If
  If rsTemp.RecordCount > 0 Then
    rsTemp.MoveFirst
  End If
  Set CreateNotesSearch = rsTemp
Exit Function
EH:
 MsgBox Err.Description & " in CreateNotesSearch."
End Function
'
Public Function CreateFromList() As ADODB.Recordset
  On Error GoTo EH
  '
  Dim rsTemp As New ADODB.Recordset
  '
  QueryString = ListQuery
  '
  If SortField <> vbNullString Then
    QueryString = QueryString & " ORDER BY TContact." & SortField
    If SortDirection = 1 Then QueryString = QueryString & " DESC"
  Else
    QueryString = QueryString & " ORDER BY TContact.State"
  End If
  rsTemp.Source = QueryString
  '
  rsTemp.Open , cnMain
  '
  If rsTemp.RecordCount > 0 Then
    rsTemp.MoveFirst
  End If
  '
  Set CreateFromList = rsTemp
  '
  Exit Function
EH:
  MsgBox Err.Description & " in CreateFromlist."
End Function
'
Public Function CreateFrontier() As ADODB.Recordset
  On Error GoTo EH
  Dim rsTemp As New ADODB.Recordset
  Notes = "B#"
  ResultsType = "Sale"
  Validate
  QueryString = "SELECT TCompany.ID as CompanyID" & _
                        ", TContact.ID as ContactID,TCompany.Name AS Company" & _
                        ",TContact.FirstName,TContact.LastName" & _
                        ",TContact.Phone1,TContact.Status" & _
                        ",TContact.Notes,TSupportActs.Date" & _
                " FROM TSupportActs LEFT OUTER JOIN" & _
                    " TContact ON" & _
                    " TSupportActs.CustRecID = dbo.TContact.ID LEFT OUTER JOIN" & _
                    " TCompany ON" & _
                    " TContact.CompanyID = TCompany.ID" & _
                   " WHERE (TContact.FirstName LIKE '" & FirstName & "' )" & _
                   " AND (TContact.LastName LIKE '" & LastName & "' )" & _
                   " AND (TCompany.Name LIKE '" & Company & "' )" & _
                   " AND (TContact.Status LIKE '" & Status & "' )" & _
                   " AND (TContact.City LIKE '" & City & "' )" & _
                   " AND (TContact.Zip LIKE '" & Zip & "' )" & _
                   " AND (TContact.State LIKE '" & State & "' )" & _
                   " AND (TContact.Notes LIKE '" & Notes & "')" & _
                   " AND (TContact.Source LIKE '" & Source & "' )" & _
                   " AND (TSupportActs.Date >= '" & ResultsDateMin & "') " & _
                   " AND (TSupportActs.Date <= '" & ResultsDateMax & "') " & _
                   " AND (TSupportActs.Type LIKE '" & ResultsType & "') " '& _
                   '" ORDER BY TContact.LastName"
   '
  If SortField <> vbNullString Then
    QueryString = QueryString & " ORDER BY TContact." & SortField
    If SortDirection = 1 Then QueryString = QueryString & " DESC"
  Else
    QueryString = QueryString & " ORDER BY TContact.LastName"
  End If
  rsTemp.Source = QueryString
   '
    rsTemp.Source = QueryString
  'MsgBox (QueryString)
  rsTemp.Open , cnMain
  Dim trs As New ADODB.Recordset
  trs.Fields.Append "CompanyID", adChar, 300, adFldIsNullable
  trs.Fields.Append "ContactID", adChar, 300, adFldIsNullable
  trs.Fields.Append "FirstName", adChar, 300, adFldIsNullable
  trs.Fields.Append "LastName", adChar, 300, adFldIsNullable
  trs.Fields.Append "Branch", adChar, 300, adFldIsNullable
  trs.Fields.Append "Date", adDate, 300, adFldIsNullable
  trs.Fields.Append "Status", adChar, 300, adFldIsNullable
  trs.Open
  Dim SingleChar As String
  Dim Complete As String
  Dim pos As Integer
  Dim NumFlag As Boolean
  '
  Do Until rsTemp.eof
    trs.AddNew
    trs("CompanyID") = rsTemp("CompanyID")
    trs("ContactID") = rsTemp("ContactID")
    trs("FirstName") = rsTemp("FirstName")
    trs("LastName") = rsTemp("LastName")
    trs("Status") = rsTemp("Status")
    SingleChar = Mid(rsTemp("Notes"), 3, 1)
    NumFlag = isNumber(SingleChar)
    Complete = ""
    pos = 1
    Do Until NumFlag = False
      Complete = Complete + SingleChar
      SingleChar = Mid(rsTemp("Notes"), 3 + pos, 1)
      NumFlag = isNumber(SingleChar)
      pos = pos + 1
    Loop
    trs("Branch") = Complete
    trs("Date") = rsTemp("Date")
   rsTemp.MoveNext
  Loop
  If trs.RecordCount > 0 Then
    trs.MoveFirst
  End If
  Set CreateFrontier = trs
Exit Function
EH:
 MsgBox Err.Description & " in CreateFrontier."
End Function
'
Private Function CallOrNot(AccusedDate As Date, DaysUntilCall As Integer) As Boolean
  On Error GoTo EH
  '
  Dim Difference As Integer
  Dim Today As Date
  '
  Today = Date
  Difference = DateDiff("d", AccusedDate, Today)
  CallOrNot = False
  '
  If Difference = DaysUntilCall Then CallOrNot = True
  '
  If Difference = DaysUntilCall + 2 And Weekday(Today) = vbMonday Then
  '
    CallOrNot = True
  '
  End If
  '
  If Difference = DaysUntilCall + 1 And Weekday(Today) = vbMonday Then
   '
   CallOrNot = True
   '
  End If
  '
Exit Function
EH:
  MsgBox Err.Description & " in CallOrNot."
End Function
'
Public Sub FillList(rs As Recordset, list As ListView)
  'On Error GoTo EH
  '
  Dim iSortOrder As Integer
  Dim lSortkey As Long
  '
  Dim LineCount As Long
  Dim FieldPos As Long
  Dim TotalCharacters As Long
  Dim sKey As String
  Dim iRecCount As Integer
  '
  TotalCharacters = 0
  LineCount = 0
  FieldPos = 2
  '
  iSortOrder = list.SortOrder
  lSortkey = list.SortKey
  list.SortKey = 0
  list.Sorted = False
  With rs
  '
    If .RecordCount > 0 Then
    '
      Do
      '
        iRecCount = 0
        TotalCharacters = 0
        .MoveFirst
        Do
        '
          If Not IsNull(.Fields(FieldPos)) Then
          '
            'TotalCharacters = Len(CStr(.Fields(FieldPos))) + TotalCharacters
            If Len(CStr(.Fields(FieldPos))) > 0 Then
              TotalCharacters = Len(CStr(.Fields(FieldPos))) + TotalCharacters
              iRecCount = iRecCount + 1
            End If
          '
          End If
          '
          .MoveNext
        '
        Loop Until .eof
        '
        If TotalCharacters = 0 Then
          list.ColumnHeaders.Add , "w1" & FieldPos, .Fields(FieldPos).Name, (Len(CStr(.Fields(FieldPos).Name)) * 100)
        Else
          list.ColumnHeaders.Add , "w1" & FieldPos, .Fields(FieldPos).Name, 400 + ((TotalCharacters / iRecCount) * 100)
        End If
        'list.ColumnHeaders.Add , .Fields, .Fields(FieldPos).Name, 400 + ((TotalCharacters / .RecordCount) * 100)
        FieldPos = FieldPos + 1
        '
      Loop Until FieldPos = .Fields.Count
      '
      .MoveFirst
      '
      FieldPos = 2
      LineCount = 0
      Do Until .eof
      '
        'list.ListItems.Add , "r1" & LineCount, .Fields(FieldPos)
        sKey = !CompanyID & "A" & !ContactID
        list.ListItems.Add , sKey, .Fields(FieldPos)
        FieldPos = FieldPos + 1
        '
        Do
'        '
        ' If Not IsNull(.Fields(FieldPos)) Then
'         '
         ' ListItems.Item(sKey).SubItems(FieldPos - 2) = .Fields(FieldPos)
          list.ListItems.Item(sKey).ListSubItems.Add , , .Fields(FieldPos) & vbNullString '.ForeColor = lColor
'
'
       '  End If
'         '
         FieldPos = FieldPos + 1
'         '
        Loop Until FieldPos = .Fields.Count
        '
        FieldPos = 2
        .MoveNext
        LineCount = LineCount + 1
      '
      Loop
      '
    End If
    '
  End With
  '
  list.SortOrder = iSortOrder
  list.SortKey = lSortkey
Exit Sub
EH:
 MsgBox Err.Description & " in FillList."
End Sub

Private Function isNumber(Value As String) As Boolean
  On Error GoTo EH
  '
  isNumber = False
  '
  If Value = "0" Then isNumber = True
  If Value = "1" Then isNumber = True
  If Value = "2" Then isNumber = True
  If Value = "3" Then isNumber = True
  If Value = "4" Then isNumber = True
  If Value = "5" Then isNumber = True
  If Value = "6" Then isNumber = True
  If Value = "7" Then isNumber = True
  If Value = "8" Then isNumber = True
  If Value = "9" Then isNumber = True
  '
Exit Function
EH:
 MsgBox Err.Description & " in isNumber."
End Function
'
Private Sub CombineRS(RSFirst As Recordset, RSSecond As Recordset)
  '
  On Error GoTo EH
  If RSSecond.RecordCount > 0 Then
'   RSSecond.MoveFirst
  End If
  '
  Do Until RSSecond.eof
    '
    RSFirst.AddNew
    RSFirst!FirstName = RSSecond!FirstName
    RSFirst!LastName = RSSecond!LastName
    RSFirst!Company = RSSecond!Company
    RSFirst!Phone1 = RSSecond!Phone1
'   RSFirst!Status = RSSecond!Status
    RSFirst!Type = RSSecond!Type
    RSFirst!DaysSinceEvent = RSSecond!DaysSinceEvent
    RSFirst.Update
    RSSecond.MoveNext
    '
  Loop
  '
  If RSSecond.RecordCount > 0 Then
    RSSecond.MoveFirst
  End If
  '
  If RSFirst.RecordCount > 0 Then
    RSFirst.MoveFirst
  End If
 '
 Exit Sub
EH:
 MsgBox Err.Description & " in CombineRS."
End Sub

