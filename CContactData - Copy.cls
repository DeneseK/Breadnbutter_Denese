VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContactData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TContact
  ID As Long
  DateEntered As Date
  LastUpdate As Date
  CompanyID As Long
  BranchID As Long
  FirstName As String
  LastName As String
  Salutation As String
  Title As String
  Address1 As String
  Address2 As String
  City As String
  State As String
  Zip As String
  MailAddress1 As String
  MailAddress2 As String
  MailCity As String
  MailState As String
  MailZip As String
  PCEmail As String
  PCEmailPassword As String
  Phone1 As String
  Phone2 As String
  Fax As String
  Notes As String
  Email As String
  Selected As Byte 'Betatester
  Source As String
  Status As String
  AuthStatus As String
  AuthDate As Date
  AuthDays As Integer
  AuthRemaining As Long
  Version As String
  VersionShipped As String
  Copies As Integer
  ShipStatus As String
  ShipDate As Date
  Rate As Currency
  ContactType As Integer
  RateExpDate As Date
  PreferredAddress As Integer
  PVAuthStatus As String
  PVDownloadStatus As String
  DownloadStatus As String
  PVAuthDate As Date
  PVAuthDays As Integer
  PVAuthRemaining As Long
  PVVersion As String
  PVVersionShipped As String
  PVCopies As Integer
  PVShipStatus As String
  PVShipDate As Date
  DownloadDate As Date
  PVDownloadDate As Date
  '
  DaysPending As Integer
  GraceDays As Integer
  OnlineAuths As Integer
  SaleDate As Date
  SaleDays As Integer
  PVDaysPending As Integer
  PVGraceDays As Integer
  PVOnlineAuths As Integer
  PVSaleDate As Date
  PVSaleDays As Integer
  WebPassword As String
  ContactByEmail As Boolean
  ChangedData As Integer
  AdjusterID As String
End Type

Private r() As TContact

Private Sub Class_Initialize()
  On Error Resume Next
  ReDim r(0)
End Sub

Public Property Get DaysPending() As Integer
  DaysPending = r(0).DaysPending
End Property

Public Property Let DaysPending(ByVal Value As Integer)
  r(0).DaysPending = Value
End Property


Public Property Get GraceDays() As Integer
  GraceDays = r(0).GraceDays
End Property

Public Property Let GraceDays(ByVal Value As Integer)
  r(0).GraceDays = Value
End Property

Public Property Get OnlineAuths() As Integer
  OnlineAuths = r(0).OnlineAuths
End Property

Public Property Let OnlineAuths(ByVal Value As Integer)
  r(0).OnlineAuths = Value
End Property

Public Property Get SaleDate() As Date
  SaleDate = r(0).SaleDate
End Property

Public Property Let SaleDate(ByVal NewValue As Date)
  r(0).SaleDate = NewValue
End Property

Public Property Get SaleDays() As Integer
  SaleDays = r(0).SaleDays
End Property

Public Property Let SaleDays(ByVal Value As Integer)
  r(0).SaleDays = Value
End Property

Public Property Get PVDaysPending() As Integer
  PVDaysPending = r(0).PVDaysPending
End Property

Public Property Let PVDaysPending(ByVal Value As Integer)
  r(0).PVDaysPending = Value
End Property

Public Property Get PVGraceDays() As Integer
  PVGraceDays = r(0).PVGraceDays
End Property

Public Property Let PVGraceDays(ByVal Value As Integer)
  r(0).PVGraceDays = Value
End Property

Public Property Get PVOnlineAuths() As Integer
  PVOnlineAuths = r(0).PVOnlineAuths
End Property

Public Property Let PVOnlineAuths(ByVal Value As Integer)
  r(0).PVOnlineAuths = Value
End Property

Public Property Get PVSaleDate() As Date
  PVSaleDate = r(0).PVSaleDate
End Property

Public Property Let PVSaleDate(ByVal NewValue As Date)
  r(0).PVSaleDate = NewValue
End Property

Public Property Get PVSaleDays() As Integer
  PVSaleDays = r(0).PVSaleDays
End Property

Public Property Let PVSaleDays(ByVal Value As Integer)
  r(0).PVSaleDays = Value
End Property

Public Property Get WebPassword() As String
  WebPassword = r(0).WebPassword
End Property

Public Property Let WebPassword(ByVal Value As String)
  r(0).WebPassword = Value
End Property

Public Property Get ContactByEmail() As Boolean
  ContactByEmail = r(0).ContactByEmail
End Property

Public Property Let ContactByEmail(ByVal NewValue As Boolean)
  r(0).ContactByEmail = NewValue
End Property

Public Property Get ChangedData() As Integer
  ChangedData = r(0).ChangedData
End Property

Public Property Let ChangedData(ByVal Value As Integer)
  r(0).ChangedData = Value
End Property

Public Property Get PCEmailPassword() As String
  PCEmailPassword = r(0).PCEmailPassword
End Property

Public Property Let PCEmailPassword(ByVal Value As String)
  r(0).PCEmailPassword = Value
End Property

Public Property Get PCEmail() As String
  PCEmail = r(0).PCEmail
End Property

Public Property Let PCEmail(ByVal Value As String)
  r(0).PCEmail = Value
End Property

Public Property Get Rate() As Currency
  Rate = r(0).Rate
End Property

Public Property Let Rate(ByVal Value As Currency)
  r(0).Rate = Value
End Property

Public Property Get Adding() As Boolean
  Adding = fAdding
End Property

Public Property Get Loaded() As Boolean
  Loaded = fLoaded
End Property

Public Property Get ID() As Long
  ID = r(0).ID
End Property

Public Property Let ID(ByVal NewValue As Long)
  r(0).ID = NewValue
End Property

'Public Property Get DateEntered() As Date
'  DateEntered = r(0).DateEntered
'End Property
'
'Public Property Let DateEntered(ByVal NewValue As Date)
'  r(0).DateEntered = NewValue
'End Property

'Public Property Get LastUpdate() As Date
'   LastUpdate = r(0).LastUpdate
'End Property
'
'Public Property Let LastUpdate(ByVal NewValue As Date)
'  r(0).LastUpdate = NewValue
'End Property

Public Property Get CompanyID() As Long
  CompanyID = r(0).CompanyID
End Property

Public Property Let CompanyID(ByVal NewValue As Long)
  r(0).CompanyID = NewValue
End Property

Public Property Get BranchID() As Long
  BranchID = r(0).BranchID
End Property

Public Property Let BranchID(ByVal NewValue As Long)
  r(0).BranchID = NewValue
End Property

Public Property Get FirstName() As String
  FirstName = r(0).FirstName
End Property

Public Property Let FirstName(ByVal NewValue As String)
  r(0).FirstName = NewValue
End Property

Public Property Get LastName() As String
  LastName = r(0).LastName
End Property

Public Property Let LastName(ByVal NewValue As String)
  r(0).LastName = NewValue
End Property

Public Property Get Salutation() As String
  Salutation = r(0).Salutation
End Property

Public Property Let Salutation(ByVal NewValue As String)
  r(0).Salutation = NewValue
End Property

Public Property Get Title() As String
  Title = r(0).Title
End Property

Public Property Let Title(ByVal NewValue As String)
  r(0).Title = NewValue
End Property

Public Property Get Address1() As String
  Address1 = r(0).Address1
End Property

Public Property Let Address1(ByVal NewValue As String)
  r(0).Address1 = NewValue
End Property

Public Property Get Address2() As String
  Address2 = r(0).Address2
End Property

Public Property Let Address2(ByVal NewValue As String)
  r(0).Address2 = NewValue
End Property

Public Property Get City() As String
  City = r(0).City
End Property

Public Property Let City(ByVal NewValue As String)
  r(0).City = NewValue
End Property

Public Property Get State() As String
  State = r(0).State
End Property

Public Property Let State(ByVal NewValue As String)
  r(0).State = NewValue
End Property

Public Property Get Zip() As String
  Zip = r(0).Zip
End Property

Public Property Let Zip(ByVal NewValue As String)
  r(0).Zip = NewValue
End Property

Public Property Get MailAddress1() As String
  MailAddress1 = r(0).MailAddress1
End Property

Public Property Let MailAddress1(ByVal NewValue As String)
  r(0).MailAddress1 = NewValue
End Property

Public Property Get MailAddress2() As String
  MailAddress2 = r(0).MailAddress2
End Property

Public Property Let MailAddress2(ByVal NewValue As String)
  r(0).MailAddress2 = NewValue
End Property

Public Property Get MailCity() As String
  MailCity = r(0).MailCity
End Property

Public Property Let MailCity(ByVal NewValue As String)
  r(0).MailCity = NewValue
End Property

Public Property Get MailState() As String
  MailState = r(0).MailState
End Property

Public Property Let MailState(ByVal NewValue As String)
  r(0).MailState = NewValue
End Property

Public Property Get MailZip() As String
  MailZip = r(0).MailZip
End Property

Public Property Let MailZip(ByVal NewValue As String)
  r(0).MailZip = NewValue
End Property

Public Property Get Phone1() As String
  If IsNumeric(r(0).Phone1) Then
    Phone1 = r(0).Phone1
  Else
    Phone1 = StripChars(Phone1) 'Replace(Replace(Replace(r(0).Phone1, "-", vbNullString), "_", vbNullString), "x", vbNullString)
  End If
End Property

Public Property Let Phone1(ByVal NewValue As String)
  r(0).Phone1 = NewValue
End Property

Public Property Get Phone2() As String
  If IsNumeric(r(0).Phone2) Then
    Phone2 = r(0).Phone2
  Else
    Phone2 = StripChars(Phone2) 'Replace(Replace(Replace(r(0).Phone2, "-", vbNullString), "_", vbNullString), "x", vbNullString)
  End If
End Property

Public Property Let Phone2(ByVal NewValue As String)
  r(0).Phone2 = NewValue
End Property

Public Property Get Fax() As String
  If IsNumeric(r(0).Fax) Then
    Fax = r(0).Fax
  Else
    Fax = StripChars(Fax) 'Replace(Replace(Replace(r(0).Fax, "-", vbNullString), "_", vbNullString), "x", vbNullString)
  End If
End Property

Public Property Let Fax(ByVal NewValue As String)
  r(0).Fax = NewValue
End Property

Public Property Get Notes() As String
  Notes = r(0).Notes
End Property

Public Property Let Notes(ByVal NewValue As String)
  r(0).Notes = NewValue
End Property

Public Property Get Email() As String
  Email = r(0).Email
End Property

Public Property Let Email(ByVal NewValue As String)
  r(0).Email = NewValue
End Property

Public Property Get Selected() As Byte
  Selected = r(0).Selected
End Property

Public Property Let Selected(ByVal NewValue As Byte)
  r(0).Selected = NewValue
End Property

Public Property Get PreferredAddress() As Integer
  PreferredAddress = r(0).PreferredAddress
End Property

Public Property Let PreferredAddress(ByVal NewValue As Integer)
  r(0).PreferredAddress = NewValue
End Property

Public Property Get Source() As String
  Source = r(0).Source
End Property

Public Property Let Source(ByVal NewValue As String)
  r(0).Source = NewValue
End Property

Public Property Get Status() As String
  Status = r(0).Status
End Property

Public Property Let Status(ByVal NewValue As String)
  r(0).Status = NewValue
End Property

Public Property Get AuthStatus() As String
  AuthStatus = r(0).AuthStatus
End Property

Public Property Let AuthStatus(ByVal NewValue As String)
  r(0).AuthStatus = NewValue
End Property

Public Property Get AuthDate() As Date
  AuthDate = r(0).AuthDate
End Property

Public Property Let AuthDate(ByVal NewValue As Date)
  r(0).AuthDate = NewValue
End Property

Public Property Get AuthDays() As Integer
  AuthDays = r(0).AuthDays
End Property

Public Property Let AuthDays(ByVal NewValue As Integer)
  r(0).AuthDays = NewValue
End Property

Public Property Get AuthRemaining() As Long
  AuthRemaining = r(0).AuthRemaining
End Property

Public Property Let AuthRemaining(ByVal NewValue As Long)
  r(0).AuthRemaining = NewValue
End Property

Public Property Get Version() As String
  Version = r(0).Version
End Property

Public Property Let Version(ByVal NewValue As String)
  r(0).Version = NewValue
End Property

Public Property Get VersionShipped() As String
  VersionShipped = r(0).VersionShipped
End Property

Public Property Let VersionShipped(ByVal NewValue As String)
  r(0).VersionShipped = NewValue
End Property

Public Property Get Copies() As Integer
  Copies = r(0).Copies
End Property

Public Property Let Copies(ByVal NewValue As Integer)
  r(0).Copies = NewValue
End Property

Public Property Get ShipStatus() As String
  ShipStatus = r(0).ShipStatus
End Property

Public Property Let ShipStatus(ByVal NewValue As String)
  r(0).ShipStatus = NewValue
End Property

Public Property Get ShipDate() As Date
  ShipDate = r(0).ShipDate
End Property

Public Property Let ShipDate(ByVal NewValue As Date)
  r(0).ShipDate = NewValue
End Property

Public Property Get ContactType() As Integer
  ContactType = r(0).ContactType
End Property

Public Property Let ContactType(ByVal NewValue As Integer)
  r(0).ContactType = NewValue
End Property

Public Property Get RateExpDate() As Date
  RateExpDate = r(0).RateExpDate
End Property

Public Property Let RateExpDate(ByVal NewValue As Date)
  r(0).RateExpDate = NewValue
End Property

Public Property Get SearchID() As Long
  SearchID = lSearchID
End Property

Public Property Let SearchID(ByVal NewValue As Long)
  lSearchID = NewValue
End Property

Public Property Get PVAuthStatus() As String
  PVAuthStatus = r(0).PVAuthStatus
End Property

Public Property Let PVAuthStatus(ByVal NewValue As String)
  r(0).PVAuthStatus = NewValue
End Property

Public Property Get PVAuthDate() As Date
  PVAuthDate = r(0).PVAuthDate
End Property

Public Property Let PVAuthDate(ByVal NewValue As Date)
  r(0).PVAuthDate = NewValue
End Property

Public Property Get PVAuthDays() As Integer
  PVAuthDays = r(0).PVAuthDays
End Property

Public Property Let PVAuthDays(ByVal NewValue As Integer)
  r(0).PVAuthDays = NewValue
End Property

Public Property Get PVAuthRemaining() As Long
  PVAuthRemaining = r(0).PVAuthRemaining
End Property

Public Property Let PVAuthRemaining(ByVal NewValue As Long)
  r(0).PVAuthRemaining = NewValue
End Property

Public Property Get PVVersion() As String
  PVVersion = r(0).Version
End Property

Public Property Let PVVersion(ByVal NewValue As String)
  r(0).PVVersion = NewValue
End Property

Public Property Get PVVersionShipped() As String
  PVVersionShipped = r(0).PVVersionShipped
End Property

Public Property Let PVVersionShipped(ByVal NewValue As String)
  r(0).PVVersionShipped = NewValue
End Property

Public Property Get PVCopies() As Integer
  PVCopies = r(0).PVCopies
End Property

Public Property Let PVCopies(ByVal NewValue As Integer)
  r(0).PVCopies = NewValue
End Property

Public Property Get PVShipStatus() As String
  PVShipStatus = r(0).PVShipStatus
End Property

Public Property Let PVShipStatus(ByVal NewValue As String)
  r(0).PVShipStatus = NewValue
End Property

Public Property Get PVShipDate() As Date
  PVShipDate = r(0).PVShipDate
End Property

Public Property Let PVShipDate(ByVal NewValue As Date)
  r(0).PVShipDate = NewValue
End Property

Public Property Get DownloadDate() As Date
  DownloadDate = r(0).DownloadDate
End Property

Public Property Let DownloadDate(ByVal NewValue As Date)
  r(0).DownloadDate = NewValue
End Property

Public Property Get PVDownloadDate() As Date
  PVDownloadDate = r(0).PVDownloadDate
End Property

Public Property Let PVDownloadDate(ByVal NewValue As Date)
  r(0).PVDownloadDate = NewValue
End Property

Public Property Get PVDownloadStatus() As String
  PVDownloadStatus = r(0).PVDownloadStatus
End Property

Public Property Let PVDownloadStatus(ByVal NewValue As String)
  r(0).PVDownloadStatus = NewValue
End Property

Public Property Get DownloadStatus() As String
  DownloadStatus = r(0).DownloadStatus
End Property

Public Property Let DownloadStatus(ByVal NewValue As String)
  r(0).DownloadStatus = NewValue
End Property

Public Property Get AdjusterID() As String
  AdjsuterID = r(0).AdjusterID
End Property

Public Property Let AdjusterID(ByVal Value As String)
  r(0).AdjusterID = Value
End Property
