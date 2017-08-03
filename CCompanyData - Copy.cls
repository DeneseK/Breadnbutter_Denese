VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCompanyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TCompany
  ID              As Long
  DateEntered     As Date
  LastUpdate      As Date
  Name            As String
  Division        As String
  Individual      As Boolean
  DoNotContact    As Boolean
  Note            As String
  InterestRank    As Integer
End Type

Private r() As TCompany

Private Sub Class_Initialize()
  On Error Resume Next
  ReDim r(0)
End Sub

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
'Public Property Get LastUpdate() As Date
'  LastUpdate = r(0).LastUpdate
'End Property



Public Property Get Name() As String
  Name = r(0).Name
End Property

Public Property Let Name(ByVal NewValue As String)
  r(0).Name = NewValue
End Property
'
'Public Property Get Division() As String
'  Division = r(0).Division
'End Property

Public Property Let Office(ByVal NewValue As String)
  r(0).Division = NewValue
End Property

Public Property Get Individual() As Boolean
  Individual = r(0).Individual
End Property

Public Property Let Individual(ByVal NewValue As Boolean)
  r(0).Individual = NewValue
End Property

Public Property Get DoNotContact() As Boolean
  DoNotContact = r(0).DoNotContact
End Property

Public Property Let DoNotContact(ByVal NewValue As Boolean)
  r(0).DoNotContact = NewValue
End Property

Public Property Get Note() As String
  Note = r(0).Note
End Property

Public Property Let Note(ByVal NewValue As String)
  r(0).Note = NewValue
End Property

Public Property Get InterestRank() As Integer
  InterestRank = r(0).InterestRank
End Property

Public Property Let InterestRank(ByVal NewValue As Integer)
  r(0).InterestRank = NewValue
End Property
