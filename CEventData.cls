VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEventData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TEvent
  RecID              As Long
  CustRecID As Long
  EventDate As Date
  EventTime As Date
  EventType As String
  EventResults As String
  EventUser As String
  EventSubject As String
  ProductID As Integer
  ClosedTime As Date
  OpenCall As Boolean
  Sticky As Boolean
End Type

Private r() As TEvent

Private Sub Class_Initialize()
  On Error Resume Next
  ReDim r(0)
End Sub

Public Property Get RecID() As Long
  RecID = r(0).RecID
End Property

Public Property Let RecID(ByVal NewValue As Long)
  r(0).RecID = NewValue
End Property

Public Property Get CustRecID() As Long
  CustRecID = r(0).CustRecID
End Property

Public Property Let CustRecID(ByVal NewValue As Long)
  r(0).CustRecID = NewValue
End Property

Public Property Get EventDate() As Date
  EventDate = r(0).EventDate
End Property

Public Property Let EventDate(ByVal NewValue As Date)
  r(0).EventDate = NewValue
End Property

Public Property Get EventTime() As Date
  EventTime = r(0).EventTime
End Property

Public Property Let EventTime(ByVal NewValue As Date)
  r(0).EventTime = NewValue
End Property

Public Property Get EventType() As String
  EventType = r(0).EventType
End Property

Public Property Let EventType(ByVal NewValue As String)
  r(0).EventType = NewValue
End Property

Public Property Get EventResults() As String
  EventResults = r(0).EventResults
End Property

Public Property Let EventResults(ByVal NewValue As String)
  r(0).EventResults = NewValue
End Property

Public Property Get EventUser() As String
  EventUser = r(0).EventUser
End Property

Public Property Let EventUser(ByVal NewValue As String)
  r(0).EventUser = NewValue
End Property

Public Property Get EventSubject() As String
  EventSubject = r(0).EventSubject
End Property

Public Property Let EventSubject(ByVal NewValue As String)
  r(0).EventSubject = NewValue
End Property

Public Property Get ProductID() As Integer
  ProductID = r(0).ProductID
End Property

Public Property Let ProductID(ByVal NewValue As Integer)
  r(0).ProductID = NewValue
End Property

Public Property Get OpenCall() As Boolean
  OpenCall = r(0).OpenCall
End Property

Public Property Let OpenCall(ByVal NewValue As Boolean)
  r(0).OpenCall = NewValue
End Property

Public Property Get Sticky() As Boolean
  Sticky = r(0).Sticky
End Property

Public Property Let Sticky(ByVal NewValue As Boolean)
  r(0).Sticky = NewValue
End Property

Public Property Get ClosedTime() As Date
  ClosedTime = r(0).ClosedTime
End Property

Public Property Let ClosedTime(ByVal NewValue As Date)
  r(0).ClosedTime = NewValue
End Property
