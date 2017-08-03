VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroupListLinkData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TGroupListLink
  ID              As Long
  ContactID       As Long
  ListID          As Long
End Type

Private r() As TGroupListLink

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

Public Property Get ContactID() As Long
  ContactID = r(0).ContactID
End Property

Public Property Let ContactID(ByVal NewValue As Long)
  r(0).ContactID = NewValue
End Property

Public Property Get ListID() As Long
  ListID = r(0).ListID
End Property

Public Property Let ListID(ByVal NewValue As Long)
  r(0).ListID = NewValue
End Property


