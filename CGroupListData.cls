VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroupListData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TGroupList
  ID              As Long
  ListName        As String
  EmployeeID      As Long
End Type

Private r() As TGroupList

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

Public Property Get ListName() As String
  ListName = r(0).ListName
End Property

Public Property Let ListName(ByVal NewValue As String)
  r(0).ListName = NewValue
End Property

Public Property Get EmployeeID() As Long
  EmployeeID = r(0).EmployeeID
End Property

Public Property Let EmployeeID(ByVal NewValue As Long)
  r(0).EmployeeID = NewValue
End Property
