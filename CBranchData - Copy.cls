VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBranchData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TBranch
  BranchID As Long
  CompanyID As Long
  Name As String
  Number As String
  ManagerFirstName As String
  ManagerLastName As String
  Address1 As String
  Address2 As String
  Address3 As String
  Zip As String
  State As String
  City As String
  PhoneNumber As String
  FaxNumber As String
  Email As String
End Type
'
Private r() As TBranch

Private Sub Class_Initialize()
  On Error Resume Next
  ReDim r(0)
End Sub

Public Property Get BranchID() As Long
  BranchID = r(0).BranchID
End Property

Public Property Let BranchID(ByVal iNewValue As Long)
  r(0).BranchID = iNewValue
End Property

Public Property Get CompanyID() As Long
  CompanyID = r(0).CompanyID
End Property

Public Property Let CompanyID(ByVal lNewValue As Long)
  r(0).CompanyID = lNewValue
End Property

Public Property Get Name() As String
  Name = r(0).Name
End Property

Public Property Let Name(ByVal lNewValue As String)
  r(0).Name = lNewValue
End Property

Public Property Get Number() As String
  Number = r(0).Number
End Property

Public Property Let Number(ByVal lNewValue As String)
  r(0).Number = lNewValue
End Property

Public Property Get ManagerFirstName() As String
  ManagerFirstName = r(0).ManagerFirstName
End Property

Public Property Let ManagerFirstName(ByVal lNewValue As String)
  r(0).ManagerFirstName = lNewValue
End Property

Public Property Get ManagerLastName() As String
  ManagerLastName = r(0).ManagerLastName
End Property

Public Property Let ManagerLastName(ByVal lNewValue As String)
  r(0).ManagerLastName = lNewValue
End Property

Public Property Get Address1() As String
  Address1 = r(0).Address1
End Property

Public Property Let Address1(ByVal lNewValue As String)
  r(0).Address1 = lNewValue
End Property

Public Property Get Address2() As String
  Address2 = r(0).Address2
End Property

Public Property Let Address2(ByVal lNewValue As String)
  r(0).Address2 = lNewValue
End Property

Public Property Get Address3() As String
  Address3 = r(0).Address3
End Property

Public Property Let Address3(ByVal lNewValue As String)
  r(0).Address3 = lNewValue
End Property

Public Property Get Zip() As String
  Zip = r(0).Zip
End Property

Public Property Let Zip(ByVal lNewValue As String)
  r(0).Zip = lNewValue
End Property

Public Property Get State() As String
  State = r(0).State
End Property

Public Property Let State(ByVal lNewValue As String)
  r(0).State = lNewValue
End Property

Public Property Get City() As String
  City = r(0).City
End Property

Public Property Let City(ByVal lNewValue As String)
  r(0).City = lNewValue
End Property

Public Property Get PhoneNumber() As String
  PhoneNumber = r(0).PhoneNumber
End Property

Public Property Let PhoneNumber(ByVal lNewValue As String)
  r(0).PhoneNumber = lNewValue
End Property

Public Property Get FaxNumber() As String
  FaxNumber = r(0).FaxNumber
End Property

Public Property Let FaxNumber(ByVal lNewValue As String)
  r(0).FaxNumber = lNewValue
End Property

Public Property Get Email() As String
  Email = r(0).Email
End Property

Public Property Let Email(ByVal lNewValue As String)
  r(0).Email = lNewValue
End Property

