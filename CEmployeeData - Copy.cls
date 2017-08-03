VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployeeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type eData
  EmployeeID As Integer
  EmployeeNumber As String
  EmployeeLast As String
  EmployeeFirst As String
  EmployeeMiddle As String
  EMailAddress As String
  Password As String
  Groups As Integer
  SecurityLevel As Integer
  EmployeeExt As Integer
  WorkGroups As Integer
End Type

Private Data As eData

Public Property Get EmployeeID() As Integer
  EmployeeID = Data.EmployeeID
End Property

Public Property Let EmployeeID(ByVal Value As Integer)
  Data.EmployeeID = Value
End Property

Public Property Get EmployeeNumber() As String
  EmployeeNumber = Data.EmployeeNumber
End Property

Public Property Let EmployeeNumber(ByVal Value As String)
  Data.EmployeeNumber = Value
End Property

Public Property Get EmployeeLast() As String
  EmployeeLast = Data.EmployeeLast
End Property

Public Property Let EmployeeLast(ByVal Value As String)
  Data.EmployeeLast = Value
End Property

Public Property Get EmployeeFirst() As String
  EmployeeFirst = Data.EmployeeFirst
End Property

Public Property Let EmployeeFirst(ByVal Value As String)
  Data.EmployeeFirst = Value
End Property

Public Property Get EmployeeMiddle() As String
  EmployeeMiddle = Data.EmployeeMiddle
End Property

Public Property Let EmployeeMiddle(ByVal Value As String)
  Data.EmployeeMiddle = Value
End Property
Public Property Get Password() As String
  Password = Data.Password
End Property

Public Property Let Password(ByVal Value As String)
  Data.Password = Value
End Property
Public Property Get EMailAddress() As String
  EMailAddress = Data.EMailAddress
End Property

Public Property Let EMailAddress(ByVal Value As String)
  Data.EMailAddress = Value
End Property

Public Property Get Groups() As Integer
  Groups = Data.Groups
End Property

Public Property Let Groups(ByVal Value As Integer)
  Data.Groups = Value
End Property

Public Property Get SecurityLevel() As Integer
  SecurityLevel = Data.SecurityLevel
End Property

Public Property Let SecurityLevel(ByVal Value As Integer)
  Data.SecurityLevel = Value
End Property

Public Property Get EmployeeExt() As Integer
  EmployeeExt = Data.EmployeeExt
End Property

Public Property Let EmployeeExt(ByVal Value As Integer)
  Data.EmployeeExt = Value
End Property

Public Property Get WorkGroups() As Integer
  WorkGroups = Data.WorkGroups
End Property

Public Property Let WorkGroups(ByVal Value As Integer)
  Data.WorkGroups = Value
End Property
  
