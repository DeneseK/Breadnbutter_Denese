VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProductData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TProduct
  ProductID As Long
  Product As String
  Color As Long
  Seed1 As Integer
  Seed2 As Integer
End Type
'
Private r() As TProduct

Private Sub Class_Initialize()
  On Error Resume Next
  ReDim r(0)
End Sub

Public Property Get ProductID() As Long
  ProductID = r(0).ProductID
End Property

Public Property Let ProductID(ByVal iNewValue As Long)
  r(0).ProductID = iNewValue
End Property

Public Property Get Product() As String
  Product = r(0).Product
End Property

Public Property Let Product(ByVal lNewValue As String)
  r(0).Product = lNewValue
End Property

Public Property Get Color() As Long
  Color = r(0).Color
End Property

Public Property Let Color(ByVal lNewValue As Long)
  r(0).Color = lNewValue
End Property

Public Property Get Seed1() As Integer
  Seed1 = r(0).Seed1
End Property

Public Property Let Seed1(ByVal iNewValue As Integer)
  r(0).Seed1 = iNewValue
End Property

Public Property Get Seed2() As Integer
  Seed2 = r(0).Seed2
End Property

Public Property Let Seed2(ByVal iNewValue As Integer)
  r(0).Seed2 = iNewValue
End Property

