VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The private collection object
Private m_CEmployees As Collection


Private Sub Class_Initialize()

  Set m_CEmployees = New Collection

End Sub


' This sub adds a new CEmployeeData item to the collection.
Sub Add(Item As CEmployeeData, Optional Key As Variant)

  'TODO: Initialize the new CEmployeeData item's properties here

  m_CEmployees.Add Item ', Key

End Sub


' This sub remove an item from the collection.
Sub Remove(Index As Variant)

  m_CEmployees.Remove Index

End Sub


' This function returns a CEmployeeData item from the collection. It's the default method.
Function Item(Index As Variant) As CEmployeeData
Attribute Item.VB_UserMemId = 0

  Set Item = m_CEmployees.Item(Index)

End Function


' This property returns the number of items in the collection.
Property Get Count() As Long

  Count = m_CEmployees.Count

End Property


' This sub remove all items from the collection.
Sub Clear()

  Set m_CEmployees = New Collection

End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

  Set NewEnum = m_CEmployees.[_NewEnum]

End Function



