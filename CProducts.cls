VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The private collection object
Private m_CProducts As Collection


Private Sub Class_Initialize()

  Set m_CProducts = New Collection

End Sub


' This sub adds a new CProductData item to the collection.
Sub Add(Item As CProductData, Optional Key As Variant)

  'TODO: Initialize the new CProductData item's properties here

  m_CProducts.Add Item ', Key

End Sub


' This sub remove an item from the collection.
Sub Remove(Index As Variant)

  m_CProducts.Remove Index

End Sub


' This function returns a CProductData item from the collection. It's the default method.
Function Item(Index As Variant) As CProductData
Attribute Item.VB_UserMemId = 0

  Set Item = m_CProducts.Item(Index)

End Function


' This property returns the number of items in the collection.
Property Get Count() As Long

  Count = m_CProducts.Count

End Property


' This sub remove all items from the collection.
Sub Clear()

  Set m_CProducts = New Collection

End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

  Set NewEnum = m_CProducts.[_NewEnum]

End Function



