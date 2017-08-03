VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCompanys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The private collection object
Private m_CCompanys As Collection


Private Sub Class_Initialize()

    Set m_CCompanys = New Collection

End Sub


' This sub adds a new CCompanyData item to the collection.
Sub Add(Item As CCompanyData, Optional Key As Variant)

    'TODO: Initialize the new CCompanyData item's properties here

    m_CCompanys.Add Item ', Key

End Sub


' This sub remove an item from the collection.
Sub Remove(Index As Variant)

    m_CCompanys.Remove Index

End Sub


' This function returns a CCompanyData item from the collection. It's the default method.
Function Item(Index As Variant) As CCompanyData

    Set Item = m_CCompanys.Item(Index)

End Function


' This property returns the number of items in the collection.
Property Get Count() As Long

    Count = m_CCompanys.Count

End Property


' This sub remove all items from the collection.
Sub Clear()

    Set m_CCompanys = New Collection

End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown

    Set NewEnum = m_CCompanys.[_NewEnum]

End Function



