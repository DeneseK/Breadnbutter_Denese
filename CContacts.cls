VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The private collection object
Private m_CContacts As Collection


Private Sub Class_Initialize()

  Set m_CContacts = New Collection

End Sub


' This sub adds a new CContactData item to the collection.
Sub Add(Item As CContactData, Optional Key As Variant)

  'TODO: Initialize the new CContactData item's properties here

  m_CContacts.Add Item ', Key

End Sub


' This sub remove an item from the collection.
Sub Remove(Index As Variant)

  m_CContacts.Remove Index

End Sub


' This function returns a CContactData item from the collection. It's the default method.
Function Item(Index As Variant) As CContactData
Attribute Item.VB_UserMemId = 0

  Set Item = m_CContacts.Item(Index)

End Function


' This property returns the number of items in the collection.
Property Get Count() As Long

  Count = m_CContacts.Count

End Property


' This sub remove all items from the collection.
Sub Clear()

  Set m_CContacts = New Collection

End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

  Set NewEnum = m_CContacts.[_NewEnum]

End Function



