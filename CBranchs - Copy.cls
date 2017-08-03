VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBranchs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The private collection object
Private m_CBranchs As Collection


Private Sub Class_Initialize()

    Set m_CBranchs = New Collection

End Sub


' This sub adds a new CBranchData item to the collection.
Sub Add(Item As CBranchData, Optional Key As Variant)

    'TODO: Initialize the new CBranchData item's properties here

    m_CBranchs.Add Item ', Key

End Sub


' This sub remove an item from the collection.
Sub Remove(Index As Variant)

    m_CBranchs.Remove Index

End Sub


' This function returns a CBranchData item from the collection. It's the default method.
Function Item(Index As Variant) As CBranchData
Attribute Item.VB_UserMemId = 0

    Set Item = m_CBranchs.Item(Index)

End Function


' This property returns the number of items in the collection.
Property Get Count() As Long

    Count = m_CBranchs.Count

End Property


' This sub remove all items from the collection.
Sub Clear()

    Set m_CBranchs = New Collection

End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Set NewEnum = m_CBranchs.[_NewEnum]

End Function



