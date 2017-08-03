VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroupListDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The private collection object
Private m_CGroupListDatas As Collection


Private Sub Class_Initialize()

  Set m_CGroupListDatas = New Collection

End Sub


' This sub adds a new CGroupListData item to the collection.
Sub Add(Item As CGroupListData, Optional Key As Variant)

  'TODO: Initialize the new CGroupListData item's properties here

  m_CGroupListDatas.Add Item, Key

End Sub


' This sub remove an item from the collection.
Sub Remove(Index As Variant)

  m_CGroupListDatas.Remove Index

End Sub


' This function returns a CGroupListData item from the collection. It's the default method.
Function Item(Index As Variant) As CGroupListData
Attribute Item.VB_UserMemId = 0

  Set Item = m_CGroupListDatas.Item(Index)

End Function


' This property returns the number of items in the collection.
Property Get Count() As Long

  Count = m_CGroupListDatas.Count

End Property


' This sub remove all items from the collection.
Sub Clear()

  Set m_CGroupListDatas = New Collection

End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

  Set NewEnum = m_CGroupListDatas.[_NewEnum]

End Function



