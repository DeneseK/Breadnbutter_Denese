VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eMode
  None
  AddNewRecord
End Enum

Private iMode As Integer

Public Enum eRecord
  FirstRecord
  PreviousRecord
  NextRecord
  LastRecord
End Enum

Public Event Fetch()
Public Event Save(ByRef Success As Boolean)
Public Event AddNew(ByRef Success As Boolean)
Public Event Delete()
Public Event Edit()
Public Event MoveRecord(Record As Integer)
Public Event MoveFirst()
Public Event MovePrevious()
Public Event MoveNext()
Public Event MoveLast()
Public Event PrintRecord()
Public Event Sort(pbAsc As Boolean)
Public Event FindRecord(Where As String)
Public Event ClearControls()
Public Event Read()
Public Event Changed(ByRef fChanged As Boolean)
Public Event Enable(pfEnable As Boolean)

Public Function Changed() As Boolean
  Dim fChanged As Boolean
  RaiseEvent Changed(fChanged)
  Changed = fChanged
End Function

Public Sub Enable(pfEnable As Boolean)
  RaiseEvent Enable(pfEnable)
End Sub

Public Function Save() As Boolean
  Dim fSuccess As Boolean
  RaiseEvent Save(fSuccess)
  Save = fSuccess
End Function

Public Function AddNew() As Boolean
  Dim fSuccess As Boolean
  RaiseEvent AddNew(fSuccess)
  AddNew = fSuccess
End Function

Public Sub Delete()
  RaiseEvent Delete
End Sub

Public Sub MoveRecord(peRecord As eRecord)
  RaiseEvent MoveRecord(CInt(peRecord))
End Sub

Public Sub MoveFirst()
  RaiseEvent MoveFirst
End Sub

Public Sub MovePrevious()
  RaiseEvent MovePrevious
End Sub

Public Sub MoveNext()
  RaiseEvent MoveNext
End Sub

Public Sub MoveLast()
  RaiseEvent MoveLast
End Sub

Public Property Get Mode() As eMode
  Mode = iMode
End Property

Public Property Let Mode(ByVal peMode As eMode)
  iMode = peMode
End Property

Public Sub Edit()
  RaiseEvent Edit
End Sub

Public Sub PrintRecord()
  RaiseEvent PrintRecord
End Sub

Public Sub Sort(pbAsc As Boolean)
  RaiseEvent Sort(pbAsc)
End Sub

Public Sub FindRecord(psWhere As String)
  RaiseEvent FindRecord(psWhere)
End Sub

Public Sub ClearControls()
  RaiseEvent ClearControls
End Sub

Public Sub Read()
  RaiseEvent Read
End Sub

Public Sub Fetch()
  RaiseEvent Fetch
End Sub

'Public Sub GetFilterTable(pdb As Database, psTable As String)
'  Dim db As Database
'  Dim sTable As String
'  RaiseEvent GetFilterTable(db, sTable)
'  Set pdb = db
'  psTable = sTable
'End Sub

'Public Sub SetFilter(Filter As clsFilterCriteria)
'  RaiseEvent SetFilter(Filter)
'End Sub
'
'Public Sub ClearFilter()
'  RaiseEvent ClearFilter
'End Sub
