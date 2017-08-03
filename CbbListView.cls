VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CbbListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ModuleName As String = "CbbListView"

Public Sub Sort(plvwControl As Bblistview1, plColumnIndex As Long, _
  Optional psSortOrder As String)
      '
      '\\ Assumptions
      '\\ PictureIndex 1: "Ascending" Arrow
      '\\ PictureIndex 2: "Descending" Arrow
      '
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim lColIdx As Long
110   Dim lvwCur  As Bblistview1
      '
120   Set lvwCur = plvwControl
130   lColIdx = plColumnIndex
140   psSortOrder = LCase(psSortOrder)
      '
150   lvwCur.LockState = True
      '
160   With lvwCur.ColumnHeaders(lColIdx)
        '
170     If .ContentType = 1 Then
180       lvwCur.ColumnHeaders(lvwCur.KeyPressBoundColumn).ContentType = 1
190       lvwCur.KeyPressBoundColumn = lColIdx
200       .ContentType = 3
210       If psSortOrder <> vbNullString Then
220         .PictureIndex = IIf(psSortOrder = "ascending", 1, 2)
230       End If
240     Else
250       If psSortOrder = vbNullString Then
260         If .PictureIndex = 1 Then
270           .PictureIndex = 2
280         Else
290           .PictureIndex = 1
300         End If
310       Else
320         .PictureIndex = IIf(psSortOrder = "ascending", 1, 2)
330       End If
340     End If
         '
350     lvwCur.SortListItems (lColIdx * IIf(.PictureIndex = 1, 1, -1))
         '
360   End With
       '
370   lvwCur.LockState = False
       '
       '\\ Deallocate Resources
380   Set lvwCur = Nothing
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise ModuleName, "Sort", Err.number, Err.Description, Erl
      '</EhFooter>
End Sub
Public Sub AddMetric(plvwControl As Bblistview1, iColumnIndex, lMetric As Long)
  '\\ Local Declarations
  'Dim iColIdx As Integer
  'Dim iRowIdx As Integer
  'Dim lvwCur  As Bblistview1
  '
  'iColIdx = iColumnIndex
  'Set lvwCur = plvwControl
  '
  'With lvwCur.ListItem(iRowIdx, iColIdex)
  '  .Text = FormatMetric(lMetric)
  '  .ItemData = lMetric
  'End With
  '
  '\\ Deallocate Resources
  'Set lvwCur = Nothing
  '<EhHeader>
  On Error GoTo EH
  '
  '</EhHeader>
  '<EhFooter>
  '
  Exit Sub
EH:
  ErrorMgr.Raise ModuleName, "AddMetric", Err.number, Err.Description, Erl
  '</EhFooter>
End Sub
