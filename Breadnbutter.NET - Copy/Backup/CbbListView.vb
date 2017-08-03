Option Strict Off
Option Explicit On
Friend Class CbbListView
	
	Private Const ModuleName As String = "CbbListView"
	
	Public Sub Sort(ByRef plvwControl As AxBBLISTVIEWLib.AxBblistview1, ByRef plColumnIndex As Integer, Optional ByRef psSortOrder As String = "")
		'
		'\\ Assumptions
		'\\ PictureIndex 1: "Ascending" Arrow
		'\\ PictureIndex 2: "Descending" Arrow
		'
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim lColIdx As Integer
110: Dim lvwCur As AxBBLISTVIEWLib.AxBblistview1
		'
120: lvwCur = plvwControl
130: lColIdx = plColumnIndex
140: psSortOrder = LCase(psSortOrder)
		'
150: lvwCur.LockState = True
		'
160: With lvwCur.get_ColumnHeaders(lColIdx)
			'
170: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders(lColIdx).ContentType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .ContentType = 1 Then
180: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders().ContentType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lvwCur.get_ColumnHeaders(lvwCur.KeyPressBoundColumn).ContentType = 1
190: lvwCur.KeyPressBoundColumn = lColIdx
200: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders().ContentType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.ContentType = 3
210: If psSortOrder <> vbNullString Then
220: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders().PictureIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.PictureIndex = IIf(psSortOrder = "ascending", 1, 2)
230: End If
240: Else
250: If psSortOrder = vbNullString Then
260: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders(lColIdx).PictureIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .PictureIndex = 1 Then
270: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders().PictureIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.PictureIndex = 2
280: Else
290: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders().PictureIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.PictureIndex = 1
300: End If
310: Else
320: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders().PictureIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.PictureIndex = IIf(psSortOrder = "ascending", 1, 2)
330: End If
340: End If
			'
350: 'UPGRADE_WARNING: Couldn't resolve default property of object lvwCur.ColumnHeaders().PictureIndex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lvwCur.SortListItems(lColIdx * IIf(.PictureIndex = 1, 1, -1))
			'
360: End With
		'
370: lvwCur.LockState = False
		'
		'\\ Deallocate Resources
380: 'UPGRADE_NOTE: Object lvwCur may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		lvwCur = Nothing
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise(ModuleName, "Sort", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	Public Sub AddMetric(ByRef plvwControl As AxBBLISTVIEWLib.AxBblistview1, ByRef iColumnIndex As Object, ByRef lMetric As Integer)
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
		ErrorMgr.Raise(ModuleName, "AddMetric", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
End Class