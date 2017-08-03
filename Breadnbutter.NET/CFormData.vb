Option Strict Off
Option Explicit On
Friend Class CFormData
	
	Public Enum eMode
		None
		AddNewRecord
	End Enum
	
	Private iMode As Short
	
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
	Public Event MoveRecord(ByRef Record As Short)
	Public Event MoveFirst()
	Public Event MovePrevious()
	Public Event MoveNext()
	Public Event MoveLast()
	Public Event PrintRecord()
	Public Event Sort(ByRef pbAsc As Boolean)
	Public Event FindRecord(ByRef Where As String)
	Public Event ClearControls()
	Public Event Read()
	Public Event Changed(ByRef fChanged As Boolean)
	Public Event Enable(ByRef pfEnable As Boolean)
	
	'UPGRADE_NOTE: Changed was upgraded to Changed_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Changed_Renamed() As Boolean
		Dim fChanged As Boolean
		RaiseEvent Changed(fChanged)
		Changed_Renamed = fChanged
	End Function
	
	'UPGRADE_NOTE: Enable was upgraded to Enable_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Enable_Renamed(ByRef pfEnable As Boolean)
		RaiseEvent Enable(pfEnable)
	End Sub
	
	'UPGRADE_NOTE: Save was upgraded to Save_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Save_Renamed() As Boolean
		Dim fSuccess As Boolean
		RaiseEvent Save(fSuccess)
		Save_Renamed = fSuccess
	End Function
	
	'UPGRADE_NOTE: AddNew was upgraded to AddNew_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function AddNew_Renamed() As Boolean
		Dim fSuccess As Boolean
		RaiseEvent AddNew(fSuccess)
		AddNew_Renamed = fSuccess
	End Function
	
	'UPGRADE_NOTE: Delete was upgraded to Delete_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Delete_Renamed()
		RaiseEvent Delete()
	End Sub
	
	'UPGRADE_NOTE: MoveRecord was upgraded to MoveRecord_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub MoveRecord_Renamed(ByRef peRecord As eRecord)
		RaiseEvent MoveRecord(CShort(peRecord))
	End Sub
	
	'UPGRADE_NOTE: MoveFirst was upgraded to MoveFirst_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub MoveFirst_Renamed()
		RaiseEvent MoveFirst()
	End Sub
	
	'UPGRADE_NOTE: MovePrevious was upgraded to MovePrevious_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub MovePrevious_Renamed()
		RaiseEvent MovePrevious()
	End Sub
	
	'UPGRADE_NOTE: MoveNext was upgraded to MoveNext_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub MoveNext_Renamed()
		RaiseEvent MoveNext()
	End Sub
	
	'UPGRADE_NOTE: MoveLast was upgraded to MoveLast_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub MoveLast_Renamed()
		RaiseEvent MoveLast()
	End Sub
	
	
	Public Property Mode() As eMode
		Get
			Mode = iMode
		End Get
		Set(ByVal Value As eMode)
			iMode = Value
		End Set
	End Property
	
	'UPGRADE_NOTE: Edit was upgraded to Edit_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Edit_Renamed()
		RaiseEvent Edit()
	End Sub
	
	'UPGRADE_NOTE: PrintRecord was upgraded to PrintRecord_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub PrintRecord_Renamed()
		RaiseEvent PrintRecord()
	End Sub
	
	'UPGRADE_NOTE: Sort was upgraded to Sort_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Sort_Renamed(ByRef pbAsc As Boolean)
		RaiseEvent Sort(pbAsc)
	End Sub
	
	'UPGRADE_NOTE: FindRecord was upgraded to FindRecord_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub FindRecord_Renamed(ByRef psWhere As String)
		RaiseEvent FindRecord(psWhere)
	End Sub
	
	'UPGRADE_NOTE: ClearControls was upgraded to ClearControls_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub ClearControls_Renamed()
		RaiseEvent ClearControls()
	End Sub
	
	'UPGRADE_NOTE: Read was upgraded to Read_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Read_Renamed()
		RaiseEvent Read()
	End Sub
	
	'UPGRADE_NOTE: Fetch was upgraded to Fetch_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Fetch_Renamed()
		RaiseEvent Fetch()
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
End Class