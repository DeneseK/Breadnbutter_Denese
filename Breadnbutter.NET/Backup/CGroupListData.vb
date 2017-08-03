Option Strict Off
Option Explicit On
Friend Class CGroupListData
	Private Structure TGroupList
		Dim ID As Integer
		Dim ListName As String
		Dim EmployeeID As Integer
	End Structure
	
	Private r() As TGroupList
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		On Error Resume Next
		ReDim r(0)
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property ID() As Integer
		Get
			ID = r(0).ID
		End Get
		Set(ByVal Value As Integer)
			r(0).ID = Value
		End Set
	End Property
	
	
	Public Property ListName() As String
		Get
			ListName = r(0).ListName
		End Get
		Set(ByVal Value As String)
			r(0).ListName = Value
		End Set
	End Property
	
	
	Public Property EmployeeID() As Integer
		Get
			EmployeeID = r(0).EmployeeID
		End Get
		Set(ByVal Value As Integer)
			r(0).EmployeeID = Value
		End Set
	End Property
End Class