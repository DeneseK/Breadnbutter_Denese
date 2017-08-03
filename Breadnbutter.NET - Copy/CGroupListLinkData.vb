Option Strict Off
Option Explicit On
Friend Class CGroupListLinkData
	Private Structure TGroupListLink
		Dim ID As Integer
		Dim ContactID As Integer
		Dim ListID As Integer
	End Structure
	
	Private r() As TGroupListLink
	
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
	
	
	Public Property ContactID() As Integer
		Get
			ContactID = r(0).ContactID
		End Get
		Set(ByVal Value As Integer)
			r(0).ContactID = Value
		End Set
	End Property
	
	
	Public Property ListID() As Integer
		Get
			ListID = r(0).ListID
		End Get
		Set(ByVal Value As Integer)
			r(0).ListID = Value
		End Set
	End Property
End Class