Option Strict Off
Option Explicit On
Friend Class CProductData
	
	Private Structure TProduct
		Dim ProductID As Integer
		Dim Product As String
		Dim Color As Integer
		Dim Seed1 As Short
		Dim Seed2 As Short
	End Structure
	'
	Private r() As TProduct
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		On Error Resume Next
		ReDim r(0)
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property ProductID() As Integer
		Get
			ProductID = r(0).ProductID
		End Get
		Set(ByVal Value As Integer)
			r(0).ProductID = Value
		End Set
	End Property
	
	
	Public Property Product() As String
		Get
			Product = r(0).Product
		End Get
		Set(ByVal Value As String)
			r(0).Product = Value
		End Set
	End Property
	
	
	Public Property Color() As Integer
		Get
			Color = r(0).Color
		End Get
		Set(ByVal Value As Integer)
			r(0).Color = Value
		End Set
	End Property
	
	
	Public Property Seed1() As Short
		Get
			Seed1 = r(0).Seed1
		End Get
		Set(ByVal Value As Short)
			r(0).Seed1 = Value
		End Set
	End Property
	
	
	Public Property Seed2() As Short
		Get
			Seed2 = r(0).Seed2
		End Get
		Set(ByVal Value As Short)
			r(0).Seed2 = Value
		End Set
	End Property
End Class