Option Strict Off
Option Explicit On
Friend Class CBranchData
	
	Private Structure TBranch
		Dim BranchID As Integer
		Dim CompanyID As Integer
		Dim Name As String
		Dim Number As String
		Dim ManagerFirstName As String
		Dim ManagerLastName As String
		Dim Address1 As String
		Dim Address2 As String
		Dim Address3 As String
		Dim Zip As String
		Dim State As String
		Dim City As String
		Dim PhoneNumber As String
		Dim FaxNumber As String
		Dim Email As String
	End Structure
	'
	Private r() As TBranch
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		On Error Resume Next
		ReDim r(0)
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property BranchID() As Integer
		Get
			BranchID = r(0).BranchID
		End Get
		Set(ByVal Value As Integer)
			r(0).BranchID = Value
		End Set
	End Property
	
	
	Public Property CompanyID() As Integer
		Get
			CompanyID = r(0).CompanyID
		End Get
		Set(ByVal Value As Integer)
			r(0).CompanyID = Value
		End Set
	End Property
	
	
	Public Property Name() As String
		Get
			Name = r(0).Name
		End Get
		Set(ByVal Value As String)
			r(0).Name = Value
		End Set
	End Property
	
	
	Public Property Number() As String
		Get
			Number = r(0).Number
		End Get
		Set(ByVal Value As String)
			r(0).Number = Value
		End Set
	End Property
	
	
	Public Property ManagerFirstName() As String
		Get
			ManagerFirstName = r(0).ManagerFirstName
		End Get
		Set(ByVal Value As String)
			r(0).ManagerFirstName = Value
		End Set
	End Property
	
	
	Public Property ManagerLastName() As String
		Get
			ManagerLastName = r(0).ManagerLastName
		End Get
		Set(ByVal Value As String)
			r(0).ManagerLastName = Value
		End Set
	End Property
	
	
	Public Property Address1() As String
		Get
			Address1 = r(0).Address1
		End Get
		Set(ByVal Value As String)
			r(0).Address1 = Value
		End Set
	End Property
	
	
	Public Property Address2() As String
		Get
			Address2 = r(0).Address2
		End Get
		Set(ByVal Value As String)
			r(0).Address2 = Value
		End Set
	End Property
	
	
	Public Property Address3() As String
		Get
			Address3 = r(0).Address3
		End Get
		Set(ByVal Value As String)
			r(0).Address3 = Value
		End Set
	End Property
	
	
	Public Property Zip() As String
		Get
			Zip = r(0).Zip
		End Get
		Set(ByVal Value As String)
			r(0).Zip = Value
		End Set
	End Property
	
	
	Public Property State() As String
		Get
			State = r(0).State
		End Get
		Set(ByVal Value As String)
			r(0).State = Value
		End Set
	End Property
	
	
	Public Property City() As String
		Get
			City = r(0).City
		End Get
		Set(ByVal Value As String)
			r(0).City = Value
		End Set
	End Property
	
	
	Public Property PhoneNumber() As String
		Get
			PhoneNumber = r(0).PhoneNumber
		End Get
		Set(ByVal Value As String)
			r(0).PhoneNumber = Value
		End Set
	End Property
	
	
	Public Property FaxNumber() As String
		Get
			FaxNumber = r(0).FaxNumber
		End Get
		Set(ByVal Value As String)
			r(0).FaxNumber = Value
		End Set
	End Property
	
	
	Public Property Email() As String
		Get
			Email = r(0).Email
		End Get
		Set(ByVal Value As String)
			r(0).Email = Value
		End Set
	End Property
End Class