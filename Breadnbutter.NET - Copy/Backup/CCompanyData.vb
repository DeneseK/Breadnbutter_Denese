Option Strict Off
Option Explicit On
Friend Class CCompanyData
	Private Structure TCompany
		Dim ID As Integer
		Dim DateEntered As Date
		Dim LastUpdate As Date
		Dim Name As String
		Dim Division As String
		Dim Individual As Boolean
		Dim DoNotContact As Boolean
		Dim Note As String
		Dim InterestRank As Short
	End Structure
	
	Private r() As TCompany
	
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
	
	'Public Property Get DateEntered() As Date
	'  DateEntered = r(0).DateEntered
	'End Property
	'
	'Public Property Get LastUpdate() As Date
	'  LastUpdate = r(0).LastUpdate
	'End Property
	
	
	
	
	Public Property Name() As String
		Get
			Name = r(0).Name
		End Get
		Set(ByVal Value As String)
			r(0).Name = Value
		End Set
	End Property
	'
	'Public Property Get Division() As String
	'  Division = r(0).Division
	'End Property
	
	Public WriteOnly Property Office() As String
		Set(ByVal Value As String)
			r(0).Division = Value
		End Set
	End Property
	
	
	Public Property Individual() As Boolean
		Get
			Individual = r(0).Individual
		End Get
		Set(ByVal Value As Boolean)
			r(0).Individual = Value
		End Set
	End Property
	
	
	Public Property DoNotContact() As Boolean
		Get
			DoNotContact = r(0).DoNotContact
		End Get
		Set(ByVal Value As Boolean)
			r(0).DoNotContact = Value
		End Set
	End Property
	
	
	Public Property Note() As String
		Get
			Note = r(0).Note
		End Get
		Set(ByVal Value As String)
			r(0).Note = Value
		End Set
	End Property
	
	
	Public Property InterestRank() As Short
		Get
			InterestRank = r(0).InterestRank
		End Get
		Set(ByVal Value As Short)
			r(0).InterestRank = Value
		End Set
	End Property
End Class