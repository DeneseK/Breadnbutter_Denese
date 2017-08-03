Option Strict Off
Option Explicit On
Friend Class CEventData
	Private Structure TEvent
		Dim RecID As Integer
		Dim CustRecID As Integer
		Dim EventDate As Date
		Dim EventTime As Date
		Dim EventType As String
		Dim EventResults As String
		Dim EventUser As String
		Dim EventSubject As String
		Dim ProductID As Short
		Dim ClosedTime As Date
		Dim OpenCall As Boolean
		Dim Sticky As Boolean
	End Structure
	
	Private r() As TEvent
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		On Error Resume Next
		ReDim r(0)
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property RecID() As Integer
		Get
			RecID = r(0).RecID
		End Get
		Set(ByVal Value As Integer)
			r(0).RecID = Value
		End Set
	End Property
	
	
	Public Property CustRecID() As Integer
		Get
			CustRecID = r(0).CustRecID
		End Get
		Set(ByVal Value As Integer)
			r(0).CustRecID = Value
		End Set
	End Property
	
	
	Public Property EventDate() As Date
		Get
			EventDate = r(0).EventDate
		End Get
		Set(ByVal Value As Date)
			r(0).EventDate = Value
		End Set
	End Property
	
	
	Public Property EventTime() As Date
		Get
			EventTime = r(0).EventTime
		End Get
		Set(ByVal Value As Date)
			r(0).EventTime = Value
		End Set
	End Property
	
	
	Public Property EventType() As String
		Get
			EventType = r(0).EventType
		End Get
		Set(ByVal Value As String)
			r(0).EventType = Value
		End Set
	End Property
	
	
	Public Property EventResults() As String
		Get
			EventResults = r(0).EventResults
		End Get
		Set(ByVal Value As String)
			r(0).EventResults = Value
		End Set
	End Property
	
	
	Public Property EventUser() As String
		Get
			EventUser = r(0).EventUser
		End Get
		Set(ByVal Value As String)
			r(0).EventUser = Value
		End Set
	End Property
	
	
	Public Property EventSubject() As String
		Get
			EventSubject = r(0).EventSubject
		End Get
		Set(ByVal Value As String)
			r(0).EventSubject = Value
		End Set
	End Property
	
	
	Public Property ProductID() As Short
		Get
			ProductID = r(0).ProductID
		End Get
		Set(ByVal Value As Short)
			r(0).ProductID = Value
		End Set
	End Property
	
	
	Public Property OpenCall() As Boolean
		Get
			OpenCall = r(0).OpenCall
		End Get
		Set(ByVal Value As Boolean)
			r(0).OpenCall = Value
		End Set
	End Property
	
	
	Public Property Sticky() As Boolean
		Get
			Sticky = r(0).Sticky
		End Get
		Set(ByVal Value As Boolean)
			r(0).Sticky = Value
		End Set
	End Property
	
	
	Public Property ClosedTime() As Date
		Get
			ClosedTime = r(0).ClosedTime
		End Get
		Set(ByVal Value As Date)
			r(0).ClosedTime = Value
		End Set
	End Property
End Class