Option Strict Off
Option Explicit On
Friend Class CEmployeeData
	
	Private Structure eData
		Dim EmployeeID As Short
		Dim EmployeeNumber As String
		Dim EmployeeLast As String
		Dim EmployeeFirst As String
		Dim EmployeeMiddle As String
		Dim EMailAddress As String
		Dim Password As String
		Dim Groups As Short
		Dim SecurityLevel As Short
		Dim EmployeeExt As Short
		Dim WorkGroups As Short
	End Structure
	
	Private Data As eData
	
	
	Public Property EmployeeID() As Short
		Get
			EmployeeID = Data.EmployeeID
		End Get
		Set(ByVal Value As Short)
			Data.EmployeeID = Value
		End Set
	End Property
	
	
	Public Property EmployeeNumber() As String
		Get
			EmployeeNumber = Data.EmployeeNumber
		End Get
		Set(ByVal Value As String)
			Data.EmployeeNumber = Value
		End Set
	End Property
	
	
	Public Property EmployeeLast() As String
		Get
			EmployeeLast = Data.EmployeeLast
		End Get
		Set(ByVal Value As String)
			Data.EmployeeLast = Value
		End Set
	End Property
	
	
	Public Property EmployeeFirst() As String
		Get
			EmployeeFirst = Data.EmployeeFirst
		End Get
		Set(ByVal Value As String)
			Data.EmployeeFirst = Value
		End Set
	End Property
	
	
	Public Property EmployeeMiddle() As String
		Get
			EmployeeMiddle = Data.EmployeeMiddle
		End Get
		Set(ByVal Value As String)
			Data.EmployeeMiddle = Value
		End Set
	End Property
	
	Public Property Password() As String
		Get
			Password = Data.Password
		End Get
		Set(ByVal Value As String)
			Data.Password = Value
		End Set
	End Property
	
	Public Property EMailAddress() As String
		Get
			EMailAddress = Data.EMailAddress
		End Get
		Set(ByVal Value As String)
			Data.EMailAddress = Value
		End Set
	End Property
	
	
	Public Property Groups() As Short
		Get
			Groups = Data.Groups
		End Get
		Set(ByVal Value As Short)
			Data.Groups = Value
		End Set
	End Property
	
	
	Public Property SecurityLevel() As Short
		Get
			SecurityLevel = Data.SecurityLevel
		End Get
		Set(ByVal Value As Short)
			Data.SecurityLevel = Value
		End Set
	End Property
	
	
	Public Property EmployeeExt() As Short
		Get
			EmployeeExt = Data.EmployeeExt
		End Get
		Set(ByVal Value As Short)
			Data.EmployeeExt = Value
		End Set
	End Property
	
	
	Public Property WorkGroups() As Short
		Get
			WorkGroups = Data.WorkGroups
		End Get
		Set(ByVal Value As Short)
			Data.WorkGroups = Value
		End Set
	End Property
End Class