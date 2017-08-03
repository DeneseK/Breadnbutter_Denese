Option Strict Off
Option Explicit On
Friend Class CContactData
	Private Structure TContact
		Dim ID As Integer
		Dim DateEntered As Date
		Dim LastUpdate As Date
		Dim CompanyID As Integer
		Dim BranchID As Integer
		Dim FirstName As String
		Dim LastName As String
		Dim Salutation As String
		Dim Title As String
		Dim Address1 As String
		Dim Address2 As String
		Dim City As String
		Dim State As String
		Dim Zip As String
		Dim MailAddress1 As String
		Dim MailAddress2 As String
		Dim MailCity As String
		Dim MailState As String
		Dim MailZip As String
		Dim PCEmail As String
		Dim PCEmailPassword As String
		Dim Phone1 As String
		Dim Phone2 As String
		Dim Fax As String
		Dim Notes As String
		Dim Email As String
		Dim Selected As Byte 'Betatester
		Dim Source As String
		Dim Status As String
		Dim AuthStatus As String
		Dim AuthDate As Date
		Dim AuthDays As Short
		Dim AuthRemaining As Integer
		Dim Version As String
		Dim VersionShipped As String
		Dim Copies As Short
		Dim ShipStatus As String
		Dim ShipDate As Date
		'UPGRADE_NOTE: Rate was upgraded to Rate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Rate_Renamed As Decimal
		Dim ContactType As Short
		Dim RateExpDate As Date
		Dim PreferredAddress As Short
		Dim PVAuthStatus As String
		Dim PVDownloadStatus As String
		Dim DownloadStatus As String
		Dim PVAuthDate As Date
		Dim PVAuthDays As Short
		Dim PVAuthRemaining As Integer
		Dim PVVersion As String
		Dim PVVersionShipped As String
		Dim PVCopies As Short
		Dim PVShipStatus As String
		Dim PVShipDate As Date
		Dim DownloadDate As Date
		Dim PVDownloadDate As Date
		'
		Dim DaysPending As Short
		Dim GraceDays As Short
		Dim OnlineAuths As Short
		Dim SaleDate As Date
		Dim SaleDays As Short
		Dim PVDaysPending As Short
		Dim PVGraceDays As Short
		Dim PVOnlineAuths As Short
		Dim PVSaleDate As Date
		Dim PVSaleDays As Short
		Dim WebPassword As String
		Dim ContactByEmail As Boolean
		Dim ChangedData As Short
		Dim AdjusterID As String
	End Structure
	
	Private r() As TContact
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		On Error Resume Next
		ReDim r(0)
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	Public Property DaysPending() As Short
		Get
			DaysPending = r(0).DaysPending
		End Get
		Set(ByVal Value As Short)
			r(0).DaysPending = Value
		End Set
	End Property
	
	
	
	Public Property GraceDays() As Short
		Get
			GraceDays = r(0).GraceDays
		End Get
		Set(ByVal Value As Short)
			r(0).GraceDays = Value
		End Set
	End Property
	
	
	Public Property OnlineAuths() As Short
		Get
			OnlineAuths = r(0).OnlineAuths
		End Get
		Set(ByVal Value As Short)
			r(0).OnlineAuths = Value
		End Set
	End Property
	
	
	Public Property SaleDate() As Date
		Get
			SaleDate = r(0).SaleDate
		End Get
		Set(ByVal Value As Date)
			r(0).SaleDate = Value
		End Set
	End Property
	
	
	Public Property SaleDays() As Short
		Get
			SaleDays = r(0).SaleDays
		End Get
		Set(ByVal Value As Short)
			r(0).SaleDays = Value
		End Set
	End Property
	
	
	Public Property PVDaysPending() As Short
		Get
			PVDaysPending = r(0).PVDaysPending
		End Get
		Set(ByVal Value As Short)
			r(0).PVDaysPending = Value
		End Set
	End Property
	
	
	Public Property PVGraceDays() As Short
		Get
			PVGraceDays = r(0).PVGraceDays
		End Get
		Set(ByVal Value As Short)
			r(0).PVGraceDays = Value
		End Set
	End Property
	
	
	Public Property PVOnlineAuths() As Short
		Get
			PVOnlineAuths = r(0).PVOnlineAuths
		End Get
		Set(ByVal Value As Short)
			r(0).PVOnlineAuths = Value
		End Set
	End Property
	
	
	Public Property PVSaleDate() As Date
		Get
			PVSaleDate = r(0).PVSaleDate
		End Get
		Set(ByVal Value As Date)
			r(0).PVSaleDate = Value
		End Set
	End Property
	
	
	Public Property PVSaleDays() As Short
		Get
			PVSaleDays = r(0).PVSaleDays
		End Get
		Set(ByVal Value As Short)
			r(0).PVSaleDays = Value
		End Set
	End Property
	
	
	Public Property WebPassword() As String
		Get
			WebPassword = r(0).WebPassword
		End Get
		Set(ByVal Value As String)
			r(0).WebPassword = Value
		End Set
	End Property
	
	
	Public Property ContactByEmail() As Boolean
		Get
			ContactByEmail = r(0).ContactByEmail
		End Get
		Set(ByVal Value As Boolean)
			r(0).ContactByEmail = Value
		End Set
	End Property
	
	
	Public Property ChangedData() As Short
		Get
			ChangedData = r(0).ChangedData
		End Get
		Set(ByVal Value As Short)
			r(0).ChangedData = Value
		End Set
	End Property
	
	
	Public Property PCEmailPassword() As String
		Get
			PCEmailPassword = r(0).PCEmailPassword
		End Get
		Set(ByVal Value As String)
			r(0).PCEmailPassword = Value
		End Set
	End Property
	
	
	Public Property PCEmail() As String
		Get
			PCEmail = r(0).PCEmail
		End Get
		Set(ByVal Value As String)
			r(0).PCEmail = Value
		End Set
	End Property
	
	
	'UPGRADE_NOTE: Rate was upgraded to Rate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Property Rate_Renamed() As Decimal
		Get
			Rate_Renamed = r(0).Rate_Renamed
		End Get
		Set(ByVal Value As Decimal)
			r(0).Rate_Renamed = Value
		End Set
	End Property
	
	Public ReadOnly Property Adding() As Boolean
		Get
			Dim fAdding As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object fAdding. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Adding = fAdding
		End Get
	End Property
	
	Public ReadOnly Property Loaded() As Boolean
		Get
			Dim fLoaded As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object fLoaded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Loaded = fLoaded
		End Get
	End Property
	
	
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
	'Public Property Let DateEntered(ByVal NewValue As Date)
	'  r(0).DateEntered = NewValue
	'End Property
	
	'Public Property Get LastUpdate() As Date
	'   LastUpdate = r(0).LastUpdate
	'End Property
	'
	'Public Property Let LastUpdate(ByVal NewValue As Date)
	'  r(0).LastUpdate = NewValue
	'End Property
	
	
	Public Property CompanyID() As Integer
		Get
			CompanyID = r(0).CompanyID
		End Get
		Set(ByVal Value As Integer)
			r(0).CompanyID = Value
		End Set
	End Property
	
	
	Public Property BranchID() As Integer
		Get
			BranchID = r(0).BranchID
		End Get
		Set(ByVal Value As Integer)
			r(0).BranchID = Value
		End Set
	End Property
	
	
	Public Property FirstName() As String
		Get
			FirstName = r(0).FirstName
		End Get
		Set(ByVal Value As String)
			r(0).FirstName = Value
		End Set
	End Property
	
	
	Public Property LastName() As String
		Get
			LastName = r(0).LastName
		End Get
		Set(ByVal Value As String)
			r(0).LastName = Value
		End Set
	End Property
	
	
	Public Property Salutation() As String
		Get
			Salutation = r(0).Salutation
		End Get
		Set(ByVal Value As String)
			r(0).Salutation = Value
		End Set
	End Property
	
	
	Public Property Title() As String
		Get
			Title = r(0).Title
		End Get
		Set(ByVal Value As String)
			r(0).Title = Value
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
	
	
	Public Property City() As String
		Get
			City = r(0).City
		End Get
		Set(ByVal Value As String)
			r(0).City = Value
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
	
	
	Public Property Zip() As String
		Get
			Zip = r(0).Zip
		End Get
		Set(ByVal Value As String)
			r(0).Zip = Value
		End Set
	End Property
	
	
	Public Property MailAddress1() As String
		Get
			MailAddress1 = r(0).MailAddress1
		End Get
		Set(ByVal Value As String)
			r(0).MailAddress1 = Value
		End Set
	End Property
	
	
	Public Property MailAddress2() As String
		Get
			MailAddress2 = r(0).MailAddress2
		End Get
		Set(ByVal Value As String)
			r(0).MailAddress2 = Value
		End Set
	End Property
	
	
	Public Property MailCity() As String
		Get
			MailCity = r(0).MailCity
		End Get
		Set(ByVal Value As String)
			r(0).MailCity = Value
		End Set
	End Property
	
	
	Public Property MailState() As String
		Get
			MailState = r(0).MailState
		End Get
		Set(ByVal Value As String)
			r(0).MailState = Value
		End Set
	End Property
	
	
	Public Property MailZip() As String
		Get
			MailZip = r(0).MailZip
		End Get
		Set(ByVal Value As String)
			r(0).MailZip = Value
		End Set
	End Property
	
	
	Public Property Phone1() As String
		Get
			If IsNumeric(r(0).Phone1) Then
				Phone1 = r(0).Phone1
			Else
				Phone1 = StripChars(Phone1) 'Replace(Replace(Replace(r(0).Phone1, "-", vbNullString), "_", vbNullString), "x", vbNullString)
			End If
		End Get
		Set(ByVal Value As String)
			r(0).Phone1 = Value
		End Set
	End Property
	
	
	Public Property Phone2() As String
		Get
			If IsNumeric(r(0).Phone2) Then
				Phone2 = r(0).Phone2
			Else
				Phone2 = StripChars(Phone2) 'Replace(Replace(Replace(r(0).Phone2, "-", vbNullString), "_", vbNullString), "x", vbNullString)
			End If
		End Get
		Set(ByVal Value As String)
			r(0).Phone2 = Value
		End Set
	End Property
	
	
	Public Property Fax() As String
		Get
			If IsNumeric(r(0).Fax) Then
				Fax = r(0).Fax
			Else
				Fax = StripChars(Fax) 'Replace(Replace(Replace(r(0).Fax, "-", vbNullString), "_", vbNullString), "x", vbNullString)
			End If
		End Get
		Set(ByVal Value As String)
			r(0).Fax = Value
		End Set
	End Property
	
	
	Public Property Notes() As String
		Get
			Notes = r(0).Notes
		End Get
		Set(ByVal Value As String)
			r(0).Notes = Value
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
	
	
	Public Property Selected() As Byte
		Get
			Selected = r(0).Selected
		End Get
		Set(ByVal Value As Byte)
			r(0).Selected = Value
		End Set
	End Property
	
	
	Public Property PreferredAddress() As Short
		Get
			PreferredAddress = r(0).PreferredAddress
		End Get
		Set(ByVal Value As Short)
			r(0).PreferredAddress = Value
		End Set
	End Property
	
	
	Public Property Source() As String
		Get
			Source = r(0).Source
		End Get
		Set(ByVal Value As String)
			r(0).Source = Value
		End Set
	End Property
	
	
	Public Property Status() As String
		Get
			Status = r(0).Status
		End Get
		Set(ByVal Value As String)
			r(0).Status = Value
		End Set
	End Property
	
	
	Public Property AuthStatus() As String
		Get
			AuthStatus = r(0).AuthStatus
		End Get
		Set(ByVal Value As String)
			r(0).AuthStatus = Value
		End Set
	End Property
	
	
	Public Property AuthDate() As Date
		Get
			AuthDate = r(0).AuthDate
		End Get
		Set(ByVal Value As Date)
			r(0).AuthDate = Value
		End Set
	End Property
	
	
	Public Property AuthDays() As Short
		Get
			AuthDays = r(0).AuthDays
		End Get
		Set(ByVal Value As Short)
			r(0).AuthDays = Value
		End Set
	End Property
	
	
	Public Property AuthRemaining() As Integer
		Get
			AuthRemaining = r(0).AuthRemaining
		End Get
		Set(ByVal Value As Integer)
			r(0).AuthRemaining = Value
		End Set
	End Property
	
	
	Public Property Version() As String
		Get
			Version = r(0).Version
		End Get
		Set(ByVal Value As String)
			r(0).Version = Value
		End Set
	End Property
	
	
	Public Property VersionShipped() As String
		Get
			VersionShipped = r(0).VersionShipped
		End Get
		Set(ByVal Value As String)
			r(0).VersionShipped = Value
		End Set
	End Property
	
	
	Public Property Copies() As Short
		Get
			Copies = r(0).Copies
		End Get
		Set(ByVal Value As Short)
			r(0).Copies = Value
		End Set
	End Property
	
	
	Public Property ShipStatus() As String
		Get
			ShipStatus = r(0).ShipStatus
		End Get
		Set(ByVal Value As String)
			r(0).ShipStatus = Value
		End Set
	End Property
	
	
	Public Property ShipDate() As Date
		Get
			ShipDate = r(0).ShipDate
		End Get
		Set(ByVal Value As Date)
			r(0).ShipDate = Value
		End Set
	End Property
	
	
	Public Property ContactType() As Short
		Get
			ContactType = r(0).ContactType
		End Get
		Set(ByVal Value As Short)
			r(0).ContactType = Value
		End Set
	End Property
	
	
	Public Property RateExpDate() As Date
		Get
			RateExpDate = r(0).RateExpDate
		End Get
		Set(ByVal Value As Date)
			r(0).RateExpDate = Value
		End Set
	End Property
	
	
	Public Property SearchID() As Integer
		Get
			Dim lSearchID As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object lSearchID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SearchID = lSearchID
		End Get
		Set(ByVal Value As Integer)
			Dim lSearchID As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object lSearchID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lSearchID = Value
		End Set
	End Property
	
	
	Public Property PVAuthStatus() As String
		Get
			PVAuthStatus = r(0).PVAuthStatus
		End Get
		Set(ByVal Value As String)
			r(0).PVAuthStatus = Value
		End Set
	End Property
	
	
	Public Property PVAuthDate() As Date
		Get
			PVAuthDate = r(0).PVAuthDate
		End Get
		Set(ByVal Value As Date)
			r(0).PVAuthDate = Value
		End Set
	End Property
	
	
	Public Property PVAuthDays() As Short
		Get
			PVAuthDays = r(0).PVAuthDays
		End Get
		Set(ByVal Value As Short)
			r(0).PVAuthDays = Value
		End Set
	End Property
	
	
	Public Property PVAuthRemaining() As Integer
		Get
			PVAuthRemaining = r(0).PVAuthRemaining
		End Get
		Set(ByVal Value As Integer)
			r(0).PVAuthRemaining = Value
		End Set
	End Property
	
	
	Public Property PVVersion() As String
		Get
			PVVersion = r(0).Version
		End Get
		Set(ByVal Value As String)
			r(0).PVVersion = Value
		End Set
	End Property
	
	
	Public Property PVVersionShipped() As String
		Get
			PVVersionShipped = r(0).PVVersionShipped
		End Get
		Set(ByVal Value As String)
			r(0).PVVersionShipped = Value
		End Set
	End Property
	
	
	Public Property PVCopies() As Short
		Get
			PVCopies = r(0).PVCopies
		End Get
		Set(ByVal Value As Short)
			r(0).PVCopies = Value
		End Set
	End Property
	
	
	Public Property PVShipStatus() As String
		Get
			PVShipStatus = r(0).PVShipStatus
		End Get
		Set(ByVal Value As String)
			r(0).PVShipStatus = Value
		End Set
	End Property
	
	
	Public Property PVShipDate() As Date
		Get
			PVShipDate = r(0).PVShipDate
		End Get
		Set(ByVal Value As Date)
			r(0).PVShipDate = Value
		End Set
	End Property
	
	
	Public Property DownloadDate() As Date
		Get
			DownloadDate = r(0).DownloadDate
		End Get
		Set(ByVal Value As Date)
			r(0).DownloadDate = Value
		End Set
	End Property
	
	
	Public Property PVDownloadDate() As Date
		Get
			PVDownloadDate = r(0).PVDownloadDate
		End Get
		Set(ByVal Value As Date)
			r(0).PVDownloadDate = Value
		End Set
	End Property
	
	
	Public Property PVDownloadStatus() As String
		Get
			PVDownloadStatus = r(0).PVDownloadStatus
		End Get
		Set(ByVal Value As String)
			r(0).PVDownloadStatus = Value
		End Set
	End Property
	
	
	Public Property DownloadStatus() As String
		Get
			DownloadStatus = r(0).DownloadStatus
		End Get
		Set(ByVal Value As String)
			r(0).DownloadStatus = Value
		End Set
	End Property
	
	
	Public Property AdjusterID() As String
		Get
			Dim AdjsuterID As Object
			'UPGRADE_WARNING: Couldn't resolve default property of object AdjsuterID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AdjsuterID = r(0).AdjusterID
		End Get
		Set(ByVal Value As String)
			r(0).AdjusterID = Value
		End Set
	End Property
End Class