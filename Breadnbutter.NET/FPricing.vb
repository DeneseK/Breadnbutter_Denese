Option Strict Off
Option Explicit On
Friend Class FPricing
	Inherits System.Windows.Forms.Form
	
	Private Function FixPhoneNumber(ByRef psNumber As String) As String
		Dim sTemp As String
		Dim l As Integer
		Dim sChar As String
		'
		For l = 1 To Len(psNumber)
			sChar = Mid(psNumber, l, 1)
			If (Asc(sChar) >= 48) And (Asc(sChar) <= 57) Then
				sTemp = sTemp & sChar
			End If
		Next 
		'
		FixPhoneNumber = sTemp
	End Function
	
	'Private Sub Command1_Click()
	'  Dim ContactData As New CContactData
	'  Dim contact As New CContact
	'  Dim rs As New Recordset
	'  '
	'  rs.Open "SELECT * FROM Tcin", cnMain, adOpenForwardOnly, adLockReadOnly
	'  '
	'  While Not rs.eof
	'    Set ContactData = New CContactData
	'    '
	'    With ContactData
	'      .CompanyID = 5728
	'      .FirstName = Trim(rs!FirstName & vbNullString)
	'      .LastName = Trim(rs!LastName & vbNullString)
	'      .Phone1 = FixPhoneNumber(rs!OfficePhone & vbNullString)
	'      .Email = rs!Emailaddress & vbNullString
	'      .MailState = rs!State & vbNullString
	'      .MailZip = rs!Zip & vbNullString
	'      .MailCity = rs!City & vbNullString
	'      .MailAddress1 = rs!POBox & vbNullString
	'      .PreferredAddress = 1
	'      .AdjusterID = rs!AdjNo & vbNullString
	'      .Status = "Customer"
	'      .ContactType = 1
	'    End With
	'    '
	'    contact.Save ContactData, True
	'    '
	'    rs.MoveNext
	'  Wend
	'End Sub
	'
	'Private Sub Command2_Click()
	'Dim ContactData As New CContactData
	'  Dim contact As New CContact
	'  Dim rs As New Recordset
	'  '
	'  rs.Open "SELECT * FROM TContact WHERE CompanyID = 5728", cnMain, adOpenKeyset, adLockOptimistic
	'  '
	'
	'  While Not rs.eof
	'    contact.Load ContactData, rs!ID
	'    ContactData.FirstName = Trim(ContactData.FirstName)
	'    ContactData.LastName = Trim(ContactData.LastName)
	'    contact.Save ContactData, False
	'    rs.MoveNext
	'  Wend
	''    Set ContactData = New CContactData
	''    '
	''    With ContactData
	''      .CompanyID = 5728
	''      .FirstName = rs!FirstName & vbNullString
	''      .LastName = rs!LastName & vbNullString
	''      .Phone1 = FixPhoneNumber(rs!OfficePhone & vbNullString)
	''      .Email = rs!Emailaddress & vbNullString
	''      .MailZip = rs!Zip & vbNullString
	''      .MailState = rs!State & vbNullString
	''      .MailZip = rs!City & vbNullString
	''      .MailAddress1 = rs!POBox & vbNullString
	''      .PreferredAddress = 1
	''      .AdjusterID = rs!AdjNo & vbNullString
	''      .Status = "Customer"
	''      .ContactType = 1
	''    End With
	''    '
	''    contact.Save ContactData, True
	''    '
	''    rs.MoveNext
	''  Wend
	'End Sub
	Private Sub FPricing_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
	End Sub
End Class