Option Strict Off
Option Explicit On
Friend Class FUtility
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdScanForLicenses_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdScanForLicenses.Click
		Dim rsContact As ADODB.Recordset
		Dim rsLicense As ADODB.Recordset
		'
		rsContact = New ADODB.Recordset
		rsLicense = New ADODB.Recordset
		'
		rsContact.Open("SELECT * FROM TContact WHERE Status = 'Customer'", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		rsLicense.Open("SELECT * FROM TLicense", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		'
		Do Until rsContact.EOF
			rsLicense.AddNew()
			'
			rsLicense.Fields("ContactID").Value = rsContact.Fields("ID").Value
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			rsLicense.Fields("LicenseDate").Value = IIf(IsDbNull(rsContact.Fields("AuthDate").Value), 0, rsContact.Fields("AuthDate").Value)
			rsLicense.Fields("Days").Value = rsContact.Fields("AuthDays").Value
			rsLicense.Fields("Amount").Value = rsContact.Fields("Rate").Value
			'
			rsLicense.Update()
			rsContact.MoveNext()
		Loop 
	End Sub
End Class