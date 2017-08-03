Option Strict Off
Option Explicit On
Module MLicense
	
	'\\ Security
	Public bLicChecked As Boolean
	Public bLicTimer As Boolean
	'UPGRADE_WARNING: Lower bound of array lSecValCode was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public lSecValCode(3) As Integer
	'UPGRADE_WARNING: Lower bound of array lSecValRslt was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public lSecValRslt(3) As Integer
	Public dSecVar As Double
	Public bLicError As Boolean
	Public bSecDisp As Boolean
	Public lSecCompID As Integer
	'
	Public Sub InitLicense()
		'
		'\\ Security
		If bLicChecked = False Then
			With FMain.License
				lSecValCode(1) = 1514044385
				lSecValRslt(1) = 6035375
				lSecValCode(2) = 2067020449
				lSecValRslt(2) = 4912112
				lSecValCode(3) = 1463088053
				lSecValRslt(3) = 8960596
				dSecVar = Int((3000 - 1 + 1) * Rnd() + 1)
				.Enabled = False
				'.CPAlgorithm = 1 + 2
				.CPAlgorithm = 65536
				.CPAlgorithmDrive = Left(My.Application.Info.DirectoryPath, 1)
				.TCSeed = 192
				.TCRegKey2Seed = 48
				.LFPassword = "D" & "uct" & "Ta" & "p" & "e"
				.LFName = My.Application.Info.DirectoryPath & "\PowerKey.lf"
				.ExpireDateHard = "12/31/2040"
				.Enabled = True
				lSecCompID = .CPCompNo
				.ForceStatusChanged()
				bLicChecked = True
			End With
		End If
	End Sub
	'
	Public Function CPCheck() As Double
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iCt As Short
		'
		CPCheck = CDbl(Today.ToOADate - (dSecVar / 3.14))
		'
		With FMain.License
			For iCt = 1 To .LicensedComputers
				If lSecCompID = .get_LicensedComputer(iCt) Then
					CPCheck = CDbl(Today.ToOADate - dSecVar)
					Exit For
				End If
			Next 
		End With
		'
		Exit Function
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.General.CPCheck")
	End Function
	
	Public Function ValidateLicense() As Double
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iSecValPair As Short
		Dim sDlgMsg As String
		'
		With FMain.License
			'
			iSecValPair = Int((3 - 1 + 1) * Rnd() + 1)
			'
			If .LibTest(lSecValCode(iSecValPair)) <> lSecValRslt(iSecValPair) Then
				ValidateLicense = CDbl(Today.ToOADate - (dSecVar / 3.14))
			ElseIf System.Date.FromOADate(CPCheck) = (System.Date.FromOADate(Today.ToOADate - dSecVar)) Then 
				If .IsExpired Then
					ValidateLicense = CDbl(Today.ToOADate - (dSecVar / 3.14))
				Else
					If .IsClockTurnedBack Then
						ValidateLicense = CDbl(Today.ToOADate - (dSecVar / 3.14))
					Else
						ValidateLicense = CDbl(Today.ToOADate - dSecVar)
					End If
				End If
			Else
				If .ExpireMode = "D" Then
					If .IsClockTurnedBack Then
						ValidateLicense = CDbl(Today.ToOADate - (dSecVar / 3.14))
					Else
						ValidateLicense = CDbl(Today.ToOADate - (dSecVar / 3.14))
					End If
				Else
					ValidateLicense = CDbl(Today.ToOADate - (dSecVar / 3.14))
				End If
			End If
			'
		End With
		'
		If ValidateLicense = Today.ToOADate - (dSecVar / 3.14)() Then
			sDlgMsg = "Your license is either damaged or has expired. You will not be" & vbCrLf & "able to authorize licenses until you contact Jason or Eric." & vbCrLf & "Click on the License Facility button for more information."
			MsgBox(sDlgMsg, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Unauthorized License")
		End If
		'
		Exit Function
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.General.ValidateLicense")
	End Function
End Module