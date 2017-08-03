Option Strict Off
Option Explicit On
Module MReporter
	
	'Public dbmInt As New DBMgr
	'Public dbmMats As New DBMgr
	
	Public Enum HeaderType
		htFull
		htPartial
		htNone
	End Enum
	
	Public Function FormatCSZ(ByRef City As Object, ByRef State As Object, ByRef Zip As Object) As String
		On Error GoTo ErrCall
		'
		Dim CSZTemp As String
		
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object City. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not IsDbNull(City) Then CSZTemp = City
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object State. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not IsDbNull(State) Then CSZTemp = CSZTemp & ", " & Trim(State)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If Not IsDbNull(Zip) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Zip. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Len(Trim(Zip)) > 5 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Zip. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CSZTemp = CSZTemp & " " & VB6.Format(Zip, "&&&&&" & "-" & "&&&&")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Zip. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CSZTemp = CSZTemp & " " & Trim(Zip)
			End If
		End If
		
		FormatCSZ = CSZTemp
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Module1.FormatCSZ", MsgBoxStyle.Critical, "Error")
	End Function
	
	
	Public Function FormatPhone(ByRef sPhone As String) As String
		On Error GoTo ErrCall
		'
		Dim PhoneStripped As String
		Dim PhoneTemp As String
		Dim i As Short
		'
		PhoneStripped = Trim(sPhone)
		PhoneTemp = VB6.Format(Left(PhoneStripped, 10), "!&&&-&&&-&&&&")
		If Len(PhoneStripped) > 10 Then
			PhoneTemp = PhoneTemp & " x " & Mid(PhoneStripped, 11)
		End If
		'
		FormatPhone = PhoneTemp
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Module1.FormatPhone", MsgBoxStyle.Critical, "Error")
	End Function
	
	
	Public Function FormatAddress(ByRef sAddress1 As String, ByRef sAddress2 As String, Optional ByRef iOrientation As Short = 0) As String
		On Error GoTo ErrCall
		'
		If sAddress2 = "" Then
			FormatAddress = sAddress1
		Else
			If sAddress1 = "" Then
				FormatAddress = sAddress2
			Else
				FormatAddress = sAddress1 & IIf(iOrientation, ", ", vbCrLf) & sAddress2
			End If
		End If
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Module1.FormatAddress", MsgBoxStyle.Critical, "Error")
	End Function
	
	Public Function FormatName(ByRef sName1 As String, ByRef sName2 As String) As String
		On Error GoTo ErrCall
		'
		If sName2 = "" Then
			FormatName = sName1
		Else
			FormatName = sName1 & vbCrLf & "Attention: " & sName2
		End If
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Module1.FormatName", MsgBoxStyle.Critical, "Error")
	End Function
End Module