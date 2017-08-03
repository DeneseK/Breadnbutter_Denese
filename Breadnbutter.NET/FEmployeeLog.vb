Option Strict Off
Option Explicit On
Friend Class FEmployeeLog
	Inherits System.Windows.Forms.Form
	
	'Public Enum enMode
	'  enLogin
	'  enLogout
	'End Enum
	'
	Private iMode As FLogon.enMode
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error GoTo ErrCall
		'
		txtPassword.Text = ""
		User.LogResults = False
		Me.Close()
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmEmployeeLog.cmdCancel_Click.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Function Submit(ByRef InOut As Short, ByRef bLog As Boolean) As Boolean
		On Error GoTo ErrCall
		'
		Dim i As Short
		Dim j As Short
		Dim strMsg As String
		Dim strDate As String
		Dim resp As Short
		Dim rsEmployee As ADODB.Recordset
		Dim rsLog As ADODB.Recordset
		Dim bEdit As Boolean
		'
		rsEmployee = New ADODB.Recordset
		'
		If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
			rsEmployee.Open("SELECT *, EmployeeFirst & ' ' & EmployeeLast AS Name FROM tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		Else 'SQL Server
			rsEmployee.Open("SELECT *, EmployeeFirst + ' ' + EmployeeLast AS Name FROM tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		End If
		'
		rsLog = New ADODB.Recordset
		'
		If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
			'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			rsLog.Open("SELECT * FROM tblHours WHERE LogDate = #" & VB6.Format(tdtDate.CtlText, "mm/dd/yy") & "# ORDER BY RecID", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		Else 'SQL Server
			'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			rsLog.Open("SELECT * FROM tblHours WHERE LogDate = '" & VB6.Format(tdtDate.CtlText, "mm/dd/yy") & "' ORDER BY RecID", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		End If
		'
		rsEmployee.Find("Name = '" & txtName.Text & "'")
		'
		If Not rsEmployee.eof Then
			If DecryptStr(rsEmployee.Fields("Password").Value & "") = txtPassword.Text Then
				If bLog Then
					If InOut = 1 Then '(In)
						If Not rsLog.eof Then rsLog.MoveLast()
						rsLog.Find("Employee = '" & txtName.Text & "'",  , ADODB.SearchDirectionEnum.adSearchBackward)
						If rsLog.BOF Then
							rsLog.AddNew()
						Else
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If IsDbNull(rsLog.Fields("actualout").Value) Then
								MsgBox("You have already logged in but not logged out for this date. Please log out before logging back in.")
							Else
								rsLog.AddNew()
							End If
						End If
					Else 'InOut% = 0 (Out)
						rsLog.MoveLast()
						rsLog.Find("Employee = '" & txtName.Text & "'",  , ADODB.SearchDirectionEnum.adSearchBackward)
						If rsLog.BOF Then
							MsgBox("You have not logged in on this date. Please log in before logging out.")
						Else
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							If IsDbNull(rsLog.Fields("actualout").Value) Then
								bEdit = True
							Else
								MsgBox("You have logged in and logged out for this date. Please log in before logging out.")
							End If
						End If
					End If
					'
					If bEdit Or rsLog.EditMode = ADODB.EditModeEnum.adEditAdd Then
						'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						strMsg = "The following will be submitted: " & vbCrLf & vbCrLf & IIf(InOut, "LOG IN", "LOG OUT") & vbCrLf & vbCrLf & "Name: " & vbTab & vbTab & txtName.Text & vbCrLf & "Time: " & vbTab & vbTab & tdtDate.CtlText & " " & ttmTime.CtlText
						resp = MsgBox(strMsg, MsgBoxStyle.OKCancel)
						'
						If resp = MsgBoxResult.OK Then
							'If bEdit Then rsLog.Edit
							'
							If InOut = 1 Then
								rsLog.Fields("actualin").Value = Now
								'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
								rsLog.Fields("logdate").Value = CDate(VB6.Format(tdtDate.CtlText, "mm/dd/yy"))
								rsLog.Fields("intime").Value = ttmTime.Value
								User.Name = txtName.Text
							Else
								rsLog.Fields("actualout").Value = Now
								rsLog.Fields("outtime").Value = ttmTime.Value
								sUserName = ""
							End If
							'
							rsLog.Fields("Employee").Value = txtName.Text
							rsLog.Update()
							'
							SaveSetting(My.Application.Info.Title, "Settings", "User", txtName.Text)
							txtPassword.Text = ""
							txtPassword.Focus()
							'
							Submit = True
						Else
							Submit = False
						End If
					Else
						Submit = False
					End If
				Else
					User.Name = txtName.Text
					Submit = True
				End If
			Else
				MsgBox("Password incorrect. Please try again.")
				txtPassword.Focus()
				Submit = False
			End If
		Else
			MsgBox("Employee name not found. Please try again.")
			txtName.Focus()
			Submit = False
		End If
		'
		DBOps.ZapRS(rsEmployee)
		DBOps.ZapRS(rsLog)
		'
		Exit Function
ErrCall: 
		Submit = False
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmEmployeeLog.Submit", MsgBoxStyle.Critical, "Error")
	End Function
	
	Private Sub cmdContinue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdContinue.Click
		On Error GoTo ErrCall
		'
		Select Case iMode
			Case FLogon.enMode.enLogin
				If Submit(1, False) Then
					User.LogResults = True
					SaveSetting(My.Application.Info.Title, "Settings", "User", txtName.Text)
					Me.Close()
				Else
					User.LogResults = False
				End If
			Case FLogon.enMode.enLogout
				User.LogResults = True
				SaveSetting(My.Application.Info.Title, "Settings", "User", txtName.Text)
				Me.Close()
		End Select
		StrUser = txtName.Text
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmEmployeeLog.cmdContinue_Click.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub cmdLog_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLog.Click
		On Error GoTo ErrCall
		'
		Select Case iMode
			Case FLogon.enMode.enLogin
				If Submit(1, True) Then
					User.LogResults = True
					SaveSetting(My.Application.Info.Title, "Settings", "User", txtName.Text)
					Me.Close()
				Else
					User.LogResults = False
				End If
			Case FLogon.enMode.enLogout
				If Submit(0, True) Then
					User.LogResults = True
					SaveSetting(My.Application.Info.Title, "Settings", "User", txtName.Text)
					Me.Close()
				Else
					User.LogResults = False
				End If
		End Select
		StrUser = txtName.Text
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmEmployeeLog.cmdLog_Click.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'UPGRADE_WARNING: Form event FEmployeeLog.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FEmployeeLog_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error GoTo ErrCall
		'
		Dim sTimeofDay As String
		'
		Select Case CShort(VB6.Format(Now, "hh"))
			Case Is < 12
				If iMode = FLogon.enMode.enLogout Then
					sTimeofDay = "day"
				Else
					sTimeofDay = "morning"
				End If
			Case Is < 17
				sTimeofDay = "afternoon"
			Case Else
				sTimeofDay = "evening"
		End Select
		'
		Select Case iMode
			Case FLogon.enMode.enLogin
				lblWelcomeMessage.Text = "Good " & sTimeofDay & ". Welcome to " & My.Application.Info.Title & ". Please provide your name and password below:"
				cmdContinue.Text = "Continue"
				cmdLog.Text = "Continue and Log In"
			Case FLogon.enMode.enLogout
				lblWelcomeMessage.Text = "Thanks for using " & My.Application.Info.Title & ". Please supply your name and password if you will be logging out. Have a good " & sTimeofDay & "."
				cmdContinue.Text = "Exit"
				cmdLog.Text = "Exit and Log Out"
		End Select
		'
		ttmTime.Value = Now
		tdtDate.Value = Now
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmEmployeeLog.Form_Activate", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub FEmployeeLog_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrCall
		'
		' Get default username
		'
		txtName.Text = GetSetting(My.Application.Info.Title, "Settings", "User", "")
		txtName.SelectionStart = 0
		txtName.SelectionLength = Len(txtName.Text)
		
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmEmployeeLog.Form_Load", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub tmrClock_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrClock.Tick
		On Error GoTo ErrCall
		'
		Static Minutes As Short
		'
		If Minutes < 3 Then
			Minutes = Minutes + 1
		Else
			ttmTime.Value = Now
			tdtDate.Value = Now
			Minutes = 0
		End If
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmEmployeeLog.tmrClock_Timer", MsgBoxStyle.Critical, "Error")
	End Sub
	
	
	Public Property Mode() As FLogon.enMode
		Get
			Mode = iMode
		End Get
		Set(ByVal Value As FLogon.enMode)
			iMode = Value
		End Set
	End Property
	
	Private Sub txtPassword_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPassword.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 39 Then KeyAscii = 180
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class