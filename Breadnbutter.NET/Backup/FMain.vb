Option Strict Off
Option Explicit On
Friend Class FMain
	Inherits System.Windows.Forms.Form
	
	'Private TrayHook As New CTrayHook
	Dim iMessageCount1 As Short
	Dim iMessageCount2 As Short
	Dim iMessageCount3 As Short
	Dim iMessageCount4 As Short
	Dim iFlasher As Short
	Dim iVMailTotal As Short
	
	
	'UPGRADE_NOTE: MDIForm_Initialize was upgraded to MDIForm_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MDIForm_Initialize_Renamed()
		On Error GoTo ErrCall
		'
		'RemoveCancelMenuItem Me
		'TrayHook.Setup Me, tmrTray, scTray, mnuTray
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.MDIForm_Initialize.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub FMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrCall
		'
		'
		Me.Text = "Bread 'n' Butter" '"Track It!   Datafile: " & DBOps.DBName
		'
		'  With sbMain
		'  .Panels.Add
		'  .Panels.Add
		'  .Panels.Add
		'  '
		'  .Panels(1).Width = (Me.Width / Screen.TwipsPerPixelX) - 200
		'  .Panels(1).Text = dbmMain.DbPath & dbmMain.DBName
		'  .Panels(2).Style = sbrDate
		'  .Panels(3).Style = sbrTime
		'  End With
		'
		RemoveCancelMenuItem(Me) 'disables the X
		'
		If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
			tbMain.ToolBars.Item(2).Tools.Item("ID_PathFile").Name = "Database: " & sAccessDB
		Else
			tbMain.ToolBars.Item(2).Tools.Item("ID_PathFile").Name = "Server: " & sSQLServerName & "  Database: " & sSQLServerDB
		End If
		tbMain.ToolBars.Item(2).Tools.Item("ID_UserName").Name = "User: " & StrUser
		
		'
		InitLicense()
		'
		tmrMessages_Tick(tmrMessages, New System.EventArgs())
		'
		'the form must be fully visible before calling Shell_NotifyIcon
		Me.Show()
		'Company.Contact.SearchID = 0
		FormMgr.ShowForm(Me.ActiveMDIChild, FContact, True)
		'Me.Refresh
		With nid
			.cbSize = Len(nid)
			.hWnd = Me.Handle.ToInt32
			.uID = VariantType.Null
			.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
			.uCallbackMessage = WM_MOUSEMOVE
			.hIcon = CInt(CObj(Me.Icon))
			.szTip = "Bread'n'Butter" & vbNullChar
		End With
		Shell_NotifyIcon(NIM_ADD, nid)
		'
		'Me.WindowState = 1
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.MDIForm_Load.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'UPGRADE_ISSUE: Form event MDIForm.MouseMove was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub MDIForm_MouseMove(ByRef Button As Short, ByRef Shift As Short, ByRef X As Single, ByRef y As Single)
		'this procedure receives the callbacks from the System Tray icon.
		Dim Result As Integer
		Dim msg As Integer
		'Debug.Print X, X / Screen.TwipsPerPixelX
		
		msg = X / VB6.TwipsPerPixelX
		'
		Select Case msg
			Case WM_LBUTTONUP '514 restore form window
				'Me.WindowState = vbMaximized
				'Result = SetForegroundWindow(Me.hwnd)
				'Me.Show
			Case WM_LBUTTONDBLCLK '515 restore form window
				Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
				Result = SetForegroundWindow(Me.Handle.ToInt32)
				Me.Show()
			Case WM_RBUTTONUP '517 display popup menu
				'Result = SetForegroundWindow(Me.hwnd)
				'Me.PopupMenu Me.mnuTray
		End Select
		If Button = 2 And Me.WindowState = 1 Then
			Result = SetForegroundWindow(Me.Handle.ToInt32)
			'UPGRADE_ISSUE: MDIForm method FMain.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Me.PopupMenu(Me.mnuTray)
		End If
	End Sub
	
	Private Sub FMain_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		On Error GoTo ErrCall
		'
		If UnloadMode <> System.Windows.Forms.CloseReason.UserClosing Then
			FLogon.Close()
			FLogon.Mode = FLogon.enMode.enLogout
			FLogon.ShowDialog()
			Cancel = Not User.LogResults
			'
			If Cancel = 0 Then
				Shell_NotifyIcon(NIM_DELETE, nid)
				'End ' Will this work?
			End If
			'
		End If
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.MDIForm_QueryUnload.", MsgBoxStyle.Critical, "Error")
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event FMain.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FMain_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error GoTo ErrCall
		'
		If Not Me.ActiveMDIChild Is Nothing Then
			FormMgr.ResizeForm(Me.ActiveMDIChild)
		End If
		'
		'sbMain.Panels(1).Width = (Me.Width / Screen.TwipsPerPixelX) - 200
		'
		'this is necessary to assure that the minimized window is hidden
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then Me.Hide()
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.MDIForm_Resize.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub FMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		' End
		'this removes the icon from the system tray
		'Shell_NotifyIcon NIM_DELETE, nid
		'Cancel = 1
	End Sub
	
	Public Sub mnuTrayExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTrayExit.Click
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.Show()
		Me.Close()
	End Sub
	
	Public Sub mnuTrayOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTrayOpen.Click
		'called when the user clicks the popup menu Restore command
		Dim Result As Integer
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Result = SetForegroundWindow(Me.Handle.ToInt32)
		Me.Show()
	End Sub
	Public Sub tbMain_Go(ByRef sResult As String)
		On Error GoTo ErrCall
		'
		Dim bActiveForm As Boolean
		'
		bActiveForm = Not Me.ActiveMDIChild Is Nothing
		If bActiveForm = True Then
			SaveSetting(My.Application.Info.Title, "Miscellaneous", "ActiveForm", Me.ActiveMDIChild.Name)
		Else
			SaveSetting(My.Application.Info.Title, "Miscellaneous", "ActiveForm", vbNullString)
		End If
		'
		'  If sResult = "ID_Lookup" Then
		'        Company.Contact.SearchID = 0
		'        FormMgr.ShowForm Me.ActiveForm, FContact
		'  End If
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.tbMain_ToolClick.", MsgBoxStyle.Critical, "Error")
		
	End Sub
	
	Public Sub tbMain_ToolClick(ByVal eventSender As System.Object, ByVal eventArgs As AxActiveToolBars.DSSToolBarsEvents_ToolClickEvent) Handles tbMain.ToolClick
		Dim RCustStatus As Object
		Dim RShipping As Object
		On Error GoTo ErrCall
		'
		Dim Employee As New CEmployee
		Dim bActiveForm As Boolean
		'
		bActiveForm = Not Me.ActiveMDIChild Is Nothing
		If bActiveForm = True Then
			SaveSetting(My.Application.Info.Title, "Miscellaneous", "ActiveForm", Me.ActiveMDIChild.Name)
		Else
			SaveSetting(My.Application.Info.Title, "Miscellaneous", "ActiveForm", vbNullString)
		End If
		'
		Select Case eventArgs.Tool.ID
			'\\ File Menu
			'  Case "ID_Cases"
			'    bCases = True
			'    FormMgr.ShowForm Me.ActiveForm, FCase, True
			Case "ID_VMail"
				' InitializeVmail
				FormMgr.ShowForm(Me.ActiveMDIChild, FVMail, True)
				'  Case "ID_Reports"
				'    FormMgr.ShowForm Me.ActiveForm, FReports, True
			Case "ID_HistoryReporter"
				FormMgr.ShowForm(Me.ActiveMDIChild, FHistory, True)
			Case "ID_OpenCalls"
				If Employee.InGroup((User.Name), "Management") = True Or Employee.InGroup((User.Name), "Development") = True Then
					FormMgr.ShowForm(Me.ActiveMDIChild, FSupportOpen, True)
				Else
					MsgBox("Access denied.", MsgBoxStyle.Critical, "")
				End If
			Case "ID_ContactReporter"
				FormMgr.ShowForm(Me.ActiveMDIChild, FReport, True)
			Case "ID_LicenseFacility"
				FormMgr.ShowForm(Me.ActiveMDIChild, FLicense, True)
			Case "ID_CallLog"
				FormMgr.ShowForm(Me.ActiveMDIChild, FSupportLog, True)
			Case "ID_AuthorizationLog"
				FormMgr.ShowForm(Me.ActiveMDIChild, FAuthLog, True)
			Case "ID_Customer"
				FormMgr.ShowForm(Me.ActiveMDIChild, FContact, True)
			Case "ID_Close"
				If bActiveForm Then
					'UPGRADE_ISSUE: Control FormControl could not be resolved because it was within the generic namespace ActiveMDIChild. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					Me.ActiveMDIChild.FormControl.SwitchFrom()
					Me.ActiveMDIChild.Close()
					Me.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000C)
				End If
				'  Case "ID_ProspectView" 'Prospecting View
				'    FormMgr.ShowForm Me.ActiveForm, FProspecting, True
			Case "ID_ProspectMgt" 'Prospect Management
				FormMgr.ShowForm(Me.ActiveMDIChild, FProspectMgt, True)
			Case "ID_Pricing" 'Pricing
				FPricing.ShowDialog()
			Case "ID_Lookup" 'Customer Lookup
				'Company.Contact.SearchID = 0
				FormMgr.ShowForm(Me.ActiveMDIChild, FContact, True)
				' Case "ID_KB"
				'  FKB.Height = Me.Height - 1500
				' FKB.Width = Me.Width - 1500
				'FKB.Show vbModal, Me
			Case "ID_TechSupport"
				'frmTechSupport.Show vbModal, Me
				'Case "ID_List"
				'FormMgr.ShowForm Me.ActiveForm, FCustomerList, True
			Case "ID_LogIn"
				FLogon.Mode = FLogon.enMode.enLogin
				FLogon.ShowDialog()
			Case "ID_Logout"
				FLogon.Mode = FLogon.enMode.enLogout
				FLogon.ShowDialog()
			Case "ID_Prefs"
				VB6.ShowForm(FPrefs, VB6.FormShowConstants.Modal, Me)
			Case "ID_Hours"
				If Employee.InGroup((User.Name), "Management") = True Or Employee.InGroup((User.Name), "Development") = True Then
					'FHours.Show vbModal
					FormMgr.ShowForm(Me.ActiveMDIChild, FHours, True)
				Else
					MsgBox("Access denied.", MsgBoxStyle.Critical, "")
				End If
			Case "ID_Password"
				SetPassword()
				'Case "ID_BatchHist"
				'FBatchHistory.Show vbModal, Me
			Case "ID_MailingLbl"
				VB6.ShowForm(FPrintLabels, VB6.FormShowConstants.Modal, Me)
			Case "ID_SelectCust"
				FormMgr.ShowForm(Me.ActiveMDIChild, FSelect, True)
			Case "ID_ShipRpt"
				'UPGRADE_WARNING: Couldn't resolve default property of object RShipping.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RShipping.Show()
			Case "ID_CustStatus"
				'UPGRADE_WARNING: Couldn't resolve default property of object RCustStatus.DBName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RCustStatus.DBName(DBOps.DBName)
				'UPGRADE_WARNING: Couldn't resolve default property of object RCustStatus.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				RCustStatus.Show()
			Case "ID_OpenDB"
				'OpenDatabase
			Case "ID_Exit"
				Me.Close()
				FVMail.Close()
			Case "ID_Utility"
				VB6.ShowForm(FUtility,  , Me)
				'  Case "ID_EMail"
				'    FEMail.Show , Me
			Case "ID_CallTimes"
				FormMgr.ShowForm(Me.ActiveMDIChild, FCallStats, True)
				'FCallStats.Show , Me
			Case "ID_PhoneChart"
				FormMgr.ShowForm(Me.ActiveMDIChild, frmMultiChart, True)
				'FCallStats.Show , Me
			Case "ID_EmployeeMgt"
				If Employee.InGroup((User.Name), "Development") = True Then ' Or Employee.InGroup(User.Name, "Management") = True Then
					FormMgr.ShowForm(Me.ActiveMDIChild, FEmployeeMgt, True)
				Else
					MsgBox("Access denied.", MsgBoxStyle.Critical, "")
				End If
				'
		End Select
		'
		'UPGRADE_NOTE: Object Employee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employee = Nothing
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.tbMain_ToolClick.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Public WriteOnly Property ControlsEnabled() As Boolean
		Set(ByVal Value As Boolean)
			On Error GoTo ErrCall
			'
			Dim i, j As Short
			'
			i = tbMain.Tools.Count
			'
			tbMain.Redraw = False
			For j = 1 To i
				tbMain.Tools.Item(j).Enabled = Value
			Next j
			tbMain.Redraw = True
			'
			'sbMain.Enabled = pbStatus
			'
			Exit Property
ErrCall: 
			MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.ControlsEnabled.", MsgBoxStyle.Critical, "Error")
		End Set
	End Property
	
	Private Sub SetPassword()
		'On Error Resume Next
		On Error GoTo EH
		'
		Dim rsEmployee As ADODB.Recordset
		rsEmployee = New ADODB.Recordset
		'
		If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
			rsEmployee.Open("Select EmployeeFirst & ' ' & EmployeeLast AS Employee, Password FROM tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		Else
			rsEmployee.Open("SELECT *, EmployeeFirst + ' ' + EmployeeLast AS Employee FROM tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
			'rsEmployee.Open "UpSelectEmployeeList", cnMain, adOpenDynamic, adLockOptimistic, adCmdStoredProc
		End If
		'
		rsEmployee.Find("Employee = '" & User.Name & "'") ', , adSearchForward
		'
		If Not rsEmployee.eof Then
			FSetPassword.Setup(DecryptStr(rsEmployee.Fields("Password").Value & ""), 0, 50, False)
			VB6.ShowForm(FSetPassword, VB6.FormShowConstants.Modal, Me)
			'
			If FSetPassword.PwdOK And Not FSetPassword.Cancelled Then
				rsEmployee.Fields("Password").Value = EncryptStr((FSetPassword.NewPwd))
				rsEmployee.Update()
				cnMain.Execute("EXEC sp_password NULL, '" & Rot39(FSetPassword.NewPwd) & "', '" & Replace(User.Name, " ", "") & "'")
			End If
			'
			FSetPassword.Close()
			'UPGRADE_NOTE: Object FSetPassword may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			FSetPassword = Nothing
		End If
		'
		rsEmployee.Close()
		'UPGRADE_NOTE: Object rsEmployee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsEmployee = Nothing
		Exit Sub
EH: 
		MsgBox(Err.Description & " in Reset Password.")
	End Sub
	
	Private Sub OpenDatabase()
		On Error GoTo EH
		'
		Dim frm As System.Windows.Forms.Form
		'
		For	Each frm In My.Application.OpenForms
			If frm.Name <> "FMain" Then
				frm.Close()
				Me.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000C)
			End If
		Next frm
		'
		Dim sPath, sFile As String
		'
		sPath = FileOps.IsolatePath((DBOps.DBName))
		sFile = FileOps.IsolateFile((DBOps.DBName))
		'
		If DBOps.GetPathFile(sPath, sFile, "PowerClaim Customers") Then
			If Not DBOps.OpenConnection(cnMain, sPath, sFile, "Bread 'n' Butter Data") Then
				MsgBox("Invalid database")
				End
			Else
				SaveSetting(My.Application.Info.Title, "File", "PCCustomersName", sFile)
				SaveSetting(My.Application.Info.Title, "File", "PCCustomersPath", sPath)
				'
				tbMain.ToolBars.Item(2).Tools.Item("ID_PathFile").Name = FileOps.IsolatePath((DBOps.DBName)) & FileOps.IsolateFile((DBOps.DBName))
			End If
		End If
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in Open Database.")
	End Sub
	
	Public Sub tmrMessages_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrMessages.Tick
		Dim rsUser As New ADODB.Recordset
		Dim rsMessages As New ADODB.Recordset
		
		iMessageCount1 = 0
		iMessageCount2 = 0
		iMessageCount3 = 0
		iMessageCount4 = 0
		'
		'StrUser = cmbUser
		rsUser.Open("select [EmployeeFirst], [EmployeeLast], [Groups] from tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rsUser
			Do While Not .eof
				If LCase(StrUser) = LCase(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) Then
					iGroupNumber = .Fields("Groups").Value '& vbNullString
				End If
				.MoveNext()
			Loop 
			.Close()
		End With
		
		rsMessages.Open("SELECT [Group], [Completed] from TVMailMessages WHERE Completed = 'False'", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		
		With rsMessages
			If Not .eof Then
				.MoveFirst()
				While Not .eof
					If .Fields("Group").Value = "Authorizations" Then
						iMessageCount1 = iMessageCount1 + 1
					End If
					If .Fields("Group").Value = "Sales" Then
						iMessageCount2 = iMessageCount2 + 1
					End If
					If .Fields("Group").Value = "Support" Then
						iMessageCount3 = iMessageCount3 + 1
					End If
					If .Fields("Group").Value = "Operator" Then
						iMessageCount4 = iMessageCount4 + 1
					End If
					.MoveNext()
				End While
			End If
		End With
		'
		If iFlasher > 2 Then
			iFlasher = 1
		End If
		'
		
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(Mid(GetLastUpdate, 1, 11)), TimeOfDay) > 30 Or DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(Mid(GetLastUpdate, 12, 11)), Today) > 0 Then
			tmrMessages.Interval = 100
			If iFlasher = 1 Then
				tbMain.ToolBars.Item(2).Tools.Item("ID_Messages").Name = " ***Server Is Down, Tell Supervisor***"
			Else
				tbMain.ToolBars.Item(2).Tools.Item("ID_Messages").Name = "       Server Is Down, Tell Supervisor"
			End If
			'
			iFlasher = iFlasher + 1
		Else
			tmrMessages.Interval = 30000
			'
			tbMain.ToolBars.Item(2).Tools.Item("ID_Messages").Name = GetUserGroups '"Msgs: Auth-" & iMessageCount1 & "  Sales-" & iMessageCount2 & "  Support-" & iMessageCount3 & "  Operator-" & iMessageCount4
			'
			'If iMessageCount1 > 0 Then
			' tbMain.ToolBars(2).Tools("ID_Messages").ForeColor = vbRed
			'Else
			'tbMain.ToolBars(2).Tools("ID_Messages").ForeColor = vbBlack
			'End If
			'iFlasher = iFlasher + 1
		End If
		
		
		
		
		
		
		'tbMain.ToolBars(2).Tools("ID_Messages").Name = "New Messages: " & iMessageCount1
		'
	End Sub
	
	Private Sub tmrSecChk_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrSecChk.Tick
		On Error GoTo ErrorHandler
		'
		bLicTimer = True
		Me.License.ForceStatusChanged()
		bLicTimer = False
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.tmrSecChk.Timer")
	End Sub
	
	Private Sub License_Trigger(ByVal eventSender As System.Object, ByVal eventArgs As AxSKCLLib._ILFileEvents_TriggerEvent) Handles License.Trigger
		On Error GoTo ErrorHandler
		'
		FLicense.NotifyResult(eventArgs.event_num, eventArgs.event_data)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.License.Trigger")
	End Sub
	
	Private Sub License_StatusChanged(ByVal eventSender As System.Object, ByVal eventArgs As AxSKCLLib._ILFileEvents_StatusChangedEvent) Handles License.StatusChanged
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iSecValPair As Short
		Dim iDlgRsp As Short
		Dim sDlgMsg As String
		'
		With Me.License
			'
			iSecValPair = Int((3 - 1 + 1) * Rnd() + 1)
			'
			If .LibTest(lSecValCode(iSecValPair)) <> lSecValRslt(iSecValPair) Then
				If bSecDisp = False Then
					If bLicTimer = False Then VB6.ShowForm(FLicense, VB6.FormShowConstants.Modal, Me)
				Else
					FLicense.NotifyStatus("Failed")
				End If
			End If
			'
			If System.Date.FromOADate(CPCheck) = (System.Date.FromOADate(Today.ToOADate - dSecVar)) Then
				If .IsExpired Then
					If bSecDisp = False Then
						If bLicTimer = False Then FormMgr.ShowForm(Me.ActiveMDIChild, FLicense)
						'If bLicTimer = False Then FLicense.Show vbModal, FMain
					Else
						FLicense.NotifyStatus("Expired")
					End If
				Else
					If .IsClockTurnedBack Then
						If bSecDisp = False Then
							If bLicTimer = False Then FormMgr.ShowForm(Me.ActiveMDIChild, FLicense)
							'If bLicTimer = False Then FLicense.Show vbModal, FMain
						Else
							FLicense.NotifyStatus("ClockTurnedBack")
						End If
					Else
						If bSecDisp = True Then
							FLicense.NotifyStatus("Licensed")
						ElseIf bLicChecked = False Then 
							Select Case .DaysLeft
								Case 1
									sDlgMsg = "Your license will expire after today!"
									sDlgMsg = sDlgMsg & vbCrLf & vbCrLf
									sDlgMsg = sDlgMsg & "Contact Hawkins Research, Inc. at 1-800-736-1246 to extend your license."
									iDlgRsp = MsgBox(sDlgMsg, MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "ATTENTION: License Expires Today")
								Case 30, 15, Is <= 10
									If CBool(GetSetting(My.Application.Info.Title, "Lic", "NotifyExp" & Trim(CStr(.DaysLeft)), CStr(True))) Then
										sDlgMsg = "You only have " & Trim(CStr(.DaysLeft)) & " remaining before your license expires."
										sDlgMsg = sDlgMsg & "Your license will expire on " & Trim(CStr(.ExpireDateSoft)) & "."
										sDlgMsg = sDlgMsg & vbCrLf & vbCrLf
										sDlgMsg = sDlgMsg & "Contact Hawkins Research, Inc. at 1-800-736-1246 to extend your license."
										sDlgMsg = sDlgMsg & vbCrLf
										sDlgMsg = sDlgMsg & vbCrLf & vbCrLf & "Do you want to see this message the next time you start " & My.Application.Info.Title & " today?"
										iDlgRsp = MsgBox(sDlgMsg, MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "ATTENTION: License Expires Soon")
										SaveSetting(My.Application.Info.Title, "Lic", "NotifyExp" & Trim(CStr(.DaysLeft)), IIf(iDlgRsp = MsgBoxResult.Yes, True, False))
									End If
							End Select
						End If
					End If
				End If
			Else
				'MsgBox .ExpireMode
				If .ExpireMode = "D" Then
					If .IsClockTurnedBack Then
						If bSecDisp = False Then
							If bLicTimer = False Then FormMgr.ShowForm(Me.ActiveMDIChild, FLicense)
							'If bLicTimer = False Then FLicense.Show vbModal, FMain
						Else
							FLicense.NotifyStatus("ClockTurnedBack")
						End If
					Else
						If bSecDisp = False Then
							'If bLicTimer = False Then FLicense.Show vbModal, FMain
							If bLicTimer = False Then FormMgr.ShowForm(Me.ActiveMDIChild, FLicense)
						Else
							FLicense.NotifyStatus("NeverAuthorized")
						End If
					End If
				Else
					If bSecDisp = False Then
						FormMgr.ShowForm(Me.ActiveMDIChild, FLicense)
						'FLicense.Show vbModal, FMain
					Else
						FLicense.NotifyStatus("SystemFailure")
					End If
				End If
			End If
			'
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.License.StatusChanged")
	End Sub
	
	Private Sub License_Error(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles License.Error
		On Error GoTo ErrorHandler
		'
		If bLicError = True Then
			bLicError = False
			FLicense.NotifyStatus("Error")
			Exit Sub
		End If
		'
		bLicError = True
		'
		If bSecDisp = False Then FormMgr.ShowForm(Me.ActiveMDIChild, FLicense)
		' If bSecDisp = False Then FLicense.Show vbModal, FMain
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.License.Error")
	End Sub
	
	Private Sub tmrTray_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrTray.Tick
		Dim iAddon As Short
		'
		If CDbl(GetSetting(My.Application.Info.Title, "Settings", "Icon", CStr(0))) = 1 Then
			iVMailTotal = 0
		End If
		'
		If iVMailTotal > 0 Then
			Select Case iVMailTotal
				Case Is < 3
					iAddon = 0
				Case 3, 4
					iAddon = 8 '4
				Case 5, 6
					iAddon = 16 '8
				Case 7, 8
					iAddon = 24 '12
				Case 9, 10
					iAddon = 32 '16
				Case Else
					iAddon = 40 '20
			End Select
			
			iFlasher = iFlasher + 1
			'
			'UPGRADE_WARNING: Lower bound of collection TrayIconList2.ListImages has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			nid.hIcon = CInt(CObj(TrayIconList2.Images.Item(iFlasher + iAddon)))
			Shell_NotifyIcon(NIM_MODIFY, nid)
			If iFlasher = 8 Then iFlasher = 0
			'
		Else
			nid.hIcon = CInt(CObj(Me.Icon))
			Shell_NotifyIcon(NIM_MODIFY, nid)
		End If
	End Sub
	
	Public Function GetUserGroups() As String
		Dim iTemp As Short
		Dim sTempMsg As String
		'
		iVMailTotal = 0
		'
		iTemp = iGroupNumber
		If iTemp >= 8 Then
			iVMailTotal = iMessageCount1
			iTemp = iTemp - 8
		End If
		'
		If iTemp >= 4 Then
			iVMailTotal = iVMailTotal + iMessageCount2
			iTemp = iTemp - 4
		End If
		'
		If iTemp >= 2 Then
			iVMailTotal = iVMailTotal + iMessageCount3
			iTemp = iTemp - 2
		End If
		'
		If iTemp >= 1 Then
			iVMailTotal = iVMailTotal + iMessageCount4
		End If
		'
		sTempMsg = "Msgs: Auth-" & iMessageCount1 & "  Sales-" & iMessageCount2 & "  Support-" & iMessageCount3 & "  Operator-" & iMessageCount4
		
		GetUserGroups = sTempMsg
		'
	End Function
End Class