Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FNote
	Inherits System.Windows.Forms.Form
	
	'Private Enum eReportType
	'  DateRestoration
	'  Reauthorization
	'  Deathorization
	'  SecondAuthorization
	'  PaidAuthorization
	'  UserLimitChanged
	'  Sale
	'  EvalAddition
	'  EvalAuthorized
	'  Normal
	'End Enum
	'
	'Private NoteType As eReportType
	Dim bNewNote As Boolean
	Dim lProductID As Short
	Dim Contact As New CContact
	Dim ContactData As CContactData
	Dim Event1 As New CEvent
	Dim EventData As CEventData
	Dim bSaved As Boolean
	Dim bAuthEvent As Boolean
	Dim ProductData As New CProductData
	Dim Product As New CProduct
	'
	Private Declare Function pp_ctcodes Lib "TrgLib32.dll" (ByVal code As Integer, ByVal cenum As Integer, ByVal computer As Integer, ByVal seed As Integer) As Integer
	Private Declare Function pp_nencrypt Lib "KeyLib32.DLL" (ByVal Number As Integer, ByVal seed As Integer) As Integer
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: day was upgraded to day_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Declare Sub pp_cedate Lib "TrgLib32.dll" (ByVal cenum As Integer, ByRef month_Renamed As Integer, ByRef day_Renamed As Integer, ByRef year_Renamed As Integer)
	
	'UPGRADE_WARNING: Event cboProduct.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cboProduct.Change was upgraded to cboProduct.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cboProduct_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboProduct.TextChanged
		SwitchProduct()
		'
		If txtDays.Visible = True Then SuggestDays()
	End Sub
	
	Private Sub SwitchProduct()
		Select Case cboProduct.Text
			Case "PowerClaim PV"
				Product.Load(ProductData, 2)
				'
				Me.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF8080)
			Case "PowerClaim XML"
				Product.Load(ProductData, 1)
				'
				Me.BackColor = System.Drawing.SystemColors.Control
		End Select
		'
	End Sub
	
	'UPGRADE_WARNING: Event cboProduct.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboProduct_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboProduct.SelectedIndexChanged
		SwitchProduct()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Dim Index As Short = cmdCancel.GetIndex(eventSender)
		Me.Close()
	End Sub
	
	Private Sub cmdExpDate_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExpDate.ClickEvent
		txtExpDate.Value = FDatePick.DateText(txtExpDate.Value)
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		Dim Employee As New CEmployee
		'
		If lstTypes.Text = "" Then
			MsgBox("Please select type.")
			Exit Sub
		End If
		'
		If bAuthEvent Then
			If txtSubject.Text <> vbNullString Then
				CommitAuth()
			Else
				MsgBox("Please complete authorization before saving")
				Exit Sub
			End If
		End If
		'
		If lstTypes.Text = "Authorization Revision" Then
			CommitAuthorizationRevision()
		End If
		'
		If lstTypes.Text = "Sale" Or lstTypes.Text = "Sale Revision" Then
			CommitSale()
		End If
		
		'If Not bOk Then
		'   Exit Sub
		' End If
		'
		EventData.CustRecID = ContactData.ID
		'UPGRADE_WARNING: Couldn't resolve default property of object mskDate.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		EventData.EventDate = mskDate.Value
		EventData.EventSubject = txtSubject.Text
		EventData.EventResults = txtResults.Text
		EventData.EventUser = lblUser.Text
		'UPGRADE_WARNING: Couldn't resolve default property of object mskTime.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		EventData.EventTime = mskTime.Value
		EventData.EventType = lstTypes.Text
		EventData.Sticky = IIf((chkSticky.CheckState = 1), True, False)
		'
		Select Case cboProduct.Text
			Case "PowerClaim XML"
				EventData.ProductID = 1
			Case "PowerClaim PV"
				EventData.ProductID = 2
		End Select
		'
		If Employee.InGroup(StrUser, "Support") = True Then
			EventData.OpenCall = True
		Else
			EventData.OpenCall = False
		End If
		'
		If bNewNote = True Then
			Event1.Save(EventData, True)
		Else
			Event1.Save(EventData, False)
		End If
		'
		bSaved = True
		'
		'UPGRADE_NOTE: Object Employee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employee = Nothing
		'
		Me.Close()
	End Sub
	
	Private Function CommitAuthorizationRevision() As Boolean
		If Not ContactData Is Nothing Then
			If ContactData.ID > 0 Then
				txtSubject.Text = txtReviseAuthDays.Text
				'
				Select Case cboProduct.Text
					Case "PowerClaim XML"
						'UPGRADE_WARNING: Couldn't resolve default property of object mskReviseAuthDate.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.AuthDate = mskReviseAuthDate.Value
						ContactData.AuthDays = CShort(txtReviseAuthDays.Text)
					Case "PowerClaim PV"
						'UPGRADE_WARNING: Couldn't resolve default property of object mskReviseAuthDate.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.PVAuthDate = mskReviseAuthDate.Value
						ContactData.PVAuthDays = CShort(txtReviseAuthDays.Text)
				End Select
				'
				Contact.Save(ContactData, False)
			End If
		End If
	End Function
	
	Private Function CommitSale() As Boolean
		If Not ContactData Is Nothing Then
			If ContactData.ID > 0 Then
				If UCase(ContactData.MailState) = "KY" Then
					MsgBox("Remember to charge Sales Tax!", MsgBoxStyle.Information, "State Sales Tax")
				End If
				txtSubject.Text = txtPendingDays.Text
				'
				ContactData.Status = "Customer"
				'
				Select Case cboProduct.Text
					Case "PowerClaim XML"
						'UPGRADE_WARNING: Couldn't resolve default property of object mskDate.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.SaleDate = mskDate.Value
						'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.SaleDays = CShort(nnNum((Me.txtPendingDays.Text)))
						'
						If ContactData.AuthRemaining > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ContactData.GraceDays = ContactData.AuthRemaining + CShort(nnNum((Me.txtGraceDays.Text)))
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ContactData.GraceDays = CShort(nnNum((Me.txtGraceDays.Text)))
						End If
						'
						ContactData.OnlineAuths = 2
					Case "PowerClaim PV"
						'UPGRADE_WARNING: Couldn't resolve default property of object mskDate.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.PVSaleDate = mskDate.Value
						'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.PVSaleDays = CShort(nnNum((Me.txtPendingDays.Text)))
						'
						If ContactData.PVAuthRemaining > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ContactData.PVGraceDays = ContactData.PVAuthRemaining + CShort(nnNum((Me.txtGraceDays.Text)))
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ContactData.PVGraceDays = CShort(nnNum((Me.txtGraceDays.Text)))
						End If
						'
						ContactData.PVOnlineAuths = 2
				End Select
				'
				Contact.Save(ContactData, False)
			End If
		End If
		'
		
	End Function
	
	Private Function CommitAuth() As Boolean
		'If iAuthDays <= 0 Then
		' MsgBox "Please enter the number of days before continuing.", vbInformation, "Authorization"
		'Exit Sub
		'Else
		If Not ContactData Is Nothing Then
			If ContactData.ID > 0 Then
				'
				Select Case lstTypes.Text
					Case "Paid Authorization"
						ContactData.Status = "Customer"
						ContactData.AuthStatus = "Purchase"
					Case "Eval Authorized"
						ContactData.AuthStatus = "Evaluation"
					Case "Eval Addition"
						ContactData.AuthStatus = "Extended Evaluation"
				End Select
				'
				Select Case cboProduct.Text
					Case "PowerClaim XML"
						Select Case lstTypes.Text
							Case "Paid Authorization"
								ContactData.Status = "Customer"
								ContactData.AuthStatus = "Purchase"
							Case "Eval Authorized"
								ContactData.AuthStatus = "Evaluation"
							Case "Eval Addition"
								ContactData.AuthStatus = "Extended Evaluation"
						End Select
						'
						ContactData.GraceDays = 0
						ContactData.SaleDate = System.Date.FromOADate(0)
						ContactData.SaleDays = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.AuthDays = CShort(nnNum((txtSubject.Text)))
						'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
						ContactData.AuthRemaining = DateDiff(Microsoft.VisualBasic.DateInterval.Day, Now, DateAdd(Microsoft.VisualBasic.DateInterval.Day, nnNum((txtSubject.Text)), mskDate.Value))
						'UPGRADE_WARNING: Couldn't resolve default property of object mskDate.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.AuthDate = mskDate.Value
					Case "PowerClaim PV"
						Select Case lstTypes.Text
							Case "Paid Authorization"
								ContactData.Status = "Customer"
								ContactData.PVAuthStatus = "Purchase"
							Case "Eval Authorized"
								ContactData.PVAuthStatus = "Evaluation"
							Case "Eval Addition"
								ContactData.PVAuthStatus = "Extended Evaluation"
						End Select
						'
						ContactData.PVGraceDays = 0
						ContactData.PVSaleDate = System.Date.FromOADate(0)
						ContactData.PVSaleDays = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.PVAuthDays = CShort(nnNum((txtSubject.Text)))
						'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
						ContactData.PVAuthRemaining = DateDiff(Microsoft.VisualBasic.DateInterval.Day, Now, DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(txtSubject.Text), mskDate.Value))
						'UPGRADE_WARNING: Couldn't resolve default property of object mskDate.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ContactData.PVAuthDate = mskDate.Value
				End Select
				'
				'ContactData.Status = "Customer"
				'
				Contact.Save(ContactData, False)
			End If
		End If
		'End If
	End Function
	
	Private Function CommitAddition() As Boolean
		
	End Function
	
	
	Private Sub cmdSpell_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSpell.Click
		On Error GoTo cmdSpell_Click_EH
		'
		cmdSpell.Enabled = False 'avoid reentrancy
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'
		With vspNote
			'lSpellStart = 0
			.Clear()
			.Text = txtResults.Text
			.CheckText()
			'
			txtResults.Text = .Text
		End With
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		cmdSpell.Enabled = True
		'
		Exit Sub
cmdSpell_Click_EH: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		cmdSpell.Enabled = True
	End Sub
	
	'Private Sub cmdSupport_Click2(Index As Integer)
	'  Dim iSupportActID As Long
	'  Dim iCaseID As Long
	'  Dim Employee As New CEmployee
	'  On Error GoTo EH
	'  '
	'  If Index = 0 Then 'Submit
	'    If fAuthorizing Then
	'      Dim iAuthDays As Integer
	'      iAuthDays = CInt(nnNum(txtSubject.Text))
	'      '
	'      If iAuthDays <= 0 Then
	'        MsgBox "Please enter the number of days before continuing.", vbInformation, "Authorization"
	'        Exit Sub
	'      Else
	'        mskAuthDate.Text = mskDate.Text
	'        '
	'        Select Case cboType.Text
	'        Case "Eval Authorized"
	'          cboAuthStatus.Text = "Evaluation"
	'        Case "Eval Addition"
	'          cboAuthStatus.Text = "Extended Evaluation"
	'        Case "Sale"
	'          cboStatus.Text = "Customer"
	'          cboAuthStatus.Text = "Purchase"
	'        End Select
	'        '
	'        txtAuthDays.Text = txtSubject.Text
	'        lblAuthRemaining.Caption = DateDiff("d", Now, DateAdd("d", CDbl(txtSubject.Text), mskDate.DateValue))
	'        lblExpires.Caption = Date + CLng(lblAuthRemaining.Caption)
	'      End If
	'    End If
	'    '
	'    'Dim test As File
	'    'test.
	'
	'  End If
	'  '
	'  fAuthorizing = False
	'  fSupport = False
	'  Me.DataContact.Save
	'  '
	'  Set Employee = Nothing
	'  Exit Sub
	'EH:
	'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FCustomer.cmdOther_Click.", vbCritical, "Error"
	'End Sub
	
	Private Sub FNote_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object ContactData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ContactData = Nothing
		'UPGRADE_NOTE: Object EventData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EventData = Nothing
		'UPGRADE_NOTE: Object Contact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Contact = Nothing
		'UPGRADE_NOTE: Object Event1 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Event1 = Nothing
	End Sub
	
	Private Sub txtExpDate_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExpDate.Change
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If txtExpDate.Value Is System.DBNull.Value Then
			txtDays.Text = "0"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
			txtDays.Text = CStr(DateDiff(Microsoft.VisualBasic.DateInterval.Day, Now, nnNum((txtExpDate.Value))))
		End If
	End Sub
	
	Private Sub txtReviseAuthDays_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReviseAuthDays.Enter
		InputNumber.Setup(txtReviseAuthDays, CInputNumber.eNumberType.NumberTypeInteger)
	End Sub
	
	Private Sub txtSiteKey_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSiteKey.Enter
		On Error GoTo ErrorHandler
		'
		SelectText(txtSiteKey)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtSiteKey.GotFocus")
	End Sub
	
	Private Sub txtSiteDays_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSiteDays.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		On Error GoTo ErrorHandler
		'
		If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
			If KeyAscii <> 32 Then
				If KeyAscii <> System.Windows.Forms.Keys.Back Then KeyAscii = 0
			End If
		End If
		'
		GoTo EventExitSub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtSiteDays.KeyPress")
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtSiteCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSiteCode.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		On Error GoTo ErrorHandler
		'
		If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
			If KeyAscii <> 32 Then '\\ Space
				If KeyAscii <> System.Windows.Forms.Keys.Back Then '\\ Backspace
					If KeyAscii <> 22 Then '\\ CTRL-V: Paste
						KeyAscii = 0
					End If
				End If
			End If
		End If
		'
		GoTo EventExitSub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtSiteCode.KeyPress")
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtConfCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtConfCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConfCode.TextChanged
		On Error GoTo ErrorHandler
		'
		txtSiteDays.Text = vbNullString
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		txtClientDate.CtlText = "__/__/____"
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		txtClientTime.CtlText = "__:__:____"
		txtClientDate.BackColor = System.Drawing.Color.White
		txtClientDate.ForeColor = System.Drawing.Color.Black
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtConfCode.Change")
	End Sub
	
	Private Sub txtConfCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConfCode.Enter
		On Error GoTo ErrorHandler
		'
		SelectText(txtConfCode)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtConfCode.GotFocus")
	End Sub
	
	Private Sub txtConfCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConfCode.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		On Error GoTo ErrorHandler
		'
		If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
			If KeyAscii <> 32 Then '\\ Space
				If KeyAscii <> System.Windows.Forms.Keys.Back Then '\\ Backspace
					If KeyAscii <> 22 Then '\\ CTRL-V: Paste
						KeyAscii = 0
					End If
				End If
			End If
		End If
		'
		GoTo EventExitSub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtConfCode.KeyPress")
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtDays.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDays_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDays.TextChanged
		On Error GoTo ErrorHandler
		'
		txtSiteKey.Text = vbNullString
		'txtExpDate.Text = "__/__/____"
		txtExpDate.Value = DateAdd(Microsoft.VisualBasic.DateInterval.Day, Val(txtDays.Text), Now)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtDays.Change")
	End Sub
	
	Private Sub txtDays_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDays.Enter
		On Error GoTo ErrorHandler
		'
		SelectText(txtDays)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtDays.GotFocus")
	End Sub
	
	'UPGRADE_WARNING: Event txtRestCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRestCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRestCode.TextChanged
		On Error GoTo ErrorHandler
		'
		txtRestKey.Text = vbNullString
		txtRestDate.BackColor = System.Drawing.Color.White
		txtRestDate.ForeColor = System.Drawing.Color.Black
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtRestCode.Change")
	End Sub
	
	Private Sub txtRestCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRestCode.Enter
		On Error GoTo ErrorHandler
		'
		SelectText(txtRestCode)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtRestCode.GotFocus")
	End Sub
	
	Private Sub txtRestCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRestCode.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		On Error GoTo ErrorHandler
		'
		If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
			If KeyAscii <> 32 Then '\\ Space
				If KeyAscii <> System.Windows.Forms.Keys.Back Then '\\ Backspace
					If KeyAscii <> 22 Then '\\ CTRL-V: Paste
						KeyAscii = 0
					End If
				End If
			End If
		End If
		'
		GoTo EventExitSub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtRestCode.KeyPress")
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtSiteCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSiteCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSiteCode.TextChanged
		On Error GoTo ErrorHandler
		'
		txtDays_TextChanged(txtDays, New System.EventArgs())
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		txtSiteDate.CtlText = "__/__/____"
		' txtExpDate.Text = "__/__/____"
		txtSiteKey.Text = vbNullString
		txtSiteDate.BackColor = System.Drawing.Color.White
		txtSiteDate.ForeColor = System.Drawing.Color.Black
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtSiteCode.Change")
	End Sub
	
	Private Sub txtSiteCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSiteCode.Enter
		On Error GoTo ErrorHandler
		'
		SelectText(txtSiteCode)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtSiteCode.GotFocus")
	End Sub
	
	Public Sub SelectText(ByRef pctrlCur As System.Windows.Forms.Control)
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim fAlt As Boolean
		'
		With pctrlCur
			'UPGRADE_WARNING: Couldn't resolve default property of object pctrlCur.SelStart. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.SelStart = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object pctrlCur.SelLength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object pctrlCur.DisplayText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.SelLength = Len(pctrlCur.DisplayText)
			'UPGRADE_WARNING: Couldn't resolve default property of object pctrlCur.SelLength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If fAlt = True Then .SelLength = Len(pctrlCur)
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		If Err.Number = 438 Then '\\ Object Doesn't Support Property Or Method
			fAlt = True
			Resume Next
		End If
		'
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.General.SelectText")
	End Sub
	
	Private Sub txtDays_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDays.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		On Error GoTo ErrorHandler
		'
		If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
			If KeyAscii <> 32 Then
				If KeyAscii <> System.Windows.Forms.Keys.Back Then KeyAscii = 0
			End If
		End If
		'
		GoTo EventExitSub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtDays.KeyPress")
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub cmdUserNumGen_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUserNumGen.ClickEvent
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iDelPos As Short
		Dim lKeyCode As Integer
		Dim lUsers As Integer
		Dim lDateDay As Integer
		Dim lDateMonth As Integer
		Dim lDateYear As Integer
		Dim lSiteKey1 As Integer
		Dim lSiteKey2 As Integer
		Dim sSiteCode As String
		Dim sSiteCodeCompacted As String
		Dim sSiteCode1 As String
		Dim sSiteCode2 As String
		'
		Dim rsLog As New ADODB.Recordset
		'
		If ValidateLicense <> Today.ToOADate - dSecVar() Then Exit Sub
		'
		'If GeneralDataSpecified = False Then Exit Sub
		'
		sSiteCode = Trim(txtUserNumCode.Text)
		'
		If InStr(1, sSiteCode, " ", CompareMethod.Binary) <= 0 Then
			MsgBox("The site code you entered does not contain a space.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Site Code Does Not Contain Space")
			txtUserNumCode.Focus()
			Exit Sub
		End If
		'
		sSiteCodeCompacted = Replace(sSiteCode, " ", vbNullString)
		'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lUsers = nnNum((txtUserNum.Text))
		'lDays = CLng(txtDays.Text)
		lKeyCode = 3 'IIf(optWrite(0) = True, 1, 2)
		'
		If sSiteCodeCompacted = vbNullString Then
			MsgBox("Please enter a valid site code.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Site Code Not Specified")
			txtUserNumCode.Focus()
			Exit Sub
		End If
		'
		If IsNumeric(sSiteCodeCompacted) = False Then
			MsgBox("Site codes cannot contain non-numeric characters. Please enter a valid site code.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Non-Numeric Site Code Specified")
			txtUserNumCode.Focus()
			Exit Sub
		End If
		'
		If lUsers < 1 Then
			MsgBox("Please enter the number of users to authorize.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Number of Days Not Specified")
			txtUserNum.Focus()
			Exit Sub
		End If
		'
		iDelPos = InStr(1, sSiteCode, " ", CompareMethod.Binary)
		If iDelPos > 0 Then
			sSiteCode2 = VB.Left(sSiteCode, iDelPos - 1)
			sSiteCode1 = Trim(VB.Right(sSiteCode, Len(sSiteCode) - iDelPos))
		Else
			sSiteCode1 = sSiteCode
		End If
		'
		pp_cedate(CInt(sSiteCode1), lDateMonth, lDateDay, lDateYear)
		txtSiteDate.Value = CStr(lDateMonth) & "/" & CStr(lDateDay) & "/" & CStr(lDateYear)
		lSiteKey1 = pp_ctcodes(lKeyCode, Val(sSiteCode1), Val(sSiteCode2), ProductData.Seed1) '173)
		lSiteKey2 = pp_nencrypt(lUsers, ProductData.Seed2) '236
		'
		txtUserNumKey.Text = CStr(lSiteKey1) & " " & CStr(lSiteKey2)
		' If lKeyCode = 1 Then txtExpDate.Value = Date + lDays
		'
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If CDate(txtSiteDate.CtlText) <> Today Then
			txtSiteDate.BackColor = System.Drawing.Color.Red
			txtSiteDate.ForeColor = System.Drawing.Color.Yellow
			'  MsgBox "The client's system date (" & txtSiteDate.Text & ") does not match HRI's date (" & txtHRIDate.Text & ")." & vbCrLf & vbCrLf & _
			''  "Please request the client to verify and -- if necessary -- correct his system's date. After a date correction, PowerClaim must be shut down and restarted in order to generate a correct site code." & vbCrLf & vbCrLf & _
			''  "If the client verifies that his date is set correctly, he must refresh the site code in the advanced licensing area.", vbCritical + vbOKOnly, "Error: Client System Date Does Not Match HRI Date"
		End If
		'
		'\\ Log Action
		rsLog.LockType = ADODB.LockTypeEnum.adLockPessimistic
		rsLog.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
		rsLog.Open("SELECT * from tbllog", cnMain)
		' rsLog.RecordCount
		With rsLog
			.AddNew()
			.Fields("ID").Value = rsLog.RecordCount + 1
			.Fields("Company").Value = ContactData.ID & vbNullString
			.Fields("User").Value = ContactData.FirstName & " " & ContactData.LastName & vbNullString
			.Fields("Employee").Value = User.Name & vbNullString
			.Fields("ActionDateTime").Value = System.Date.FromOADate(Today.ToOADate + TimeOfDay.ToOADate) & vbNullString
			.Fields("ActionType").Value = "User Count Change"
			.Fields("ActionSubType").Value = "None"
			'    If optWrite(0) = True Then
			'      .Fields("SiteExpirationDate").Value = txtExpDate.Value & vbNullString
			'    End If
			'.Fields("SiteDays").Value = txtDays.Text & vbNullString
			.Fields("SiteCompID").Value = CInt(sSiteCode2) & vbNullString
			.Fields("SiteSessionID").Value = CInt(sSiteCode1) & vbNullString
			.Fields("SiteKey").Value = txtSiteKey.Text & vbNullString
			.Fields("SiteConfCode").Value = "N/A"
			.Fields("SiteDateTime").Value = txtSiteDate.Value & vbNullString
			.Fields("ProductID").Value = ProductData.ProductID
			'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Fields("UsersAllowed").Value = nnNum((Me.txtUserNum.Text))
			.UpdateBatch()
		End With
		'
		txtSubject.Text = txtUserNum.Text
		'
		rsLog.Close()
		'UPGRADE_NOTE: Object rsLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLog = Nothing
		'
		txtUserNumCode.Focus()
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.cmdUserNumGen.Click")
	End Sub
	
	Private Sub cmdDecrypt_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDecrypt.ClickEvent
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim lDelPos As Integer
		Dim sCodes() As String
		Dim sConfCode As String
		Dim sDateCode As String
		Dim sDaysCode As String
		'
		Dim rsLog As New ADODB.Recordset
		'
		If ValidateLicense <> Today.ToOADate - dSecVar() Then Exit Sub
		'
		'If GeneralDataSpecified = False Then Exit Sub
		'
		sConfCode = Trim(txtConfCode.Text)
		'
		If sConfCode = vbNullString Then
			MsgBox("Please enter a valid confirmation code.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Confirmation Code Not Specified")
			txtConfCode.Focus()
			Exit Sub
		End If
		'
		sCodes = Split(sConfCode)
		'
		If sCodes(0) = vbNullString Or sCodes(1) = vbNullString Or sCodes(2) = vbNullString Or sCodes(3) = vbNullString Then
			MsgBox("The confirmation code you entered is not valid. A valid confirmation code consists of four numbers separated by spaces.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Site Code Does Not Contain Space")
			txtConfCode.Focus()
			Exit Sub
		End If
		
		sDateCode = sCodes(0) & "." & sCodes(1)
		sDaysCode = sCodes(2) & "." & sCodes(3)
		'
		If IsNumeric(sDateCode) = False Or IsNumeric(sDaysCode) = False Then
			MsgBox("Confirmation codes cannot contain non-numeric characters. Please enter a valid confirmation code.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Non-Numeric Confirmation Code Specified")
			txtConfCode.Focus()
			Exit Sub
		End If
		'
		txtClientDate.Value = System.Date.FromOADate(CDbl(sDateCode))
		txtClientTime.Value = System.Date.FromOADate(CDbl(sDateCode))
		txtSiteDays.Text = CStr(System.Math.Round(CDbl(sDaysCode) * 1.27))
		'
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If CDate(txtClientDate.CtlText) <> Today Then
			txtClientDate.BackColor = System.Drawing.Color.Red
			txtClientDate.ForeColor = System.Drawing.Color.Yellow
			'MsgBox "The client's system date (" & txtClientDate.Text & ") does not match HRI's date (" & txtHRIDate.Text & ") Please request the client to verify and -- if necessary -- correct his system's date and deauthorize his license again. It is not necessary to shut down and restart PowerClaim in order to generate a correct confirmation code after the client corrects his system's date.", vbCritical + vbOKOnly, "Error: Client System Date Does Not Match HRI Date"
			'txtConfCode.SetFocus
		End If
		'
		'\\ Log Action
		rsLog.LockType = ADODB.LockTypeEnum.adLockPessimistic
		rsLog.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
		rsLog.Open("SELECT * from tbllog", cnMain)
		With rsLog
			.AddNew()
			.Fields("ID").Value = .RecordCount + 1
			.Fields("Company").Value = ContactData.ID & vbNullString
			.Fields("User").Value = ContactData.FirstName & " " & ContactData.LastName & vbNullString
			.Fields("Employee").Value = User.Name & vbNullString
			.Fields("ActionDateTime").Value = Today
			.Fields("ActionType").Value = "Deauthorization"
			.Fields("ActionSubType").Value = "N/A"
			' .Fields("SiteExpirationDate").Value = "N/A" 'Date '+ txtSiteDays.Text
			.Fields("SiteDays").Value = txtSiteDays.Text & vbNullString
			.Fields("SiteCompID").Value = 0
			.Fields("SiteSessionID").Value = 0
			.Fields("SiteKey").Value = "N/A"
			.Fields("SiteConfCode").Value = txtConfCode.Text & vbNullString
			.Fields("SiteDateTime").Value = CDate(txtClientDate.Value) ' + CDate(txtClientTime.Value)
			.UpdateBatch()
		End With
		'
		txtSubject.Text = CStr(0)
		'
		rsLog.Close()
		'UPGRADE_NOTE: Object rsLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLog = Nothing
		'
		Exit Sub
		'
ErrorHandler: 
		If Err.Number = 9 Then '\\ Subscript out of Range
			MsgBox("The confirmation code you entered is not valid. A valid confirmation code consists of four numbers separated by spaces.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Site Code Does Not Contain Space")
			txtConfCode.Focus()
			Exit Sub
		End If
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.cmdDecrypt.Click")
	End Sub
	
	Private Sub cmdRestGen_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRestGen.ClickEvent
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iDelPos As Short
		Dim lSiteKey1 As Integer
		Dim lDateDay As Integer
		Dim lDateMonth As Integer
		Dim lDateYear As Integer
		Dim sSiteCode As String
		Dim sSiteCodeCompacted As String
		Dim sSiteCode1 As String
		Dim sSiteCode2 As String
		'
		Dim rsLog As New ADODB.Recordset
		'
		If ValidateLicense <> Today.ToOADate - dSecVar() Then Exit Sub
		'
		'If GeneralDataSpecified = False Then Exit Sub
		'
		sSiteCode = Trim(txtRestCode.Text)
		'
		If InStr(1, sSiteCode, " ", CompareMethod.Binary) <= 0 Then
			MsgBox("The site code you entered does not contain a space.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Site Code Does Not Contain Space")
			txtSiteCode.Focus()
			Exit Sub
		End If
		'
		sSiteCodeCompacted = Replace(sSiteCode, " ", vbNullString)
		'
		If sSiteCodeCompacted = vbNullString Then
			MsgBox("Please enter a valid site code.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Site Code Not Specified")
			txtRestCode.Focus()
			Exit Sub
		End If
		'
		If IsNumeric(sSiteCodeCompacted) = False Then
			MsgBox("Site codes cannot contain non-numeric characters. Please enter a valid site code.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Non-Numeric Site Code Specified")
			txtRestCode.Focus()
			Exit Sub
		End If
		'
		iDelPos = InStr(1, sSiteCode, " ", CompareMethod.Binary)
		If iDelPos > 0 Then
			sSiteCode2 = VB.Left(sSiteCode, iDelPos - 1)
			sSiteCode1 = Trim(VB.Right(sSiteCode, Len(sSiteCode) - iDelPos))
		Else
			sSiteCode1 = sSiteCode
		End If
		'
		pp_cedate(CInt(sSiteCode1), lDateMonth, lDateDay, lDateYear)
		txtRestDate.Value = lDateMonth & "/" & lDateDay & "/" & lDateYear
		lSiteKey1 = pp_ctcodes(7, CInt(sSiteCode1), CInt(sSiteCode2), ProductData.Seed1) '173)
		txtRestKey.Text = CStr(lSiteKey1)
		'
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If CDate(txtRestDate.CtlText) <> Today Then
			txtRestDate.BackColor = System.Drawing.Color.Red
			txtRestDate.ForeColor = System.Drawing.Color.Yellow
			'  MsgBox "The client's system date (" & txtSiteDate.Text & ") does not match HRI's date (" & txtHRIDate.Text & ")." & vbCrLf & vbCrLf & _
			''  "Please request the client to verify and -- if necessary -- correct his system's date. After a date correction, PowerClaim must be shut down and restarted in order to generate a correct site code." & vbCrLf & vbCrLf & _
			''  "If the client verifies that his date is set correctly, he must refresh the site code in the advanced licensing area.", vbCritical + vbOKOnly, "Error: Client System Date Does Not Match HRI Date"
		End If
		'
		txtRestKey.Focus()
		'
		rsLog.LockType = ADODB.LockTypeEnum.adLockPessimistic
		rsLog.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
		rsLog.Open("SELECT * from tbllog", cnMain)
		'
		With rsLog
			.AddNew()
			.Fields("ID").Value = .RecordCount + 1
			.Fields("Company").Value = ContactData.ID & vbNullString
			.Fields("User").Value = ContactData.FirstName & " " & ContactData.LastName & vbNullString
			.Fields("Employee").Value = User.Name
			.Fields("ActionDateTime").Value = System.Date.FromOADate(Today.ToOADate + TimeOfDay.ToOADate)
			.Fields("ActionType").Value = "Restoration"
			.Fields("ActionSubType").Value = "N/A"
			' .Fields("SiteExpirationDate").Value = vbNullString
			' .Fields("SiteDays").Value = vbNull
			.Fields("SiteCompID").Value = CInt(sSiteCode2) & vbNullString
			.Fields("SiteSessionID").Value = CInt(sSiteCode1) & vbNullString
			.Fields("SiteKey").Value = txtRestKey.Text & vbNullString
			.Fields("SiteConfCode").Value = "N/A"
			'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			.Fields("SiteDateTime").Value = txtRestDate.CtlText
			.UpdateBatch()
		End With
		'
		txtSubject.Text = CStr(ContactData.AuthRemaining)
		'
		rsLog.Close()
		'UPGRADE_NOTE: Object rsLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLog = Nothing
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.cmdRestGen.Click")
	End Sub
	
	Private Sub txtRestKey_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRestKey.Enter
		On Error GoTo ErrorHandler
		'
		SelectText(txtRestKey)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.txtRestKey.GotFocus")
	End Sub
	
	Private Sub cmdGen_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGen.ClickEvent
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iDelPos As Short
		Dim lKeyCode As Integer
		Dim lDays As Integer
		Dim lDateDay As Integer
		Dim lDateMonth As Integer
		Dim lDateYear As Integer
		Dim lSiteKey1 As Integer
		Dim lSiteKey2 As Integer
		Dim sSiteCode As String
		Dim sSiteCodeCompacted As String
		Dim sSiteCode1 As String
		Dim sSiteCode2 As String
		'
		Dim rsLog As New ADODB.Recordset
		'
		If ValidateLicense <> Today.ToOADate - dSecVar() Then Exit Sub
		'
		'If GeneralDataSpecified = False Then Exit Sub
		'
		sSiteCode = Trim(txtSiteCode.Text)
		'
		If InStr(1, sSiteCode, " ", CompareMethod.Binary) <= 0 Then
			MsgBox("The site code you entered does not contain a space.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Site Code Does Not Contain Space")
			txtSiteCode.Focus()
			Exit Sub
		End If
		'
		sSiteCodeCompacted = Replace(sSiteCode, " ", vbNullString)
		lDays = CInt(txtDays.Text)
		lKeyCode = IIf(optWrite(0).Checked = True, 1, 2)
		'
		If sSiteCodeCompacted = vbNullString Then
			MsgBox("Please enter a valid site code.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Site Code Not Specified")
			txtSiteCode.Focus()
			Exit Sub
		End If
		'
		If IsNumeric(sSiteCodeCompacted) = False Then
			MsgBox("Site codes cannot contain non-numeric characters. Please enter a valid site code.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Non-Numeric Site Code Specified")
			txtSiteCode.Focus()
			Exit Sub
		End If
		'
		If lDays < 1 Then
			MsgBox("Please enter the number of days to authorize.", MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: Number of Days Not Specified")
			txtDays.Focus()
			Exit Sub
		End If
		'
		iDelPos = InStr(1, sSiteCode, " ", CompareMethod.Binary)
		If iDelPos > 0 Then
			sSiteCode2 = VB.Left(sSiteCode, iDelPos - 1)
			sSiteCode1 = Trim(VB.Right(sSiteCode, Len(sSiteCode) - iDelPos))
		Else
			sSiteCode1 = sSiteCode
		End If
		'
		pp_cedate(CInt(sSiteCode1), lDateMonth, lDateDay, lDateYear)
		txtSiteDate.Value = CStr(lDateMonth) & "/" & CStr(lDateDay) & "/" & CStr(lDateYear)
		lSiteKey1 = pp_ctcodes(lKeyCode, Val(sSiteCode1), Val(sSiteCode2), ProductData.Seed1) '173)
		lSiteKey2 = pp_nencrypt(lDays, ProductData.Seed2) '236
		'
		txtSiteKey.Text = CStr(lSiteKey1) & " " & CStr(lSiteKey2)
		If lKeyCode = 1 Then txtExpDate.Value = System.Date.FromOADate(Today.ToOADate + lDays)
		'
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		If CDate(txtSiteDate.CtlText) <> Today Then
			txtSiteDate.BackColor = System.Drawing.Color.Red
			txtSiteDate.ForeColor = System.Drawing.Color.Yellow
			'  MsgBox "The client's system date (" & txtSiteDate.Text & ") does not match HRI's date (" & txtHRIDate.Text & ")." & vbCrLf & vbCrLf & _
			''  "Please request the client to verify and -- if necessary -- correct his system's date. After a date correction, PowerClaim must be shut down and restarted in order to generate a correct site code." & vbCrLf & vbCrLf & _
			''  "If the client verifies that his date is set correctly, he must refresh the site code in the advanced licensing area.", vbCritical + vbOKOnly, "Error: Client System Date Does Not Match HRI Date"
		End If
		'
		'\\ Log Action
		rsLog.LockType = ADODB.LockTypeEnum.adLockPessimistic
		rsLog.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
		rsLog.Open("SELECT * from tbllog", cnMain)
		' rsLog.RecordCount
		With rsLog
			.AddNew()
			.Fields("ID").Value = rsLog.RecordCount + 1
			.Fields("Company").Value = ContactData.ID & vbNullString
			.Fields("User").Value = ContactData.FirstName & " " & ContactData.LastName & vbNullString
			.Fields("Employee").Value = User.Name & vbNullString
			.Fields("ActionDateTime").Value = System.Date.FromOADate(Today.ToOADate + TimeOfDay.ToOADate) & vbNullString
			.Fields("ActionType").Value = "Authorization"
			.Fields("ActionSubType").Value = IIf(lKeyCode = 1, "New", "Extension")
			If optWrite(0).Checked = True Then
				.Fields("SiteExpirationDate").Value = txtExpDate.Value & vbNullString
			End If
			.Fields("SiteDays").Value = txtDays.Text & vbNullString
			.Fields("SiteCompID").Value = CInt(sSiteCode2) & vbNullString
			.Fields("SiteSessionID").Value = CInt(sSiteCode1) & vbNullString
			.Fields("SiteKey").Value = txtSiteKey.Text & vbNullString
			.Fields("SiteConfCode").Value = "N/A"
			.Fields("SiteDateTime").Value = txtSiteDate.Value & vbNullString
			.Fields("ProductID").Value = ProductData.ProductID
			.UpdateBatch()
		End With
		'
		txtSubject.Text = txtDays.Text
		'
		rsLog.Close()
		'UPGRADE_NOTE: Object rsLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLog = Nothing
		'
		txtSiteKey.Focus()
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: Fcontact.cmdGen.Click")
	End Sub
	'End Sub
	
	Public Function LoadNote(ByRef plCustomerID As Integer, ByRef plNoteID As Integer) As Boolean
		Dim iPos As Short
		'
		bNewNote = False
		'
		ContactData = New CContactData
		'
		EventData = New CEventData
		'
		Contact.Load(ContactData, plCustomerID)
		'
		Event1.Load(EventData, plNoteID)
		'
		Me.mskDate.Value = EventData.EventDate
		Me.mskDate.Enabled = False
		'
		Me.mskTime.Value = EventData.EventTime
		Me.mskTime.Enabled = False
		'
		Me.lblUser.Text = EventData.EventUser
		'
		Me.txtSubject.Text = EventData.EventSubject
		Me.txtResults.Text = EventData.EventResults
		'
		Me.chkSticky.CheckState = IIf(EventData.Sticky, 1, 0)
		'
		Select Case EventData.ProductID
			Case 1
				cboProduct.Text = "PowerClaim XML"
			Case 2
				cboProduct.Text = "PowerClaim PV"
		End Select
		'
		Product.Load(ProductData, (EventData.ProductID))
		'
		SwitchProduct()
		'
		For iPos = 0 To lstTypes.Items.Count - 1
			If VB6.GetItemString(lstTypes, iPos) = EventData.EventType Then
				lstTypes.SelectedIndex = iPos
			End If
		Next 
		'
		Me.ShowDialog()
		'
		LoadNote = bSaved
	End Function
	
	Public Function NewNote(ByRef plCustomerID As Integer) As Boolean
		bNewNote = True
		'
		ContactData = New CContactData
		'
		EventData = New CEventData
		'
		Contact.Load(ContactData, plCustomerID)
		'
		cboProduct.Text = "PowerClaim XML"
		'
		lblUser.Text = User.Name
		'
		mskDate.Value = CDate(Now)
		'
		mskTime.Value = VB6.Format(Now, "hh:nn AM/PM")
		'
		Me.ShowDialog()
		'
		NewNote = bSaved
	End Function
	
	Private Sub FNote_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		mskDate._Value = Today
		mskTime._Value = TimeOfDay
		'
		LoadTypes()
		'
		cboProduct.Items.Add("PowerClaim XML")
		cboProduct.Items.Add("PowerClaim PV")
		'
		Product.Load(ProductData, 1)
		'
		bSaved = False
	End Sub
	
	Private Sub LoadTypes()
		Dim rsType As New ADODB.Recordset
		'
		rsType.Open("SELECT * FROM tblActivities WHERE ActivityType = 1 ORDER BY Activity", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		Do While Not rsType.eof
			lstTypes.Items.Add(rsType.Fields("Activity").Value & vbNullString)
			rsType.MoveNext()
		Loop 
		'
		rsType.Close()
		'
		lstTypes.Items.Add(vbNullString)
		'
		rsType.Open("SELECT * FROM tblActivities WHERE ActivityType = 0 ORDER BY Activity", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		Do While Not rsType.eof
			lstTypes.Items.Add(rsType.Fields("Activity").Value & vbNullString)
			rsType.MoveNext()
		Loop 
		'
		rsType.Close()
		'
		'UPGRADE_NOTE: Object rsType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsType = Nothing
	End Sub
	
	Private Sub SuggestAuthorizationRevisionData()
		If cboProduct.Text = "PowerClaim XML" Then
			'If ContactData.DaysPending > 0 Then
			Me.txtReviseAuthDays.Text = CStr(ContactData.AuthDays)
			Me.mskReviseAuthDate._Value = ContactData.AuthDate
			'End If
		End If
		'
		If cboProduct.Text = "PowerClaim PV" Then
			'If ContactData.PVDaysPending > 0 Then
			Me.txtReviseAuthDays.Text = CStr(ContactData.PVAuthDays)
			Me.mskReviseAuthDate._Value = ContactData.PVAuthDate
			' End If
		End If
	End Sub
	
	Private Sub SuggestDays()
		If cboProduct.Text = "PowerClaim XML" Then
			If ContactData.DaysPending > 0 Then
				txtDays.Text = CStr(ContactData.DaysPending)
			End If
		End If
		'
		If cboProduct.Text = "PowerClaim PV" Then
			If ContactData.PVDaysPending > 0 Then
				txtDays.Text = CStr(ContactData.PVDaysPending)
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event lstTypes.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstTypes_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstTypes.SelectedIndexChanged
		If bNewNote Then
			txtResults.Height = VB6.TwipsToPixelsY(5000)
			fmeAuth.Visible = False
			fmeDeathorization.Visible = False
			fmeUserNum.Visible = False
			fmeAuthRevise.Visible = False
			fmeDateRestore.Visible = False
			fmeSale.Visible = False
			txtSubject.Enabled = True
			bAuthEvent = False
			'
			Select Case lstTypes.Text
				Case "Eval Authorized", "Eval Addition", "Second Authorization", "Reauthorization", "Paid Authorization"
					bAuthEvent = True
					txtResults.Height = VB6.TwipsToPixelsY(4000)
					fmeAuth.Top = VB6.TwipsToPixelsY(5500)
					fmeAuth.Left = VB6.TwipsToPixelsX(1920)
					fmeAuth.Visible = True
					fmeAuth.BackColor = System.Drawing.SystemColors.Control
					txtSubject.Text = vbNullString
					txtSubject.Enabled = False
					'
					SuggestDays()
				Case "Authorization Revision"
					txtResults.Height = VB6.TwipsToPixelsY(4000)
					fmeAuthRevise.Top = VB6.TwipsToPixelsY(5500)
					fmeAuthRevise.Left = VB6.TwipsToPixelsX(1920)
					fmeAuthRevise.Visible = True
					fmeAuthRevise.BackColor = System.Drawing.SystemColors.Control
					txtSubject.Text = vbNullString
					txtSubject.Enabled = False
					'
					SuggestAuthorizationRevisionData()
				Case "Deathorization"
					txtResults.Height = VB6.TwipsToPixelsY(4000)
					fmeDeathorization.Top = VB6.TwipsToPixelsY(5500)
					fmeDeathorization.Left = VB6.TwipsToPixelsX(1920)
					fmeDeathorization.Visible = True
					fmeDeathorization.BackColor = System.Drawing.SystemColors.Control
					txtSubject.Text = vbNullString
					txtSubject.Enabled = False
				Case "User Limit Changed"
					txtResults.Height = VB6.TwipsToPixelsY(4000)
					fmeUserNum.Top = VB6.TwipsToPixelsY(5500)
					fmeUserNum.Left = VB6.TwipsToPixelsX(1920)
					fmeUserNum.Visible = True
					fmeUserNum.BackColor = System.Drawing.SystemColors.Control
					txtSubject.Enabled = False
					txtSubject.Text = vbNullString
				Case "Date Restoration"
					txtResults.Height = VB6.TwipsToPixelsY(4000)
					fmeDateRestore.Top = VB6.TwipsToPixelsY(5500)
					fmeDateRestore.Left = VB6.TwipsToPixelsX(1920)
					fmeDateRestore.Visible = True
					fmeDateRestore.BackColor = System.Drawing.SystemColors.Control
					txtSubject.Enabled = False
					txtSubject.Text = vbNullString
				Case "Sale", "Sale Revision"
					txtResults.Height = VB6.TwipsToPixelsY(4000)
					fmeSale.Top = VB6.TwipsToPixelsY(5500)
					fmeSale.Left = VB6.TwipsToPixelsX(1920)
					fmeSale.Visible = True
					fmeSale.BackColor = System.Drawing.SystemColors.Control
					txtSubject.Enabled = False
					txtSubject.Text = vbNullString
					txtGraceDays.Text = CStr(14)
			End Select
		End If
	End Sub
	
	'Private Sub txtSubject_KeyPress(KeyAscii As Integer)
	'  If fAuthorizing Then
	'    Select Case KeyAscii
	'    Case 8, 48 To 57 ' 48 to 57 0-9   8=backspace
	'    Case Else
	'      KeyAscii = 0
	'    End Select
	'  End If
	'End Sub
	
	Private Sub txtPendingDays_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPendingDays.Enter
		InputNumber.Setup(txtPendingDays, CInputNumber.eNumberType.NumberTypeInteger)
	End Sub
	
	Private Sub txtGraceDays_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGraceDays.Enter
		InputNumber.Setup(txtGraceDays, CInputNumber.eNumberType.NumberTypeInteger)
	End Sub
End Class