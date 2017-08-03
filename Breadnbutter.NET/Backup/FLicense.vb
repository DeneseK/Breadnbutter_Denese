Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FLicense
	Inherits System.Windows.Forms.Form
	'
	
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	'
	'\\ General
	Private lSessionID As Integer
	Private sngFrmHtStd As Single
	Private sngFrmHtAdv As Single
	Private sLicExpireDays As String
	Private sLicExpireDate As String
	Private ctrlAct As System.Windows.Forms.Control
	'
	'\\ License Messages
	Private Const csLicMsgHeader As String = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman Times New Roman;}{\f3\froman Times New Roman;}}{\colortbl\red0\green0\blue0;}\deflang1033\pard\plain\f2\fs20"
	'
	Private Const csLicMsgApp As String = "\plain\f2\fs20\b\i PowerKey\plain\f2\fs20"
	'
	Private Const csLicMsgContact As String = "ontact the IT Manager"
	'
	Private Const csLicMsgError As String = csLicMsgContact & " to report this error."
	'
	Private Const csLicMsgLimited As String = "You cannot authorize or extend licenses."
	'
	Private Const csLicMsgUnauthorized As String = csLicMsgHeader & "\b Your license is not authorized.\par\par\plain\f2\fs20 " & "You cannot authorize or extend licenses.\par\par\plain\f2\fs20\b " & "C" & csLicMsgContact & " to authorize your license and unlock all the power of " & csLicMsgApp & ".\par }"
	'
	Private Const csLicMsgNeverAuthorized As String = csLicMsgHeader & "\b Welcome to " & csLicMsgApp & "!\par\par\plain\f2\fs20 " & "To authorize " & csLicMsgApp & " , c" & csLicMsgContact & "\plain\f2\fs20.\par\par " & "Until then, you will not be able to authorize licenses.\par }"
	'
	Private Const csLicMsg30DayEval As String = csLicMsgHeader & "\b Your license could not be authorized because it has previously been authorized for a 30-day trial.\par\par\par }"
	'
	Private Const csLicMsg15DayEval As String = csLicMsgHeader & "\b Your license could not be authorized because it has previously been authorized for 15-day trial extension.\par\par\par }"
	'
	Private Const csLicMsgValidDeauthorization As String = csLicMsgHeader & " Your license has been deauthorized.\par\par\par }"
	'
	Private Const csLicMsgInvalidDeauthorization As String = csLicMsgHeader & "\b Your license cannot be deauthorized because it is either currently unauthorized or in trial mode.\par\par\par }"
	'
	Private Const csLicMsgKeyValidAuthorization As String = csLicMsgHeader & "\b Your license has been authorized.\par\par\par }"
	'
	Private Const csLicMsgKeyValidExtension As String = csLicMsgHeader & "\b Your license has been extended.\par\par\par }"
	'
	Private Const csLicMsgInvalidExtension As String = csLicMsgHeader & "\b Your license cannot be extended because it is not authorized.\par\par\par }"
	'
	Private Const csLicMsgSiteKeyNotSpecified As String = csLicMsgHeader & "\b You must enter a site key obtained from Jason or Eric in order to authorize a license.\par\par\par }"
	'
	Private Const csLicMsgSiteCodeCompacted As String = csLicMsgHeader & "\b Your license could not be authorized because the site code you entered does not contain a space. Please check you code and attempt to authorize your license again.\par\par\par }"
	'
	Private Const csLicMsgKeyInvalid As String = csLicMsgHeader & "\b Your license could not be authorized because the site key you entered is invalid. Please verify your site key and attempt to authorize your license again.\par\par\par }"
	'
	Private Const csLicMsgKey1Invalid As String = csLicMsgHeader & "\b Your license could not be authorized because part one of the site key you entered is invalid. Please verify your site key and attempt to authorize your license again.\par\par\par }"
	'
	Private Const csLicMsgKey2Invalid As String = csLicMsgHeader & "\b Your license could not be authorized because part two of the site key you entered is invalid. Please verify your site key and attempt to authorize your license again.\par\par\par }"
	'
	Private Const csLicMsgSystemFailure As String = csLicMsgHeader & "\b Your license is damaged or security has been compromised. " & "C" & csLicMsgError & "\par\par " & csLicMsgUnauthorized & "\par }"
	'
	Private Const csLicMsgClockTurnedBack As String = csLicMsgHeader & "\b Your system calendar and/or clock has been turned back. " & "Please correct your system's date and time and restart " & csLicMsgApp & ".\par\par " & "If this error persists, c" & csLicMsgContact & "\par\par " & csLicMsgUnauthorized & "\par }"
	'
	Private Const csLicMsgSiteCodeReset As String = csLicMsgHeader & "\b Your site code has been refreshed.\par\par\par }"
	'
	Private Const csLicMsgClockReset As String = csLicMsgHeader & "\b Your license has been synchronized with your system's date and time.\par\par\par }"
	'
	Private Const csLicMsgAuthorized As String = csLicMsgHeader & " To extend your license, c" & csLicMsgContact & ".\par }"
	'
	Private Const csLicMsgExpired As String = csLicMsgHeader & "\plain\f2\fs20\  " & csLicMsgLimited & "\par\par " & "To renew your license, c" & csLicMsgContact & "\par }"
	Public Sub NotifyStatus(ByRef psLicSts As String)
		On Error GoTo ErrorHandler
		'
		With FMain.License
			'
			If psLicSts = "Error" Then
				DisplayStatus(csLicMsgHeader & "\b The following error has occurred: Error Number " & CStr(.LastErrorNumber) & " -- " & .LastErrorString & ".\par\par\plain\f2\fs20\  " & "C" & csLicMsgError & "\par\par " & "Your license may not be authorized. \par }")
				Exit Sub
			End If
			'
			'\\ Calculate and Display Site Code
			lSessionID = .get_UserNumber(5)
			If lSessionID = 0 Then lSessionID = .TCSessionCode
			'txtSiteCode.Text = Trim(CStr(FMain.lSecCompID)) & " " & Trim(CStr(lSessionID))
			txtSiteCode.Text = Trim(CStr(lSecCompID)) & " " & Trim(CStr(lSessionID))
			'
			sLicExpireDays = CStr(.DaysLeft)
			sLicExpireDate = .ExpireDateSoft
			'
			Select Case psLicSts
				Case "Licensed"
					Select Case .DaysLeft
						Case Is < 0
							DisplayStatus(csLicMsgExpired)
						Case Is > 1
							DisplayStatus(csLicMsgHeader & "\b You have " & .DaysLeft & " days remaining before your license expires. The expiration date is " & sLicExpireDate & ".\par\par" & csLicMsgAuthorized)
						Case 1
							DisplayStatus(csLicMsgHeader & "\b Your license will expire today. The expiration date is " & sLicExpireDate & ".\par\par" & csLicMsgAuthorized)
					End Select
				Case "Expired"
					DisplayStatus(csLicMsgHeader & "\b Your license expired on " & FMain.License.ExpireDateSoft & ".\par\par\ }" & csLicMsgExpired)
				Case "30DayEval"
					DisplayStatus(csLicMsg30DayEval)
				Case "15DayEval"
					DisplayStatus(csLicMsg15DayEval)
				Case "NeverAuthorized"
					DisplayStatus(csLicMsgNeverAuthorized)
				Case "Unauthorized"
					DisplayStatus(csLicMsgUnauthorized)
				Case "ClockTurnedBack"
					DisplayStatus(csLicMsgClockTurnedBack)
				Case "SystemFailure"
					DisplayStatus(csLicMsgSystemFailure)
				Case "Error"
					DisplayStatus(csLicMsgHeader & "\b The following error has occurred: Error Number " & CStr(.LastErrorNumber) & " -- " & .LastErrorString & ".\par\par\plain\f2\fs20\  " & "C" & csLicMsgError & "\par\par " & "Your license may not be authorized. \par }")
			End Select
			'
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.General.NotifyStatus")
	End Sub
	Public Sub DisplayStatus(ByRef psStatus As String)
		On Error GoTo ErrorHandler
		'
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		txtDescript.SelStart = Len(txtDescript.CtlText)
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		txtDescript.SelLength = Len(txtDescript.CtlText)
		txtDescript.RTFSelText = psStatus
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.General.DisplayStatus")
	End Sub
	Private Sub cmdAdv_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdv.ClickEvent
		On Error GoTo ErrorHandler
		'
		Select Case Trim(cmdAdv.Text)
			Case "Advanced >>"
				cmdAdv.Text = "<< Standard"
				Me.Height = VB6.TwipsToPixelsY(sngFrmHtAdv)
				CenterForm(Me)
			Case "<< Standard"
				cmdAdv.Text = "Advanced >>"
				Me.Height = VB6.TwipsToPixelsY(sngFrmHtStd)
				CenterForm(Me)
		End Select
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.cmdAdv.Click")
	End Sub
	Private Sub cmdAuthorize_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAuthorize.ClickEvent
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iDelPos As Short
		Dim sSiteCode As String
		Dim sRegKey1 As String
		Dim sRegKey2 As String
		'
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		sSiteCode = Trim(mskSiteKey.CtlText)
		'
		If sSiteCode = vbNullString Then
			DisplayResult(csLicMsgSiteKeyNotSpecified)
			FMain.License.ForceStatusChanged()
			Exit Sub
		End If
		'
		With FMain.License
			iDelPos = InStr(1, sSiteCode, " ", CompareMethod.Binary)
			If iDelPos > 0 Then
				sRegKey1 = VB.Left(sSiteCode, iDelPos - 1)
				sRegKey2 = Trim(VB.Right(sSiteCode, Len(sSiteCode) - iDelPos))
			Else
				sRegKey1 = sSiteCode
			End If
			.TCode(Val(sRegKey1), Val(sRegKey2), lSessionID, 0, 0)
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		If Err.Number = 6 Then '\\ No Space in Site Code
			If InStr(1, sSiteCode, " ", CompareMethod.Binary) > 0 Then
				DisplayResult(csLicMsgKeyInvalid)
			Else
				DisplayResult(csLicMsgSiteCodeCompacted)
			End If
			FMain.License.ForceStatusChanged()
			Exit Sub
		End If
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.cmdAuthorize.Click")
	End Sub
	Private Sub cmdDeauthorize_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDeauthorize.ClickEvent
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim lRes As Integer
		Dim sDaysLeft As String
		'
		If MsgBox("Are you certain you want to deauthorize the license? If you do, you will have to contact Hawkins Research before you will be able to create new claims using PowerClaim.", MsgBoxStyle.YesNo, "CONFIRM: Deauthorize License") = MsgBoxResult.No Then Exit Sub
		'
		With FMain.License
			'
			lRes = FMain.License.CPDelete(-1)
			'
			Select Case lRes
				Case 1
					If .ExpireMode = "P" Then
						If .ExpireDateSoft <> "00/00/00" Then
							If .IsExpired = False Then
								DisplayResult(csLicMsgValidDeauthorization)
								sDaysLeft = CStr(.DaysLeft / 1.27)
								If InStr(1, sDaysLeft, ".", CompareMethod.Binary) = 0 Then sDaysLeft = sDaysLeft & ".0"
								'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
								mskConfirm.CtlText = Replace(CStr(System.Date.FromOADate(Today.ToOADate + TimeOfDay.ToOADate)()), ".", " ") & " " & Replace(sDaysLeft, ".", " ")
								.ExpireMode = "D"
								.ExpireDateSoft = "00/00/00"
								ResetSiteCode()
								.ForceStatusChanged()
							Else
								'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
								mskConfirm.CtlText = vbNullString
								DisplayResult(csLicMsgInvalidDeauthorization)
								.ForceStatusChanged()
							End If
						Else
							'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
							mskConfirm.CtlText = vbNullString
							DisplayResult(csLicMsgInvalidDeauthorization)
							.ForceStatusChanged()
						End If
					Else
						'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						mskConfirm.CtlText = vbNullString
						DisplayResult(csLicMsgInvalidDeauthorization)
						.ForceStatusChanged()
					End If
				Case Else
					'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
					mskConfirm.CtlText = vbNullString
					DisplayResult(csLicMsgError)
			End Select
			'
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.cmdDeauthorize.Click")
	End Sub
	Private Sub cmdDone_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDone.ClickEvent
		On Error GoTo ErrorHandler
		'
		FMain.License.set_UserNumber(5, lSessionID)
		FMain.tmrSecChk.Enabled = True
		'FPrimary.bSecDisp = False
		bSecDisp = False
		Me.Close()
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.cmdDone.Click")
	End Sub
	Private Sub cmdRefresh_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRefresh.ClickEvent
		On Error GoTo ErrorHandler
		'
		ResetSiteCode()
		DisplayResult(csLicMsgSiteCodeReset)
		FMain.License.ForceStatusChanged()
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.cmdRefresh.Click")
	End Sub
	'UPGRADE_WARNING: Form event FLicense.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FLicense_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error GoTo ErrorHandler
		'
		FMain.License.ForceStatusChanged()
		'If FPrimary.bLicError = True Then NotifyStatus "Error"
		If bLicError = True Then NotifyStatus("Error")
		'
		'mskSiteKey.SetFocus
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.Form.Activate")
	End Sub
	
	'UPGRADE_NOTE: Form_Initialize was upgraded to Form_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Form_Initialize_Renamed()
		On Error GoTo ErrorHandler
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.MinHeight = 1965
		FormControl.MinWidth = VB6.PixelsToTwipsX(Me.Width)
		FormControl.DataForm = True
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.Form_Initialize")
		
	End Sub
	
	Private Sub FLicense_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrorHandler
		'
		bSecDisp = True
		'FPrimary.bSecDisp = True
		FMain.tmrSecChk.Enabled = False
		'
		sngFrmHtStd = VB6.PixelsToTwipsY(cmdAdv.Top) + VB6.PixelsToTwipsY(cmdAdv.Height) + 255
		sngFrmHtAdv = VB6.PixelsToTwipsY(fmeAdv.Top) + VB6.PixelsToTwipsY(fmeAdv.Height) + 255
		'
		Me.Height = VB6.TwipsToPixelsY(sngFrmHtStd)
		CenterForm(Me)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.Form.Load")
	End Sub
	Public Sub CenterForm(ByRef pFCur As System.Windows.Forms.Form, Optional ByRef pFRef As System.Windows.Forms.Form = Nothing, Optional ByRef pbCmnDlg As Boolean = False)
		On Error GoTo ErrorHandler
		'
		With pFCur
			If Not pbCmnDlg Then
				If pFRef Is Nothing Then
					.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(.Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Else
					.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(pFRef.Left) + (VB6.PixelsToTwipsX(pFRef.Width) - VB6.PixelsToTwipsX(.Width)) / 2), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(pFRef.Top) + (VB6.PixelsToTwipsY(pFRef.Height) - VB6.PixelsToTwipsY(.Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				End If
			Else
				.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(.Width) - 1000) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(.Height) - 175) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			End If
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.General.CenterForm")
	End Sub
	Private Sub mskConfirm_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mskConfirm.Enter
		On Error GoTo ErrorHandler
		'
		ctrlAct = mskConfirm
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.mskConfirm.GotFocus")
	End Sub
	Private Sub mskSiteKey_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mskSiteKey.Enter
		On Error GoTo ErrorHandler
		'
		ctrlAct = mskSiteKey
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.mskSiteKey.GotFocus")
	End Sub
	Private Sub mskSiteKey_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxTDBMask6.ITDBMaskEvents_KeyPressEvent) Handles mskSiteKey.KeyPressEvent
		On Error GoTo ErrorHandler
		'
		If eventArgs.KeyAscii = 22 Then
			'UPGRADE_ISSUE: Clipboard method Clipboard.GetText was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			My.Computer.Clipboard.GetText()
			eventArgs.KeyAscii = 0
		End If
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.mskSiteKey.KeyPress")
	End Sub
	Private Sub tbLic_ToolClick(ByVal eventSender As System.Object, ByVal eventArgs As AxActiveToolBars.DSSToolBarsEvents_ToolClickEvent) Handles tbLic.ToolClick
		On Error GoTo ErrorHandler
		'
		Select Case eventArgs.Tool.ID
			Case "MnuCtxTxtCut"
				'UPGRADE_WARNING: Couldn't resolve default property of object ctrlAct.SelText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				My.Computer.Clipboard.SetText(ctrlAct.SelText)
				'UPGRADE_WARNING: Couldn't resolve default property of object ctrlAct.SelText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ctrlAct.SelText = vbNullString
			Case "MnuCtxTxtCopy"
				'UPGRADE_WARNING: Couldn't resolve default property of object ctrlAct.SelText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				My.Computer.Clipboard.SetText(ctrlAct.SelText)
			Case "MnuCtxTxtPaste"
				If ctrlAct.Name = "mskSiteCode" Or ctrlAct.Name = "txtConfirmCode" Then Exit Sub
				'UPGRADE_WARNING: Couldn't resolve default property of object ctrlAct.SelText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ctrlAct.SelText = My.Computer.Clipboard.GetText
		End Select
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.tbLic.ToolClick")
	End Sub
	Private Sub txtSiteCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSiteCode.Enter
		On Error GoTo ErrorHandler
		'
		ctrlAct = txtSiteCode
		SelectText(txtSiteCode)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.txtSiteCode.GotFocus")
	End Sub
	Public Sub TransferLicense(ByRef sTransferMode As String)
		'On Error GoTo ErrorHandler
		''
		''\\ Local Declarations
		'Dim lRes      As String
		'Dim sFolder   As String
		'Dim sTxfrMode As String
		'Dim cdLic     As CCommonDialog
		''
		''\\ Obtain Directory From User
		'Select Case sTransferMode
		'  Case "Imprint"
		'    sTxfrMode = "Imprint License"
		'  Case "Export"
		'    sTxfrMode = "Export License"
		'  Case "Import"
		'    sTxfrMode = "Import License"
		'End Select
		''
		'Set cdLic = New CCommonDialog
		'cdLic.hWnd = FLicense.hWnd
		'cdLic.DialogTitle = sTxfrMode & ": Select Folder"
		'cdLic.ShowOpen
		'sFolder = cdLic.FileName
		'If sFolder = vbNullString Then Exit Sub
		''
		'sFolder = sFolder & "\sample.ini"
		'MsgBox sFolder
		''
		''\\ Perform Operation
		''lstOps.AddItem "Authorize License" & vbTab & "Failed" & vbTab & License.ErrorMessage
		'Select Case sTransferMode
		'  Case "Imprint"
		'    lRes = FPrimary.License.Transfer(1, sFolder)
		'    'If lRes = 1 Then MsgBox("Proceed to Step 2, "safsdf")
		'  Case "Export"
		'    lRes = FPrimary.License.Transfer(2, sFolder)
		'  Case "Import"
		'    lRes = 1
		'    FileCopy sFolder, App.Path & "\sample.ini"
		'End Select
		''
		'Select Case lRes
		'  Case 1
		'
		'  Case Else
		'
		'End Select
		'
		'Exit Sub
		''
		'ErrorHandler:
		'  ErrorMgr.Raise "FLicense", "cmdImprint.Click", Err.Number, Err.Description
	End Sub
	Public Sub DisplayResult(ByRef psResult As String)
		On Error GoTo ErrorHandler
		'
		txtDescript.SelStart = 0
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		txtDescript.SelLength = Len(txtDescript.CtlText)
		txtDescript.RTFSelText = psResult
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.General.DisplayResult")
	End Sub
	Public Sub NotifyResult(ByRef plAction As Integer, ByRef plData As Integer)
		On Error GoTo ErrorHandler
		'
		With FMain.License
			'
			.Enabled = False
			'
			If plAction <> 7 Then
				If plData = -1 Then
					DisplayResult(csLicMsgKey2Invalid)
					.Enabled = True
					.ForceStatusChanged()
					Exit Sub
				End If
			End If
			'
			Select Case plAction
				Case 0
					DisplayResult(csLicMsgKey1Invalid)
				Case 1 '\\ Authorize License
					If plData = 30 Then
						If CBool(GetSetting(My.Application.Info.Title, "License", "30DayEval", CStr(False))) = True Then
							'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
							mskSiteKey.CtlText = vbNullString
							DisplayResult(csLicMsg30DayEval)
							.Enabled = True
							Exit Sub
						End If
					ElseIf plData = 15 Then 
						If CBool(GetSetting(My.Application.Info.Title, "License", "15DayEval", CStr(False))) = True Then
							'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
							mskSiteKey.CtlText = vbNullString
							DisplayResult(csLicMsg15DayEval)
							.Enabled = True
							Exit Sub
						End If
					End If
					.CPAdd(0, 0)
					.ExpireMode = "P"
					.ExpireDateSoft = CStr(System.Date.FromOADate(Today.ToOADate + plData))
					If plData = 30 Then
						SaveSetting(My.Application.Info.Title, "License", "30DayEval", CStr(True))
					ElseIf plData = 15 Then 
						SaveSetting(My.Application.Info.Title, "License", "15DayEval", CStr(True))
					End If
					ResetExpirationNotifications()
					ResetSiteCode()
					DisplayResult(csLicMsgKeyValidAuthorization)
				Case 2 '\\ Extend License
					If .ExpireDateSoft <> "0/0/0" Then
						If .IsExpired = False Then
							.CPAdd(0, 0)
							.ExpireMode = "P"
							.ExpireDateSoft = CStr(System.Date.FromOADate(CDate(.ExpireDateSoft).ToOADate + plData))
							ResetExpirationNotifications()
							ResetSiteCode()
							DisplayResult(csLicMsgKeyValidExtension)
						Else
							'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
							mskSiteKey.CtlText = vbNullString
							DisplayResult(csLicMsgInvalidExtension)
						End If
					Else
						'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
						mskSiteKey.CtlText = vbNullString
						DisplayResult(csLicMsgInvalidExtension)
					End If
				Case 7
					.ResetLastUsedInfo()
					DisplayResult(csLicMsgClockReset)
				Case Else
					DisplayResult(csLicMsgError)
			End Select
			'
			.Enabled = True
			'
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.General.NotifyResult")
	End Sub
	Public Sub ResetExpirationNotifications()
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iCt As Short
		'
		SaveSetting(My.Application.Info.Title, "License", "NotifyExp30", CStr(True))
		SaveSetting(My.Application.Info.Title, "License", "NotifyExp15", CStr(True))
		'
		For iCt = 2 To 9
			SaveSetting(My.Application.Info.Title, "Lic", "NotifyExp" & CStr(iCt), CStr(True))
		Next 
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.General.ResetExpirationNotifications")
	End Sub
	Public Sub ResetSiteCode()
		On Error GoTo ErrorHandler
		'
		FMain.License.set_UserNumber(5, 0)
		lSessionID = FMain.License.TCSessionCode
		'txtSiteCode.Text = Trim(CStr(FPrimary.lSecCompID)) & " " & Trim(CStr(lSessionID))
		txtSiteCode.Text = Trim(CStr(lSecCompID)) & " " & Trim(CStr(lSessionID))
		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		mskSiteKey.CtlText = vbNullString
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FLicense.General.ResetSiteCode")
	End Sub
End Class