Option Strict Off
Option Explicit On
Friend Class FSetPassword
	Inherits System.Windows.Forms.Form
	
	Private iMinPwdLen As Short
	Private iMaxPwdLen As Short
	
	Public NewPwd As String
	Public Cancelled As Boolean
	Public PwdOK As Boolean
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error GoTo ErrCall
		'
		Dim bPwdOK As Boolean
		bPwdOK = True
		'
		bPwdOK = (txtNewPwd.Text = txtVerifyPwd.Text)
		'
		If iMaxPwdLen > 0 Then If bPwdOK Then bPwdOK = Len(txtNewPwd.Text) <= iMaxPwdLen
		'
		If iMinPwdLen > 0 Then If bPwdOK Then bPwdOK = Len(txtNewPwd.Text) >= iMinPwdLen
		'
		If bPwdOK Then
			NewPwd = txtNewPwd.Text
			PwdOK = True
			Cancelled = False
			Me.Hide()
		Else
			txtNewPwd.Focus()
			txtNewPwd.SelectionStart = 0
			txtNewPwd.SelectionLength = Len(txtNewPwd.Text)
			'
			MsgBox("Valid password not supplied.", MsgBoxStyle.Information, "Invalid Password")
			PwdOK = False
		End If
		'
		Exit Sub
ErrCall: 
		PwdOK = False
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSetPassword.cmdOK_Click", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		On Error GoTo ErrCall
		'
		Cancelled = True
		Me.Hide()
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSetPassword.Command1_Click", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'UPGRADE_WARNING: Form event FSetPassword.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FSetPassword_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error GoTo ErrCall
		'
		txtNewPwd.Focus()
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSetPassword.Form_Activate", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Public Sub Setup(ByRef psOldPwd As String, ByRef piMinPwdLen As Short, ByRef piMaxPwdLen As Short, ByRef pbHideOldPwd As Boolean)
		On Error GoTo ErrCall
		'
		iMinPwdLen = piMinPwdLen
		iMaxPwdLen = piMaxPwdLen
		'
		'UPGRADE_WARNING: TextBox property txtNewPwd.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		txtNewPwd.Maxlength = piMaxPwdLen
		'UPGRADE_WARNING: TextBox property txtVerifyPwd.MaxLength has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		txtVerifyPwd.Maxlength = piMaxPwdLen
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSetPassword.Setup", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub txtNewPwd_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNewPwd.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 39 Then KeyAscii = 180
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtOldPwd_KeyPress(ByRef KeyAscii As Short)
		If KeyAscii = 39 Then KeyAscii = 180
	End Sub
	
	Private Sub txtVerifyPwd_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVerifyPwd.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 39 Then KeyAscii = 180
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class