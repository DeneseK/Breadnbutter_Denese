Option Strict Off
Option Explicit On
Friend Class FSelectDB
	Inherits System.Windows.Forms.Form
	
	
	Public Cancelled As Boolean
	
	Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
		Cancelled = True
		Me.Hide()
	End Sub
	
	Private Sub cmdSelectDB_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSelectDB.Click
		
		Dim sDBPath As String
		Dim sDBName As String
		'
		DBOps.GetPathFile(sDBPath, sDBName, "Bread 'n' Butter Data")
		'
		Me.txtDatabase.Text = sDBPath & sDBName
		
	End Sub
	
	Private Sub FSelectDB_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim ConnectionType As MMain.ConnectionTypeEnum
		'
		ConnectionType = CShort(GetSetting(My.Application.Info.Title, "Database", "Type", CStr(MMain.ConnectionTypeEnum.SQL)))
		'
		If ConnectionType = MMain.ConnectionTypeEnum.SQL Then
			Me.optDBType(0).Checked = True
		Else
			Me.optDBType(1).Checked = True
		End If
		'
		Me.cboServer.Text = GetSetting(My.Application.Info.Title, "Database", "Server", "HAWKINS-MAIN")
		Me.cboDatabase.Text = GetSetting(My.Application.Info.Title, "Database", "SQLDB", "BNB_DATA")
		Me.txtDatabase.Text = GetSetting(My.Application.Info.Title, "Database", "AccessDB", vbNullString)
		Me.chkLogin.CheckState = CShort(GetSetting(My.Application.Info.Title, "Database", "Login", "0"))
		Me.txtPassword.Text = GetSetting(My.Application.Info.Title, "Database", "Password", "")
		GetLoginSettings()
		'
	End Sub
	
	Private Sub FSelectDB_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		If UnloadMode = System.Windows.Forms.CloseReason.UserClosing Then
			Me.Hide()
			Cancel = True
		End If
		
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub OKButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKButton.Click
		
		sLogin = CStr(chkLogin.CheckState)
		sLoginName = txtUserID.Text
		sPassword = txtPassword.Text
		
		If optDBType(0).Checked = True Then
			MMain.ConnType = MMain.ConnectionTypeEnum.SQL
		Else
			MMain.ConnType = MMain.ConnectionTypeEnum.Access
		End If
		'
		SaveSetting(My.Application.Info.Title, "Database", "Type", CStr(MMain.ConnType))
		SaveSetting(My.Application.Info.Title, "Database", "Server", Me.cboServer.Text)
		SaveSetting(My.Application.Info.Title, "Database", "SQLDB", Me.cboDatabase.Text)
		SaveSetting(My.Application.Info.Title, "Database", "AccessDB", Me.txtDatabase.Text)
		SaveSetting(My.Application.Info.Title, "Database", "Login", CStr(Me.chkLogin.CheckState))
		SaveSetting(My.Application.Info.Title, "Database", "Password", Me.txtPassword.Text)
		'
		Me.Hide()
		
	End Sub
	
	'UPGRADE_WARNING: Event optDBType.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optDBType_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDBType.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optDBType.GetIndex(eventSender)
			
			If Me.optDBType(0).Checked = True Then
				Me.cboServer.Enabled = True
				Me.cboDatabase.Enabled = True
				Me.txtDatabase.Enabled = False
			Else
				Me.cboServer.Enabled = False
				Me.cboDatabase.Enabled = False
				Me.txtDatabase.Enabled = True
			End If
			
			
		End If
	End Sub
	
	Private Sub GetLoginSettings()
		
		If chkLogin.CheckState = 1 Then
			txtUserID.Text = UCase(GetSetting(My.Application.Info.Title, "Settings", "User", ""))
			EditTextBox()
			txtUserID.Visible = True
			txtPassword.Visible = True
			lblUserID.Visible = True
			lblPassword.Visible = True
			txtDatabase.SetBounds(VB6.TwipsToPixelsX(1470), VB6.TwipsToPixelsY(3120), VB6.TwipsToPixelsX(2775), VB6.TwipsToPixelsY(315))
			Label3.SetBounds(VB6.TwipsToPixelsX(660), VB6.TwipsToPixelsY(3150), VB6.TwipsToPixelsX(795), VB6.TwipsToPixelsY(285))
			optDBType(1).SetBounds(VB6.TwipsToPixelsX(300), VB6.TwipsToPixelsY(2640), VB6.TwipsToPixelsX(3555), VB6.TwipsToPixelsY(315))
			cmdSelectDB.SetBounds(VB6.TwipsToPixelsX(4290), VB6.TwipsToPixelsY(3150), VB6.TwipsToPixelsX(315), VB6.TwipsToPixelsY(315))
			Me.SetBounds(VB6.TwipsToPixelsX(2715), VB6.TwipsToPixelsY(3420), VB6.TwipsToPixelsX(6345), VB6.TwipsToPixelsY(4080))
		Else
			txtUserID.Visible = False
			txtPassword.Visible = False
			lblUserID.Visible = False
			lblPassword.Visible = False
			txtDatabase.SetBounds(VB6.TwipsToPixelsX(1470), VB6.TwipsToPixelsY(2160), VB6.TwipsToPixelsX(2775), VB6.TwipsToPixelsY(315))
			Label3.SetBounds(VB6.TwipsToPixelsX(660), VB6.TwipsToPixelsY(2190), VB6.TwipsToPixelsX(795), VB6.TwipsToPixelsY(285))
			optDBType(1).SetBounds(VB6.TwipsToPixelsX(300), VB6.TwipsToPixelsY(1800), VB6.TwipsToPixelsX(3555), VB6.TwipsToPixelsY(315))
			cmdSelectDB.SetBounds(VB6.TwipsToPixelsX(4290), VB6.TwipsToPixelsY(2190), VB6.TwipsToPixelsX(315), VB6.TwipsToPixelsY(315))
			Me.SetBounds(VB6.TwipsToPixelsX(2715), VB6.TwipsToPixelsY(3420), VB6.TwipsToPixelsX(6345), VB6.TwipsToPixelsY(3030))
		End If
		'
	End Sub
	
	'UPGRADE_WARNING: Event chkLogin.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkLogin_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLogin.CheckStateChanged
		If chkLogin.CheckState = 1 Then
			txtUserID.Text = UCase(GetSetting(My.Application.Info.Title, "Settings", "User", ""))
			EditTextBox()
			txtUserID.Visible = True
			txtPassword.Visible = True
			lblUserID.Visible = True
			lblPassword.Visible = True
			txtDatabase.SetBounds(VB6.TwipsToPixelsX(1470), VB6.TwipsToPixelsY(3120), VB6.TwipsToPixelsX(2775), VB6.TwipsToPixelsY(315))
			Label3.SetBounds(VB6.TwipsToPixelsX(660), VB6.TwipsToPixelsY(3150), VB6.TwipsToPixelsX(795), VB6.TwipsToPixelsY(285))
			optDBType(1).SetBounds(VB6.TwipsToPixelsX(300), VB6.TwipsToPixelsY(2640), VB6.TwipsToPixelsX(3555), VB6.TwipsToPixelsY(315))
			cmdSelectDB.SetBounds(VB6.TwipsToPixelsX(4290), VB6.TwipsToPixelsY(3150), VB6.TwipsToPixelsX(315), VB6.TwipsToPixelsY(315))
			Me.SetBounds(VB6.TwipsToPixelsX(2715), VB6.TwipsToPixelsY(3420), VB6.TwipsToPixelsX(6345), VB6.TwipsToPixelsY(4080))
		Else
			txtUserID.Visible = False
			txtPassword.Visible = False
			lblUserID.Visible = False
			lblPassword.Visible = False
			txtDatabase.SetBounds(VB6.TwipsToPixelsX(1470), VB6.TwipsToPixelsY(2160), VB6.TwipsToPixelsX(2775), VB6.TwipsToPixelsY(315))
			Label3.SetBounds(VB6.TwipsToPixelsX(660), VB6.TwipsToPixelsY(2190), VB6.TwipsToPixelsX(795), VB6.TwipsToPixelsY(285))
			optDBType(1).SetBounds(VB6.TwipsToPixelsX(300), VB6.TwipsToPixelsY(1800), VB6.TwipsToPixelsX(3555), VB6.TwipsToPixelsY(315))
			cmdSelectDB.SetBounds(VB6.TwipsToPixelsX(4290), VB6.TwipsToPixelsY(2190), VB6.TwipsToPixelsX(315), VB6.TwipsToPixelsY(315))
			Me.SetBounds(VB6.TwipsToPixelsX(2715), VB6.TwipsToPixelsY(3420), VB6.TwipsToPixelsX(6345), VB6.TwipsToPixelsY(3030))
		End If
	End Sub
	Private Sub EditTextBox()
		Dim i As Short
		Dim sLetter As String
		Dim sName As String
		'
		For i = 1 To Len(txtUserID.Text)
			sLetter = Mid(txtUserID.Text, i, 1)
			If Not sLetter = " " Then
				sName = sName & sLetter
			End If
		Next i
		txtUserID.Text = UCase(sName)
		'
	End Sub
End Class