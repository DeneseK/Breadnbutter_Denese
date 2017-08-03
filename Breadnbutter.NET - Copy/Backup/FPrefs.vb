Option Strict Off
Option Explicit On
Friend Class FPrefs
	Inherits System.Windows.Forms.Form
	
	Private Sub cboAuthStatus_InitColumnProps(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboAuthStatus.InitColumnProps
		On Error GoTo ErrCall
		'
		Dim rs As ADODB.Recordset
		'
		'Set rs = dbMain.OpenRecordset("SELECT * FROM tblAuthStatus ORDER BY RecID", dbOpenForwardOnly)
		rs = New ADODB.Recordset
		rs.Open("SELECT * FROM tblAuthStatus ORDER BY RecID", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		Do While Not rs.EOF
			cboAuthStatus.AddItem(rs.Fields("Status").Value)
			rs.MoveNext()
		Loop 
		'
		DBOps.ZapRS(rs)
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmPrefs.cboAuthStatus_InitColumnProps.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub cboStatus_InitColumnProps(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboStatus.InitColumnProps
		On Error GoTo ErrCall
		'
		Dim rsStatus As ADODB.Recordset
		'
		'Set rsStatus = dbMain.OpenRecordset("SELECT * FROM tblStatus", dbOpenForwardOnly)
		rsStatus = New ADODB.Recordset
		rsStatus.Open("SELECT * FROM tblStatus", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		Do While Not rsStatus.EOF
			cboStatus.AddItem("" & rsStatus.Fields("Status").Value)
			rsStatus.MoveNext()
		Loop 
		'
		DBOps.ZapRS(rsStatus)
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmPrefs.cboStatus_InitColumnProps.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		SaveSetting(My.Application.Info.Title, "Preferences", "InitStatus", cboStatus.Text)
		SaveSetting(My.Application.Info.Title, "Preferences", "InitShipStatus", cboShipStatus.Text)
		SaveSetting(My.Application.Info.Title, "Preferences", "InitAuthStatus", cboAuthStatus.Text)
		Me.Close()
	End Sub
	
	Private Sub FPrefs_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		cboStatus.Text = GetSetting(My.Application.Info.Title, "Preferences", "InitStatus", "Prospect")
		cboShipStatus.Text = GetSetting(My.Application.Info.Title, "Preferences", "InitShipStatus", "Not Shipped")
		cboAuthStatus.Text = GetSetting(My.Application.Info.Title, "Preferences", "InitAuthStatus", "Not Authorized")
	End Sub
	
	Private Sub cboShipStatus_InitColumnProps(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboShipStatus.InitColumnProps
		On Error GoTo ErrCall
		'
		Dim rsStatus As ADODB.Recordset
		'
		'Set rsStatus = dbMain.OpenRecordset("SELECT * FROM tblShipStatus", dbOpenForwardOnly)
		rsStatus = New ADODB.Recordset
		rsStatus.Open("SELECT * FROM tblShipStatus", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
		'
		Do While Not rsStatus.EOF
			cboShipStatus.AddItem("" & rsStatus.Fields("Status").Value)
			rsStatus.MoveNext()
		Loop 
		'
		DBOps.ZapRS(rsStatus)
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmPrefs.SSDBCombo1_InitColumnProps.", MsgBoxStyle.Critical, "Error")
	End Sub
End Class