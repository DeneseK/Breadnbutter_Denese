Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FSendTo
	Inherits System.Windows.Forms.Form
	Private boolAttachment As Boolean
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Hide()
	End Sub
	
	Private Sub cmdSend_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSend.Click
		SendMail()
		cboSendTo.Text = vbNullString
		lblEmailAddress.Text = vbNullString
		Me.Hide()
	End Sub
	
	Private Sub FSendTo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim rs As New ADODB.Recordset
		Me.Text = StrUser & " Forward Message To:"
		rs.Open("SELECT * FROM tblEmployees ORDER BY EmployeeLast", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rs
			While Not .eof
				cboSendTo.Items.Add((.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value))
				If LCase(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) = LCase(StrUser) Then
					'          sFromAddress = !EMailAddress & "@powerclaim.com"
				End If
				.MoveNext()
			End While
		End With
		'
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
		'
	End Sub
	
	'UPGRADE_WARNING: Event cboSendTo.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboSendTo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSendTo.SelectedIndexChanged
		Dim rs As New ADODB.Recordset
		'
		rs.Open("SELECT * FROM tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rs
			.MoveFirst()
			While Not .eof
				If (.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) = cboSendTo.Text Then
					sEmailAddress = .Fields("EMailAddress").Value '& "@powerclaim.com"
					lblEmailAddress.Text = sEmailAddress
					Exit Sub
				Else
					sEmailAddress = cboSendTo.Text
				End If
				.MoveNext()
			End While
		End With
		lblEmailAddress.Text = sEmailAddress
		'
	End Sub
	
	Private Sub cboSendTo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSendTo.Leave
		Dim rs As New ADODB.Recordset
		'
		rs.Open("SELECT * FROM tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rs
			.MoveFirst()
			While Not .eof
				If (.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) = cboSendTo.Text Then
					sEmailAddress = .Fields("EMailAddress").Value '& "@powerclaim.com"
					lblEmailAddress.Text = sEmailAddress
					Exit Sub
				Else
					sEmailAddress = cboSendTo.Text
				End If
				.MoveNext()
			End While
		End With
		lblEmailAddress.Text = sEmailAddress
		'
	End Sub
	Private Sub getAttachment()
		Dim strStream As Object
		Dim FileSys As New Scripting.FileSystemObject
		Dim myStr As String
		Dim rsAttach As New ADODB.Recordset
		Dim sMessageName As String
		
		'
		If Not FileSys.FolderExists(My.Application.Info.DirectoryPath & "\Attachment") Then
			FileSys.CreateFolder(My.Application.Info.DirectoryPath & "\Attachment")
		End If
		'rsAttach.Open "SELECT * FROM TVMailMessages WHERE MessageName like '" & sMessageName & "'", cnMain, adOpenKeyset, adLockBatchOptimistic
		rsAttach.Open("SELECT * FROM TVMailMessages WHERE MessageID like '" & CStr(sMessageID) & "'", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		
		'
		With rsAttach
			sMessageName = "" & .Fields("MessageName").Value
			If sMessageName <> "" And UCase(VB.Right(sMessageName, 4)) = ".WAV" Then
				
				' If !Attachment <> Null Then
				strStream = New ADODB.Stream
				'UPGRADE_WARNING: Couldn't resolve default property of object strStream.Type. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strStream.Type = ADODB.StreamTypeEnum.adTypeBinary
				'UPGRADE_WARNING: Couldn't resolve default property of object strStream.Open. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strStream.Open()
				'UPGRADE_WARNING: Couldn't resolve default property of object strStream.Write. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strStream.Write(.Fields("Attachment"))
				'UPGRADE_WARNING: Couldn't resolve default property of object strStream.SaveToFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strStream.SaveToFile(My.Application.Info.DirectoryPath & "\Attachment\" & .Fields("MessageName").Value, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
				'UPGRADE_WARNING: Couldn't resolve default property of object strStream.Close. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strStream.Close()
				'UPGRADE_NOTE: Object strStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				strStream = Nothing
				boolAttachment = True
			Else
				'     Debug.Print !MessageName
				boolAttachment = False
			End If
			
		End With
		
		
		'
		'getAttachment = boolAttachment
	End Sub
	
	Private Sub SendMail()
		Dim SMTP As Object
		Dim FileSys As New Scripting.FileSystemObject
		Dim X As Short
		
		SMTP = CreateObject("EasyMail.SMTP.6")
		'UPGRADE_WARNING: Couldn't resolve default property of object SMTP.LicenseKey. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SMTP.LicenseKey = "Hawkins Research (Single Developer)/00B0630C10151C00BC30"
		'UPGRADE_WARNING: Couldn't resolve default property of object SMTP.MailServer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SMTP.MailServer = "HRI-svr-02"
		'UPGRADE_WARNING: Couldn't resolve default property of object SMTP.FromAddr. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sFromAddress. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SMTP.FromAddr = sFromAddress
		'UPGRADE_WARNING: Couldn't resolve default property of object SMTP.AddRecipient. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SMTP.AddRecipient("", sEmailAddress, 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object SMTP.Subject. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SMTP.Subject = sSubject
		'UPGRADE_WARNING: Couldn't resolve default property of object SMTP.BodyText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object sBody. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SMTP.BodyText = sBody
		getAttachment()
		If boolAttachment Then
			'UPGRADE_WARNING: Couldn't resolve default property of object SMTP.AddAttachment. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			X = SMTP.AddAttachment(My.Application.Info.DirectoryPath & "\Attachment\" & sMessageName, 0)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object SMTP.Send. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		X = SMTP.Send
		If X = 0 Then
			MsgBox("Message sent successfully.")
		Else
			MsgBox("There was an error sending your message.  Error: " & CStr(X))
		End If
		
		If FileSys.FolderExists(My.Application.Info.DirectoryPath & "\Attachment") Then
			FileSys.DeleteFolder((My.Application.Info.DirectoryPath & "\Attachment"))
		End If
		'UPGRADE_NOTE: Object SMTP may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		SMTP = Nothing
	End Sub
	
	
	Private Function sBody() As Object
		Dim TSTemp As Scripting.TextStream
		Dim fso As New Scripting.FileSystemObject
		'
		TSTemp = fso.OpenTextFile(My.Application.Info.DirectoryPath & "\tempMessage.txt", Scripting.IOMode.ForReading, True, Scripting.Tristate.TristateUseDefault)
		'UPGRADE_WARNING: Couldn't resolve default property of object sBody. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sBody = TSTemp.ReadAll
		'
		
	End Function
	
	Private Function sFromAddress() As Object
		Dim rsUser As New ADODB.Recordset
		'
		rsUser.Open("select * from tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		
		With rsUser
			.MoveFirst()
			While Not .eof
				If LCase(StrUser) = LCase(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object sFromAddress. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sFromAddress = .Fields("EMailAddress").Value '& "@powerclaim.com"
				End If
				.MoveNext()
			End While
		End With
		'
		
	End Function
End Class