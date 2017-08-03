Option Strict Off
Option Explicit On
Friend Class FUser
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOk_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOk.Click
		Dim rsUser As New ADODB.Recordset
		'
		If cmbUser.Text <> "" Then
			StrUser = cmbUser.Text
			rsUser.Open("select * from tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
			'
			With rsUser
				Do While Not .eof
					If LCase(StrUser) = LCase(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) Then
						iGroupNumber = .Fields("Groups").Value
					End If
					.MoveNext()
				Loop 
				.Close()
			End With
			'
			FVMail.Show()
			Me.Close()
		Else
			MsgBox("Please enter user Name")
		End If
		'
		
		
		
	End Sub
	
	Private Sub FUser_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim rsUser As New ADODB.Recordset
		Dim x As Integer
		'
		rsUser.Open("select * from tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		With rsUser
			Do While Not .eof
				cmbUser.Items.Add(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value)
				.MoveNext()
			Loop 
			.Close()
		End With
	End Sub
End Class