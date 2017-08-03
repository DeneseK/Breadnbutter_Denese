Option Strict Off
Option Explicit On
Friend Class FUserGroups
	Inherits System.Windows.Forms.Form
	'Dim iNum As Integer
	
	'UPGRADE_WARNING: Event chkOperator.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkOperator_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOperator.CheckStateChanged
		'
		
	End Sub
	
	'UPGRADE_WARNING: Event chkAuthorizations.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkAuthorizations_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAuthorizations.CheckStateChanged
		'
		
	End Sub
	
	'UPGRADE_WARNING: Event chkSales.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkSales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSales.CheckStateChanged
		'
		
	End Sub
	
	'UPGRADE_WARNING: Event chkSupport.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkSupport_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSupport.CheckStateChanged
		'
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		FUserGroups_Load(Me, New System.EventArgs())
		Me.Hide()
		FVMail.GetUserGroups()
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		'
		iGroupNumber = 0
		'
		If chkOperator.CheckState = 1 Then
			iGroupNumber = iGroupNumber + 1
		End If
		'
		If chkSupport.CheckState = 1 Then
			iGroupNumber = iGroupNumber + 2
		End If
		'
		If chkSales.CheckState = 1 Then
			iGroupNumber = iGroupNumber + 4
		End If
		'
		If chkAuthorizations.CheckState = 1 Then
			iGroupNumber = iGroupNumber + 8
		End If
		'
		If chkAuthorizations.CheckState = 0 And chkSales.CheckState = 0 And chkSupport.CheckState = 0 And chkOperator.CheckState = 0 Then
			MsgBox("You must check at least 1 Group.", MsgBoxStyle.Information)
		Else
			SaveData()
			FUserGroups_Load(Me, New System.EventArgs())
			Me.Hide()
			FVMail.RefreshMessages()
			FVMail.GetUserGroups()
			FVMail.listview1_Click(Nothing, New System.EventArgs())
		End If
	End Sub
	
	Private Sub FUserGroups_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim iTemp As Short
		'
		Me.Text = "User Groups For: " & StrUser
		'
		iTemp = iGroupNumber
		If iTemp >= 8 Then
			chkAuthorizations.CheckState = System.Windows.Forms.CheckState.Checked
			iTemp = iTemp - 8
		Else
			chkAuthorizations.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If iTemp >= 4 Then
			chkSales.CheckState = System.Windows.Forms.CheckState.Checked
			iTemp = iTemp - 4
		Else
			chkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If iTemp >= 2 Then
			chkSupport.CheckState = System.Windows.Forms.CheckState.Checked
			iTemp = iTemp - 2
		Else
			chkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If iTemp >= 1 Then
			chkOperator.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkOperator.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
	End Sub
	
	Private Sub SaveData()
		Dim rsUser As New ADODB.Recordset
		'
		rsUser.Open("select * from tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		With rsUser
			Do While Not .eof
				If LCase(StrUser) = LCase(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) Then
					.Fields("Groups").Value = iGroupNumber
				End If
				.UpdateBatch()
				.MoveNext()
			Loop 
			.Close()
		End With
		'StrGroups = iGroupNumber
	End Sub
End Class