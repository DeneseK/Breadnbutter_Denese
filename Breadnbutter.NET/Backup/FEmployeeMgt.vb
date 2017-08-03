Option Strict Off
Option Explicit On
Friend Class FEmployeeMgt
	Inherits System.Windows.Forms.Form
	
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	Private iTempID As Short
	Dim bAddNew As Boolean
	
	Private Sub LoadEmployeeList()
		Dim Employees As New CEmployees
		Dim Employee As New CEmployee
		Dim iCounter As Short
		Dim sKey As String
		'
		Employee.LoadCollection(Employees)
		'
		ListView1.Items.Clear()
		'
		For iCounter = 1 To Employees.Count
			sKey = "A" & CStr(Employees.Item(iCounter).EmployeeID)
			ListView1.Items.Add(sKey, CStr(Employees.Item(iCounter).EmployeeID), "")
			'UPGRADE_ISSUE: MSComctlLib.ListSubItems method ListView1.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			ListView1.Items.Item(sKey).SubItems.Add( ,  , Employees.Item(iCounter).EmployeeFirst & " " & Employees.Item(iCounter).EmployeeLast,  , Employees.Item(iCounter).EmployeeFirst & " " & Employees.Item(iCounter).EmployeeLast).ForeColor = System.Drawing.Color.Black
			'UPGRADE_ISSUE: MSComctlLib.ListSubItems method ListView1.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			ListView1.Items.Item(sKey).SubItems.Add( ,  , CStr(Employees.Item(iCounter).EmployeeExt),  , CStr(Employees.Item(iCounter).EmployeeExt)).ForeColor = System.Drawing.Color.Black
			'UPGRADE_ISSUE: MSComctlLib.ListSubItems method ListView1.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			ListView1.Items.Item(sKey).SubItems.Add( ,  , CStr(Employees.Item(iCounter).SecurityLevel),  , CStr(Employees.Item(iCounter).SecurityLevel)).ForeColor = System.Drawing.Color.Black
			'UPGRADE_ISSUE: MSComctlLib.ListSubItems method ListView1.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			ListView1.Items.Item(sKey).SubItems.Add( ,  , CStr(Employees.Item(iCounter).Groups),  , CStr(Employees.Item(iCounter).Groups)).ForeColor = System.Drawing.Color.Black
			'UPGRADE_ISSUE: MSComctlLib.ListSubItems method ListView1.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			ListView1.Items.Item(sKey).SubItems.Add( ,  , CStr(Employees.Item(iCounter).WorkGroups),  , CStr(Employees.Item(iCounter).WorkGroups)).ForeColor = System.Drawing.Color.Black
			'UPGRADE_ISSUE: MSComctlLib.ListSubItems method ListView1.ListItems.Item.ListSubItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			ListView1.Items.Item(sKey).SubItems.Add( ,  , Employees.Item(iCounter).EMailAddress,  , Employees.Item(iCounter).EMailAddress).ForeColor = System.Drawing.Color.Black
		Next 
		'UPGRADE_NOTE: Object Employees may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employees = Nothing
		'UPGRADE_NOTE: Object Employee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employee = Nothing
	End Sub
	
	Private Sub LoadGroups(ByRef piTemp As Short)
		'
		If piTemp >= 8 Then
			chkAuthorizations.CheckState = System.Windows.Forms.CheckState.Checked
			piTemp = piTemp - 8
		Else
			chkAuthorizations.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If piTemp >= 4 Then
			chkSales.CheckState = System.Windows.Forms.CheckState.Checked
			piTemp = piTemp - 4
		Else
			chkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If piTemp >= 2 Then
			chkSupport.CheckState = System.Windows.Forms.CheckState.Checked
			piTemp = piTemp - 2
		Else
			chkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If piTemp >= 1 Then
			chkOperator.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkOperator.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
	End Sub
	
	Private Function SaveGroup() As Short
		Dim iGroupNumber As Short
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
		SaveGroup = iGroupNumber
	End Function
	
	Private Sub LoadWorkGroups(ByRef piTemp As Short)
		'
		If piTemp >= 8 Then
			chkManagement.CheckState = System.Windows.Forms.CheckState.Checked
			piTemp = piTemp - 8
		Else
			chkManagement.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If piTemp >= 4 Then
			chkWorkSales.CheckState = System.Windows.Forms.CheckState.Checked
			piTemp = piTemp - 4
		Else
			chkWorkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If piTemp >= 2 Then
			chkWorkSupport.CheckState = System.Windows.Forms.CheckState.Checked
			piTemp = piTemp - 2
		Else
			chkWorkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
		If piTemp >= 1 Then
			chkDev.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkDev.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		'
	End Sub
	
	Private Function SaveWorkGroup() As Short
		Dim iGroupNumber As Short
		'
		iGroupNumber = 0
		'
		If chkDev.CheckState = 1 Then
			iGroupNumber = iGroupNumber + 1
		End If
		'
		If chkWorkSupport.CheckState = 1 Then
			iGroupNumber = iGroupNumber + 2
		End If
		'
		If chkWorkSales.CheckState = 1 Then
			iGroupNumber = iGroupNumber + 4
		End If
		'
		If chkManagement.CheckState = 1 Then
			iGroupNumber = iGroupNumber + 8
		End If
		SaveWorkGroup = iGroupNumber
	End Function
	
	
	
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		bAddNew = True
		EnableEdit()
		cmdDelete.Enabled = False
		txtFirst.Focus()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		DisableEdit()
		bAddNew = False
		cmdDelete.Enabled = True
	End Sub
	
	Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
		Dim Employee As New CEmployee
		bAddNew = False
		If MsgBox("Confirm Delete", MsgBoxStyle.YesNo, "Delete Employee") = MsgBoxResult.Yes Then
			Employee.Delete(CShort(ListView1.FocusedItem.Text))
		End If
		'UPGRADE_NOTE: Object Employee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employee = Nothing
		LoadEmployeeList()
	End Sub
	
	Private Sub cmdEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEdit.Click
		Dim Employee As New CEmployee
		Dim EmployeeData As New CEmployeeData
		'
		Employee.Load(EmployeeData, CShort(ListView1.FocusedItem.Text))
		'
		bAddNew = False
		EnableEdit()
		'
		iTempID = EmployeeData.EmployeeID
		txtFirst.Text = EmployeeData.EmployeeFirst
		txtMid.Text = EmployeeData.EmployeeMiddle
		txtLast.Text = EmployeeData.EmployeeLast
		txtExt.Text = CStr(EmployeeData.EmployeeExt)
		txtPassword.Text = DecryptStr((EmployeeData.Password))
		LoadGroups((EmployeeData.Groups))
		txtMail.Text = EmployeeData.EMailAddress
		LoadWorkGroups((EmployeeData.WorkGroups))
		If EmployeeData.SecurityLevel = 1 Then
			optLow.Checked = True
		Else
			optHigh.Checked = True
		End If
		'
		'UPGRADE_NOTE: Object EmployeeData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EmployeeData = Nothing
		'UPGRADE_NOTE: Object Employee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employee = Nothing
		cmdDelete.Enabled = False
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		Dim Employee As New CEmployee
		Dim EmployeeData As New CEmployeeData
		'
		EmployeeData.EmployeeID = iTempID
		EmployeeData.EmployeeFirst = txtFirst.Text
		EmployeeData.EmployeeMiddle = txtMid.Text
		EmployeeData.EmployeeLast = txtLast.Text
		EmployeeData.EmployeeExt = Val(txtExt.Text)
		EmployeeData.Groups = SaveGroup
		EmployeeData.WorkGroups = SaveWorkGroup
		EmployeeData.Password = txtPassword.Text
		EmployeeData.EMailAddress = txtMail.Text
		
		
		If optLow.Checked = True Then
			EmployeeData.SecurityLevel = 1
		Else
			EmployeeData.SecurityLevel = 2
		End If
		'
		If bAddNew = False Then
			Employee.Save(EmployeeData, iTempID)
		Else
			Employee.AddNew(EmployeeData)
		End If
		'
		'UPGRADE_NOTE: Object EmployeeData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		EmployeeData = Nothing
		'UPGRADE_NOTE: Object Employee may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Employee = Nothing
		'
		LoadEmployeeList()
		DisableEdit()
		bAddNew = False
		cmdDelete.Enabled = True
	End Sub
	
	Private Sub FEmployeeMgt_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.Setup(Me, False,  ,  , "Employee Management")
		'
		LoadEmployeeList()
		bAddNew = False
		DisableEdit()
	End Sub
	
	'UPGRADE_WARNING: Event FEmployeeMgt.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FEmployeeMgt_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error GoTo EH
		'
		Frame1.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(Frame1.Width)) / 2), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FEmployeeMgt.Form_Resize.")
	End Sub
	
	Private Sub DisableEdit()
		txtFirst.Text = ""
		txtMid.Text = ""
		txtLast.Text = ""
		txtExt.Text = ""
		txtPassword.Text = ""
		txtMail.Text = ""
		optLow.Checked = True
		chkManagement.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkWorkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkWorkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkDev.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkAuthorizations.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkOperator.CheckState = System.Windows.Forms.CheckState.Unchecked
		'
		txtFirst.Enabled = False
		txtMid.Enabled = False
		txtLast.Enabled = False
		txtExt.Enabled = False
		txtMail.Enabled = False
		'
		txtPassword.Visible = False
		lblPass.Visible = False
		'
		optLow.Enabled = False
		optHigh.Enabled = False
		chkManagement.Enabled = False
		chkWorkSales.Enabled = False
		chkWorkSupport.Enabled = False
		chkDev.Enabled = False
		chkAuthorizations.Enabled = False
		chkSales.Enabled = False
		chkSupport.Enabled = False
		chkOperator.Enabled = False
		cmdSave.Enabled = False
		cmdCancel.Enabled = False
	End Sub
	
	Private Sub EnableEdit()
		txtFirst.Text = ""
		txtMid.Text = ""
		txtLast.Text = ""
		txtExt.Text = ""
		txtPassword.Text = ""
		txtMail.Text = ""
		optLow.Checked = True
		chkManagement.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkWorkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkWorkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkDev.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkAuthorizations.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkSupport.CheckState = System.Windows.Forms.CheckState.Unchecked
		chkOperator.CheckState = System.Windows.Forms.CheckState.Unchecked
		'
		txtFirst.Enabled = True
		txtMid.Enabled = True
		txtLast.Enabled = True
		txtExt.Enabled = True
		txtMail.Enabled = True
		optLow.Enabled = True
		optHigh.Enabled = True
		chkManagement.Enabled = True
		chkWorkSales.Enabled = True
		chkWorkSupport.Enabled = True
		chkDev.Enabled = True
		chkAuthorizations.Enabled = True
		chkSales.Enabled = True
		chkSupport.Enabled = True
		chkOperator.Enabled = True
		cmdSave.Enabled = True
		cmdCancel.Enabled = True
		'
		'If bAddNew = True Then
		txtPassword.Visible = True
		lblPass.Visible = True
		'End If
		'
	End Sub
	
	Private Sub ListView1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ListView1.Click
		cmdCancel_Click(cmdCancel, New System.EventArgs())
	End Sub
	
	Private Sub txtIcon_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtIcon.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If Shift = 7 And KeyCode = 73 Then
			If CDbl(GetSetting(My.Application.Info.Title, "Settings", "Icon", CStr(0))) = 1 Then
				SaveSetting(My.Application.Info.Title, "Settings", "Icon", CStr(0))
				MsgBox("Icon Enabled")
				FMain.GetUserGroups()
			Else
				SaveSetting(My.Application.Info.Title, "Settings", "Icon", CStr(1))
				MsgBox("Icon Disabled")
			End If
		End If
		Shift = 0
		KeyCode = 0
	End Sub
End Class