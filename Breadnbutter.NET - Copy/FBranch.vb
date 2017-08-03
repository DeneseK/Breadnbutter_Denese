Option Strict Off
Option Explicit On
Friend Class FBranch
	Inherits System.Windows.Forms.Form
	
	Private lResultID As Integer
	Private bAddNew As Boolean
	Private BranchData As New CBranchData
	Private Branch As New CBranch
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		'
		BranchData.Name = txtName.Text
		BranchData.Number = txtNumber.Text
		BranchData.ManagerFirstName = txtManagerFirstName.Text
		BranchData.ManagerLastName = txtManagerLastName.Text
		BranchData.Address1 = txtAddress1.Text
		BranchData.Address2 = txtAddress2.Text
		BranchData.Address3 = txtAddress3.Text
		BranchData.City = txtCity.Text
		BranchData.State = txtState.Text
		BranchData.Zip = txtZip.Text
		BranchData.PhoneNumber = txtPhoneNumber.Text
		BranchData.FaxNumber = txtFaxNumber.Text
		BranchData.Email = txtEmail.Text
		'
		If Branch.Save(BranchData, bAddNew) Then
			lResultID = BranchData.BranchID
		Else
			lResultID = 0
		End If
		'
		Me.Close()
	End Sub
	
	Public Function NewBranch(ByRef pCompanyID As Short) As Integer
		Branch = New CBranch
		BranchData = New CBranchData
		'
		bAddNew = True
		'
		BranchData.CompanyID = pCompanyID
		'
		Me.ShowDialog()
		'
		NewBranch = lResultID
		'
		'UPGRADE_NOTE: Object BranchData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		BranchData = Nothing
		'UPGRADE_NOTE: Object Branch may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Branch = Nothing
	End Function
	
	Public Function EditBranch(ByRef pBranchID As Integer) As Integer
		Branch = New CBranch
		BranchData = New CBranchData
		'
		bAddNew = False
		'
		Branch.Load(BranchData, pBranchID)
		'
		'lCompanyID = BranchData.CompanyID
		txtName.Text = BranchData.Name
		txtNumber.Text = BranchData.Number
		txtManagerFirstName.Text = BranchData.ManagerFirstName
		txtManagerLastName.Text = BranchData.ManagerLastName
		txtAddress1.Text = BranchData.Address1
		txtAddress2.Text = BranchData.Address2
		txtAddress3.Text = BranchData.Address3
		txtCity.Text = BranchData.City
		txtState.Text = BranchData.State
		txtZip.Text = BranchData.Zip
		txtPhoneNumber.Text = BranchData.PhoneNumber
		txtFaxNumber.Text = BranchData.FaxNumber
		txtEmail.Text = BranchData.Email
		'
		Me.ShowDialog()
		'
		EditBranch = lResultID
		'
		'UPGRADE_NOTE: Object BranchData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		BranchData = Nothing
		'UPGRADE_NOTE: Object Branch may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Branch = Nothing
	End Function
End Class