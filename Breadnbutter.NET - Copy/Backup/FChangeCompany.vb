Option Strict Off
Option Explicit On
Friend Class FChangeCompany
	Inherits System.Windows.Forms.Form
	
	Private ContactData As CContactData
	Private CompanyData As CCompanyData
	Private Company As CCompany
	Private Contact As CContact
	Private Companys As New CCompanys
	
	Private lContactID As Integer
	Private bChanged As Boolean
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		bChanged = False
		'
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		If Not lvwCompany.FocusedItem Is Nothing Then
			ContactData.CompanyID = GetIDFromKey(lvwCompany.FocusedItem.Name)
			'
			If Contact.Save(ContactData, False) Then
				bChanged = True
				'
				Me.Close()
			Else
				MsgBox("Contact could not be saved")
			End If
		Else
			MsgBox("Invalid Selection")
		End If
	End Sub
	
	Private Sub FChangeCompany_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		ContactData = New CContactData
		CompanyData = New CCompanyData
		Contact = New CContact
		Company = New CCompany
		Companys = New CCompanys
		'
		Dim lPos As Integer
		'
		Contact.Load(ContactData, lContactID)
		Company.Load(CompanyData, (ContactData.CompanyID))
		'
		Company.LoadCollection(Companys)
		'
		lvwCompany.Visible = False
		'
		lvwCompany.Items.Clear()
		'
		For lPos = 1 To Companys.Count
			lvwCompany.Items.Add("A" & Companys.Item(lPos).ID, Companys.Item(lPos).Name, "")
		Next 
		'
		lvwCompany.FocusedItem = lvwCompany.Items.Item("A" & CompanyData.ID)
		'
		'UPGRADE_WARNING: MSComctlLib.IListItem method lvwCompany.SelectedItem.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		lvwCompany.FocusedItem.EnsureVisible()
		'
		lvwCompany.Visible = True
		'
		
	End Sub
	
	Public Function ChangeCompany(ByRef plContactID As Integer) As Boolean
		lContactID = plContactID
		'
		Me.ShowDialog()
		'
		ChangeCompany = bChanged
	End Function
	
	Private Sub FChangeCompany_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object Companys may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Companys = Nothing
		'UPGRADE_NOTE: Object Contact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Contact = Nothing
		'UPGRADE_NOTE: Object Company may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Company = Nothing
		'UPGRADE_NOTE: Object CompanyData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		CompanyData = Nothing
		'UPGRADE_NOTE: Object ContactData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ContactData = Nothing
	End Sub
End Class