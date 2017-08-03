Option Strict Off
Option Explicit On
Friend Class FCompany2
	Inherits System.Windows.Forms.Form
	
	Dim lResultID As Integer
	Dim bNew As Boolean
	Dim Company As CCompany
	Dim CompanyData As CCompanyData
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Me.Close()
		'<EhFooter>
		'
		Exit Sub
		'
EH: 
		ErrorMgr.Raise("FCompany2", "cmdCancel_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: CompanyData.Name = Me.txtName.Text
110: CompanyData.Note = Me.txtNote.Text
115: CompanyData.InterestRank = GetRank
		'
120: CompanyData.DoNotContact = IIf((Me.chkDoNotContact.CheckState = 1), True, False)
		'
130: If Company.Save(CompanyData, bNew) Then
140: lResultID = CompanyData.ID
150: Else
160: lResultID = 0
170: End If
		'
180: Me.Close()
		'<EhFooter>
		'
		Exit Sub
		'
EH: 
		ErrorMgr.Raise("FCompany2", "cmdSave_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Public Function NewCompany() As Integer
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: CompanyData = New CCompanyData
110: Company = New CCompany
		'
120: bNew = True
		'
130: Me.Text = "Add New Company"
135: SetStar(0)
140: Me.ShowDialog()
		'
150: NewCompany = lResultID
		'
160: 'UPGRADE_NOTE: Object CompanyData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		CompanyData = Nothing
170: 'UPGRADE_NOTE: Object Company may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Company = Nothing
		'<EhFooter>
		'
		Exit Function
		'
EH: 
		ErrorMgr.Raise("FCompany2", "NewCompany", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Function
	
	Public Function LoadCompany(ByRef plCompanyID As Integer) As Integer
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: CompanyData = New CCompanyData
110: Company = New CCompany
		'
120: bNew = False
		'
130: Company.Load(CompanyData, plCompanyID)
		'
140: Me.txtName.Text = CompanyData.Name
150: Me.txtNote.Text = CompanyData.Note
160: Me.chkDoNotContact.CheckState = IIf(CompanyData.DoNotContact, 1, 0)
165: SetStar((CompanyData.InterestRank))
		'
170: Me.ShowDialog()
		'
180: LoadCompany = lResultID
		'
190: 'UPGRADE_NOTE: Object CompanyData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		CompanyData = Nothing
200: 'UPGRADE_NOTE: Object Company may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Company = Nothing
		'<EhFooter>
		'
		Exit Function
		'
EH: 
		ErrorMgr.Raise("FCompany2", "LoadCompany", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Function
	
	Private Sub SetStar(ByRef Index As Short)
		Dim x As Short
		For x = 0 To 5
			'UPGRADE_WARNING: Lower bound of collection Me.ImageList2.ListImages has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			StarPicture(x).Image = Me.ImageList2.Images.Item(1)
		Next x
		'
		For x = 1 To Index
			'UPGRADE_WARNING: Lower bound of collection Me.ImageList2.ListImages has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			StarPicture(x).Image = Me.ImageList2.Images.Item(2)
		Next x
	End Sub
	
	Private Function GetRank() As Short
		Dim x As Short
		GetRank = 0
		For x = 0 To 5
			'UPGRADE_WARNING: Lower bound of collection Me.ImageList2.ListImages has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If StarPicture(x).Image.equals(Me.ImageList2.Images.Item(2)) Then GetRank = x
		Next x
		'
	End Function
	
	Private Sub StarPicture_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles StarPicture.Click
		Dim Index As Short = StarPicture.GetIndex(eventSender)
		SetStar(Index)
	End Sub
End Class