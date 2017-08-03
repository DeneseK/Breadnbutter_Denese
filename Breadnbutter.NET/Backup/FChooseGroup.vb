Option Strict Off
Option Explicit On
Friend Class FChooseGroup
	Inherits System.Windows.Forms.Form
	Private lListID As Integer
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		lListID = 0
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		If lstCustGroups.SelectedIndex <> -1 Then
			lListID = VB6.GetItemData(lstCustGroups, lstCustGroups.SelectedIndex)
			Me.Close()
		End If
	End Sub
	
	Private Sub FChooseGroup_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		LoadCustGroupsList()
	End Sub
	
	Private Sub LoadCustGroupsList()
		Dim GroupList As CGroupList
		Dim GroupLists As CGroupListDatas
		Dim GroupListLink As CGroupListLink
		Dim X As Short
		Dim sKey As String
		'
		GroupList = New CGroupList
		GroupLists = New CGroupListDatas
		GroupListLink = New CGroupListLink
		'
		GroupList.LoadCollection(GroupLists)
		'
		'lstCustGroups.Clear
		'
		For X = 1 To GroupLists.Count
			sKey = "A" & GroupLists.Item(X).ID
			lstCustGroups.Items.Add(GroupLists.Item(X).ListName)
			VB6.SetItemData(lstCustGroups, X - 1, GroupLists.Item(X).ID)
		Next 
		'
		'UPGRADE_NOTE: Object GroupList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		GroupList = Nothing
		'UPGRADE_NOTE: Object GroupLists may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		GroupLists = Nothing
		'UPGRADE_NOTE: Object GroupListLink may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		GroupListLink = Nothing
	End Sub
	
	Public Function GetGroup() As Integer
		Me.ShowDialog()
		GetGroup = lListID
	End Function
End Class