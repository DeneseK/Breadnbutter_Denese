Option Strict Off
Option Explicit On
Friend Class FProspectMgt
	Inherits System.Windows.Forms.Form
	
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	
	Private fShowTotals As Boolean
	Private fSettingPrefs As Boolean
	
	Private lProspectGroupID As Integer
	
	Private rsGroupCategories As ADODB.Recordset
	Private rsProspectGroup As ADODB.Recordset
	Private rsHistory As ADODB.Recordset
	
	Private Enum eFilter
		FilterNone
		FilterStandard
		FilterAMBest
		FilterAll
	End Enum
	
	Private Enum eSortColumn
		SortByGroup
		SortByLabel
	End Enum
	
	'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Filter_Renamed As eFilter
	Private SortColumn As eSortColumn
	
	'UPGRADE_WARNING: Event chkFilter.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkFilter_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkFilter.CheckStateChanged
		Dim Index As Short = chkFilter.GetIndex(eventSender)
		'*
		'
		On Error GoTo ErrHndlr
		'
		If fSettingPrefs = True Then Exit Sub
		'
		If chkFilter(0).CheckState = System.Windows.Forms.CheckState.Checked And chkFilter(1).CheckState = System.Windows.Forms.CheckState.Checked Then
			Filter_Renamed = eFilter.FilterAll
		ElseIf chkFilter(0).CheckState = System.Windows.Forms.CheckState.Checked Then 
			Filter_Renamed = eFilter.FilterStandard
		ElseIf chkFilter(1).CheckState = System.Windows.Forms.CheckState.Checked Then 
			Filter_Renamed = eFilter.FilterAMBest
		Else
			Filter_Renamed = eFilter.FilterNone
		End If
		'
		SaveSetting(My.Application.Info.Title, "ProspectMgt", "FilterGroups", CStr(Filter_Renamed))
		'
		SetupGroups()
		SelectGroup()
		'
		Exit Sub
		'
ErrHndlr: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.chkFilter.Click", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'UPGRADE_WARNING: Event chkTtls.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkTtls_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTtls.CheckStateChanged
		'*
		'
		On Error GoTo ErrHndlr
		'
		If fSettingPrefs Then Exit Sub
		'
		If chkTtls.CheckState = System.Windows.Forms.CheckState.Checked Then
			MsgBox("WARNING: This option is fast as dirt. You've been warned.", MsgBoxStyle.Information, "I Wouldn't Do That If I Were You")
		End If
		'
		fShowTotals = IIf(chkTtls.CheckState = System.Windows.Forms.CheckState.Checked, True, False)
		SaveSetting(My.Application.Info.Title, "ProspectMgt", "ShowTotals", CStr(fShowTotals))
		'
		SetupGroups()
		SelectGroup()
		'
		Exit Sub
ErrHndlr: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.chkTtls.Click", MsgBoxStyle.Critical, "Error")
	End Sub
	'UPGRADE_WARNING: Form event FProspectMgt.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FProspectMgt_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'*
		'
		On Error GoTo ErrHndlr
		'
		Dim lLastContactID As Integer
		'
		ReadPreferences()
		'
		SetupGroups()
		SelectGroup()
		'
		On Error Resume Next
		rsProspectGroup.MoveFirst()
		lLastContactID = CInt(GetSetting(My.Application.Info.Title, "ProspectMgt", "CurrentRow", CStr(-1)))
		'
		Dim iRow As Short
		With Me.grdProspectGroup
			.Redraw = False
			'
			'
			For iRow = 0 To .Rows - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Me.grdProspectGroup.AddItemBookmark(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Bookmark = .AddItemBookmark(iRow)
				'
				If .Columns(0).Value = lLastContactID Then
					Exit For
				End If
			Next iRow
			'
			.Redraw = True
		End With
		'
		Exit Sub
		'
ErrHndlr: 
		Me.grdProspectGroup.Redraw = True
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.Form_Activate.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub FProspectMgt_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'*
		'
		On Error GoTo ErrHndlr
		'
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.MinHeight = VB6.PixelsToTwipsY(Me.Height)
		FormControl.MinWidth = VB6.PixelsToTwipsX(Me.Width)
		FormControl.DataForm = False
		'
		rsProspectGroup = New ADODB.Recordset
		rsHistory = New ADODB.Recordset
		'
		Exit Sub
ErrHndlr: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.Form_Load.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Public Sub SetupGroups()
		'*
		'
		On Error GoTo ErrHndlr
		'
		Dim rsGroups As ADODB.Recordset
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		lstGroups.Items.Clear()
		'
		rsGroupCategories = New ADODB.Recordset
		'
		SortColumn = CShort(GetSetting(My.Application.Info.Title, "ProspectMgt", "SortColumn", CStr(eSortColumn.SortByGroup)))
		'
		If Filter_Renamed = eFilter.FilterAll Then
			rsGroupCategories.Open("SELECT * FROM [tblGroupCategories] ORDER BY [" & SortColumnText(SortColumn) & "]", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		ElseIf Filter_Renamed = eFilter.FilterStandard Then 
			rsGroupCategories.Open("SELECT * FROM [tblGroupCategories] WHERE [Label] NOT LIKE 'AM Best%' ORDER BY [" & SortColumnText(SortColumn) & "]", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		ElseIf Filter_Renamed = eFilter.FilterAMBest Then 
			rsGroupCategories.Open("SELECT * FROM [tblGroupCategories] WHERE [Label] LIKE 'AM Best%' ORDER BY [" & SortColumnText(SortColumn) & "]", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		ElseIf Filter_Renamed = eFilter.FilterNone Then 
			rsGroupCategories.Open("SELECT * FROM [tblGroupCategories] WHERE [Label] = 'Dummy'", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		End If
		'
		rsGroups = New ADODB.Recordset
		'
		Dim sSQL As String
		With rsGroupCategories
			Do While .eof = False
				If rsGroups.State = ADODB.ObjectStateEnum.adStateOpen Then rsGroups.Close()
				If fShowTotals Then
					If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
						rsGroups.Open("SELECT Count(*) as GroupCount FROM QProspectMgt WHERE " & .Fields("Formula").Value, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
						'
					Else
						'
						sSQL = "SELECT Count(*) as GroupCount " & "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & "WHERE " & ConvertFormula(.Fields("Formula").Value)
						'
						rsGroups.Open(sSQL, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
					End If
					'
					lstGroups.Items.Add(.Fields("Label").Value & " (" & rsGroups.Fields("GroupCount").Value & ")")
				Else
					lstGroups.Items.Add(.Fields("Label").Value)
				End If
				'UPGRADE_ISSUE: ListBox property lstGroups.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
				VB6.SetItemData(lstGroups, lstGroups.NewIndex, .Fields("RecID").Value)
				.MoveNext()
			Loop 
		End With
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		'
		Exit Sub
		'
ErrHndlr: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		System.Windows.Forms.Application.DoEvents()
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.SetupGroups.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'UPGRADE_WARNING: Event FProspectMgt.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FProspectMgt_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		'*
		'
		On Error Resume Next
		grdHistory.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(grdHistory.Top) - 160)
	End Sub
	
	Private Sub FProspectMgt_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'*
		'
		On Error Resume Next
		SaveSetting(My.Application.Info.Title, "ProspectMgt", "CurrentRow", grdProspectGroup.Columns(0).Value)
	End Sub
	
	Private Sub grdProspectGroup_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdProspectGroup.ClickEvent
		'*
		'
		On Error GoTo ErrHndlr
		'
		LoadHistory(Me.grdProspectGroup.Columns(0).Value)
		'
		Exit Sub
ErrHndlr: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.grdProspectGroup_Click.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub grdProspectGroup_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdProspectGroup.DblClick
		'*
		'
		On Error GoTo ErrHndlr
		'
		' Company.Fetch grdProspectGroup.Columns(1).Value
		'Company.Contact.Fetch grdProspectGroup.Columns(0).Value
		'
		FContact.LoadContact((grdProspectGroup.Columns(0).Value), True)
		FormMgr.ShowForm(Me, FContact)
		'
		Exit Sub
ErrHndlr: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.grdProspectGroup_DblClick.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'UPGRADE_WARNING: Event lstGroups.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstGroups_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstGroups.SelectedIndexChanged
		'*
		'
		On Error GoTo ErrHndlr
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		System.Windows.Forms.Application.DoEvents()
		'
		rsGroupCategories.MoveFirst()
		rsGroupCategories.Find("RecID = " & VB6.GetItemData(lstGroups, lstGroups.SelectedIndex),  , ADODB.SearchDirectionEnum.adSearchForward)
		'
		Dim sSQL As String
		If Not rsGroupCategories.eof Then
			Me.grdProspectGroup.Redraw = False
			Me.grdProspectGroup.RemoveAll()
			'
			If rsProspectGroup.State = ADODB.ObjectStateEnum.adStateOpen Then
				rsProspectGroup.Close()
			End If
			'
			lProspectGroupID = VB6.GetItemData(lstGroups, lstGroups.SelectedIndex)
			SaveSetting(My.Application.Info.Title, "ProspectMgt", "CurrentGroupRecID", CStr(lProspectGroupID))
			'
			If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
				rsProspectGroup.Open("SELECT * FROM QProspectMgt WHERE " & rsGroupCategories.Fields("Formula").Value, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
			Else 'SQL
				'
				sSQL = "SELECT  TCompany.ID AS CompanyID, TContact.ID AS ContactID, " & "TCompany.Name AS Company, " & "TContact.[FirstName] + ' ' + [LastName] AS FullName, " & "TContact.State, TContact.Status, TContact.AuthDate, " & "TContact.Status, TContact.ShipStatus, TContact.AuthStatus, " & "TContact.AuthDate, TContact.AuthDays, TContact.ShipDate " & "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & "WHERE " & ConvertFormula(rsGroupCategories.Fields("Formula").Value) & " ORDER BY TContact.State"
				'
				rsProspectGroup.Open(sSQL, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
			End If
			'
			With rsProspectGroup
				Do Until .eof
					If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
						Me.grdProspectGroup.AddItem(.Fields("TContact.ID").Value & vbTab & .Fields("TCompany.ID").Value & vbTab & .Fields("FullName").Value & vbTab & .Fields("Company").Value & vbTab & .Fields("State").Value & vbTab & .Fields("Status").Value & vbTab & .Fields("AuthDate").Value)
					Else
						Me.grdProspectGroup.AddItem(.Fields("ContactID").Value & vbTab & .Fields("CompanyID").Value & vbTab & .Fields("FullName").Value & vbTab & .Fields("Company").Value & vbTab & .Fields("State").Value & vbTab & .Fields("Status").Value & vbTab & .Fields("AuthDate").Value)
					End If
					'
					.MoveNext()
				Loop 
			End With
			'
			Me.grdProspectGroup.Redraw = True
		End If
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		System.Windows.Forms.Application.DoEvents()
		'
		Exit Sub
		'
ErrHndlr: 
		Me.grdProspectGroup.Redraw = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		System.Windows.Forms.Application.DoEvents()
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.lstGroups_Click.", MsgBoxStyle.Critical, "Error")
	End Sub
	Public Sub ReadPreferences()
		'*
		'
		On Error GoTo ErrHndlr
		'
		fSettingPrefs = True
		'
		'\\ Filter
		Filter_Renamed = GetSetting(My.Application.Info.Title, "ProspectMgt", "FilterGroups", CStr(eFilter.FilterAll))
		'
		If Filter_Renamed = eFilter.FilterAll Then
			chkFilter(0).CheckState = System.Windows.Forms.CheckState.Checked
			chkFilter(1).CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkFilter(0).CheckState = IIf(Filter_Renamed = eFilter.FilterStandard, System.Windows.Forms.CheckState.Checked, System.Windows.Forms.CheckState.Unchecked)
			chkFilter(1).CheckState = IIf(Filter_Renamed = eFilter.FilterAMBest, System.Windows.Forms.CheckState.Checked, System.Windows.Forms.CheckState.Unchecked)
		End If
		'
		'\\ Totals
		fShowTotals = CBool(GetSetting(My.Application.Info.Title, "ProspectMgt", "ShowTotals", "False"))
		chkTtls.CheckState = IIf(fShowTotals, System.Windows.Forms.CheckState.Checked, System.Windows.Forms.CheckState.Unchecked)
		'
		'\\ Sort
		SortColumn = CShort(GetSetting(My.Application.Info.Title, "ProspectMgt", "SortColumn", CStr(eSortColumn.SortByGroup)))
		'
		If SortColumn = eSortColumn.SortByGroup Then
			optSort(0).Checked = True
		Else
			optSort(1).Checked = True
		End If
		'
		fSettingPrefs = False
		'
		Exit Sub
		'
ErrHndlr: 
		fSettingPrefs = False
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.General.ReadPreferences", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'UPGRADE_WARNING: Event optSort.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSort_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSort.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optSort.GetIndex(eventSender)
			'*
			'
			On Error GoTo ErrHndlr
			'
			If fSettingPrefs = True Then Exit Sub
			'
			SortColumn = Index
			SaveSetting(My.Application.Info.Title, "ProspectMgt", "SortColumn", CStr(Index))
			'
			SetupGroups()
			SelectGroup()
			'
			Exit Sub
			'
ErrHndlr: 
			MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.optSort.Click", MsgBoxStyle.Critical, "Error")
		End If
	End Sub
	
	Public Sub SelectGroup()
		'*
		'
		On Error GoTo ErrHndlr
		'
		'\\ Local Declarations
		Dim iCur As Short
		Dim iCt As Short
		'
		lProspectGroupID = CInt(GetSetting(My.Application.Info.Title, "ProspectMgt", "CurrentGroupRecID", CStr(0)))
		'
		With lstGroups
			iCt = lstGroups.Items.Count - 1
			If iCt < 0 Then Exit Sub
			'
			For iCur = 0 To iCt
				If VB6.GetItemData(lstGroups, iCur) = lProspectGroupID Then
					.SelectedIndex = iCur
					Exit Sub
				End If
			Next 
			'
			.SelectedIndex = 0
		End With
		'
		Exit Sub
		'
ErrHndlr: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.General.SelectGroup", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub LoadHistory(ByVal plCustomerID As Integer)
		'*
		'
		On Error GoTo EH
		'
		grdHistory.Redraw = False
		'
		Me.grdHistory.RemoveAll()
		'
		Dim rsHistory As ADODB.Recordset
		Dim cmdHistory As New ADODB.Command
		If plCustomerID <> -1 Then
			'
			rsHistory = New ADODB.Recordset
			'
			If MMain.ConnType = MMain.ConnectionTypeEnum.Access Then
				rsHistory.Open("SELECT * FROM tblSupportActs WHERE CustRecID = " & plCustomerID & " ORDER BY Date DESC, Time DESC", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
			Else
				'
				With cmdHistory
					.ActiveConnection = cnMain
					.CommandText = "dbo.UpParmSelSupportActs"
					.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
					.Parameters.Append(.CreateParameter("CustomerID", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput,  , plCustomerID))
					rsHistory = .Execute
				End With 'cmdSupportAct
				'
				'UPGRADE_NOTE: Object cmdHistory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				cmdHistory = Nothing
			End If
			'
			With rsHistory
				Do While Not .eof
					grdHistory.AddItem(.Fields("RecID").Value & vbTab & .Fields("CustRecID").Value & vbTab & .Fields("Date").Value & vbTab & .Fields("Time").Value & vbTab & .Fields("Type").Value & vbTab & .Fields("User").Value & vbTab & .Fields("Subject").Value & vbTab & .Fields("Results").Value)
					.MoveNext()
				Loop 
			End With 'rsHistory
			'
			rsHistory.Close()
			'UPGRADE_NOTE: Object rsHistory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsHistory = Nothing
		End If
		'
		grdHistory.Redraw = True
		'
		Exit Sub
EH: 
		grdHistory.Redraw = True
		MsgBox(Err.Description)
	End Sub
	
	Private Function SortColumnText(ByVal SortType As eSortColumn) As String
		'*
		'
		If SortType = eSortColumn.SortByGroup Then
			SortColumnText = "Priority"
		ElseIf SortType = eSortColumn.SortByLabel Then 
			SortColumnText = "Label"
		End If
	End Function
	
	Private Function ConvertFormula(ByVal psFormula As String) As String
		On Error GoTo EH
		'
		Dim sFormula As String
		'
		sFormula = psFormula
		'
		sFormula = Replace(sFormula, "ShipDays", "DateDiff(Day,[ShipDate],GETDATE())")
		sFormula = Replace(sFormula, "AuthDaysRemaining", "([AuthDays] - DateDiff(Day, [AuthDate], GETDATE()))")
		sFormula = Replace(sFormula, "isnull(shipdate)", "(ShipDate = Null)")
		'
		ConvertFormula = sFormula
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in Convert Formula.)")
	End Function
End Class