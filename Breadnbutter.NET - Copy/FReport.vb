Option Strict Off
Option Explicit On
Friend Class FReport
	Inherits System.Windows.Forms.Form
	Private Report As New CReport
	'
	Private Enum eFilter
		FilterNone
		FilterStandard
		FilterAMBest
		FilterAll
	End Enum
	'
	Private fShowTotals As Boolean
	Private fSettingPrefs As Boolean
	
	Private lProspectGroupID As Integer
	'
	Private rsGroupCategories As ADODB.Recordset
	Private rsProspectGroup As ADODB.Recordset
	'
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	'
	'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Filter_Renamed As eFilter
	'
	Private WithEvents objLvPrint As clsPrintLV
	'
	Private iLastKey As Short
	'
	'Private SortColumn        As eSortColumn
	'
	Private Sub Command2_Click()
		On Error GoTo EH
		'
		FHistory.Show()
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.Command2_Click.")
	End Sub
	
	'UPGRADE_WARNING: Event chkAlpha.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkAlpha_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAlpha.CheckStateChanged
		If chkAlpha.CheckState = System.Windows.Forms.CheckState.Checked Then
			SaveSetting(My.Application.Info.Title, "ProspectMgt", "SortColumn", CStr(1))
		Else
			SaveSetting(My.Application.Info.Title, "ProspectMgt", "SortColumn", CStr(0))
		End If
		'
		SetupGroups()
		SelectGroup()
	End Sub
	
	Private Sub cmdCopyToClipBoard_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopyToClipBoard.Click
		Dim sText As String
		Dim lItemCount As Integer
		Dim lSubItemCount As Integer
		Dim lColumnCount As Integer
		'
		My.Computer.Clipboard.Clear()
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		SetupData()
		ListView1.Items.Clear()
		ListView1.Columns.Clear()
		Report.FillList((Report.rsReport), ListView1)
		'
		For lColumnCount = 1 To ListView1.Columns.Count
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			sText = sText & ListView1.Columns.Item(lColumnCount).Text & vbTab
		Next 
		'
		sText = sText & vbCrLf
		'
		For lItemCount = 1 To ListView1.Items.Count
			'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			sText = sText & ListView1.Items.Item(lItemCount).Text & vbTab
			'
			'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			For lSubItemCount = 1 To ListView1.Items.Item(lItemCount).SubItems.Count
				'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems().ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				sText = sText & ListView1.Items.Item(lItemCount).SubItems.Item(lSubItemCount).Text & vbTab
			Next 
			'
			sText = sText & vbCrLf
		Next 
		'
		'MsgBox sText
		My.Computer.Clipboard.SetText(sText)
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	'Private Sub DateSet1_Click()
	'  On Error GoTo EH
	'  '
	'  Me.Date1.Caption = FDatePick.DateText(Date1.Caption)
	'  '
	'  Exit Sub
	'EH:
	' MsgBox Err.Description & " in FReport.DateSet1_Click."
	'End Sub
	
	'Private Sub DateSet2_Click()
	'  On Error GoTo EH
	'  '
	'  Me.Date2.Caption = FDatePick.DateText(Date2.Caption)
	'  '
	'  Exit Sub
	'EH:
	' MsgBox Err.Description & " in FReport.DateSet2_Click."
	'End Sub
	
	'UPGRADE_WARNING: Form event FReport.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FReport_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error GoTo EH
		'
		frame1.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(frame1.Width)) / 2), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
		frame1.Visible = True
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.Form_Activate.")
	End Sub
	
	'
	Private Sub FReport_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo EH
		'
		iLastKey = 1
		'
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.Setup(Me, False,  ,  , "Contact Reports")
		'
		'Date1.Caption = Date
		'Date2.Caption = Date
		'optChoice(0) = True
		'
		ReadPreferences()
		'
		SetupGroups()
		SelectGroup()
		'
		'  textNotes.AddItem "(Close 1)"
		'  textNotes.AddItem "(Close 2)"
		'  textNotes.AddItem "(Close 3)"
		'  textNotes.AddItem "(Close 4)"
		'  textNotes.AddItem "(Close 5)"
		'  textNotes.AddItem "(Tech)"
		'
		Dim rsStatus As ADODB.Recordset
		rsStatus = New ADODB.Recordset
		rsStatus.Open("SELECT * FROM tblStatus", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		ComboStatus.Items.Add("Everyone")
		'
		Do While Not rsStatus.eof
			ComboStatus.Items.Add("" & rsStatus.Fields("Status").Value)
			rsStatus.MoveNext()
		Loop 
		'UPGRADE_NOTE: Object rsStatus may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsStatus = Nothing
		'
		FillProductBox()
		'
		StateCombo.Items.Add("All")
		StateCombo.Items.Add("AL")
		StateCombo.Items.Add("AK")
		StateCombo.Items.Add("AZ")
		StateCombo.Items.Add("AR")
		StateCombo.Items.Add("CA")
		StateCombo.Items.Add("CO")
		StateCombo.Items.Add("CT")
		StateCombo.Items.Add("DE")
		StateCombo.Items.Add("DC")
		StateCombo.Items.Add("FL")
		StateCombo.Items.Add("GA")
		StateCombo.Items.Add("HI")
		StateCombo.Items.Add("ID")
		StateCombo.Items.Add("IL")
		StateCombo.Items.Add("IN")
		StateCombo.Items.Add("IA")
		StateCombo.Items.Add("KS")
		StateCombo.Items.Add("KY")
		StateCombo.Items.Add("LA")
		StateCombo.Items.Add("ME")
		StateCombo.Items.Add("MD")
		StateCombo.Items.Add("MA")
		StateCombo.Items.Add("MI")
		StateCombo.Items.Add("MN")
		StateCombo.Items.Add("MS")
		StateCombo.Items.Add("MO")
		StateCombo.Items.Add("MT")
		StateCombo.Items.Add("NE")
		StateCombo.Items.Add("NV")
		StateCombo.Items.Add("NH")
		StateCombo.Items.Add("NJ")
		StateCombo.Items.Add("NM")
		StateCombo.Items.Add("NY")
		StateCombo.Items.Add("NC")
		StateCombo.Items.Add("ND")
		StateCombo.Items.Add("OH")
		StateCombo.Items.Add("OK")
		StateCombo.Items.Add("OR")
		StateCombo.Items.Add("PA")
		StateCombo.Items.Add("PR")
		StateCombo.Items.Add("RI")
		StateCombo.Items.Add("SC")
		StateCombo.Items.Add("SD")
		StateCombo.Items.Add("TN")
		StateCombo.Items.Add("TX")
		StateCombo.Items.Add("UT")
		StateCombo.Items.Add("VT")
		StateCombo.Items.Add("WA")
		StateCombo.Items.Add("WV")
		StateCombo.Items.Add("WI")
		StateCombo.Items.Add("WY")
		'
		LoadcboAction()
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.Form_Load.")
	End Sub
	
	Public Sub ReadPreferences()
		'*
		'
		On Error GoTo ErrHndlr
		'
		'fSettingPrefs = True
		'
		'\\ Filter
		Filter_Renamed = GetSetting(My.Application.Info.Title, "ProspectMgt", "FilterGroups", CStr(eFilter.FilterAll))
		'
		Dim iChoice As Short
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iChoice = nnNum(GetSetting(My.Application.Info.Title, "Reports", "TypeSelect", "0"))
		'
		optChoice(iChoice).Checked = True
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
		If CDbl(GetSetting(My.Application.Info.Title, "ProspectMgt", "SortColumn", CStr(0))) = 0 Then
			chkAlpha.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkAlpha.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		'
		'  '\\ Sort
		'  SortColumn = GetSetting(App.Title, "ProspectMgt", "SortColumn", SortByGroup)
		'  '
		'  If SortColumn = SortByGroup Then
		'    optSort(0).value = True
		'  Else
		'    optSort(1).value = True
		'  End If
		'
		'fSettingPrefs = False
		'
		Exit Sub
		'
ErrHndlr: 
		fSettingPrefs = False
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.General.ReadPreferences", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub SetupData()
		On Error GoTo EH
		'
		Dim DaysMaxTemp As Short
		Dim DaysMinTemp As Short
		'
		Report.SortDirection = GetSortDirection
		Report.SortField = GetSortField
		'
		Report.ProductID = Product.GetProductID(cboProduct.Text)
		Report.FirstName = Me.TextFirstName.Text
		Report.LastName = Me.TextLastName.Text
		Report.Company = Me.TextCompany.Text
		Report.Status = Me.ComboStatus.Text
		Report.City = Me.TextCity.Text
		Report.Zip = Me.TextZip.Text
		Report.State = Me.StateCombo.Text
		Report.Source = Me.TextSource.Text
		'
		If Val(Me.TextDaysMin.Text) > 10000 Or Val(Me.TextDaysMin.Text) < -10000 Then
			DaysMinTemp = 0
		Else
			DaysMinTemp = Val(Me.TextDaysMin.Text)
		End If
		'
		If Val(Me.TextDaysMax.Text) > 10000 Or Val(Me.TextDaysMax.Text) < -10000 Then
			DaysMaxTemp = 0
		Else
			DaysMaxTemp = Val(Me.TextDaysMax.Text)
		End If
		'
		If optChoice(0).Checked = True Then
			Report.Rtype = CReport.ReportType.SimpleContact
		End If
		If optChoice(1).Checked = True Then
			Report.DaysMax = DaysMaxTemp ',Val(Me.TextDaysMax.Text)
			Report.DaysMin = DaysMinTemp 'Val(Me.TextDaysMin.Text)
			Report.Rtype = CReport.ReportType.DaysRemaining
		End If
		If optChoice(2).Checked = True Then
			Report.DaysMin = DaysMinTemp 'Val(Me.TextDaysMin.Text)
			Report.DaysMax = DaysMaxTemp 'Val(Me.TextDaysMax.Text)
			Report.Rtype = CReport.ReportType.DaysNotAuth
		End If
		If optChoice(3).Checked = True Then
			Report.Notes = Me.textNotes.Text
			Report.Rtype = CReport.ReportType.NotesSearch
		End If
		If optChoice(4).Checked = True Then
			Report.DaysMin = Val(Me.txtDays.Text)
			Report.ActionType = cboAction.Text
			Report.Rtype = CReport.ReportType.NoContactIn
		End If
		If optChoice(5).Checked = True Then
			Report.ListQuery = GetGroupQuery
			Report.Rtype = CReport.ReportType.FromList
		End If
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.SetupData.")
	End Sub
	
	'UPGRADE_WARNING: Event FReport.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FReport_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error GoTo EH
		'
		frame1.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(frame1.Width)) / 2), 0, 0, 0, Windows.Forms.BoundsSpecified.X)
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.Form_Resize.")
	End Sub
	
	Private Sub grdHistory_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdHistory.DblClick
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(FResult)
		FResult.TextResult.Text = grdHistory.Columns(7).Value
		FResult.ShowDialog()
		Exit Sub
	End Sub
	
	Private Sub ListView1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ListView1.Click
		On Error GoTo EH
		'
		'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		ListView1.Items.Item(iLastKey).ForeColor = System.Drawing.Color.Black
		If optChoice(5).Checked = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LoadHistory(nnNum(GetCurrentContactID))
		End If
		ListView1.FocusedItem.ForeColor = System.Drawing.Color.Blue
		iLastKey = ListView1.FocusedItem.Index
		Exit Sub
EH: 
	End Sub
	
	Private Sub ListView1_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ListView1.DoubleClick
		On Error GoTo EH
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FContact.LoadContact(nnNum(GetCurrentContactID), True)
		'Company.Fetch nnNum(GetCurrentCompanyID)
		'Company.Contact.Fetch nnNum(GetCurrentContactID)
		FormMgr.ShowForm(FMain.ActiveMDIChild, FContact, True)
		'FormMgr.ShowForm Me, FContact
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.ListView1_DblClick.")
	End Sub
	
	Private Function GetGroupQuery() As String
		Dim sSQL As String
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		System.Windows.Forms.Application.DoEvents()
		'
		rsGroupCategories.MoveFirst()
		rsGroupCategories.Find("RecID = " & VB6.GetItemData(lstGroups, lstGroups.SelectedIndex),  , ADODB.SearchDirectionEnum.adSearchForward)
		'
		If Not rsGroupCategories.eof Then
			'
			sSQL = "SELECT  TCompany.ID AS CompanyID, TContact.ID AS ContactID, " & "TCompany.Name AS Company " & ",TContact.FirstName" & ",TContact.LastName" & ",TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & ",TContact.State, TContact.Status, TContact.AuthDate, TContact.Phone1 " & ",TContact.VersionShipped, TContact.PVVersionShipped " & ", TContact.Source, DATEADD(day,TContact.AuthRemaining,TContact.AuthDate) AS ExpDate" & " FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & "WHERE " & ConvertFormula(rsGroupCategories.Fields("Formula").Value) '& " ORDER BY TContact.State"
			'
			'MsgBox sSQL
			'", TContact.ShipStatus, TContact.AuthStatus, " & _
			'" TContact.AuthDays, TContact.ShipDate "
			GetGroupQuery = sSQL
		End If
	End Function
	
	'UPGRADE_WARNING: Event lstGroups.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstGroups_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstGroups.SelectedIndexChanged
		If lstGroups.Enabled = True Then
			ShowResults_Click(ShowResults, New System.EventArgs())
			grdHistory.RemoveAll()
		End If
		'optChoice(5).value = True
	End Sub
	
	'UPGRADE_WARNING: Event optChoice.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optChoice_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optChoice.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optChoice.GetIndex(eventSender)
			'
			SaveSetting(My.Application.Info.Title, "Reports", "TypeSelect", CStr(Index))
			'
			ListView1.Items.Clear()
			ListView1.Columns.Clear()
			'
			TextDaysMin.Enabled = False
			TextDaysMax.Enabled = False
			textNotes.Enabled = False
			'DateSet1.Enabled = False
			'DateSet2.Enabled = False
			lstGroups.Enabled = False
			chkAlpha.Enabled = False
			chkFilter(0).Enabled = False
			chkFilter(1).Enabled = False
			chkTtls.Enabled = False
			frmCriteria.Visible = True
			ListView1.Height = VB6.TwipsToPixelsY(5295) '4695
			grdHistory.Visible = False
			txtDays.Enabled = False
			cboAction.Enabled = False
			TextFirstName.Enabled = False
			TextLastName.Enabled = False
			TextCity.Enabled = False
			StateCombo.Enabled = False
			TextZip.Enabled = False
			TextSource.Enabled = False
			'txtDays.BackColor = &H80000011
			'cboAction.BackColor = &H80000011
			TextFirstName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000011)
			TextLastName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000011)
			TextCity.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000011)
			StateCombo.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000011)
			TextZip.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000011)
			TextSource.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000011)
			
			If optChoice(0).Checked = True Then
				TextDaysMin.Enabled = True
				TextDaysMax.Enabled = True
				TextFirstName.Enabled = True
				TextLastName.Enabled = True
				TextCity.Enabled = True
				StateCombo.Enabled = True
				TextZip.Enabled = True
				TextSource.Enabled = True
				txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				cboAction.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextFirstName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextLastName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextCity.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				StateCombo.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextZip.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextSource.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			End If
			If optChoice(1).Checked = True Then
				TextDaysMin.Enabled = True
				TextDaysMax.Enabled = True
				TextFirstName.Enabled = True
				TextLastName.Enabled = True
				TextCity.Enabled = True
				StateCombo.Enabled = True
				TextZip.Enabled = True
				TextSource.Enabled = True
				txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				cboAction.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextFirstName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextLastName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextCity.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				StateCombo.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextZip.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextSource.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			End If
			If optChoice(2).Checked = True Then
				TextDaysMin.Enabled = True
				TextDaysMax.Enabled = True
				TextFirstName.Enabled = True
				TextLastName.Enabled = True
				TextCity.Enabled = True
				StateCombo.Enabled = True
				TextZip.Enabled = True
				TextSource.Enabled = True
				txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				cboAction.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextFirstName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextLastName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextCity.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				StateCombo.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextZip.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextSource.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			End If
			If optChoice(3).Checked = True Then
				'    TextDaysMin.Enabled = False
				'    TextDaysMax.Enabled = False
				textNotes.Enabled = True
				TextFirstName.Enabled = True
				TextLastName.Enabled = True
				TextCity.Enabled = True
				StateCombo.Enabled = True
				TextZip.Enabled = True
				TextSource.Enabled = True
				txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				cboAction.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextFirstName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextLastName.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextCity.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				StateCombo.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextZip.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
				TextSource.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
			End If
			If optChoice(4).Checked = True Then
				txtDays.Enabled = True
				cboAction.Enabled = True
			End If
			'
			If optChoice(5).Checked = True Then
				grdHistory.Visible = True
				grdHistory.RemoveAll()
				'ListView1.Height = 6500
				frmCriteria.Visible = False
				lstGroups.Enabled = True
				chkAlpha.Enabled = True
				chkFilter(0).Enabled = True
				chkFilter(1).Enabled = True
				chkTtls.Enabled = True
			End If
			grdHistory.RemoveAll()
		End If
	End Sub
	
	Private Sub PreviewReport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PreviewReport.Click
		On Error GoTo EH
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		SetupData()
		'If optChoice(4).Value = True Then
		'  Report.PreviewReport ("Frontier")
		'End If
		If (optChoice(1).Checked = True) Or (optChoice(2).Checked = True) Then
			Report.PreviewReport(("DaysLeft"))
		End If
		If (optChoice(0).Checked = True) Or (optChoice(3).Checked = True) Or (optChoice(5).Checked = True) Then
			Report.PreviewReport(("Simple Contact"))
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.PreviewReport_Click.")
	End Sub
	
	Private Sub PrintButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PrintButton.Click
		Dim X As Short
		On Error GoTo EH
		If GetPrinter = True Then Exit Sub
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		SetupData()
		ListView1.Items.Clear()
		ListView1.Columns.Clear()
		Report.FillList((Report.rsReport), ListView1)
		'
		'new code'''''''''''''''''''''''''''
		'Instantiate the PrintListView class
		objLvPrint = New clsPrintLV
		'
		For X = 1 To iNumofCopies
			'Call the Print command from the class
			objLvPrint.PrintListView(ListView1, 0.1, 8, "ListView Report", clsPrintLV.Orientation.Landscape, True, False)
		Next X
		'
		'Destroy the object
		'UPGRADE_NOTE: Object objLvPrint may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objLvPrint = Nothing
		'
		'End Code'''''''''''''''''''''''''''
		'
		'  'If (optChoice(4).Value = True) = True Then
		'  '  Report.PrintReport ("Frontier")
		'  'End If
		'  If (optChoice(1).Value = True) Or (optChoice(2).Value = True) Then
		'    Report.PrintReport ("DaysLeft")
		'  End If
		'  If (optChoice(0).Value = True) Or (optChoice(3).Value = True) Or (optChoice(5).Value = True) Then
		'    Report.PrintReport ("Simple Contact")
		'  End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.PrintButton_Click.")
	End Sub
	
	Private Sub ShowResults_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ShowResults.Click
		On Error GoTo EH
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		SetupData()
		ListView1.Items.Clear()
		ListView1.Columns.Clear()
		Report.FillList((Report.rsReport), ListView1)
		LabelResults.Text = Report.rsReport.RecordCount & " Results"
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FReport.ShowResults_Click.")
	End Sub
	
	Private Function GetCurrentCompanyID() As String
		On Error GoTo EH
		'
		Dim sTemp As String
		Dim sTempChar As String
		Dim iCount As Short
		Dim iLength As Short
		iLength = Len(ListView1.FocusedItem.Name)
		Do 
			iCount = iCount + 1
			sTempChar = Mid(ListView1.FocusedItem.Name, iCount, 1)
			sTemp = sTemp & sTempChar
		Loop While (sTempChar <> "A") And (iCount <= Len(ListView1.FocusedItem.Name))
		iLength = Len(sTemp) - 1
		GetCurrentCompanyID = Mid(sTemp, 1, iLength)
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in FReport.GetCurrentCompanyID.")
	End Function
	
	Private Function GetCurrentContactID() As String
		On Error GoTo EH
		'
		Dim sTemp As String
		Dim sTempChar As String
		Dim iCount As Short
		Dim iLength As Short
		iLength = Len(ListView1.FocusedItem.Name)
		Do 
			iCount = iCount + 1
			sTempChar = Mid(ListView1.FocusedItem.Name, iCount, 1)
			sTemp = sTemp & sTempChar
		Loop While (sTempChar <> "A") And (iCount <= Len(ListView1.FocusedItem.Name))
		'
		sTemp = vbNullString
		Do 
			iCount = iCount + 1
			sTempChar = Mid(ListView1.FocusedItem.Name, iCount, 1)
			sTemp = sTemp & sTempChar
		Loop While (iCount <= iLength)
		GetCurrentContactID = sTemp
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in FReport.GetCurrentContactID.")
	End Function
	
	Private Sub FillProductBox()
		Dim Products As New CProducts
		Dim i As Short
		'
		cboProduct.Items.Clear()
		'
		Product.LoadCollection(Products)
		'
		'cboProduct.AddItem "All Products"
		'
		For i = 1 To Products.Count
			cboProduct.Items.Add(Products.Item(i).Product)
		Next i
		'
		cboProduct.SelectedIndex = 0
		'
		'UPGRADE_NOTE: Object Products may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Products = Nothing
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
		'SortColumn = GetSetting(App.Title, "ProspectMgt", "SortColumn", SortByGroup)
		'
		'  If Filter = FilterAll Then
		'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] ORDER BY [Priority]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
		'  ElseIf Filter = FilterStandard Then
		'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] NOT LIKE 'AM Best%' ORDER BY [Priority]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
		'  ElseIf Filter = FilterAMBest Then
		'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] LIKE 'AM Best%' ORDER BY [Priority]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
		'  ElseIf Filter = FilterNone Then
		'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] = 'Dummy'", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
		'  End If
		'
		Dim sOrderField As String
		'
		If chkAlpha.CheckState = System.Windows.Forms.CheckState.Checked Then
			sOrderField = "Label"
		Else
			sOrderField = "Priority"
		End If
		If Filter_Renamed = eFilter.FilterAll Then
			rsGroupCategories.Open("SELECT * FROM [tblGroupCategories] ORDER BY [" & sOrderField & "]", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		ElseIf Filter_Renamed = eFilter.FilterStandard Then 
			rsGroupCategories.Open("SELECT * FROM [tblGroupCategories] WHERE [Label] NOT LIKE 'AM Best%' ORDER BY [" & sOrderField & "]", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		ElseIf Filter_Renamed = eFilter.FilterAMBest Then 
			rsGroupCategories.Open("SELECT * FROM [tblGroupCategories] WHERE [Label] LIKE 'AM Best%' ORDER BY [" & sOrderField & "]", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmReports.SetupGroups.", MsgBoxStyle.Critical, "Error")
	End Sub
	
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
		'  If chkTtls.value = vbChecked Then
		'    MsgBox "WARNING: This option is fast as dirt. You've been warned.", _
		''           vbInformation, "I Wouldn't Do That If I Were You"
		'  End If
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
	
	Private Function GetSortField() As String
		'UPGRADE_ISSUE: MSComctlLib.ListView property ListView1.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If ListView1.Columns.Count >= (ListView1.SortKey + 1) Then
			'UPGRADE_ISSUE: MSComctlLib.ListView property ListView1.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Lower bound of collection ListView1.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			GetSortField = ListView1.Columns.Item(ListView1.SortKey + 1).Text
		Else
			GetSortField = vbNullString
		End If
	End Function
	
	Private Function GetSortDirection() As Short
		With ListView1
			If .Sorting = System.Windows.Forms.SortOrder.Ascending Then
				GetSortDirection = 0
			Else
				GetSortDirection = 1
			End If
		End With
	End Function
	
	Private Sub SortListView(ByVal lvwCur As System.Windows.Forms.ListView, ByVal colHdr As System.Windows.Forms.ColumnHeader, Optional ByVal sSortOrder As String = "")
		On Error GoTo ErrorHandler
		'
		With lvwCur
			'
			'If .SortKey > -1 Then .ColumnHeaders.Item(.SortKey + 1).Icon = 0
			'
			'UPGRADE_ISSUE: MSComctlLib.ListView property lvwCur.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.SortKey = colHdr.Index - 1
			'
			If sSortOrder <> vbNullString Then
				.Sorting = IIf(sSortOrder = "Ascending", System.Windows.Forms.SortOrder.Ascending, System.Windows.Forms.SortOrder.Descending)
			Else
				.Sorting = IIf(.Sorting = System.Windows.Forms.SortOrder.Ascending, System.Windows.Forms.SortOrder.Descending, System.Windows.Forms.SortOrder.Ascending)
			End If
			'
			.Sort()
			'
			'.ColumnHeaders.Item(colHdr.Index).Icon = IIf(.SortOrder = lvwAscending, "imgAscending", "imgDescending")
			'
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox(ErrorToString(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: FReports.SortListView")
	End Sub
	
	Private Sub ListView1_ColumnClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
		Dim ColumnHeader As System.Windows.Forms.ColumnHeader = ListView1.Columns(eventArgs.Column)
		On Error GoTo ErrorHandler
		'
		SortListView(ListView1, ColumnHeader)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.ListView1.ColumnClick")
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
	
	Private Function GetPrinter() As Boolean
		FPrinterSelect.ShowDialog()
		GetPrinter = FPrinterSelect.bPrintCancel
	End Function
	
	Private Sub LoadcboAction()
		Dim rsType As New ADODB.Recordset
		'
		rsType.Open("SELECT * FROM tblActivities WHERE ActivityType = 1 ORDER BY Activity", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		Do While Not rsType.eof
			cboAction.Items.Add(rsType.Fields("Activity").Value & vbNullString)
			rsType.MoveNext()
		Loop 
		'
		rsType.Close()
		'
		'cboAction.AddItem vbNullString
		'
		rsType.Open("SELECT * FROM tblActivities WHERE ActivityType = 0 ORDER BY Activity", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		Do While Not rsType.eof
			cboAction.Items.Add(rsType.Fields("Activity").Value & vbNullString)
			rsType.MoveNext()
		Loop 
		'
		rsType.Close()
		'
		'UPGRADE_NOTE: Object rsType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsType = Nothing
		'
		cboAction.SelectedIndex = 1
	End Sub
End Class