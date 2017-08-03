Option Strict Off
Option Explicit On
Friend Class FAuthLog
	Inherits System.Windows.Forms.Form
	
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	'
	Public rsLog As New ADODB.Recordset
	
	'UPGRADE_NOTE: Form_Initialize was upgraded to Form_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Form_Initialize_Renamed()
		On Error GoTo ErrCall
		'
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.MinHeight = 1965
		FormControl.MinWidth = VB6.PixelsToTwipsX(Me.Width)
		FormControl.DataForm = True
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmAuthLog.Form_Initialize.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub OpenLog()
		On Error GoTo ErrorHandler
		'
		'\\ Log
		Dim rsLog As New ADODB.Recordset
		'
		rsLog.LockType = ADODB.LockTypeEnum.adLockPessimistic
		rsLog.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
		rsLog.Open("SELECT * from tbllog", cnMain) '
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: OpenLog")
	End Sub
	
	Public Sub SortListView(ByVal lvwCur As System.Windows.Forms.ListView, ByVal colHdr As System.Windows.Forms.ColumnHeader, Optional ByVal sSortOrder As String = "")
		On Error GoTo ErrorHandler
		'
		With lvwCur
			'
			'UPGRADE_ISSUE: MSComctlLib.ListView property lvwCur.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Lower bound of collection lvwCur.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_ISSUE: MSComctlLib.ColumnHeader property ColumnHeaders.Item.Icon was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			If .SortKey > -1 Then .Columns.Item(.SortKey + 1).Icon = 0
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
			'UPGRADE_WARNING: Lower bound of collection lvwCur.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_ISSUE: MSComctlLib.ColumnHeader property ColumnHeaders.Item.Icon was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.Columns.Item(colHdr.Index).Icon = IIf(.Sorting = System.Windows.Forms.SortOrder.Ascending, "imgAscending", "imgDescending")
			'
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox(ErrorToString(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: FPrimary.General.SortListView")
	End Sub
	
	Public Sub RefreshLogDisplay()
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim iEntryNo As Short
		Dim iActionIcon As Short
		Dim lActTtl As Integer
		Dim lActCur As Integer
		Dim sEmp() As String
		'
		rsLog.LockType = ADODB.LockTypeEnum.adLockPessimistic
		rsLog.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
		rsLog.Open("SELECT * from tblLog", cnMain) '
		lblActs.Text = "(0 of 0)"
		'
		With rsLog
			If ((.BOF = True) And (.eof = True)) = False Then
				.MoveLast()
				lActTtl = .RecordCount
				lblActs.Text = "(" & CStr(.RecordCount) & " of " & CStr(.RecordCount) & ")"
				.MoveFirst()
				rsLog.Close()
				Select Case cboFilter._Text
					Case "None"
						'
						rsLog.Open("SELECT * from tblLog", cnMain) '
						'
					Case "Authorizations, All"
						'
						rsLog.Open("SELECT * FROM [tblLog] WHERE [ActionType] = 'Authorization'", cnMain)
						'
					Case "Authorizations, New"
						'
						rsLog.Open("SELECT * FROM [tblLog] WHERE [ActionSubType] = 'New'", cnMain)
						'
					Case "Authorizations, Extensions"
						'
						rsLog.Open("SELECT * FROM [tblLog] WHERE [ActionSubType] = 'Extension'", cnMain)
						'
					Case "Deauthorizations"
						'
						rsLog.Open("SELECT * FROM [tblLog] WHERE [ActionType] = 'Deauthorization'", cnMain)
						'
					Case "Restorations"
						'
						rsLog.Open("SELECT * FROM [tblLog] WHERE [ActionType] = 'Restoration'", cnMain)
						'
				End Select
			Else
				Exit Sub
			End If
		End With
		'
		With rsLog
			If ((.BOF = True) And (.eof = True)) = False Then
				.MoveLast()
				lblActs.Text = "(" & CStr(.RecordCount) & " of " & CStr(lActTtl) & ")"
				.MoveFirst()
				lvwLog.Items.Clear()
				Do Until .eof
					sEmp = Split(.Fields("Employee").Value)
					Select Case .Fields("ActionType").Value
						Case "Authorization"
							iActionIcon = 3
						Case "Deauthorization"
							iActionIcon = 4
						Case "Restoration"
							iActionIcon = 5
					End Select
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
					lvwLog.Items.Add("r" & CStr(.Fields("ID").Value), VB6.Format(.Fields("ActionDateTime").Value, "YYYY.Mm.Dd") & "  " & VB6.Format(.Fields("ActionDateTime").Value, "Hh:Nn:Ss"), iActionIcon)
					iEntryNo = lvwLog.Items.Count
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems.Item() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					If lvwLog.Items.Item(iEntryNo).SubItems.Count > 1 Then
						lvwLog.Items.Item(iEntryNo).SubItems(1).Text = sEmp(0)
					Else
						lvwLog.Items.Item(iEntryNo).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, sEmp(0)))
					End If
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems.Item(iEntryNo) has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					If lvwLog.Items.Item(iEntryNo).SubItems.Count > 2 Then
						lvwLog.Items.Item(iEntryNo).SubItems(2).Text = .Fields("Company").Value
					Else
						lvwLog.Items.Item(iEntryNo).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, .Fields("Company").Value))
					End If
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems.Item(iEntryNo) has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					If lvwLog.Items.Item(iEntryNo).SubItems.Count > 3 Then
						lvwLog.Items.Item(iEntryNo).SubItems(3).Text = .Fields("User").Value
					Else
						lvwLog.Items.Item(iEntryNo).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, .Fields("User").Value))
					End If
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Lower bound of collection lvwLog.ListItems.Item(iEntryNo) has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					If lvwLog.Items.Item(iEntryNo).SubItems.Count > 4 Then
						lvwLog.Items.Item(iEntryNo).SubItems(4).Text = .Fields("ActionType").Value & ": " & IIf(.Fields("ActionType").Value <> "Restoration", VB6.Format(.Fields("SiteDays").Value, "0000") & " Days", VB6.Format(.Fields("SiteDateTime").Value, "Short Date"))
					Else
						lvwLog.Items.Item(iEntryNo).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, .Fields("ActionType").Value & ": " & IIf(.Fields("ActionType").Value <> "Restoration", VB6.Format(.Fields("SiteDays").Value, "0000") & " Days", VB6.Format(.Fields("SiteDateTime").Value, "Short Date"))))
					End If
					.MoveNext()
				Loop 
			End If
			'
			.Close()
			'
		End With
		'
		'SortListView lvwLog, lvwLog.ColumnHeaders(1), "Descending"
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.General.RefreshLogDisplay")
	End Sub
	
	Private Sub FAuthLog_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo ErrorHandler
		'
		cboFilter.Text = GetSetting(My.Application.Info.Title, "General", "Filter", "None")
		'
		RefreshLogDisplay()
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.lvwLog.ColumnClick")
		
	End Sub
	
	'UPGRADE_WARNING: Event FAuthLog.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FAuthLog_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error GoTo ErrorHandler
		'
		If VB6.PixelsToTwipsX(Width) > 1000 And VB6.PixelsToTwipsY(Height) > 1000 Then
			lvwLog.SetBounds(VB6.TwipsToPixelsX(30), VB6.TwipsToPixelsY(1320), VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Width) - 100), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Height) - 1275))
		End If
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.lvwLog.ColumnClick")
		
	End Sub
	
	Private Sub lvwLog_ColumnClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ColumnClickEventArgs) Handles lvwLog.ColumnClick
		Dim ColumnHeader As System.Windows.Forms.ColumnHeader = lvwLog.Columns(eventArgs.Column)
		On Error GoTo ErrorHandler
		'
		SortListView(lvwLog, ColumnHeader)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.lvwLog.ColumnClick")
	End Sub
	
	Private Sub lvwLog_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvwLog.DoubleClick
		On Error GoTo ErrorHandler
		'
		FActivity.ShowDialog()
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.lvwLog.DoubleClick")
	End Sub
	
	Private Sub cboFilter_CloseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFilter.CloseUp
		On Error GoTo ErrorHandler
		'
		SaveSetting(My.Application.Info.Title, "General", "Filter", cboFilter.Text)
		RefreshLogDisplay()
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.cboFilter.CloseUp")
	End Sub
	Private Sub cboFilter_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFilter.Enter
		On Error GoTo ErrorHandler
		'
		SelectText(cboFilter)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.cboFilter.GotFocus")
	End Sub
	Private Sub cboFilter_InitColumnProps(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFilter.InitColumnProps
		On Error GoTo ErrorHandler
		'
		With cboFilter
			.AddItem("None")
			.AddItem("Authorizations, All")
			.AddItem("Authorizations, New")
			.AddItem("Authorizations, Extensions")
			.AddItem("Deauthorizations")
			.AddItem("Restorations")
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.cboFilter.InitColumnProps")
	End Sub
End Class