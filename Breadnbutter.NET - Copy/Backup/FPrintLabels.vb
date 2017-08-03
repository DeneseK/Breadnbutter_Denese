Option Strict Off
Option Explicit On
Friend Class FPrintLabels
	Inherits System.Windows.Forms.Form
	Private Sheet(20) As Boolean
	
	Private Sub chkPageNum_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkPageNum.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: If chkPageNum.CheckState = True Then
110: EnableLabelButtons()
120: Else
130: EnableLabelButtons()
140: End If
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "chkPageNum_KeyUp", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub chkPageNum_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkPageNum.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: If chkPageNum.CheckState = 1 Then
110: EnableLabelButtons()
120: Else
130: DisableLabelButtons()
140: End If
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "chkPageNum_MouseUp", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub DisableLabelButtons()
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Short
110: For i = 0 To 19
120: cmdLabel(i).Enabled = False
130: Next i
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "DisableLabelButtons", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub EnableLabelButtons()
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Short
110: For i = 0 To 19
120: cmdLabel(i).Enabled = True
130: Next i
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "EnableLabelButtons", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	
	Private Sub FlipEnableLabelButtons()
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Short
110: For i = 0 To 19
120: cmdLabel(i).Enabled = Not cmdLabel(i).Enabled
130: Next i
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "FlipEnableLabelButtons", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Me.Close()
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "cmdCancel_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdClearAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearAll.Click
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Short
110: For i = 0 To 19
120: cmdLabel(i).Text = ""
130: Next i
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "cmdClearAll_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdClearSelected_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClearSelected.Click
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim rslabels As New ADODB.Recordset
110: rslabels.Open("SELECT BetaTester FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID WHERE BetaTester = 1", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
120: With rslabels
130: While Not .eof
140: .Fields("betatester").Value = 0
150: .UpdateBatch()
160: .MoveNext()
170: End While
180: End With
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "cmdClearSelected_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdLabel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLabel.Click
		Dim Index As Short = cmdLabel.GetIndex(eventSender)
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: If cmdLabel(Index).Text = "Print" Then
110: cmdLabel(Index).Text = ""
120: Else
130: cmdLabel(Index).Text = "Print"
140: End If
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "cmdLabel_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdLoadGroup_Click()
		Dim RLabels As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.Clear. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.Clear()
		CreateSheetArray()
		'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.SetSheet. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.SetSheet(Sheet)
		'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.SetPages. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.SetPages(Me.chkPageNum)
		'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.SetDBGroup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.SetDBGroup(FChooseGroup.GetGroup)
		'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.Show(VB6.FormShowConstants.Modal, FMain)
	End Sub
	
	Private Sub cmdPreview_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreview.Click
		Dim RLabels As Object
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: SetupLabels()
110: 'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.Show(VB6.FormShowConstants.Modal, FMain)
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "cmdPreview_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
		Dim RLabels As Object
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: SetupLabels()
110: 'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.PrintReport. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.PrintReport(True)
		'<EhFooter>
		'
		Exit Sub
		'
EH: 
		ErrorMgr.Raise("FPrintLabels", "cmdPrint_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub SetupLabels()
		Dim RLabels As Object
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
90: 'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.Clear. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.Clear()
100: CreateSheetArray()
110: 'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.SetSheet. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.SetSheet(Sheet)
120: 'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.SetPages. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.SetPages(Me.chkPageNum)
130: 'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.SetDBGroup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.SetDBGroup(FChooseGroup.GetGroup) 'SetDB
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "SetupLabels", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdPrintAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrintAll.Click
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Short
110: For i = 0 To 19
120: cmdLabel(i).Text = "Print"
130: Next i
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "cmdPrintAll_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub FPrintLabels_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Short
110: For i = 0 To 19
120: cmdLabel(i).Text = "Print"
130: Next i
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "Form_Load", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub CreateSheetArray()
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Short
110: For i = 0 To 19
120: If Me.cmdLabel(i).Text = "Print" Then
130: Sheet(i + 1) = True
140: Else
150: Sheet(i + 1) = False
160: End If
170: Next i
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FPrintLabels", "CreateSheetArray", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
End Class