Option Strict Off
Option Explicit On
Friend Class FSingleLabel
	Inherits System.Windows.Forms.Form
	
	Private Sheet(20) As Boolean
	Private lContactID As Integer
	
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
		ErrorMgr.Raise("FSingleLabel", "cmdCancel_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub cmdLabel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLabel.Click
		Dim Index As Short = cmdLabel.GetIndex(eventSender)
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Object
110: For i = 0 To 19
120: cmdLabel(i).Text = ""
130: Next i
140: cmdLabel(Index).Text = "Print"
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FSingleLabel", "cmdLabel_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
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
		ErrorMgr.Raise("FSingleLabel", "cmdPreview_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Public Sub PrintSingle(ByRef plContactID As Integer)
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: If plContactID > 0 Then
110: lContactID = plContactID
120: Me.ShowDialog()
130: Else
140: MsgBox("Contact not ready")
150: End If
		'<EhFooter>
		'
		Exit Sub
		'
EH: 
		ErrorMgr.Raise("FSingleLabel", "PrintSingle", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	
	Private Sub FSingleLabel_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim RLabels As Object
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
90: RLabels = New RLabels
100: Dim i As Short
110: For i = 0 To 19
120: cmdLabel(i).Text = ""
130: Next i
140: cmdLabel(0).Text = "Print"
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FSingleLabel", "Form_Load", Err.Number, Err.Description, Erl())
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
		'RLabels.Show vbModal, FMain
120: Me.Close()
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FSingleLabel", "cmdPrint_Click", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Private Sub CreateSheetArray()
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: Dim i As Short
110: For i = 0 To 19
120: If cmdLabel(i).Text = "Print" Then
130: Sheet(i + 1) = True
140: Else
150: Sheet(i + 1) = False
160: End If
170: Next i
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FSingleLabel", "CreateSheetArray", Err.Number, Err.Description, Erl())
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
		RLabels.SetPages(1)
130: 'UPGRADE_WARNING: Couldn't resolve default property of object RLabels.SetDBCurrent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RLabels.SetDBCurrent(lContactID)
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise("FSingleLabel", "SetupLabels", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
End Class