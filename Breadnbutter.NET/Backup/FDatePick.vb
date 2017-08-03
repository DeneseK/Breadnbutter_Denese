Option Strict Off
Option Explicit On
Friend Class FDatePick
	Inherits System.Windows.Forms.Form
	
	Private dtValue As Date
	
	'UPGRADE_NOTE: DateValue was upgraded to DateValue_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public ReadOnly Property DateValue_Renamed(ByVal CurrentValue As Object) As Date
		Get
			On Error GoTo ErrCall
			'
			Me.ShowDialog()
			'
			DateValue_Renamed = dtValue
			'
			Exit Property
ErrCall: 
			MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.DateValue", MsgBoxStyle.Critical, "Error")
			
		End Get
	End Property
	
	Public ReadOnly Property DateText(ByVal CurrentValue As Object) As String
		Get
			On Error GoTo ErrCall
			'
			If IsDate(CurrentValue) Then
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Calendar1.Value = CurrentValue
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dtValue = CurrentValue
				'
			Else
				'
				Calendar1.Value = Today
				dtValue = #12:00:00 AM#
				'
			End If
			'
			'dtValue = CurrentValue 'Calendar1.Value
			Me.ShowDialog()
			'
			'UPGRADE_WARNING: Couldn't resolve default property of object CurrentValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (dtValue <> #12:00:00 AM#) And (dtValue <> CurrentValue) Then
				DateText = VB6.Format(dtValue, "mm/dd/yyyy")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrentValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DateText = CurrentValue
			End If
			'
			Exit Property
ErrCall: 
			MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.DateText", MsgBoxStyle.Critical, "Error")
			
		End Get
	End Property
	
	Private Sub Calendar1_DateClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DMonthViewEvents_DateClickEvent) Handles Calendar1.DateClick
		Me.cmdSet.Focus()
	End Sub
	
	Private Sub Calendar1_DateDblClick(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DMonthViewEvents_DateDblClickEvent) Handles Calendar1.DateDblClick
		On Error GoTo ErrCall
		'
		dtValue = Calendar1.Value
		Me.Close()
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.Calendar1_DblClick", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'Private Sub Calendar1_DblClick()
	'  On Error GoTo ErrCall
	'  '
	'  dtValue = Calendar1.Value
	'  Unload Me
	'  '
	'  Exit Sub
	'ErrCall:
	'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.Calendar1_DblClick", vbCritical, "Error"
	'
	'End Sub
	
	Private Sub cmdSet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSet.Click
		On Error GoTo ErrCall
		'
		dtValue = Calendar1.Value
		'
		Me.Close()
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.cmdSet_Click", MsgBoxStyle.Critical, "Error")
		
	End Sub
End Class