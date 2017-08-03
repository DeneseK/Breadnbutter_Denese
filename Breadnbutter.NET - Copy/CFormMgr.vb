Option Strict Off
Option Explicit On
Friend Class CFormMgr
	
	Private frmMainForm As System.Windows.Forms.Form
	
	Public Event SetStatus(ByRef Status As String) '0=Working 1=Ready
	Public Event SetDescription(ByRef FormName As String, ByRef FormDescription As String)
	
	Public Function ShowForm(ByRef pCurForm As System.Windows.Forms.Form, ByRef pShowForm As System.Windows.Forms.Form, Optional ByRef bHideCurForm As Boolean = False) As Boolean
		On Error GoTo ErrCall
		'
		Dim bCancel As Boolean
		'
		RaiseEvent SetStatus(CStr(0))
		'
		Dim iRsp As Short
		Dim sMsg As String
		If Not pCurForm Is Nothing Then
			
			'UPGRADE_ISSUE: Control FormControl could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			bCancel = Not pCurForm.FormControl.SwitchFrom
			
			'
			If bCancel Then
				'
				sMsg = "Data could not be saved or contains invalid entries. Would you like to continue editing?" & vbCrLf & vbCrLf & "NOTE: If you choose no, changes to your data may not be saved."
				'
				iRsp = MsgBox(sMsg, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
				If iRsp = MsgBoxResult.No Then bCancel = False
			End If
		End If
		'
		Dim frm As System.Windows.Forms.Form
		Dim bFormLoaded As Object
		If Not bCancel Then
			frmMainForm.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
			'
			'
			For	Each frm In My.Application.OpenForms
				If frm.Name = pShowForm.Name Then
					'UPGRADE_WARNING: Couldn't resolve default property of object bFormLoaded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					bFormLoaded = True
					Exit For
				End If
			Next frm
			'
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			If Not bFormLoaded Then Load(pShowForm)
			ResizeForm(pShowForm)
			pShowForm.Show()
			'UPGRADE_ISSUE: Control FormControl could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			pShowForm.FormControl.SwitchTo()
			'UPGRADE_ISSUE: Control FormControl could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			RaiseEvent SetDescription((pShowForm.Name), pShowForm.FormControl.Description)
			'
			If Not pCurForm Is Nothing Then
				If pCurForm.Name <> pShowForm.Name Then
					If bHideCurForm Then
						pCurForm.Hide()
					Else
						pCurForm.Close()
					End If
				End If
			End If
			'
			ShowForm = True
		Else
			ShowForm = False
		End If
		'
		RaiseEvent SetStatus(CStr(1))
		'
		Exit Function
ErrCall: 
		ShowForm = False
		MsgBox(Err.Description)
		
	End Function
	
	Public Sub Setup(ByRef pfrmMDI As System.Windows.Forms.Form)
		frmMainForm = pfrmMDI
	End Sub
	
	Public Sub ResizeForm(ByRef pForm As System.Windows.Forms.Form)
		On Error GoTo ErrCall
		'
		Dim sngHeight, sngWidth As Single
		'
		If Not pForm Is Nothing Then
			sngHeight = VB6.PixelsToTwipsY(frmMainForm.ClientRectangle.Height)
			'UPGRADE_ISSUE: Control FormControl could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			If (sngHeight < pForm.FormControl.MinHeight) And (sngHeight > 0) Then
				'UPGRADE_ISSUE: Control FormControl could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
				sngHeight = pForm.FormControl.MinHeight
			End If
			'
			'If frmMainForm.Width - frmMainForm.ScaleWidth > 181 Then
			'  sngWidth = frmMainForm.ScaleWidth + 195
			'Else
			sngWidth = VB6.PixelsToTwipsX(frmMainForm.ClientRectangle.Width) - (2 * VB6.TwipsPerPixelX)
			'End If
			'
			'UPGRADE_ISSUE: Control FormControl could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			If sngWidth < pForm.FormControl.MinWidth Then sngWidth = pForm.FormControl.MinWidth
			'
			pForm.SetBounds(0, 0, VB6.TwipsToPixelsX(sngWidth), VB6.TwipsToPixelsY(sngHeight))
		End If
		'
		Exit Sub
ErrCall: 
		If Not Err.Number = 384 Then
			MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormMgr.ResizeForm.", MsgBoxStyle.Critical, "Error")
		End If
	End Sub
	
	Public Function CloseForm(ByRef pForm As System.Windows.Forms.Form) As Boolean
		On Error GoTo ErrCall
		'
		Dim bCancel As Boolean
		'
		CloseForm = False
		'
		Dim iRsp As Short
		Dim sMsg As String
		If Not pForm Is Nothing Then
			'UPGRADE_ISSUE: Control FormControl could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
			bCancel = Not pForm.FormControl.SwitchFrom
			'
			If bCancel Then
				'
				sMsg = "Data could not be saved or contains invalid entries. Would you like to continue editing?" & vbCrLf & vbCrLf & "NOTE: If you choose no, changes to your data may not be saved."
				'
				iRsp = MsgBox(sMsg, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
				If iRsp = MsgBoxResult.No Then bCancel = False
			End If
			'
			If Not bCancel Then
				pForm.Close()
				RaiseEvent SetDescription("", "Main")
				'
				If frmMainForm.ActiveMDIChild Is Nothing Then
					frmMainForm.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000C)
				End If
				'
				CloseForm = True
			Else
				CloseForm = False
			End If
		Else
			CloseForm = True
		End If
		'
		' CSErrorHandler begin - please do not modify or remove this line
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormMgr.CloseForm.", MsgBoxStyle.Critical, "Error")
		CloseForm = False
	End Function
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		On Error GoTo ErrCall
		'
		'UPGRADE_NOTE: Object frmMainForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmMainForm = Nothing
		'
		Exit Sub
ErrCall: 
		MsgBox("Error: " & Err.Description & " in Form Manager Terminate.")
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class