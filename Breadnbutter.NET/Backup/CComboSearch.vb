Option Strict Off
Option Explicit On
Friend Class CComboSearch
	
	Private WithEvents cmbBox As AxSSDataWidgets_B_OLEDB.AxSSOleDBCombo
	Private KeyCount As Short
	Private iSearchCol As Short
	Private dcSearch As VB6.ADODC
	Private sOriginalText As String
	
	Private Declare Function SendMessageBynum Lib "USER32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	Const EM_LIMITTEXT As Integer = &HC5s
	
	Public Sub Setup(ByRef NewCmbBox As AxSSDataWidgets_B_OLEDB.AxSSOleDBCombo, Optional ByRef piSearchCol As Short = 0, Optional ByRef pdcSearch As VB6.ADODC = Nothing)
		On Error GoTo ErrCall
		'
		Dim lTxtMax As Integer
		'
		cmbBox = NewCmbBox
		KeyCount = 0
		'
		lTxtMax = CInt(0 & NewCmbBox.TagVariant)
		SendMessageBynum(NewCmbBox.HwndEdit, EM_LIMITTEXT, lTxtMax, 0)
		'
		iSearchCol = piSearchCol
		dcSearch = pdcSearch
		'
		sOriginalText = cmbBox.Text
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.Setup", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub cmbBox_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbBox.Change
		'UPGRADE_ISSUE: VBControlExtender property cmbBox.DataChanged is not supported at runtime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="74E732F3-CAD8-417B-8BC9-C205714BB4A7"'
		cmbBox.DataChanged = True
	End Sub
	
	Private Sub cmbBox_CloseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbBox.CloseUp
		On Error GoTo ErrCall
		'
		If Not dcSearch Is Nothing Then
			'UPGRADE_WARNING: Couldn't resolve default property of object CVar(cmbBox.Bookmark). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CVar(dcSearch.Recordset.Bookmark). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CObj(dcSearch.Recordset.Bookmark) <> CObj(cmbBox.Bookmark) Then
				'UPGRADE_ISSUE: VBControlExtender property cmbBox.DataChanged is not supported at runtime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="74E732F3-CAD8-417B-8BC9-C205714BB4A7"'
				cmbBox.DataChanged = True
				'UPGRADE_WARNING: Couldn't resolve default property of object cmbBox.Bookmark. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				dcSearch.Recordset.Bookmark = cmbBox.Bookmark
			End If
		Else
			If cmbBox.Text <> sOriginalText Then
				'UPGRADE_ISSUE: VBControlExtender property cmbBox.DataChanged is not supported at runtime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="74E732F3-CAD8-417B-8BC9-C205714BB4A7"'
				cmbBox.DataChanged = True
			End If
		End If
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboSearch.cmbBox_CloseUp.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub cmbBox_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxSSDataWidgets_B_OLEDB._DSSDBComboEvents_KeyPressEvent) Handles cmbBox.KeyPressEvent
		On Error GoTo ErrCall
		'
		KeyCount = KeyCount + 1
		'
		If eventArgs.KeyAscii = 34 Then
			eventArgs.KeyAscii = 148
		ElseIf eventArgs.KeyAscii = 39 Then 
			eventArgs.KeyAscii = 146
		End If
		'
		If eventArgs.KeyAscii = 32 And cmbBox.Text = "" Then
			cmbBox.DroppedDown = True
			eventArgs.KeyAscii = 0
		End If
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.cmbBox_KeyPress", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Private Sub cmbBox_KeyUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxSSDataWidgets_B_OLEDB._DSSDBComboEvents_KeyUpEvent) Handles cmbBox.KeyUpEvent
		On Error GoTo ErrCall
		'
		Dim LenTarget As Short
		'
		If KeyCount < 1 Or eventArgs.KeyCode = System.Windows.Forms.Keys.Back Then KeyCount = 1
		KeyCount = KeyCount - 1
		'
		Dim tempText As String
		If KeyCount = 0 Then
			LenTarget = Len(cmbBox._Text)
			'
			With cmbBox
				If IsCharKeyCode(eventArgs.KeyCode) Then
					.DroppedDown = True
					'
					tempText = .Columns(iSearchCol).Text
					'
					If LCase(Left(tempText, LenTarget)) = LCase(.Text) Then
						.Text = tempText
						.SelStart = LenTarget
						.SelLength = Len(.Text) - LenTarget
					End If
				End If
			End With 'cmbBox
		End If
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.cmbBox_KeyUp", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Public Function IsCharKeyCode(ByRef pKeyCode As Short) As Boolean
		On Error GoTo ErrCall
		'
		Dim fTemp As Boolean
		'
		fTemp = False
		Select Case pKeyCode
			Case 32, 48 To 57, 65 To 90, 96 To 111, 186 To 192, 219 To 222
				If pKeyCode <> 108 Then fTemp = True
		End Select
		IsCharKeyCode = fTemp
		
		' 32 space
		' 48 to 57 0-9
		' 65 to 90 a-z
		
		' 96 to 111 (not 108) key pad keys (not enter)
		
		'  ; 186
		'  = 187
		'  , 188
		'  - 189
		'  . 190
		'  / 191
		'  ` 192
		
		'  [ 219
		'  \ 220
		'  ] 221
		'  ' 222
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.IsCharKeyCode", MsgBoxStyle.Critical, "Error")
	End Function
End Class