Option Strict Off
Option Explicit On
Friend Class CTextSearch
	
	Private WithEvents txtBox As System.Windows.Forms.TextBox
	'Private WithEvents cboBox As cboBox
	Private KeyCount As Short
	Private iSearchCol As Short
	Private rsSearch As ADODB.Recordset
	Private sOriginalText As String
	
	Public Sub Setup(ByRef NewTextBox As System.Windows.Forms.TextBox, Optional ByRef piSearchCol As Short = 0, Optional ByRef prsSearch As ADODB.Recordset = Nothing)
		On Error GoTo ErrCall
		'
		txtBox = NewTextBox
		KeyCount = 0
		'
		iSearchCol = piSearchCol
		rsSearch = prsSearch
		'
		sOriginalText = txtBox.Text
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.Setup", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'UPGRADE_WARNING: Event txtBox.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBox_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBox.TextChanged
		'UPGRADE_ISSUE: TextBox property txtBox.DataChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		txtBox.DataChanged = True
	End Sub
	
	Private Sub txtBox_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtBox.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		On Error GoTo EH
		'
		If KeyCode = System.Windows.Forms.Keys.Up Then
			If Not rsSearch.BOF Then
				rsSearch.MovePrevious()
				'
				If Not rsSearch.BOF Then
					txtBox.Text = rsSearch.Fields(iSearchCol).Value & vbNullString
				Else
					rsSearch.MoveNext()
				End If
			End If
		ElseIf KeyCode = System.Windows.Forms.Keys.Down Then 
			If Not rsSearch.eof Then
				rsSearch.MoveNext()
				'
				If Not rsSearch.eof Then
					txtBox.Text = rsSearch.Fields(iSearchCol).Value & vbNullString
				Else
					rsSearch.MovePrevious()
				End If
			End If
		End If
		'
		Exit Sub
EH: 
		MsgBox(Err.Description & " in Class Text Search: Text Box Key Down.")
	End Sub
	
	'Private Sub cmbBox_CloseUp()
	'  On Error GoTo ErrCall
	'  '
	'  If Not dcSearch Is Nothing Then
	'    If CVar(dcSearch.Recordset.Bookmark) <> CVar(cmbBox.Bookmark) Then
	'      cmbBox.DataChanged = True
	'      dcSearch.Recordset.Bookmark = cmbBox.Bookmark
	'    End If
	'  Else
	'    If cmbBox.Text <> sOriginalText Then
	'      cmbBox.DataChanged = True
	'    End If
	'  End If
	'  '
	'  Exit Sub
	'ErrCall:
	'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboSearch.cmbBox_CloseUp.", vbCritical, "Error"
	'End Sub
	
	Private Sub txtBox_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBox.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		On Error GoTo ErrCall
		'
		KeyCount = KeyCount + 1
		'
		GoTo EventExitSub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in CTextSearch.txtBox_KeyPress", MsgBoxStyle.Critical, "Error")
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtBox_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtBox.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		On Error GoTo ErrCall
		'
		Dim LenTarget As Short
		'
		If KeyCount < 1 Or KeyCode = System.Windows.Forms.Keys.Back Then KeyCount = 1
		KeyCount = KeyCount - 1
		'
		Dim tempText As String
		If (KeyCode = System.Windows.Forms.Keys.Up) Or (KeyCode = System.Windows.Forms.Keys.Down) Then
			txtBox.SelectionStart = 0
			txtBox.SelectionLength = Len(txtBox.Text)
		Else
			If KeyCount = 0 Then
				LenTarget = Len(txtBox.Text)
				'
				With txtBox
					If IsCharKeyCode(KeyCode) Then
						'
						rsSearch.MoveFirst()
						rsSearch.Find(rsSearch.Fields(iSearchCol).Name & " LIKE '" & Replace(.Text, "'", "''") & "%'")
						'
						If Not rsSearch.eof Then
							tempText = rsSearch.Fields(iSearchCol).Value
							'
							If LCase(Left(tempText, LenTarget)) = LCase(.Text) Then
								.Text = tempText
								.SelectionStart = LenTarget
								.SelectionLength = Len(.Text) - LenTarget
							End If
						Else
							'* match not found
							'rsSearch.MoveFirst
						End If
					End If
				End With 'txtBox
			End If
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