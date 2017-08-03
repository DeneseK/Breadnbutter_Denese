Option Strict Off
Option Explicit On
Friend Class CInputNumber
	
	Public Enum eNumberType
		NumberTypeByte
		NumberTypeInteger
		NumberTypeLong
		NumberTypeSingle
		NumberTypeDouble
		StringType
	End Enum
	
	Private WithEvents NumTextBox As System.Windows.Forms.TextBox
	Private WithEvents NumCBO As System.Windows.Forms.ComboBox
	Private iNumberType As eNumberType
	
	Private Const ModuleName As String = "CInputNumber"
	
	Public Sub Setup(ByRef pTextbox As System.Windows.Forms.TextBox, ByVal NumberType As eNumberType, Optional ByVal MinValue As Object = Nothing, Optional ByVal MaxValue As Object = Nothing)
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: NumTextBox = pTextbox
		'
110: NumTextBox.SelectionStart = 0
120: NumTextBox.SelectionLength = Len(NumTextBox.Text)
		'
130: iNumberType = NumberType
		'<EhFooter>
		'
		Exit Sub
EH: 
		ErrorMgr.Raise(ModuleName, "Setup", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Public Sub SetupCBO(ByRef pCombbox As System.Windows.Forms.ComboBox, ByVal NumberType As eNumberType, Optional ByVal MinValue As Object = Nothing, Optional ByVal MaxValue As Object = Nothing)
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: NumCBO = pCombbox
		'
110: NumCBO.SelectionStart = 0
120: NumCBO.SelectionLength = Len(NumCBO.Text)
		'
130: iNumberType = NumberType
		'<EhFooter>
		'
		Exit Sub
		'
EH: 
		ErrorMgr.Raise("CInputNumber", "SetupCBO", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Sub
	
	Public Function ValueByte(ByVal psNumText As String) As Byte
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: On Error Resume Next
		'
110: ValueByte = CByte(psNumText)
		'
120: If Err.Number Then
130: ValueByte = 0
140: End If
		'<EhFooter>
		'
		Exit Function
EH: 
		ErrorMgr.Raise(ModuleName, "ValueByte", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Function
	
	Public Function ValueInt(ByVal psNumText As String) As Short
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: On Error Resume Next
		'
110: ValueInt = CShort(psNumText)
		'
120: If Err.Number Then
130: ValueInt = 0
140: End If
		'<EhFooter>
		'
		Exit Function
EH: 
		ErrorMgr.Raise(ModuleName, "ValueInt", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Function
	
	Public Function ValueLng(ByVal psNumText As String) As Integer
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: On Error Resume Next
		'
110: ValueLng = CInt(psNumText)
		'
120: If Err.Number Then
130: ValueLng = 0
140: End If
		'<EhFooter>
		'
		Exit Function
EH: 
		ErrorMgr.Raise(ModuleName, "ValueLng", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Function
	
	Public Function ValueSng(ByVal psNumText As String) As Single
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: On Error Resume Next
		'
110: ValueSng = CSng(psNumText)
		'
120: If Err.Number Then
130: ValueSng = 0
140: End If
		'<EhFooter>
		'
		Exit Function
EH: 
		ErrorMgr.Raise(ModuleName, "ValueSng", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Function
	
	Public Function ValueDbl(ByVal psNumText As String) As Double
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: On Error Resume Next
		'
110: ValueDbl = CDbl(psNumText)
		'
120: If Err.Number Then
130: ValueDbl = 0
140: End If
		'<EhFooter>
		'
		Exit Function
EH: 
		ErrorMgr.Raise(ModuleName, "ValueDbl", Err.Number, Err.Description, Erl())
		'</EhFooter>
	End Function
	
	Private Sub NumTextBox_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles NumTextBox.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'\\ Disallow more than one decimal point
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: If KeyAscii = 46 Then
110: If InStr(1, NumTextBox.Text, ".") > 0 Then
120: KeyAscii = 0
130: GoTo EventExitSub
140: End If
150: End If
		'
		'\\ General Validation
160: If KeyAscii < 48 Or KeyAscii > 57 Then '\\ 0-9
170: If KeyAscii <> 45 Then '\\ -
180: If KeyAscii <> System.Windows.Forms.Keys.Back Then '\\ Backspace
190: If KeyAscii <> System.Windows.Forms.Keys.Delete Then '\\ Delete
200: KeyAscii = 0
210: End If
220: End If
230: End If
240: Else
			'\\ Disallow more than two digits after decimal point
250: If InStr(1, NumTextBox.Text, ".") > 0 Then
260: If NumTextBox.SelectionStart >= InStr(1, NumTextBox.Text, ".") Then
270: If Len(NumTextBox.Text) >= InStr(1, NumTextBox.Text, ".") + 3 Then
280: KeyAscii = 0
290: End If
300: End If
310: End If
320: End If
		'<EhFooter>
		'
		GoTo EventExitSub
EH: 
		ErrorMgr.Raise(ModuleName, "NumTextBox_KeyPress", Err.Number, Err.Description, Erl())
		'</EhFooter>
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub NumTextBox_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles NumTextBox.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: On Error GoTo NumTextBox_EH
		'
110: Dim vConverted As Object
		'
120: On Error Resume Next
		'
130: Select Case iNumberType
			Case eNumberType.NumberTypeByte
140: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CByte(NumTextBox.Text)
150: Case eNumberType.NumberTypeInteger
160: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CShort(NumTextBox.Text)
170: Case eNumberType.NumberTypeLong
180: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CInt(NumTextBox.Text)
190: Case eNumberType.NumberTypeSingle
200: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CSng(NumTextBox.Text)
210: Case eNumberType.NumberTypeDouble
220: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CDbl(NumTextBox.Text)
230: End Select
		'
240: If Err.Number <> 0 Then
250: Beep()
260: NumTextBox.Text = CStr(0)
270: End If
		'
280: GoTo EventExitSub
NumTextBox_EH: 
290: NumTextBox.Text = CStr(0)
		'<EhFooter>
		'
		GoTo EventExitSub
EH: 
		ErrorMgr.Raise(ModuleName, "NumTextBox_Validate", Err.Number, Err.Description, Erl())
		'</EhFooter>
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub NumCBO_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles NumCBO.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		'<EhHeader>
		On Error GoTo EH
		'
		'</EhHeader>
100: On Error GoTo NumTextBox_EH
		'
110: Dim vConverted As Object
120: Dim sConverted As String
		'
130: On Error Resume Next
		'
140: Select Case iNumberType
			Case eNumberType.NumberTypeByte
150: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CByte(NumCBO.Text)
160: Case eNumberType.NumberTypeInteger
170: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CShort(NumCBO.Text)
				
180: Case eNumberType.NumberTypeLong
190: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CInt(NumCBO.Text)
200: Case eNumberType.NumberTypeSingle
210: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CSng(NumCBO.Text)
220: Case eNumberType.NumberTypeDouble
230: 'UPGRADE_WARNING: Couldn't resolve default property of object vConverted. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vConverted = CDbl(NumCBO.Text)
240: Case eNumberType.StringType
250: sConverted = CStr(NumCBO.Text)
260: End Select
		'
270: Dim i As Short
280: Dim bool As Boolean
290: bool = False
300: For i = 0 To NumCBO.Items.Count - 1
			
310: If VB6.GetItemString(NumCBO, i) = NumCBO.Text Then
320: bool = True
330: End If
340: Next 
		'
350: If bool = False Then
360: If iNumberType = eNumberType.StringType Then
				' NumCBO.Text = ""
370: NumCBO.SelectedIndex = 0
380: Else
390: NumCBO.Text = "0"
400: End If
410: End If
420: If Err.Number <> 0 Then
430: Beep()
440: If iNumberType = eNumberType.StringType Then
				'NumCBO.Text = ""
450: NumCBO.SelectedIndex = 0
460: Else
470: NumCBO.Text = "0"
480: End If
490: End If
		'
500: GoTo EventExitSub
NumTextBox_EH: 
510: NumCBO.Text = ""
		'<EhFooter>
		'
		GoTo EventExitSub
		'
EH: 
		ErrorMgr.Raise("CInputNumber", "NumCBO_Validate", Err.Number, Err.Description, Erl())
		'</EhFooter>
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
End Class