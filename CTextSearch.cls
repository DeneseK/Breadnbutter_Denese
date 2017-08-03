VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txtBox As TextBox
Attribute txtBox.VB_VarHelpID = -1
'Private WithEvents cboBox As cboBox
Private KeyCount As Integer
Private iSearchCol As Integer
Private rsSearch As ADODB.Recordset
Private sOriginalText As String

Public Sub Setup(NewTextBox As TextBox, Optional piSearchCol As Integer, Optional prsSearch As ADODB.Recordset)
  On Error GoTo ErrCall
  '
  Set txtBox = NewTextBox
  KeyCount = 0
  '
  iSearchCol = piSearchCol
  Set rsSearch = prsSearch
  '
  sOriginalText = txtBox.Text
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.Setup", vbCritical, "Error"
End Sub

Private Sub txtBox_Change()
  txtBox.DataChanged = True
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo EH
  '
  If KeyCode = vbKeyUp Then
    If Not rsSearch.BOF Then
      rsSearch.MovePrevious
      '
      If Not rsSearch.BOF Then
        txtBox.Text = rsSearch.Fields(iSearchCol).Value & vbNullString
      Else
        rsSearch.MoveNext
      End If
    End If
  ElseIf KeyCode = vbKeyDown Then
    If Not rsSearch.eof Then
      rsSearch.MoveNext
      '
      If Not rsSearch.eof Then
        txtBox.Text = rsSearch.Fields(iSearchCol).Value & vbNullString
      Else
        rsSearch.MovePrevious
      End If
    End If
  End If
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in Class Text Search: Text Box Key Down."
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

Private Sub txtBox_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrCall
  '
  KeyCount = KeyCount + 1
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in CTextSearch.txtBox_KeyPress", vbCritical, "Error"
End Sub

Private Sub txtBox_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrCall:
  '
  Dim LenTarget As Integer
  '
  If KeyCount < 1 Or KeyCode = vbKeyBack Then KeyCount = 1
  KeyCount = KeyCount - 1
  '
  If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)
  Else
    If KeyCount = 0 Then
      LenTarget = Len(txtBox)
      '
      With txtBox
      If IsCharKeyCode(KeyCode) Then
        Dim tempText As String
        '
        rsSearch.MoveFirst
        rsSearch.Find rsSearch.Fields(iSearchCol).Name & " LIKE '" & Replace(.Text, "'", "''") & "%'"
        '
        If Not rsSearch.eof Then
          tempText = rsSearch.Fields(iSearchCol).Value
          '
          If LCase(Left$(tempText, LenTarget)) = LCase(.Text) Then
            .Text = tempText
            .SelStart = LenTarget
            .SelLength = Len(.Text) - LenTarget
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
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.cmbBox_KeyUp", vbCritical, "Error"
End Sub

Public Function IsCharKeyCode(pKeyCode As Integer) As Boolean
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
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.IsCharKeyCode", vbCritical, "Error"
End Function

