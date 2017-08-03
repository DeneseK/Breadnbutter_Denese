VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComboSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cmbBox As SSOleDBCombo
Attribute cmbBox.VB_VarHelpID = -1
Private KeyCount As Integer
Private iSearchCol As Integer
Private dcSearch As Adodc
Private sOriginalText As String

Private Declare Function SendMessageBynum Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const EM_LIMITTEXT As Long = &HC5

Public Sub Setup(NewCmbBox As SSOleDBCombo, Optional piSearchCol As Integer, Optional pdcSearch As Adodc)
  On Error GoTo ErrCall
  '
  Dim lTxtMax As Long
  '
  Set cmbBox = NewCmbBox
  KeyCount = 0
  '
  lTxtMax = CLng(0 & NewCmbBox.TagVariant)
  SendMessageBynum NewCmbBox.HwndEdit, EM_LIMITTEXT, lTxtMax, 0
  '
  iSearchCol = piSearchCol
  Set dcSearch = pdcSearch
  '
  sOriginalText = cmbBox.Text
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.Setup", vbCritical, "Error"
End Sub

Private Sub cmbBox_Change()
  cmbBox.DataChanged = True
End Sub

Private Sub cmbBox_CloseUp()
  On Error GoTo ErrCall
  '
  If Not dcSearch Is Nothing Then
    If CVar(dcSearch.Recordset.Bookmark) <> CVar(cmbBox.Bookmark) Then
      cmbBox.DataChanged = True
      dcSearch.Recordset.Bookmark = cmbBox.Bookmark
    End If
  Else
    If cmbBox.Text <> sOriginalText Then
      cmbBox.DataChanged = True
    End If
  End If
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboSearch.cmbBox_CloseUp.", vbCritical, "Error"
End Sub

Private Sub cmbBox_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrCall
  '
  KeyCount = KeyCount + 1
  '
  If KeyAscii = 34 Then
    KeyAscii = 148
  ElseIf KeyAscii = 39 Then
    KeyAscii = 146
  End If
  '
  If KeyAscii = 32 And cmbBox.Text = "" Then
    cmbBox.DroppedDown = True
    KeyAscii = 0
  End If
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsComboEvents.cmbBox_KeyPress", vbCritical, "Error"
End Sub

Private Sub cmbBox_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrCall:
  '
  Dim LenTarget As Integer
  '
  If KeyCount < 1 Or KeyCode = vbKeyBack Then KeyCount = 1
  KeyCount = KeyCount - 1
  '
  If KeyCount = 0 Then
    LenTarget = Len(cmbBox)
    '
    With cmbBox
    If IsCharKeyCode(KeyCode) Then
      .DroppedDown = True
      Dim tempText As String
      '
      tempText = .Columns(iSearchCol).Text
      '
      If LCase(Left$(tempText, LenTarget)) = LCase(.Text) Then
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

