VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInputNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eNumberType
  NumberTypeByte
  NumberTypeInteger
  NumberTypeLong
  NumberTypeSingle
  NumberTypeDouble
  StringType
End Enum

Private WithEvents NumTextBox As TextBox
Attribute NumTextBox.VB_VarHelpID = -1
Private WithEvents NumCBO As ComboBox
Attribute NumCBO.VB_VarHelpID = -1
Private iNumberType As eNumberType

Private Const ModuleName As String = "CInputNumber"

Public Sub Setup(ByRef pTextbox As TextBox, ByVal NumberType As eNumberType, Optional ByVal MinValue As Variant, Optional ByVal MaxValue As Variant)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Set NumTextBox = pTextbox
      '
110   NumTextBox.SelStart = 0
120   NumTextBox.SelLength = Len(NumTextBox.Text)
      '
130   iNumberType = NumberType
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise ModuleName, "Setup", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Sub SetupCBO(ByRef pCombbox As ComboBox, ByVal NumberType As eNumberType, Optional ByVal MinValue As Variant, Optional ByVal MaxValue As Variant)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Set NumCBO = pCombbox
      '
110   NumCBO.SelStart = 0
120   NumCBO.SelLength = Len(NumCBO)
      '
130   iNumberType = NumberType
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "CInputNumber", "SetupCBO", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Function ValueByte(ByVal psNumText As String) As Byte
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   On Error Resume Next
      '
110   ValueByte = CByte(psNumText)
      '
120   If Err Then
130     ValueByte = 0
140   End If
      '<EhFooter>
      '
      Exit Function
EH:
      ErrorMgr.Raise ModuleName, "ValueByte", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Public Function ValueInt(ByVal psNumText As String) As Integer
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   On Error Resume Next
      '
110   ValueInt = CInt(psNumText)
      '
120   If Err Then
130     ValueInt = 0
140   End If
      '<EhFooter>
      '
      Exit Function
EH:
      ErrorMgr.Raise ModuleName, "ValueInt", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Public Function ValueLng(ByVal psNumText As String) As Long
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   On Error Resume Next
      '
110   ValueLng = CLng(psNumText)
      '
120   If Err Then
130     ValueLng = 0
140   End If
      '<EhFooter>
      '
      Exit Function
EH:
      ErrorMgr.Raise ModuleName, "ValueLng", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Public Function ValueSng(ByVal psNumText As String) As Single
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   On Error Resume Next
      '
110   ValueSng = CSng(psNumText)
      '
120   If Err Then
130     ValueSng = 0
140   End If
      '<EhFooter>
      '
      Exit Function
EH:
      ErrorMgr.Raise ModuleName, "ValueSng", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Public Function ValueDbl(ByVal psNumText As String) As Double
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   On Error Resume Next
      '
110   ValueDbl = CDbl(psNumText)
      '
120   If Err Then
130     ValueDbl = 0
140   End If
      '<EhFooter>
      '
      Exit Function
EH:
      ErrorMgr.Raise ModuleName, "ValueDbl", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Private Sub NumTextBox_KeyPress(KeyAscii As Integer)
      '\\ Disallow more than one decimal point
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If KeyAscii = 46 Then
110     If InStr(1, NumTextBox.Text, ".") > 0 Then
120       KeyAscii = 0
130       Exit Sub
140     End If
150   End If
      '
      '\\ General Validation
160   If KeyAscii < 48 Or KeyAscii > 57 Then            '\\ 0-9
170     If KeyAscii <> 45 Then                          '\\ -
180       If KeyAscii <> vbKeyBack Then                 '\\ Backspace
190         If KeyAscii <> vbKeyDelete Then             '\\ Delete
200           KeyAscii = 0
210         End If
220       End If
230     End If
240   Else
         '\\ Disallow more than two digits after decimal point
250     If InStr(1, NumTextBox.Text, ".") > 0 Then
260       If NumTextBox.SelStart >= InStr(1, NumTextBox.Text, ".") Then
270         If Len(NumTextBox.Text) >= InStr(1, NumTextBox.Text, ".") + 3 Then
280           KeyAscii = 0
290         End If
300       End If
310     End If
320   End If
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise ModuleName, "NumTextBox_KeyPress", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub NumTextBox_Validate(Cancel As Boolean)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   On Error GoTo NumTextBox_EH
      '
110   Dim vConverted As Variant
      '
120   On Error Resume Next
      '
130   Select Case iNumberType
        Case NumberTypeByte
140       vConverted = CByte(NumTextBox.Text)
150     Case NumberTypeInteger
160       vConverted = CInt(NumTextBox.Text)
170     Case NumberTypeLong
180       vConverted = CLng(NumTextBox.Text)
190     Case NumberTypeSingle
200       vConverted = CSng(NumTextBox.Text)
210     Case NumberTypeDouble
220       vConverted = CDbl(NumTextBox.Text)
230   End Select
      '
240   If Err.Number <> 0 Then
250     Beep
260     NumTextBox.Text = 0
270   End If
      '
280   Exit Sub
NumTextBox_EH:
290   NumTextBox.Text = 0
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise ModuleName, "NumTextBox_Validate", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub NumCBO_Validate(Cancel As Boolean)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   On Error GoTo NumTextBox_EH
      '
110   Dim vConverted As Variant
120   Dim sConverted As String
      '
130   On Error Resume Next
      '
140   Select Case iNumberType
        Case NumberTypeByte
150       vConverted = CByte(NumCBO.Text)
160     Case NumberTypeInteger
170       vConverted = CInt(NumCBO.Text)
      
180     Case NumberTypeLong
190       vConverted = CLng(NumCBO.Text)
200     Case NumberTypeSingle
210       vConverted = CSng(NumCBO.Text)
220     Case NumberTypeDouble
230       vConverted = CDbl(NumCBO.Text)
240     Case StringType
250       sConverted = CStr(NumCBO.Text)
260   End Select
      '
270   Dim i As Integer
280   Dim bool As Boolean
290   bool = False
300   For i = 0 To NumCBO.ListCount - 1
   
310     If NumCBO.List(i) = NumCBO.Text Then
320       bool = True
330     End If
340   Next
      '
350   If bool = False Then
360     If iNumberType = StringType Then
       ' NumCBO.Text = ""
370     NumCBO.ListIndex = 0
380     Else
390     NumCBO.Text = "0"
400     End If
410   End If
420   If Err.Number <> 0 Then
430     Beep
440     If iNumberType = StringType Then
        'NumCBO.Text = ""
450     NumCBO.ListIndex = 0
460     Else
470     NumCBO.Text = "0"
480     End If
490   End If
      '
500   Exit Sub
NumTextBox_EH:
510   NumCBO.Text = ""
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "CInputNumber", "NumCBO_Validate", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

