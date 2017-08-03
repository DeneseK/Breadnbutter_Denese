Attribute VB_Name = "MUtility"
Option Explicit

Public Enum eFileOpenType
  foOpen
  foSave
End Enum
'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
 cbSize As Long
 hWnd As Long
 uID As Long
 uFlags As Long
 uCallbackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Declare Function SetForegroundWindow Lib "USER32" _
(ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public nid As NOTIFYICONDATA
'for disabling X
Private Declare Function GetSystemMenu Lib "USER32" _
            (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "USER32" _
            (ByVal hMenu As Long, ByVal nPosition As Long, _
             ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&

Public Sub RemoveCancelMenuItem(frm As Form)
        Dim hSysMenu As Long

        'get the system menu for this form
        hSysMenu = GetSystemMenu(frm.hWnd, 0)

        'remove the close item
        Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)

        'remove the separator that was over the close item
        Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub

Public Function FullPath(psPath As String) As String
  On Error GoTo ErrCall
  '
  If Right$(psPath, 1) <> "\" Then
    FullPath = psPath & "\"
  Else
    FullPath = psPath
  End If
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.FullPath", vbCritical, "Error"
End Function

Public Function DecryptStr(psTarget As String, Optional psKey As String, Optional pbCase As Boolean) As String
  On Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer, iC As Integer
  Dim sBfr As String
  '
  If IsMissing(psKey) Or psKey = vbNullString Then psKey = "HRPass"
  If IsMissing(pbCase) Or Not pbCase Then psKey = UCase$(psKey)
  '
  For liN = 1 To Len(psTarget)
    iC = Asc(Mid$(psTarget, liN, 1))
    iC = iC - Asc(Mid$(psKey, (liN Mod Len(psKey)) + 1, 1))
    sBfr = sBfr & Chr$(iC And &HFF)
  Next
  '
  DecryptStr = sBfr
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.fsDecrypt.", vbCritical, "Error"
End Function

Public Function EncryptStr(psTarget As String, Optional psKey As String, Optional pbCase As Boolean) As String
  On Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer, iC As Integer
  Dim sBfr As String
  '
  If IsMissing(psKey) Or psKey = vbNullString Then psKey = "HRPass"
  If IsMissing(pbCase) Or Not pbCase Then psKey = UCase$(psKey)
  '
  For liN = 1 To Len(psTarget)
    iC = Asc(Mid$(psTarget, liN, 1))
    iC = iC + Asc(Mid$(psKey, (liN Mod Len(psKey)) + 1, 1))
    sBfr = sBfr & Chr$(iC And &HFF)
  Next
  '
  EncryptStr = sBfr
  '
  ' CSErrorHandler begin - please do not modify or remove this line
  Exit Function
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.EncryptStr.", vbCritical, "Error"
End Function

Public Sub ResizeGrid(pGrd As SSOleDBGrid, psngHeight As Single, psngWidth As Single, Optional piColStart As Integer, Optional piColEnd As Integer)
  On Error GoTo ErrCall
  '
  Dim sngColRatios() As Single
  Dim sngOldWidth As Single
  Dim sngTempWidth As Single
  Dim i As Integer
  Dim iColStart As Integer, iColEnd As Integer
  '
  If piColEnd = 0 Then iColEnd = pGrd.Cols - 1
  '
  pGrd.Redraw = False
  '
  ReDim sngColRatios(iColStart To iColEnd)
  '
  For i = iColStart To iColEnd
    If pGrd.Columns(i).Visible Then sngOldWidth = sngOldWidth + pGrd.Columns(i).Width
  Next i
  '
  For i = iColStart To iColEnd
    sngColRatios(i) = pGrd.Columns(i).Width / sngOldWidth
  Next i
  '
  sngTempWidth = psngWidth - 570
  '
  pGrd.Width = psngWidth
  pGrd.Height = psngHeight
  '
  For i = iColStart To iColEnd
    pGrd.Columns(i).Width = sngTempWidth * sngColRatios(i)
  Next i
  '
  pGrd.Redraw = True
  '
  Exit Sub
ErrCall:
  pGrd.Redraw = True
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.StretchGrid.", vbCritical, "Error"
End Sub

'Public Sub FillAiList(pCombo As SSDBCombo, sTable As String, sOrder As String)
'
'  Dim rsSearch As Recordset
'  Dim sFields() As String
'  Dim i As Integer, iCols As Integer
'  Dim sColumns As String
'  '
'  iCols = pCombo.Cols
'  ReDim sFields(iCols)
'  '
'  sColumns = ""
'  For i = 0 To iCols - 1
'    sFields(i) = pCombo.Columns(i).Name
'    '
'    If i = 0 Then
'      sColumns = sFields(i)
'    Else
'      sColumns = sColumns & ", " & sFields(i)
'    End If
'  Next i
'  '
'  Set rsSearch = dbMain.OpenRecordset("SELECT " & sColumns & " FROM " & sTable & " ORDER BY " & sOrder, dbOpenDynaset)
'  rsSearch.MoveFirst
'  '
'  pCombo.Redraw = False
'  pCombo.RemoveAll
'  '
'  With rsSearch
'  Do While Not .EOF
'    sColumns = ""
'    For i = 0 To iCols - 1
'      If sColumns = "" Then
'        sColumns = .Fields(sFields(i))
'      Else
'        sColumns = sColumns & ";" & .Fields(sFields(i))
'      End If
'    Next i
'    pCombo.AddItem sColumns
'    .MoveNext
'  Loop
'  End With
'  '
'  pCombo.Redraw = True
'End Sub

Public Function nnNum(vVar As Variant) As Variant
  On Error GoTo ErrCall
  '
  If VarType(vVar) = vbNull Then
    nnNum = 0
  Else
    If vVar = "" Then
      nnNum = 0
    Else
      nnNum = vVar
    End If
  End If
  '
  Exit Function
ErrCall:
  MsgBox Err.Description
End Function

Public Function GetFileName(ByRef psPath As String, ByRef psFile As String, _
                           Optional peType As eFileOpenType, _
                           Optional psFlags As String, _
                           Optional vOwner As Variant, _
                           Optional psFilter As String, _
                           Optional piFilterIndex As Integer, _
                           Optional psInitDir As String, _
                           Optional psTitle As String, _
                           Optional psDefExt As String) As Boolean
  
  On Error GoTo ErrCall
  '
  Dim dlgMain As New CCommonDialog
  Dim sExt As String
  '
  sExt = psDefExt
  '
  If psFilter = "" Then
    psFilter = "Database Files(" & sExt & ")|" & sExt
  Else
    If piFilterIndex = 0 Then piFilterIndex = 1
  End If
  '
  With dlgMain
  .DialogTitle = psTitle
  .DefaultExt = sExt
  .Filter = psFilter
  .FilterIndex = piFilterIndex
  .FileTitle = psFile
  .FileName = psFile
  .InitDir = psInitDir
  .CancelError = True
  .FLAGS = IIf(psFlags = "", cdlOFNHideReadOnly + cdlOFNFileMustExist, psFlags)
  '
  Select Case peType
  Case 0
    .ShowOpen
  Case 1
    .ShowSave
  End Select
  '
  If .FileName = "" Then
    MsgBox "File not located."
    GetFileName = False
  Else
    FileOps.SplitPathFile .FileName, psPath, psFile
    GetFileName = True
  End If
  End With
  '
  Exit Function
ErrCall:
  If Err.Number = -2147219503 Then 'User cancelled
    GetFileName = False
  Else
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.GetOpenFileName.", vbCritical, "Error"
  End If
End Function

Public Function IsCharKeyCode(pKeyCode As Integer) As Boolean
  On Error GoTo ErrCall
  '
  Dim booTemp As Boolean
  '
  booTemp = False
  Select Case pKeyCode
  Case 32, 48 To 57, 65 To 90, 96 To 111, 186 To 192, 219 To 222
    If pKeyCode <> 108 Then booTemp = True
  End Select
  IsCharKeyCode = booTemp
  
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
  ' CSErrorHandler begin - please do not modify or remove this line
  Exit Function
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.IsCharKeyCode.", vbCritical, "Error"
End Function

Public Function NextID(psFieldName As String, psTableName As String, pCN As ADODB.Connection) As Long
  On Error GoTo EH
  '
  Dim rsID As New ADODB.Recordset
  Dim rsMax As New ADODB.Recordset
  Dim IDTemp As Long
  '
  rsID.Open "SELECT * FROM IDMAX WHERE TableName = '" & psTableName & "'", pCN, adOpenKeyset, adLockOptimistic, adCmdText
  '
  If rsID.eof Then
    rsID.AddNew
    rsID!TableName = psTableName
    rsID!MaxID = 1
    rsID.Update
    IDTemp = 1
  Else
    IDTemp = nnNum(rsID!MaxID) + 1
    If IDTemp = 0 Then IDTemp = 1
  End If
  '
  rsMax.Open "Select MAX(" & psFieldName & ") AS FieldMax FROM " & psTableName, pCN, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  If rsMax!fieldmax + 1 > IDTemp Then
    IDTemp = rsMax!fieldmax + 1
  End If
  '
  rsMax.Close
  Set rsMax = Nothing
  '
  rsID!MaxID = IDTemp
  rsID.Update
  '
  rsID.Close
  Set rsID = Nothing
  '
  NextID = IDTemp
  '
  Exit Function
EH:
  MsgBox Err.Description
End Function

Public Sub KillTime(sngSeconds As Single)
  On Error Resume Next
  '
  Dim sngStart As Single
  sngStart = Timer
  Do While (Timer - sngStart) < sngSeconds
    DoEvents
  Loop
End Sub

Public Sub SelectText(pctrlCur As Control)
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim fAlt As Boolean
  '
  With pctrlCur
    .SelStart = 0
    .SelLength = Len(pctrlCur.DisplayText)
    If fAlt = True Then .SelLength = Len(pctrlCur)
  End With
  '
  Exit Sub
  '
ErrorHandler:
  If Err.Number = 438 Then '\\ Object Doesn't Support Property Or Method
    fAlt = True
    Resume Next
  End If
  '
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.General.SelectText"
End Sub

Public Function GetIDFromKey(psKey As String) As Long
  If Len(psKey) > 1 Then
    GetIDFromKey = Right(psKey, Len(psKey) - 1)
  End If
End Function

Public Function CalculatePendingDays(pdSaleDate As Date, plGraceDays As Long, plSaleDays As Long) As Long
  Dim lTempPending As Long
  Dim lDaysPassed As Long
  '
  lDaysPassed = Abs(DateDiff("y", pdSaleDate, Now))
  '
  If plGraceDays < 0 Then
    CalculatePendingDays = plSaleDays
  Else
    If lDaysPassed < plGraceDays Then
      CalculatePendingDays = plSaleDays
    Else
      lTempPending = plSaleDays - (lDaysPassed - plGraceDays)
      '
      If lTempPending >= 0 Then
        CalculatePendingDays = lTempPending
      Else
        CalculatePendingDays = 0
      End If
    End If
  End If
End Function

Public Function FormatPhoneNumber(ByVal pText As String) As String
  ' Modify a phone-number to the format "XXX-XXXX" or "(XXX) XXX-XXXX".
  Dim i As Long
  Dim sExt As String
  '
  pText = StripChars(pText)
  FormatPhoneNumber = pText
  '
  'setup for old all number format
  If IsNumeric(pText) Then
    'pText = "1234567890123"
    If Len(pText) > 10 Then
      pText = Left(pText, 10) & "x" & Right(pText, Len(pText) - 10)
    End If
    'pText = Format$(pText, "!@@@-@@@-@@@@x")
  End If
  '
  ' ignore empty strings
  If Len(pText) = 0 Then Exit Function
  'Look for extension x, X or #
  For i = Len(pText) To 1 Step -1
      If InStr("xX#", Mid$(pText, i, 1)) <> 0 Then
          sExt = Right$(pText, Len(pText) - i)
          pText = Left$(pText, i - 1)
          Exit For
      End If
  Next
  ' get rid of dashes and invalid chars in Ext
  For i = Len(sExt) To 1 Step -1
      If InStr("0123456789", Mid$(sExt, i, 1)) = 0 Then
          sExt = Left$(sExt, i - 1) & Mid$(sExt, i + 1)
      End If
  Next
  ' get rid of dashes and invalid chars
  For i = Len(pText) To 1 Step -1
      If InStr("0123456789", Mid$(pText, i, 1)) = 0 Then
          pText = Left$(pText, i - 1) & Mid$(pText, i + 1)
      End If
  Next
  'look for proper length, bad, international numbers
  If Len(pText) > 11 Or Len(pText) < 7 Then
  
  Else
    ' then, re-insert them in the correct position
    If Len(pText) <= 7 Then
        FormatPhoneNumber = Format$(pText, "!@@@-@@@@")
    Else
        FormatPhoneNumber = Format$(pText, "!(@@@) @@@-@@@@")
    End If
    If sExt <> "" Then
      FormatPhoneNumber = FormatPhoneNumber & " Ext. " & sExt
    End If
  End If
End Function

Public Function StripChars(ByVal pText As String) As String
  Dim i As Integer
  For i = Len(pText) To 1 Step -1
      If InStr("0123456789", Mid$(pText, i, 1)) = 0 Then
          pText = Left$(pText, i - 1) & Mid$(pText, i + 1)
      End If
  Next
  StripChars = pText
End Function
