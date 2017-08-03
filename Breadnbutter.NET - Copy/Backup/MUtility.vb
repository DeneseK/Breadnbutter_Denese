Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module MUtility
	
	Public Enum eFileOpenType
		foOpen
		foSave
	End Enum
	'user defined type required by Shell_NotifyIcon API call
	Public Structure NOTIFYICONDATA
		Dim cbSize As Integer
		Dim hWnd As Integer
		Dim uID As Integer
		Dim uFlags As Integer
		Dim uCallbackMessage As Integer
		Dim hIcon As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(64),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=64)> Public szTip() As Char
	End Structure
	
	'constants required by Shell_NotifyIcon API call:
	Public Const NIM_ADD As Short = &H0s
	Public Const NIM_MODIFY As Short = &H1s
	Public Const NIM_DELETE As Short = &H2s
	Public Const NIF_MESSAGE As Short = &H1s
	Public Const NIF_ICON As Short = &H2s
	Public Const NIF_TIP As Short = &H4s
	Public Const WM_MOUSEMOVE As Short = &H200s
	Public Const WM_LBUTTONDOWN As Short = &H201s 'Button down
	Public Const WM_LBUTTONUP As Short = &H202s 'Button up
	Public Const WM_LBUTTONDBLCLK As Short = &H203s 'Double-click
	Public Const WM_RBUTTONDOWN As Short = &H204s 'Button down
	Public Const WM_RBUTTONUP As Short = &H205s 'Button up
	Public Const WM_RBUTTONDBLCLK As Short = &H206s 'Double-click
	
	Public Declare Function SetForegroundWindow Lib "USER32" (ByVal hWnd As Integer) As Integer
	'UPGRADE_WARNING: Structure NOTIFYICONDATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function Shell_NotifyIcon Lib "shell32"  Alias "Shell_NotifyIconA"(ByVal dwMessage As Integer, ByRef pnid As NOTIFYICONDATA) As Boolean
	
	Public nid As NOTIFYICONDATA
	'for disabling X
	Private Declare Function GetSystemMenu Lib "USER32" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
	
	Private Declare Function RemoveMenu Lib "USER32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
	
	Private Const MF_BYPOSITION As Integer = &H400
	
	Public Sub RemoveCancelMenuItem(ByRef frm As System.Windows.Forms.Form)
		Dim hSysMenu As Integer
		
		'get the system menu for this form
		hSysMenu = GetSystemMenu(frm.Handle.ToInt32, 0)
		
		'remove the close item
		Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
		
		'remove the separator that was over the close item
		Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
	End Sub
	
	Public Function FullPath(ByRef psPath As String) As String
		On Error GoTo ErrCall
		'
		If Right(psPath, 1) <> "\" Then
			FullPath = psPath & "\"
		Else
			FullPath = psPath
		End If
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.FullPath", MsgBoxStyle.Critical, "Error")
	End Function
	
	Public Function DecryptStr(ByRef psTarget As String, Optional ByRef psKey As String = "", Optional ByRef pbCase As Boolean = False) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN, iC As Short
		Dim sBfr As String
		'
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(psKey) Or psKey = vbNullString Then psKey = "HRPass"
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(pbCase) Or Not pbCase Then psKey = UCase(psKey)
		'
		For liN = 1 To Len(psTarget)
			iC = Asc(Mid(psTarget, liN, 1))
			iC = iC - Asc(Mid(psKey, (liN Mod Len(psKey)) + 1, 1))
			sBfr = sBfr & Chr(iC And &HFFs)
		Next 
		'
		DecryptStr = sBfr
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.fsDecrypt.", MsgBoxStyle.Critical, "Error")
	End Function
	
	Public Function EncryptStr(ByRef psTarget As String, Optional ByRef psKey As String = "", Optional ByRef pbCase As Boolean = False) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN, iC As Short
		Dim sBfr As String
		'
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(psKey) Or psKey = vbNullString Then psKey = "HRPass"
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(pbCase) Or Not pbCase Then psKey = UCase(psKey)
		'
		For liN = 1 To Len(psTarget)
			iC = Asc(Mid(psTarget, liN, 1))
			iC = iC + Asc(Mid(psKey, (liN Mod Len(psKey)) + 1, 1))
			sBfr = sBfr & Chr(iC And &HFFs)
		Next 
		'
		EncryptStr = sBfr
		'
		' CSErrorHandler begin - please do not modify or remove this line
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.EncryptStr.", MsgBoxStyle.Critical, "Error")
	End Function
	
	Public Sub ResizeGrid(ByRef pGrd As AxSSDataWidgets_B_OLEDB.AxSSOleDBGrid, ByRef psngHeight As Single, ByRef psngWidth As Single, Optional ByRef piColStart As Short = 0, Optional ByRef piColEnd As Short = 0)
		On Error GoTo ErrCall
		'
		Dim sngColRatios() As Single
		Dim sngOldWidth As Single
		Dim sngTempWidth As Single
		Dim i As Short
		Dim iColStart, iColEnd As Short
		'
		If piColEnd = 0 Then iColEnd = pGrd.Cols - 1
		'
		pGrd.Redraw = False
		'
		'UPGRADE_WARNING: Lower bound of array sngColRatios was changed from iColStart to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim sngColRatios(iColEnd)
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
		pGrd.Width = VB6.TwipsToPixelsX(psngWidth)
		pGrd.Height = VB6.TwipsToPixelsY(psngHeight)
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
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.StretchGrid.", MsgBoxStyle.Critical, "Error")
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
	
	Public Function nnNum(ByRef vVar As Object) As Object
		On Error GoTo ErrCall
		'
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If VarType(vVar) = VariantType.Null Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nnNum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nnNum = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object vVar. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If vVar = "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nnNum = 0
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object vVar. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nnNum = vVar
			End If
		End If
		'
		Exit Function
ErrCall: 
		MsgBox(Err.Description)
	End Function
	
	Public Function GetFileName(ByRef psPath As String, ByRef psFile As String, Optional ByRef peType As eFileOpenType = 0, Optional ByRef psFlags As String = "", Optional ByRef vOwner As Object = Nothing, Optional ByRef psFilter As String = "", Optional ByRef piFilterIndex As Short = 0, Optional ByRef psInitDir As String = "", Optional ByRef psTitle As String = "", Optional ByRef psDefExt As String = "") As Boolean
		
		On Error GoTo ErrCall
		'
		Dim dlgMain As New CDlgComp.CCommonDialog
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
			.FLAGS = IIf(psFlags = "", CDlgComp.ccdControlContants.cdlOFNHideReadOnly + CDlgComp.ccdControlContants.cdlOFNFileMustExist, psFlags)
			'
			Select Case peType
				Case 0
					.ShowOpen()
				Case 1
					.ShowSave()
			End Select
			'
			If .FileName = "" Then
				MsgBox("File not located.")
				GetFileName = False
			Else
				FileOps.SplitPathFile(.FileName, psPath, psFile)
				GetFileName = True
			End If
		End With
		'
		Exit Function
ErrCall: 
		If Err.Number = -2147219503 Then 'User cancelled
			GetFileName = False
		Else
			MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.GetOpenFileName.", MsgBoxStyle.Critical, "Error")
		End If
	End Function
	
	Public Function IsCharKeyCode(ByRef pKeyCode As Short) As Boolean
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
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in MUtility.IsCharKeyCode.", MsgBoxStyle.Critical, "Error")
	End Function
	
	Public Function NextID(ByRef psFieldName As String, ByRef psTableName As String, ByRef pCN As ADODB.Connection) As Integer
		On Error GoTo EH
		'
		Dim rsID As New ADODB.Recordset
		Dim rsMax As New ADODB.Recordset
		Dim IDTemp As Integer
		'
		rsID.Open("SELECT * FROM IDMAX WHERE TableName = '" & psTableName & "'", pCN, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		'
		If rsID.eof Then
			rsID.AddNew()
			rsID.Fields("TableName").Value = psTableName
			rsID.Fields("MaxID").Value = 1
			rsID.Update()
			IDTemp = 1
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			IDTemp = nnNum(rsID.Fields("MaxID")) + 1
			If IDTemp = 0 Then IDTemp = 1
		End If
		'
		rsMax.Open("Select MAX(" & psFieldName & ") AS FieldMax FROM " & psTableName, pCN, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		If rsMax.Fields("fieldmax").Value + 1 > IDTemp Then
			IDTemp = rsMax.Fields("fieldmax").Value + 1
		End If
		'
		rsMax.Close()
		'UPGRADE_NOTE: Object rsMax may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsMax = Nothing
		'
		rsID.Fields("MaxID").Value = IDTemp
		rsID.Update()
		'
		rsID.Close()
		'UPGRADE_NOTE: Object rsID may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsID = Nothing
		'
		NextID = IDTemp
		'
		Exit Function
EH: 
		MsgBox(Err.Description)
	End Function
	
	Public Sub KillTime(ByRef sngSeconds As Single)
		On Error Resume Next
		'
		Dim sngStart As Single
		sngStart = VB.Timer()
		Do While (VB.Timer() - sngStart) < sngSeconds
			System.Windows.Forms.Application.DoEvents()
		Loop 
	End Sub
	
	Public Sub SelectText(ByRef pctrlCur As System.Windows.Forms.Control)
		On Error GoTo ErrorHandler
		'
		'\\ Local Declarations
		Dim fAlt As Boolean
		'
		With pctrlCur
			'UPGRADE_WARNING: Couldn't resolve default property of object pctrlCur.SelStart. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.SelStart = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object pctrlCur.SelLength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object pctrlCur.DisplayText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.SelLength = Len(pctrlCur.DisplayText)
			'UPGRADE_WARNING: Couldn't resolve default property of object pctrlCur.SelLength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.General.SelectText")
	End Sub
	
	Public Function GetIDFromKey(ByRef psKey As String) As Integer
		If Len(psKey) > 1 Then
			GetIDFromKey = CInt(Right(psKey, Len(psKey) - 1))
		End If
	End Function
	
	Public Function CalculatePendingDays(ByRef pdSaleDate As Date, ByRef plGraceDays As Integer, ByRef plSaleDays As Integer) As Integer
		Dim lTempPending As Integer
		Dim lDaysPassed As Integer
		'
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		lDaysPassed = System.Math.Abs(DateDiff(Microsoft.VisualBasic.DateInterval.DayOfYear, pdSaleDate, Now))
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
		Dim i As Integer
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
			If InStr("xX#", Mid(pText, i, 1)) <> 0 Then
				sExt = Right(pText, Len(pText) - i)
				pText = Left(pText, i - 1)
				Exit For
			End If
		Next 
		' get rid of dashes and invalid chars in Ext
		For i = Len(sExt) To 1 Step -1
			If InStr("0123456789", Mid(sExt, i, 1)) = 0 Then
				sExt = Left(sExt, i - 1) & Mid(sExt, i + 1)
			End If
		Next 
		' get rid of dashes and invalid chars
		For i = Len(pText) To 1 Step -1
			If InStr("0123456789", Mid(pText, i, 1)) = 0 Then
				pText = Left(pText, i - 1) & Mid(pText, i + 1)
			End If
		Next 
		'look for proper length, bad, international numbers
		If Len(pText) > 11 Or Len(pText) < 7 Then
			
		Else
			' then, re-insert them in the correct position
			If Len(pText) <= 7 Then
				FormatPhoneNumber = VB6.Format(pText, "!@@@-@@@@")
			Else
				FormatPhoneNumber = VB6.Format(pText, "!(@@@) @@@-@@@@")
			End If
			If sExt <> "" Then
				FormatPhoneNumber = FormatPhoneNumber & " Ext. " & sExt
			End If
		End If
	End Function
	
	Public Function StripChars(ByVal pText As String) As String
		Dim i As Short
		For i = Len(pText) To 1 Step -1
			If InStr("0123456789", Mid(pText, i, 1)) = 0 Then
				pText = Left(pText, i - 1) & Mid(pText, i + 1)
			End If
		Next 
		StripChars = pText
	End Function
End Module