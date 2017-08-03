Option Strict Off
Option Explicit On
Module MVMail
	Declare Function sndPlaySound Lib "winmm.dll"  Alias "sndPlaySoundA"(ByVal lpszSoundName As String, ByVal uFlags As Integer) As Integer
	Public CN As New ADODB.Connection
	Public iFlasher As Short
	Public choice As Short
	Public StrUser As String
	Public sFromAddress As String
	Public sEmailAddress As String
	Public sSubject As String
	Public sBody As String
	Public sPhone As String
	Public sCaller As String
	Public sReceived As String
	Public sMessageID As Integer
	Public sMessageName As String
	Public StrGroups As String
	Public iGroupNumber As Short
	Public RefreshSpeed As Short
	Dim strDatapath As String
	Public bVMail As Boolean
	'
	'
	'
	Public iLenGroup As Double
	Public iLenMessage As Double
	Public iLenPhone As Double
	Public iLenFrom As Double
	Public iLenSubject As Double
	Public iLenDateRec As Double
	Public iLenTimeRec As Double
	Public iLenMessageNum As Double
	Public iLenUser As Double
	Public iLenCaller As Double
	Public iLenComments As Double
	Public iLenDateCom As Double
	Public iLenTimeCom As Double
	Public iFromAddress As Double
	Public bLoad As Boolean
	
	
	Public sLastName As String
	Public sFirstName As String
	Public iContact As Integer
	Public iCompany As Integer
	Public sContact As String
	'
	'
	'
	Public Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hWnd As Integer, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Integer) As Integer
	
	Public Declare Function GetDesktopWindow Lib "USER32" () As Integer
	
	Public Const SW_SHOWNORMAL As Short = 1
	
	Public Const SE_ERR_FNF As Short = 2
	Public Const SE_ERR_PNF As Short = 3
	Public Const SE_ERR_ACCESSDENIED As Short = 5
	Public Const SE_ERR_OOM As Short = 8
	Public Const SE_ERR_DLLNOTFOUND As Short = 32
	Public Const SE_ERR_SHARE As Short = 26
	Public Const SE_ERR_ASSOCINCOMPLETE As Short = 27
	Public Const SE_ERR_DDETIMEOUT As Short = 28
	Public Const SE_ERR_DDEFAIL As Short = 29
	Public Const SE_ERR_DDEBUSY As Short = 30
	Public Const SE_ERR_NOASSOC As Short = 31
	Public Const ERROR_BAD_FORMAT As Short = 11
	
	Public Const ALLCALLS As Short = 1
	Public Const NEWCALLS As Short = 2
	Public Const OLDCALLS As Short = 3
	Public OldRecordCount As Short
	Public NewRecordCount As Short
	Public FromTimer As Boolean
	Public RefreshList As Boolean
	Public SavedIndex As String
	Public strKey As String
	Public DeleteDays As Short
	
	Public Sub InitializeVmail()
		'
		Dim TSSettings As Scripting.TextStream
		Dim fso As New Scripting.FileSystemObject
		Dim rs As New ADODB.Recordset
		Dim i As Short
		'
		'RefreshSpeed = 30000
		
		Dim rsUser As New ADODB.Recordset
		'
		rsUser.Open("select * from tblEmployees", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rsUser
			Do While Not .eof
				If LCase(StrUser) = LCase(.Fields("EmployeeFirst").Value & " " & .Fields("EmployeeLast").Value) Then
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If iGroupNumber Is System.DBNull.Value Then
						.Fields("Groups").Value = 15
						.UpdateBatch()
						iGroupNumber = 15
					Else
						iGroupNumber = .Fields("Groups").Value
					End If
				End If
				.MoveNext()
			Loop 
			.Close()
		End With
		
		'  If Not SavedIndex = "" Then
		'    For i = 1 To FVMail.ListView1.ListItems.Count
		'      If FVMail.ListView1.SelectedItem.Key = SavedIndex Then
		'        Exit Sub
		'      Else
		'        With FVMail.ListView1
		'          Set .SelectedItem = .ListItems(.SelectedItem.Index + 1)
		'        End With
		'      End If
		'    Next i
		'  End If
		'
	End Sub
	
	
	Public Function GetLastUpdate() As String
		Dim rs As New ADODB.Recordset
		'
		rs.Open("SELECT * FROM TVMailSettings", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		GetLastUpdate = rs.Fields("LastUpdateTime").Value & " " & rs.Fields("LastUpdateDate").Value
		'
		rs.Close()
		'
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
	End Function
	
	Public Function GetMessageRecord(ByVal plMessageID As Integer) As ADODB.Recordset
		Dim rs As New ADODB.Recordset
		'
		Dim sQuery As String
		'
		sQuery = "SELECT [MessageID], [Group], [MessageName], [PhoneNumber], " & "[From], [Subject], " & "[DateReceived], [TimeReceived], " & "[MessageSize], [Completed], " & "[User], [Caller], " & "[Comments], [DateCompleted], " & "[TimeCompleted], " & "[FromAddress], " & "[Body], [Checked] " & "From TVMailMessages WHERE MessageID = " & plMessageID
		'
		rs.Open(sQuery, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		GetMessageRecord = rs
	End Function
	
	Public Function GetRS(ByRef ListType As Short) As ADODB.Recordset
		Dim rs As New ADODB.Recordset
		Dim qsource As String
		'
		qsource = "SELECT [MessageID], [Group], [MessageName], [PhoneNumber], " & "[From], [Subject], " & "[DateReceived], [TimeReceived], " & "[MessageSize], [Completed], " & "[User], [Caller], " & "[Comments], [DateCompleted], " & "[TimeCompleted], " & "[FromAddress], " & "[Checked] " & "From TVMailMessages "
		'    qsource = "SELECT [MessageID], [Group], [MessageName], [PhoneNumber], " & _
		''              "[From], [Subject], " & _
		''              "[DateReceived], [TimeReceived], " & _
		''              "[MessageSize], [Completed], " & _
		''              "[User], [Caller], " & _
		''              "[Comments], [DateCompleted], " & _
		''              "[TimeCompleted], " & _
		''              "[FromAddress], " & _
		''              "[Body], [Checked] " & _
		''            "From TVMailMessages "
		'
		Select Case ListType
			'Case ALLCALLS
			' qsource = qsource & "WHERE (((TMessages.Completed)=True)) "
			Case NEWCALLS
				qsource = qsource & "WHERE (((TVMailMessages.Completed)='False')) "
			Case OLDCALLS
				qsource = qsource & "WHERE (((TVMailMessages.Completed)='True')) "
			Case Else
				qsource = qsource
		End Select
		
		
		qsource = qsource & " ORDER BY TVMailMessages.DateReceived DESC , TVMailMessages.TimeReceived DESC;"
		
		rs.Open(qsource, cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		NewRecordCount = rs.RecordCount
		If NewRecordCount = OldRecordCount Then
			RefreshList = False
		Else
			RefreshList = True
			OldRecordCount = NewRecordCount
		End If
		GetRS = rs
		' rs.Close
		'Set rs = Nothing
	End Function
	
	Public Sub FillListOLD(ByRef rs As ADODB.Recordset, ByRef list As System.Windows.Forms.ListView)
		On Error GoTo EH
		'
		
		Dim LineCount As Short
		Dim FieldPos As Short
		Dim TotalCharacters As Short
		Dim Color As Object
		'
		list.Items.Clear()
		list.Columns.Clear()
		'
		TotalCharacters = 0
		LineCount = 0
		FieldPos = 0
		With rs
			If .RecordCount > 0 Then
				Do 
					TotalCharacters = 0
					.MoveFirst()
					Do 
						TotalCharacters = Len(CStr(.Fields(FieldPos).Value & vbNullString)) + TotalCharacters
						.MoveNext()
					Loop Until .eof
					list.Columns.Add("w1" & FieldPos, .Fields(FieldPos).Name, CInt(VB6.TwipsToPixelsX(400 + ((TotalCharacters / .RecordCount) * 100))))
					FieldPos = FieldPos + 1
				Loop Until FieldPos = .Fields.Count
				.MoveFirst()
				FieldPos = 0
				LineCount = 0
				Do Until .eof
					If .Fields("Completed").Value = True Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Color = &H80 'vbRed
					End If
					strKey = "r" & .Fields(FieldPos).Value
					list.Items.Add(strKey, Trim(.Fields(FieldPos).Value & vbNullString), "")
					'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					list.Items.Item(strKey).ForeColor = System.Drawing.ColorTranslator.FromOle(Color)
					FieldPos = FieldPos + 1
					Do 
						'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						list.Items.Item(strKey).SubItems.Add(.Fields(FieldPos).Value & vbNullString).ForeColor = System.Drawing.ColorTranslator.FromOle(Color)
						FieldPos = FieldPos + 1
					Loop Until FieldPos = .Fields.Count
					FieldPos = 0
					.MoveNext()
					LineCount = LineCount + 1
				Loop 
			End If
		End With
		Exit Sub
EH: 
		MsgBox(Err.Description & " in FillList.")
	End Sub
	
	Public Sub PlayTextFile(ByRef strFilename As String)
		Dim r As Integer
		Dim msg As String
		'Dim StrFileName As String
		
		r = StartDoc(My.Application.Info.DirectoryPath & "\" & strFilename)
		If r <= 32 Then
			'There was an error
			Select Case r
				Case SE_ERR_FNF
					msg = "File not found"
				Case SE_ERR_PNF
					msg = "Path not found"
				Case SE_ERR_ACCESSDENIED
					msg = "Access denied"
				Case SE_ERR_OOM
					msg = "Out of memory"
				Case SE_ERR_DLLNOTFOUND
					msg = "DLL not found"
				Case SE_ERR_SHARE
					msg = "A sharing violation occurred"
				Case SE_ERR_ASSOCINCOMPLETE
					msg = "Incomplete or invalid file association"
				Case SE_ERR_DDETIMEOUT
					msg = "DDE Time out"
				Case SE_ERR_DDEFAIL
					msg = "DDE transaction failed"
				Case SE_ERR_DDEBUSY
					msg = "DDE busy"
				Case SE_ERR_NOASSOC
					msg = "No association for file extension"
				Case ERROR_BAD_FORMAT
					msg = "Invalid EXE file or error in EXE image"
				Case Else
					msg = "Unknown error"
			End Select
			MsgBox(msg)
		End If
	End Sub
	Public Sub ClosePlayer(ByRef strFilename As String)
		Dim r As Integer
		Dim msg As String
		'
		r = CloseDoc(My.Application.Info.DirectoryPath & "\" & strFilename)
		If r <= 32 Then
			'There was an error
			Select Case r
				Case SE_ERR_FNF
					msg = "File not found"
				Case SE_ERR_PNF
					msg = "Path not found"
				Case SE_ERR_ACCESSDENIED
					msg = "Access denied"
				Case SE_ERR_OOM
					msg = "Out of memory"
				Case SE_ERR_DLLNOTFOUND
					msg = "DLL not found"
				Case SE_ERR_SHARE
					msg = "A sharing violation occurred"
				Case SE_ERR_ASSOCINCOMPLETE
					msg = "Incomplete or invalid file association"
				Case SE_ERR_DDETIMEOUT
					msg = "DDE Time out"
				Case SE_ERR_DDEFAIL
					msg = "DDE transaction failed"
				Case SE_ERR_DDEBUSY
					msg = "DDE busy"
				Case SE_ERR_NOASSOC
					msg = "No association for file extension"
				Case ERROR_BAD_FORMAT
					msg = "Invalid EXE file or error in EXE image"
				Case Else
					msg = "Unknown error"
			End Select
			MsgBox(msg)
		End If
	End Sub
	Public Sub PlaySound(ByRef strFilename As String)
		Dim r As Integer
		Dim msg As String
		'Dim StrFileName As String
		
		r = StartDoc(My.Application.Info.DirectoryPath & "\Temp\" & strFilename) '(strDatapath & "messages\" & strFilename)
		If r <= 32 Then
			'There was an error
			Select Case r
				Case SE_ERR_FNF
					msg = "File not found"
				Case SE_ERR_PNF
					msg = "Path not found"
				Case SE_ERR_ACCESSDENIED
					msg = "Access denied"
				Case SE_ERR_OOM
					msg = "Out of memory"
				Case SE_ERR_DLLNOTFOUND
					msg = "DLL not found"
				Case SE_ERR_SHARE
					msg = "A sharing violation occurred"
				Case SE_ERR_ASSOCINCOMPLETE
					msg = "Incomplete or invalid file association"
				Case SE_ERR_DDETIMEOUT
					msg = "DDE Time out"
				Case SE_ERR_DDEFAIL
					msg = "DDE transaction failed"
				Case SE_ERR_DDEBUSY
					msg = "DDE busy"
				Case SE_ERR_NOASSOC
					msg = "No association for file extension"
				Case ERROR_BAD_FORMAT
					msg = "Invalid EXE file or error in EXE image"
				Case Else
					msg = "Unknown error"
			End Select
			MsgBox(msg)
		End If
	End Sub
	
	
	Public Function FillList(ByRef rs As ADODB.Recordset, ByRef list As System.Windows.Forms.ListView) As Integer
		'On Error GoTo EH
		'
		Dim iCount As Integer
		Dim strKey As String
		Dim LineCount As Integer
		Dim FieldPos As Integer
		Dim TotalCharacters As Integer
		Dim Color As Object
		Dim pos As Object
		Dim sTemp As String
		Dim sGroupChecker1 As String
		Dim sGroupChecker2 As String
		Dim sGroupChecker3 As String
		Dim sGroupChecker4 As String
		Dim iTempGroup As Short
		'  Dim sTempKey As String
		'
		' If list.ListItems.Count > 0 Then
		'  pos = list.SelectedItem.Key
		' End If
		'  sTempKey = ""
		'  If Not list.SelectedItem Is Nothing Then
		'    sTempKey = list.SelectedItem
		'  End If
		'
		list.Visible = False
		'
		'  If Not bLoad Then
		'    iLenGroup = list.ColumnHeaders(1).Width
		'    iLenMessage = list.ColumnHeaders(2).Width
		'    iLenPhone = list.ColumnHeaders(3).Width
		'    iLenFrom = list.ColumnHeaders(4).Width
		'    iLenSubject = list.ColumnHeaders(5).Width
		'    iLenDateRec = list.ColumnHeaders(6).Width
		'    iLenTimeRec = list.ColumnHeaders(7).Width
		'    iLenMessageNum = list.ColumnHeaders(8).Width
		'    iLenUser = list.ColumnHeaders(9).Width
		'    iLenCaller = list.ColumnHeaders(10).Width
		'    iLenComments = list.ColumnHeaders(11).Width
		'    iLenDateCom = list.ColumnHeaders(12).Width
		'    iLenTimeCom = list.ColumnHeaders(13).Width
		'    iFromAddress = list.ColumnHeaders(14).Width
		'    bLoad = False
		'  End If
		'  bLoad = False
		
		
		
		
		
		
		list.Items.Clear()
		list.Columns.Clear()
		'
		TotalCharacters = 0
		LineCount = 0
		FieldPos = 1
		With rs
			If .RecordCount > 0 Then
				Do 
					TotalCharacters = 0
					.MoveFirst()
					Do 
						If FieldPos = 3 Or FieldPos = 4 And .Fields(FieldPos).Value & vbNullString = "" Then
							sTemp = "QQQQQQQ"
						Else
							sTemp = .Fields(FieldPos).Value & vbNullString
						End If
						TotalCharacters = Len(CStr(sTemp)) + TotalCharacters
						.MoveNext()
					Loop Until .eof
					'
					'If (400 + ((TotalCharacters / .RecordCount) * 100)) > 491504 Then TotalCharacters = ((491504 * .RecordCount) / 100) - 400
					'
					list.Columns.Add("w1" & FieldPos, .Fields(FieldPos).Name, CInt(VB6.TwipsToPixelsX(400 + ((TotalCharacters / .RecordCount) * 100))))
					FieldPos = FieldPos + 1
				Loop Until FieldPos = .Fields.Count - 2
				
				.MoveFirst()
				FieldPos = 1
				LineCount = 0
				'
				iTempGroup = iGroupNumber
				'
				If iTempGroup >= 8 Then
					sGroupChecker1 = "Authorizations"
					iTempGroup = iTempGroup - 8
				Else
					sGroupChecker1 = "no"
				End If
				'
				If iTempGroup >= 4 Then
					sGroupChecker2 = "Sales"
					iTempGroup = iTempGroup - 4
				Else
					sGroupChecker2 = "no"
				End If
				'
				If iTempGroup >= 2 Then
					sGroupChecker3 = "Support"
					iTempGroup = iTempGroup - 2
				Else
					sGroupChecker3 = "no"
				End If
				'
				If iTempGroup >= 1 Then
					sGroupChecker4 = "Operator"
				Else
					sGroupChecker4 = "no"
				End If
				
				'For iCount = 1 To rs.RecordCount
				Do Until .eof
					
					If (.Fields("Group").Value = sGroupChecker1) Or (.Fields("Group").Value = sGroupChecker2) Or (.Fields("Group").Value = sGroupChecker3) Or (.Fields("Group").Value = sGroupChecker4) Or (iGroupNumber = 15) Then
						' If Right$(.Fields(0), 3) = "WAV" Then
						If .Fields("Completed").Value = True Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
						Else
							If Right(.Fields(2).Value & "", 3) = "WAV" Or Right(.Fields(2).Value & "", 3) = "wav" Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Color = &H80 'vbRed
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Color = &H8000000D
							End If
							
						End If
						strKey = "r" & .Fields(0).Value
						list.Items.Add(strKey, .Fields(FieldPos).Value & vbNullString, "")
						'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						list.Items.Item(strKey).ForeColor = System.Drawing.ColorTranslator.FromOle(Color)
						FieldPos = FieldPos + 1
						Do 
							'UPGRADE_WARNING: Couldn't resolve default property of object Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							list.Items.Item(strKey).SubItems.Add(.Fields(FieldPos).Value & vbNullString).ForeColor = System.Drawing.ColorTranslator.FromOle(Color)
							FieldPos = FieldPos + 1
						Loop Until FieldPos = .Fields.Count - 2
						'------------------
						'  Debug.Print !Checked
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDbNull(.Fields("Checked").Value) Then
							' list.ListItems.Item.Checked = False
							list.Items.Item(strKey).Checked = False
						Else
							If .Fields("Checked").Value = True Then
								'  Debug.Print !MessageID
								
								' list.ListItems.Item.Checked = True
								list.Items.Item(strKey).Checked = True
							Else
								If .Fields("Checked").Value = False Then
									'Debug.Print !MessageID
									' list.ListItems.Item.Checked = False
									list.Items.Item(strKey).Checked = False
								End If
							End If
						End If
						'------------------
						FieldPos = 1
						'.MoveNext
						LineCount = LineCount + 1
						' Else
						' .MoveNext
						'End If
						
						'Next
					End If
					.MoveNext()
				Loop 
			End If
		End With
		FillList = LineCount
		'
		' If list.ListItems.Count > 0 Then
		'   pos = list.SelectedItem.Key
		' End If
		'list.SelectedItem.Key = pos
		'If sTempKey <> "" Then
		'  list.SelectedItem = sTempKey
		'End If
		
		'
		If list.Items.Count = 0 Then
			With FVMail
				.txtBody.Text = ""
				.txtPhone.Text = ""
				.txtsubject.Text = ""
				.chkComp.CheckState = System.Windows.Forms.CheckState.Unchecked
				.cmbCaller.Items.Clear()
				'.cmbCaller.Index = 0
				.cmbComment.Text = ""
				.cmdGetNames.Enabled = False
				.cmdContactInfo.Enabled = False
			End With
		End If
		'If Not iLenGroup = 0 Then
		'  list.ColumnHeaders(1).Width = iLenGroup
		'  list.ColumnHeaders(2).Width = iLenMessage
		'  list.ColumnHeaders(3).Width = iLenPhone
		'  list.ColumnHeaders(4).Width = iLenFrom
		'  list.ColumnHeaders(5).Width = iLenSubject
		'  list.ColumnHeaders(6).Width = iLenDateRec
		'  list.ColumnHeaders(7).Width = iLenTimeRec
		'  list.ColumnHeaders(8).Width = iLenMessageNum
		'  list.ColumnHeaders(9).Width = iLenUser
		'  list.ColumnHeaders(10).Width = iLenCaller
		'  list.ColumnHeaders(11).Width = iLenComments
		'  list.ColumnHeaders(12).Width = iLenDateCom
		'  list.ColumnHeaders(13).Width = iLenTimeCom
		'  list.ColumnHeaders(14).Width = iFromAddress
		'End If
		
		
		
		
		
		
		
		'list.ColumnHeaders(9).Width = 0
		'list.ColumnHeaders(1).Width = 800
		'list.ColumnHeaders(2).Width = 600
		'list.ColumnHeaders(10).Width = 100
		'list.ColumnHeaders(11).Width = 2000
		'list.ColumnHeaders(12).Width = 1200
		list.Visible = True
		Exit Function
		'EH:
		'MsgBox Err.Description & " in FillList."
	End Function
	
	Function StartDoc(ByRef DocName As String) As Integer
		Dim Scr_hDC As Integer
		Scr_hDC = GetDesktopWindow()
		StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
	End Function
	
	Function CloseDoc(ByRef DocName As String) As Integer
		Dim Scr_hDC As Integer
		Scr_hDC = GetDesktopWindow()
		CloseDoc = ShellExecute(Scr_hDC, "Close", DocName, "", "C:\", SW_SHOWNORMAL)
	End Function
	
	'Public Function AddListItem(pRecordSet As Recordset, pList As ListView, PIndex As Long) As Long
	'  On Error GoTo EH
	'  '
	'  Dim StrKey As String
	'  Dim LineCount As Integer
	'  Dim FieldPos As Integer
	'  Dim TotalCharacters As Integer
	'  Dim color As Variant
	'  Dim pos As Variant
	'  '
	'  FieldPos = 1
	'  With pRecordSet
	'      .MoveFirst
	'      FieldPos = 1
	'      LineCount = 0
	'      Do Until .EOF
	'
	'        StrKey = "r" & .Fields(0)
	'        If Not CheckForKey(StrKey, pList) Then
	'          pList.ListItems.Add , StrKey, .Fields(FieldPos) & vbNullString
	'          pList.ListItems.Item(StrKey).ForeColor = color
	'          FieldPos = FieldPos + 1
	'          Do
	'            pList.ListItems.Item(StrKey).ListSubItems.Add(, , .Fields(FieldPos) & vbNullString).ForeColor = color
	'            FieldPos = FieldPos + 1
	'          Loop Until FieldPos = .Fields.Count
	'          FieldPos = 1
	'
	'          LineCount = LineCount + 1
	'        End If
	'        .MoveNext
	'      Loop
	'   ' End If
	'  End With
	'AddListItem = LineCount
	'
	'' If list.ListItems.Count > 0 Then
	' '   pos = list.SelectedItem.Key
	'' End If
	''list.SelectedItem.Key = pos
	'Exit Function
	'EH:
	' MsgBox Err.Description & " in FillList."
	'End Function
	
	Public Function CheckForKey(ByRef pKey As String, ByRef pList As System.Windows.Forms.ListView) As Boolean
		Dim i As Short
		CheckForKey = False
		For i = 1 To pList.Items.Count
			'UPGRADE_WARNING: Lower bound of collection pList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If pKey = pList.Items.Item(i).Name Then
				CheckForKey = True
			End If
			' pList.ListItems.Item(i).Key
		Next i
	End Function
	
	'Public Sub FindNames()
	'Dim rs As New ADODB.Recordset
	'
	'rs.Open "SELECT * FROM TContact", cnMain, adOpenKeyset, adLockBatchOptimistic
	
	'With rs
	'If Not .EOF Then
	
	
	
	'End Sub
	Public Sub GetContactInfo()
		Dim rsContact As New ADODB.Recordset
		'
		GetFirstName()
		rsContact.Open("Select [FirstName], [LastName], [ID],[ComPanyID] from TContact Where [FirstName] = '" & sFirstName & "'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		With rsContact
			While Not .eof
				If sContact = .Fields("FirstName").Value & " " & .Fields("LastName").Value Then
					iContact = .Fields("ID").Value
					iCompany = .Fields("CompanyID").Value
				End If
				.MoveNext()
			End While
		End With
		'
	End Sub
	
	Public Sub GetFirstName()
		Dim i As Short
		Dim sLetter As String
		'
		sFirstName = ""
		i = 1
		sLetter = Mid(sContact, i, 1)
		While Not sLetter = " "
			sFirstName = sFirstName & sLetter
			i = i + 1
			sLetter = Mid(sContact, i, 1)
		End While
		If sFirstName = "Dr." Then
			sFirstName = sFirstName & " "
			GetRestOfName()
		End If
		'
	End Sub
	
	Public Sub GetRestOfName()
		Dim i As Short
		Dim sLetter As String
		'
		i = 5
		sLetter = Mid(sContact, i, 1)
		While Not sLetter = " "
			sFirstName = sFirstName & sLetter
			i = i + 1
			sLetter = Mid(sContact, i, 1)
		End While
		'
	End Sub
End Module