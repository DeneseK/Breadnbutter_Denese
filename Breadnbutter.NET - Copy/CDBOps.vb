Option Strict Off
Option Explicit On
Friend Class CDBOps
	
	Private FileOps As New CFileOps
	
	Public Enum eReturnPos
		eFirst
		eLast
	End Enum
	
	Public Enum eDelete
		eAll
		eUseSQL
	End Enum
	
	Private sCNDBName As String
	
	Public ReadOnly Property DBName() As String
		Get
			DBName = sCNDBName
		End Get
	End Property
	
	'Public Function TableExists(pDB As Database, psTableName As String) As Boolean
	'  On Error Resume Next
	'  '
	'  Dim tbl As TableDef
	'  '
	'  Set tbl = pDB.TableDefs(psTableName)
	'  If Err.Number = 3265 Then
	'    TableExists = False
	'  ElseIf Err.Number Then
	'    TableExists = False
	'    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.TableExists", vbCritical, "Error"
	'  Else
	'    TableExists = True
	'  End If
	'  On Error GoTo 0
	'End Function
	
	Public Function HasRecords(ByRef pRS As ADODB.Recordset, ByRef peReturnPos As eReturnPos) As Boolean
		'\\ Checks if a recordset has records
		'
		On Error GoTo ErrCall
		'
		Dim bHasRecords As Boolean
		'
		With pRS
			If Not .BOF Then
				.MoveFirst()
				'
				bHasRecords = Not .EOF
			Else
				If Not .EOF Then
					.MoveNext()
					'
					bHasRecords = Not .EOF
				Else
					HasRecords = False
				End If
			End If
			'
			If bHasRecords Then
				If peReturnPos = eReturnPos.eFirst Then
					.MoveFirst()
				Else
					.MoveLast()
				End If
			End If
		End With
		'
		HasRecords = bHasRecords
		'
		Exit Function
ErrCall: 
		HasRecords = False
	End Function
	
	Public Sub ZapRS(ByRef pRS As ADODB.Recordset)
		On Error Resume Next
		pRS.Close()
		'UPGRADE_NOTE: Object pRS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRS = Nothing
	End Sub
	
	'Public Sub ZapDB(pDB As Database)
	'  On Error Resume Next
	'  pDB.Close
	'  Set pDB = Nothing
	'End Sub
	
	'Public Function FieldExists(pDB As Database, psTableName As String, psFieldName) As Boolean
	'  On Error Resume Next
	'  '
	'  Dim fld As DAO.Field
	'  '
	'  If Me.TableExists(pDB, psTableName) Then
	'    Set fld = pDB.TableDefs(psTableName).Fields(psFieldName)
	'    If Err.Number = 3265 Then
	'      FieldExists = False
	'    ElseIf Err.Number Then
	'      FieldExists = False
	'      MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.FieldExists", vbCritical, "Error"
	'    Else
	'      FieldExists = True
	'    End If
	'  Else
	'    FieldExists = False
	'  End If
	'End Function
	
	Public Sub DeleteRecords(ByRef pRS As ADODB.Recordset)
		On Error GoTo ErrCall
		'
		If Not pRS Is Nothing Then
			If Me.HasRecords(pRS, eReturnPos.eLast) Then
				Do While Not pRS.BOF
					pRS.Delete()
					pRS.MovePrevious()
				Loop 
			End If
		End If
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.DeleteRecords.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Public Function GetPathFile(ByRef psPath As String, ByRef psFile As String, ByRef psTitle As String, Optional ByRef pbNewFile As Boolean = False) As Boolean
		On Error GoTo ErrCall
		'
		Dim dlgMain As New CDlgComp.CCommonDialog
		Dim sExt As String
		'
		sExt = "*.mdb"
		'
		With dlgMain
			.DialogTitle = psTitle
			.DefaultExt = sExt
			.Filter = "Database Files(" & sExt & ")|" & sExt
			.FileTitle = psFile
			.FileName = psFile
			.InitDir = psPath
			.CancelError = True
			'
			If pbNewFile Then
				.FLAGS = CDlgComp.ccdControlContants.cdlOFNHideReadOnly + CDlgComp.ccdControlContants.cdlOFNPathMustExist + CDlgComp.ccdControlContants.cdlOFNOverwritePrompt
				.ShowSave()
			Else
				.FLAGS = CDlgComp.ccdControlContants.cdlOFNHideReadOnly + CDlgComp.ccdControlContants.cdlOFNFileMustExist
				.ShowOpen()
			End If
			'
			If .FileName = "" Then
				MsgBox("Database not located.")
				GetPathFile = False
			Else
				FileOps.SplitPathFile(.FileName, psPath, psFile)
				'Settings.SaveSetting App.Title, "File", sDBDescript & "Path", sDbPath
				'Settings.SaveSetting App.Title, "File", sDBDescript & "Name", sDBName
				GetPathFile = True
			End If
		End With
		'
		Exit Function
ErrCall: 
		If Err.Number = -2147219503 Then 'User cancelled
			GetPathFile = False
		Else
			MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.GetPathFile.", MsgBoxStyle.Critical, "Error")
		End If
	End Function
	
	'Public Function GetPathFile(psPath As String, psFile As String, psDBDescript As String) As Boolean
	'  On Error GoTo ErrCall
	'  '
	'  Dim dlgMain As New CCommonDialog
	'  Dim sExt As String
	'  '
	'  sExt = "*.mdb"
	'  '
	'  With dlgMain
	'  .DialogTitle = "Locate " & psDBDescript & " Database"
	'  .DefaultExt = sExt
	'  .Filter = "Database Files(" & sExt & ")|" & sExt & "|All Files|*.*"
	'  '.FileTitle = psDBName
	'  .filename = psFile
	'  .InitDir = psPath
	'  .CancelError = True
	'  .FLAGS = cdlOFNHideReadOnly + cdlOFNFileMustExist
	'  .ShowOpen
	'  '
	'  If .filename = "" Then
	'    MsgBox "Database not located."
	'    GetPathFile = False
	'  Else
	'    FileOps.SplitPathFile .filename, psPath, psFile
	'    GetPathFile = True
	'  End If
	'  End With
	'  '
	'  Exit Function
	'  '
	'ErrCall:
	'  If Err.Number = -2147219503 Then 'User cancelled
	'    '
	'  Else
	'    MsgBox Err.Description
	'  End If
	'  '
	'  GetPathFile = False
	'End Function
	
	'Public Function SetupSecurity(psWorkgroup As String, psUserID As String, psUserPwd As String) As Boolean
	'  On Error GoTo ErrCall
	'  '
	'  Dim wsMain As Workspace
	'  Dim sDbWorkgroup As String
	'  Dim sDlgMsg As String
	'  '
	'  If psWorkgroup <> "" Then
	'    If DAO.DBEngine.SystemDB <> "" Then
	'      SetupSecurity = True
	'      Exit Function
	'      '
	'      If LCase(psWorkgroup) <> LCase(DBEngine.SystemDB) Then
	'        sDlgMsg = "A workgroup is already defined for this process." & vbCrLf & _
	''                  App.Title & " may not be able to access the data" & vbCrLf & _
	''                  "files correctly." & vbCrLf & vbCrLf & _
	''                  "SystemDB:  " & DBEngine.SystemDB & vbCrLf & _
	''                  "Workgroup: " & psWorkgroup
	'        '
	'        MsgBox sDlgMsg, vbInformation, "Workgroup Conflict"
	'        '
	'        sDbWorkgroup = DBEngine.SystemDB
	'      End If
	'    Else
	'      DBEngine.SystemDB = psWorkgroup
	'      sDbWorkgroup = psWorkgroup
	'    End If
	'  End If
	'  '
	'  '\\ Database Security Check
	'  If psUserID <> "" Then
	'    Dim wsTmp As Workspace
	'    '
	'    On Error Resume Next
	'    Set wsTmp = DBEngine.CreateWorkspace("Test", psUserID, psUserPwd)
	'    '
	'    If Err <> 0 Then '\\ Security Check Failed
	'      sDlgMsg = "Database initialization failed. Please report this error" & vbCrLf & _
	''                "by contacting Hawkins Research, Inc. at (800) 736-1246." & vbCrLf & vbCrLf & _
	''                "LAST ERROR: " & Err.Description & "(" & Err.Number & ")"
	'      '
	'      MsgBox sDlgMsg, vbCritical + vbOKOnly, "ERROR: Database Security Check Failed."
	'      '
	'      On Error GoTo ErrCall
	'      SetupSecurity = False
	'    Else '\\ Security Check Suceeded
	'      On Error GoTo ErrCall
	'      '
	'      wsTmp.Close
	'      Set wsTmp = Nothing
	'      '
	'      Set wsMain = DBEngine.CreateWorkspace("WS1", psUserID, psUserPwd)
	'      '
	'      SetupSecurity = True
	'    End If
	'  End If
	'  '
	'  Exit Function
	'ErrCall:
	'  SetupSecurity = False
	'  '
	'  If Err.Number = 3028 Then
	'    MsgBox Err.Description & vbCrLf & vbCrLf & "Please contact tecnical support. Reinstalling this " & vbCrLf & "application may allow you to continue normally.", vbCritical, "Workgroup File Missing"
	'  Else
	'    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in Database Operations Setup Security.", vbCritical, "Error"
	'  End If
	'End Function
	
	Public Sub First(ByRef pRS As ADODB.Recordset)
		On Error GoTo ErrCall
		'
		With pRS
			If Not (.BOF And .EOF) Then
				.MoveFirst()
			End If
		End With
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.First.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	'Public Function CreateField(ptblDef As TableDef, psName As String, plType As Long, plSize As Long, Optional plAttributes As Long, Optional pvDefaultValue As Variant) As Boolean
	'  On Error GoTo ErrCall
	'  '
	'  Dim fldDef As DAO.Field
	'  '
	'  Set fldDef = ptblDef.CreateField
	'  '
	'  fldDef.Name = psName
	'  fldDef.Type = plType
	'  '
	'  If plType = dbText Or plType = dbMemo Then
	'    fldDef.AllowZeroLength = True
	'  End If
	'  '
	'  If plSize > 0 Then
	'    fldDef.Size = plSize
	'  End If
	'  '
	'  fldDef.Attributes = plAttributes
	'  '
	'  If Not IsMissing(pvDefaultValue) Then
	'    fldDef.DefaultValue = pvDefaultValue
	'  End If
	'  '
	'  ptblDef.Fields.Append fldDef
	'  '
	'  Set fldDef = Nothing
	'  '
	'  ' CSErrorHandler begin - please do not modify or remove this line
	'  Exit Function
	'ErrCall:
	'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.CreateField.", vbCritical, "Error"
	'End Function
	
	'Public Function CreateIndex(ptblDef As TableDef, psName As String, psFields As String, pbUnique As Boolean, pbIgnoreNulls As Boolean, pbPrimary As Boolean) As Boolean
	'  On Error GoTo ErrCall
	'  '
	'  Dim idx As Index
	'  '
	'  Set idx = ptblDef.CreateIndex
	'  '
	'  With idx
	'  .Name = psName
	'  .Fields = psFields
	'  .Unique = pbUnique
	'  .IgnoreNulls = pbIgnoreNulls
	'  .Primary = pbPrimary
	'  End With
	'  '
	'  ptblDef.Indexes.Append idx
	'  '
	'  Set idx = Nothing
	'  '
	'  ' CSErrorHandler begin - please do not modify or remove this line
	'  Exit Function
	'ErrCall:
	'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.CreateIndex.", vbCritical, "Error"
	'End Function
	
	'Public Sub CopyTable(pDBFrom As Database, pDBTo As Database, psTblFrom As String, Optional psTblTo As String, Optional pbShowMsg As Boolean)
	'  On Error GoTo ErrCall
	'  '
	'  Dim tblDefFrom As TableDef, tblDefTo As TableDef
	'  Dim fld As DAO.Field
	'  '
	'  If Not TableExists(pDBTo, psTblFrom) Then
	'    Set tblDefFrom = pDBFrom.TableDefs(psTblFrom)
	'    Set tblDefTo = pDBTo.CreateTableDef(psTblFrom)
	'    '
	'    For Each fld In tblDefFrom.Fields
	'      CopyField tblDefTo, fld
	'    Next
	'    '
	'    pDBTo.TableDefs.Append tblDefTo
	'    '
	'    CopyIndexes pDBFrom, pDBTo, psTblFrom, psTblFrom
	'  End If
	'  '
	'  'have option to delete table
	'  'pDBTo.TableDefs.Delete psTblFrom
	'  '
	'  CopyRecords pDBFrom, pDBTo, psTblFrom, psTblFrom
	'  '
	'  If pbShowMsg Then MsgBox "Copy successful!", vbInformation, "Copy Status"
	'  '
	'  ' CSErrorHandler begin - please do not modify or remove this line
	'  Exit Sub
	'ErrCall:
	'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.CopyTable.", vbCritical, "Error"
	'End Sub
	
	'Private Sub CopyField(ptblDefTo As TableDef, pCurField As DAO.Field, Optional psName As String)
	'  On Error GoTo ErrCall
	'  '
	'  Dim fldNew As DAO.Field
	'  '
	'  Set fldNew = ptblDefTo.CreateField(IIf(psName = "", pCurField.Name, psName), pCurField.Type, pCurField.Size)
	'  '
	'  fldNew.Attributes = pCurField.Attributes
	'  '
	'  If fldNew.Type = dbText Then
	'    fldNew.AllowZeroLength = pCurField.AllowZeroLength
	'  End If
	'  '
	'  fldNew.Required = pCurField.Required
	'  fldNew.ValidationText = pCurField.ValidationText
	'  fldNew.ValidationRule = pCurField.ValidationRule
	'  fldNew.DefaultValue = pCurField.DefaultValue
	'  fldNew.OrdinalPosition = pCurField.OrdinalPosition
	'  '
	'  ptblDefTo.Fields.Append fldNew
	'  '
	'  Exit Sub
	'  '
	'ErrCall:
	'  MsgBox Err.Description
	'End Sub
	
	'Public Sub CopyRecords(pDBFrom As Database, pDBTo As Database, psTblFrom As String, psTblTo As String)
	'  On Error GoTo ErrCall
	'  '
	'  Dim rsFrom As Recordset, rsTo As Recordset
	'  Dim iCntFld As Integer, lCntRecords As Long, lRecCount As Long
	'  '
	'  Set rsFrom = pDBFrom.OpenRecordset("SELECT * FROM " & psTblFrom)
	'  Set rsTo = pDBTo.OpenRecordset("SELECT * FROM " & psTblFrom)
	'  '
	'  If HasRecords(rsFrom, eLast) Then
	'    lRecCount = rsFrom.RecordCount
	'    '
	'    rsFrom.MoveFirst
	'    For lCntRecords = 0 To lRecCount - 1
	'      rsTo.AddNew
	'      '
	'      For iCntFld = 0 To rsFrom.Fields.Count - 1
	'        rsTo.Fields(iCntFld).Value = rsFrom.Fields(iCntFld).Value
	'      Next iCntFld
	'      '
	'      rsTo.Update
	'      '
	'      rsFrom.MoveNext
	'    Next lCntRecords
	'  End If
	'  '
	'  ' CSErrorHandler begin - please do not modify or remove this line
	'  Exit Sub
	'ErrCall:
	'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.CopyRecords.", vbCritical, "Error"
	'End Sub
	
	'Public Sub CopyIndexes(pDBFrom As Database, pDBTo As Database, psTblFrom As String, psTblTo As String)
	'  On Error GoTo ErrCall
	'  '
	'  Dim iCntIDX As Integer
	'  Dim idxNew As Index
	'  Dim tblDefFrom As TableDef, tblDefTo As TableDef
	'  '
	'  Set tblDefFrom = pDBFrom.TableDefs(psTblFrom)
	'  Set tblDefTo = pDBTo.TableDefs(psTblTo)
	'  '
	'  For iCntIDX = 0 To tblDefFrom.Indexes.Count - 1
	'    Set idxNew = New Index
	'    idxNew.Name = tblDefFrom.Indexes(iCntIDX).Name
	'    idxNew.Fields = tblDefFrom.Indexes(iCntIDX).Fields
	'    idxNew.Unique = tblDefFrom.Indexes(iCntIDX).Unique
	'    idxNew.Primary = tblDefFrom.Indexes(iCntIDX).Primary
	'    idxNew.IgnoreNulls = tblDefFrom.Indexes(iCntIDX).IgnoreNulls
	'    tblDefTo.Indexes.Append idxNew
	'    Set idxNew = Nothing
	'  Next iCntIDX
	'  '
	'  ' CSErrorHandler begin - please do not modify or remove this line
	'  Exit Sub
	'ErrCall:
	'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.CopyIndexes.", vbCritical, "Error"
	'End Sub
	
	'Public Function Compact(pDB As Database) As Boolean
	'  On Error GoTo ErrCall
	'  '
	'  Dim sDBName As String
	'  Dim sDBTmpName As String
	'  Dim sPathMain As String
	'  Dim sPathBackup As String
	'  Dim sExt As String
	'  '
	'  '\\ Local Declarations
	'  Dim rsTmp As Recordset
	'  '
	'  Compact = False
	'  '\\ Close All DAOs Before Attempting Compact Operation
	'  For Each rsTmp In pDB.Recordsets
	'    Me.ZapRS rsTmp
	'  Next
	'  '
	'  FileOps.SplitPathFile pDB.Name, sPathMain, sDBName
	'  sExt = FileOps.IsolateExtension(sDBName)
	'  If Len(sExt) > 0 Then sDBName = Left(sDBName, Len(sDBName) - Len(sExt))
	'  '
	'  sDBTmpName = sDBName & " " & Format(Now, "yyyymmmdd-hhnnss") & sExt
	'  sDBName = sDBName & sExt
	'  '
	'  On Error Resume Next
	'  pDB.Close
	'  Set pDB = Nothing
	'  '
	'  sPathBackup = sPathMain & "Backup\"
	'  MkDir sPathBackup
	'  Kill sPathBackup & sDBName
	'  '
	'  On Error GoTo ErrCall
	'  '
	'  '\\ Compact Database
	'  DBEngine.CompactDatabase sPathMain & sDBName, sPathMain & sDBTmpName
	'  FileCopy sPathMain & sDBName, sPathBackup & sDBName
	'  '
	'  If GetSetting(App.Title, "File", "RcyBin", True) Then
	'    With FileOps
	'      .ClearSourceFiles
	'      .AddSourceFile sPathMain & sDBName
	'      .DeleteFiles
	'    End With
	'  Else
	'    Kill sPathMain & sDBName
	'  End If
	'  '
	'  Name sPathMain & sDBTmpName As sPathMain & sDBName
	'  '
	'  Compact = True
	'  '
	'  Exit Function
	'ErrCall:
	'  Compact = False
	'  '
	'  If Err.Number <> 3356 Then
	'    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsDBOps.Compact.", vbCritical, "Error"
	'  End If
	'End Function
	
	'Public Sub SetProperty(pdbTemp As Database, sName As String, vValue As Variant)
	'  Dim prpNew As Property
	'  Dim errLoop As Error
	'
	'  ' Attempt to set the specified property.
	'  On Error GoTo Err_Property
	'  pdbTemp.Properties(sName) = vValue
	'  On Error GoTo 0
	'
	'  Exit Sub
	'
	'Err_Property:
	'
	'  ' Error 3270 means that the property was not found.
	'
	'If DBEngine.Errors(0).Number = 3270 Then
	'    Set prpNew = dbsTemp.CreateProperty(strName, dbBoolean, booTemp)
	'    dbsTemp.Properties.Append prpNew
	'    Resume Next
	'  Else
	'    ' If different error has occurred, display message.
	'    For Each errLoop In DBEngine.Errors
	'      MsgBox "Error number: " & errLoop.Number & vbCr & _
	''        errLoop.Description
	'    Next errLoop
	'
	'End
	'  End If
	'End Sub
	
	'Public Sub DeleteTable(pDB As Database, psTableName As String)
	'  On Error GoTo ErrCall
	'  '
	'  pDB.TableDefs.Delete psTableName
	'  '
	'  Exit Sub
	'ErrCall:
	'  MsgBox "Error " & Err.Description & " in DB Ops Delete Table.", vbCritical
	'End Sub
	
	'Public Sub SetDatDB(pDat As Data, pDB As Database, Optional psRecordSource As String, Optional pbRefresh As Boolean)
	'  On Error GoTo ErrCall
	'  '
	'  pDat.DatabaseName = pDB.Name
	'  If psRecordSource <> vbNullString Then pDat.RecordSource = psRecordSource
	'  pDat.Enabled = True
	'  If pbRefresh Then pDat.Refresh
	'  '
	'  Exit Sub
	'ErrCall:
	'  MsgBox Err.Description
	'End Sub
	
	'Public Function OpenDb(ByRef pDB As Database, ByRef psDBPath As String, ByRef psDBName As String, psDBTitle As String, Optional pbExclusive As Boolean) As Boolean
	'  On Local Error GoTo ErrCall:
	'  '
	'  Screen.MousePointer = vbHourglass
	'  OpenDb = False
	'  '
	'  Dim bFileExists As Boolean
	'  Dim sDBPath As String
	'  Dim sDBName As String
	'  '
	'  sDBPath = psDBPath
	'  sDBName = psDBName
	'  '
	'  '\\ Check if file exists
	'  bFileExists = FileOps.Exists(psDBPath & psDBName)
	'  '
	'  If Not bFileExists Then
	'    bFileExists = DBOps.GetPathFile(psDBPath, psDBName, psDBTitle)
	'    '
	'    If bFileExists Then
	'      If psDBPath <> sDBPath Then
	'        SaveSetting App.Title, "File", psDBTitle & "Path", sDBPath
	'      End If
	'      '
	'      If psDBName <> sDBName Then
	'        SaveSetting App.Title, "File", psDBTitle & "Name", sDBName
	'      End If
	'    End If
	'  End If
	'  '
	'  If Not bFileExists Then
	'    MsgBox psDBTitle & " database not found at " & psDBPath & psDBName & ".", vbCritical, "Database not found"
	'  Else
	'    On Error Resume Next
	'    Set pDB = OpenDatabase(psDBPath & psDBName, pbExclusive)
	'    On Error GoTo ErrCall
	'    OpenDb = Not (pDB Is Nothing)
	'  End If
	'  '
	'  Screen.MousePointer = vbDefault
	'  '
	'  Exit Function
	'ErrCall:
	'  Screen.MousePointer = vbDefault
	'  MsgBox Err.Description, vbInformation + vbOKOnly, "Error Opening Database (" & psDBPath & psDBName & ")"
	'  OpenDb = False
	'End Function
	
	Public Function OpenConnection(ByRef pCN As ADODB.Connection, ByRef psDBPath As String, ByRef psDBName As String, ByRef psDBTitle As String, Optional ByRef pbExclusive As Boolean = False) As Boolean
		On Error GoTo EH
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		OpenConnection = False
		'
		Dim bFileExists As Boolean
		Dim sDBPath As String
		Dim sDBName As String
		Dim sConnectionString As String
		'
		sDBPath = psDBPath
		sDBName = psDBName
		'
		'\\ Check if file exists
		bFileExists = FileOps.Exists(psDBPath & psDBName)
		'
		If Not bFileExists Then
			bFileExists = DBOps.GetPathFile(psDBPath, psDBName, psDBTitle)
			'
			If bFileExists Then
				If psDBPath <> sDBPath Then
					SaveSetting(My.Application.Info.Title, "File", psDBTitle & "Path", sDBPath)
				End If
				'
				If psDBName <> sDBName Then
					SaveSetting(My.Application.Info.Title, "File", psDBTitle & "Name", sDBName)
				End If
			End If
		End If
		'
		If Not bFileExists Then
			MsgBox(psDBTitle & " database not found at " & psDBPath & psDBName & ".", MsgBoxStyle.Critical, "Database not found")
		Else
			sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Password=CTRALEQDISMC;User ID=Eric;Data Source=" & sDBPath & sDBName & ";Persist Security Info=True;Jet OLEDB:System database=" & FileOps.SystemPath & "hr.mdw"
			'
			pCN = New ADODB.Connection
			pCN.CursorLocation = ADODB.CursorLocationEnum.adUseClient
			pCN.Open(sConnectionString)
			sCNDBName = psDBPath & psDBName
			OpenConnection = (pCN.State = ADODB.ObjectStateEnum.adStateOpen)
		End If
		'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		'
		Exit Function
EH: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		MsgBox(Err.Description, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error Connecting to Database (" & psDBPath & psDBName & ")")
		OpenConnection = False
	End Function
End Class