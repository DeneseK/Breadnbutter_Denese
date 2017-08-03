Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class CFileOps
	'
	'**********************************************************************************************'
	'\\ NAME:         clsFileOps
	'\\ DESCRIPTION:  Wraps common file functions
	'\\ DEVELOPER:    Jason M. Purcell (jpurcell@prefer.net)
	'\\ VERSION:      1.31
	'\\ CREATED:      1998.07.29
	'\\ REVISED:      2000.02.04
	'\\ DEPENDENCIES: VB5StKit.dll; Version.dll
	'\\ USAGE NOTES:  N/A
	'**********************************************************************************************'
	'
	'\\ General
	Private iDlgRsp As Short
	Private lAPIRes As Integer
	Private Const csZeroLen As String = ""
	Private blnUndo As Boolean
	Private blnConfirmMakeDir As Boolean
	Private blnConfirmOp As Boolean
	Private strCustomText As String
	Private blnIncDirs As Boolean
	Private lngParentWnd As Integer
	Private blnRenameCollision As Boolean
	Private blnSilent As Boolean
	Private FilesSource As New Collection
	Private FilesDest As New Collection
	'
	'\\ DLL: VB5STKIT.dll
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub lmemcpy Lib "VB5STKIT.DLL" (ByRef StrDest As Any, ByVal StrSrc As Any, ByVal lBytes As Integer)
	'
	'\\ DLL: Version.dll
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function VerInstallFile Lib "VERSION.DLL"  Alias "VerInstallFileA"(ByVal FLAGS As Integer, ByVal SrcName As String, ByVal DestName As String, ByVal SrcDir As String, ByVal DestDir As String, ByVal CurrDir As Any, ByVal TmpName As String, ByRef lpTmpFileLen As Integer) As Integer
	Private Declare Function GetFileVersionInfoSize Lib "VERSION.DLL"  Alias "GetFileVersionInfoSizeA"(ByVal strFilename As String, ByRef lVerHandle As Integer) As Integer
	Private Declare Function GetFileVersionInfo Lib "VERSION.DLL"  Alias "GetFileVersionInfoA"(ByVal strFilename As String, ByVal lVerHandle As Integer, ByVal lcbSize As Integer, ByRef lpvData As Byte) As Integer
	Private Declare Function VerQueryValue Lib "VERSION.DLL"  Alias "VerQueryValueA"(ByRef lpvVerData As Byte, ByVal lpszSubBlock As String, ByRef lplpBuf As Integer, ByRef lpcb As Integer) As Integer
	Private Declare Function OSGetShortPathName Lib "kernel32"  Alias "GetShortPathNameA"(ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Integer) As Integer
	'Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
	Private Structure VERINFO 'Version FIXEDFILEINFO
		Dim strPad1 As Integer '\\ Pad out struct version
		Dim strPad2 As Integer '\\ Pad out struct signature
		Dim nMSLo As Short '\\ Low word of ver # MS DWord
		Dim nMSHi As Short '\\ High word of ver # MS DWord
		Dim nLSLo As Short '\\ Low word of ver # LS DWord
		Dim nLSHi As Short '\\ High word of ver # LS DWord
		<VBFixedArray(36)> Dim strPad3() As Byte '\\ Pad out rest of VERINFO struct (36 bytes)
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			'UPGRADE_WARNING: Lower bound of array strPad3 was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim strPad3(36)
		End Sub
	End Structure
	'
	'\\ Win32API: General
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef pTo As Any, ByRef pFrom As Any, ByVal lCount As Integer)
	'UPGRADE_WARNING: Structure SECURITY_ATTRIBUTES may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreateDirectory Lib "kernel32"  Alias "CreateDirectoryA"(ByVal lpPathName As String, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES) As Integer
	Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Integer) As Integer
	'UPGRADE_WARNING: Structure WIN32_FIND_DATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FindFirstFile Lib "kernel32"  Alias "FindFirstFileA"(ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
	Private Declare Function GetSystemDirectory Lib "kernel32"  Alias "GetSystemDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	Private Declare Function GetTempPath Lib "kernel32"  Alias "GetTempPathA"(ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer
	Private Declare Function GetWindowsDirectory Lib "kernel32"  Alias "GetWindowsDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	Private Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function WritePrivateProfileString Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Integer
	Private Const INVALID_FILE_HANDLE As Short = -1
	Private Const INVALID_HANDLE_VALUE As Short = -1
	Private Const MAX_PATH As Short = 260
	Private Const gintNOVERINFO As Short = 32767
	Private Structure SECURITY_ATTRIBUTES
		Dim nLength As Integer
		Dim lpSecurityDescriptor As Integer
		Dim bInheritHandle As Integer
	End Structure
	Private Structure FILETIME
		Dim dwLowDateTime As Integer
		Dim dwHighDateTime As Integer
	End Structure
	Private Structure WIN32_FIND_DATA
		Dim dwFileAttributes As Integer
		Dim ftCreationTime As FILETIME
		Dim ftLastAccessTime As FILETIME
		Dim ftLastWriteTime As FILETIME
		Dim nFileSizeHigh As Integer
		Dim nFileSizeLow As Integer
		Dim dwReserved0 As Integer
		Dim dwReserved1 As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(MAX_PATH),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=MAX_PATH)> Public cFileName() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(14),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=14)> Public cAlternate() As Char
	End Structure
	'
	'\\ Win32API: SHFileOperation
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SHFileOperation Lib "shell32.dll"  Alias "SHFileOperationA"(ByRef lpFileOp As Any) As Integer
	Private Const FO_MOVE As Short = 1
	Private Const FO_COPY As Short = 2
	Private Const FO_DELETE As Short = 3
	Private Const FO_RENAME As Short = 4
	Private Const FOF_MULTIFilesDest As Short = &H1s
	Private Const FOF_SILENT As Short = &H4s
	Private Const FOF_RENAMEONCOLLISION As Short = &H8s
	Private Const FOF_NOCONFIRMATION As Short = &H10s
	Private Const FOF_WANTMAPPINGHANDLE As Short = &H20s
	Private Const FOF_ALLOWUNDO As Short = &H40s
	Private Const FOF_FILESONLY As Short = &H80s
	Private Const FOF_SIMPLEPROGRESS As Short = &H100s
	Private Const FOF_NOCONFIRMMKDIR As Short = &H200s
	Private Structure SHFILEOPSTRUCT
		Dim hWnd As Integer
		Dim wFunc As Integer
		Dim pFrom As String
		Dim pTo As String
		Dim fFlags As Short
		Dim fAnyOperationsAborted As Integer
		Dim hNameMappings As Integer
		Dim lpszProgressTitle As String
	End Structure
	Public Function EnsureTrailingBackslash(ByRef sPath As String) As String
		On Error GoTo ErrCall
		'
		sPath = sPath & IIf(Right(sPath, 1) <> "\", "\", csZeroLen)
		'
		EnsureTrailingBackslash = sPath
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_EnsureTrailingBackslash")
	End Function
	Public Property AllowUndo() As Boolean
		Get
			On Error GoTo ErrCall
			'
			AllowUndo = blnUndo
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_AllowUndo_GLet")
		End Get
		Set(ByVal Value As Boolean)
			On Error GoTo ErrCall
			'
			blnUndo = Value
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_AllowUndo_Let")
		End Set
	End Property
	
	
	Public Property ConfirmMakeDir() As Boolean
		Get
			On Error GoTo ErrCall
			'
			ConfirmMakeDir = blnConfirmMakeDir
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_ConfirmMakeDir_Get")
		End Get
		Set(ByVal Value As Boolean)
			On Error GoTo ErrCall
			'
			blnConfirmMakeDir = Value
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_ConfirmMakeDir_Let")
		End Set
	End Property
	
	
	Public Property ConfirmOperation() As Boolean
		Get
			On Error GoTo ErrCall
			'
			ConfirmOperation = blnConfirmOp
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_ConfirmOperation_Get")
		End Get
		Set(ByVal Value As Boolean)
			On Error GoTo ErrCall
			'
			blnConfirmOp = Value
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_ConfirmOperation_Let")
		End Set
	End Property
	
	
	
	Public Property CustomText() As String
		Get
			On Error GoTo ErrCall
			'
			CustomText = strCustomText
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_CustomText_Get")
		End Get
		Set(ByVal Value As String)
			On Error GoTo ErrCall
			'
			'\\ NOTE: If sCustomText = clsStrNull then dialog box displays
			'\\   the name of each file as it is processed.
			'
			strCustomText = Value
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_CustomText_Let")
		End Set
	End Property
	
	
	Public Property IncludeDirectories() As Boolean
		Get
			On Error GoTo ErrCall
			'
			'\\ Determines if operation affects directories when wildcards are specified
			'
			IncludeDirectories = blnIncDirs
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_IncludeDirectories_Get")
		End Get
		Set(ByVal Value As Boolean)
			On Error GoTo ErrCall
			'
			'\\ Determines if operation affects directories when wildcards are specified
			'
			blnIncDirs = Value
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_IncludeDirectories_Let")
		End Set
	End Property
	
	
	Public Property ParentWnd() As Integer
		Get
			On Error GoTo ErrCall
			'
			ParentWnd = lngParentWnd
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_ParentWnd_Get")
		End Get
		Set(ByVal Value As Integer)
			On Error GoTo ErrCall
			'
			lngParentWnd = Value
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_ParentWnd_Let")
		End Set
	End Property
	
	Public Property RenameOnCollision() As Boolean
		Get
			On Error GoTo ErrCall
			'
			RenameOnCollision = blnRenameCollision
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_RenameOnCollision_Get")
		End Get
		Set(ByVal Value As Boolean)
			On Error GoTo ErrCall
			'
			blnRenameCollision = Value
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_RenameOnCollision_Let")
		End Set
	End Property
	
	
	
	Public Property SilentMode() As Boolean
		Get
			On Error GoTo ErrCall
			'
			SilentMode = blnSilent
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_SilentMode_Get")
		End Get
		Set(ByVal Value As Boolean)
			On Error GoTo ErrCall
			'
			blnSilent = Value
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_SilentMode_Let")
		End Set
	End Property
	
	Public ReadOnly Property WindowsPath() As String
		Get
			On Error GoTo ErrCall
			'
			'\\ Local Declarations
			Dim sPathTmp As New VB6.FixedLengthString(512)
			'
			lAPIRes = GetWindowsDirectory(sPathTmp.Value, 512)
			WindowsPath = Left(sPathTmp.Value, lAPIRes)
			WindowsPath = WindowsPath & IIf(Right(WindowsPath, 1) = "\", csZeroLen, "\")
			'
			'\\ Alternative Method of Identifying Windows and System Paths
			'sPathTmp = String$(512, Chr$(0))
			'GetSystemDirectory sPathTmp, 512
			'sPathWinSys = Left$(sPathTmp, InStr(sPathTmp, Chr$(0)) - 1)
			'If Right$(sPathWinSys, 1) <> "\" Then sPathWinSys = sPathWinSys & "\"
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_WinDir_Get")
		End Get
	End Property
	Public ReadOnly Property SystemPath() As String
		Get
			On Error GoTo ErrCall
			'
			'\\ Local Declarations
			Dim sPathTmp As New VB6.FixedLengthString(512)
			'
			lAPIRes = GetSystemDirectory(sPathTmp.Value, 512)
			SystemPath = Left(sPathTmp.Value, lAPIRes)
			SystemPath = SystemPath & IIf(Right(SystemPath, 1) = "\", csZeroLen, "\")
			'
			'\\ Alternative Method of Setting Windows and System Paths
			'sPathTmp = String$(512, Chr$(0))
			'GetSystemDirectory sPathTmp, 512
			'sPathWinSys = Left$(sPathTmp, InStr(sPathTmp, Chr$(0)) - 1)
			'If Right$(sPathWinSys, 1) <> "\" Then sPathWinSys = sPathWinSys & "\"
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_WinSysDir_Get")
		End Get
	End Property
	
	Public ReadOnly Property DesktopPath() As String
		Get
			DesktopPath = WindowsPath & "Desktop\"
		End Get
	End Property
	
	Public ReadOnly Property TimeDateStamp() As String
		Get
			TimeDateStamp = Year(Today) & "."
			TimeDateStamp = TimeDateStamp & IIf(Len(Month(Today)) < 2, "0" & Month(Today), Month(Today)) & "."
			TimeDateStamp = TimeDateStamp & IIf(Len(VB.Day(Today)) < 2, "0" & VB.Day(Today), VB.Day(Today)) & "."
			TimeDateStamp = TimeDateStamp & IIf(Len(Hour(TimeOfDay)) < 2, "0" & Hour(TimeOfDay), Hour(TimeOfDay)) & "."
			TimeDateStamp = TimeDateStamp & IIf(Len(Minute(TimeOfDay)) < 2, "0" & Minute(TimeOfDay), Minute(TimeOfDay)) & "."
			TimeDateStamp = TimeDateStamp & IIf(Len(Second(TimeOfDay)) < 2, "0" & Second(TimeOfDay), Second(TimeOfDay))
		End Get
	End Property
	
	
	
	Public ReadOnly Property TemporaryPath() As String
		Get
			On Error GoTo ErrCall
			'
			'\\ Local Declarations
			Dim sPathTmp As New VB6.FixedLengthString(512)
			'
			lAPIRes = GetTempPath(512, sPathTmp.Value)
			TemporaryPath = Left(sPathTmp.Value, lAPIRes)
			TemporaryPath = EnsureTrailingBackslash(TemporaryPath)
			'
			Exit Property
			'
ErrCall: 
			iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_WinDir_Get")
		End Get
	End Property
	
	Public Function CreateBackupDir(ByRef dtTmp As Date) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim bN As Byte
		Dim sChr, sDT As String
		'
		sDT = CStr(dtTmp)
		For bN = 1 To Len(sDT) - 1
			sChr = Mid(sDT, bN, 1)
			If sChr <> "/" Then
				If sChr <> ":" Then
					If sChr <> " " Then
						CreateBackupDir = CreateBackupDir & sChr
					End If
				End If
			End If
		Next 
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_CreateBackupDir")
	End Function
	
	
	Public Function MoveFiles() As Boolean
		On Error GoTo ErrCall
		'
		MoveFiles = DoOperation(FO_MOVE)
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_MoveFiles")
	End Function
	
	Public Function CopyFiles() As Boolean
		On Error GoTo ErrCall
		'
		CopyFiles = DoOperation(FO_COPY)
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_CopyFiles")
	End Function
	
	Public Function DeleteFiles() As Boolean
		On Error GoTo ErrCall
		'
		DeleteFiles = DoOperation(FO_DELETE)
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_DeleteFiles")
	End Function
	
	Public Function RenameFiles() As Boolean
		On Error GoTo ErrCall
		'
		RenameFiles = DoOperation(FO_RENAME)
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_RenameFiles")
	End Function
	
	Private Sub PackVerInfo(ByVal strVersion As String, ByRef sVerInfo As VERINFO)
		'\\ Parses file version number string of the form x[.x[.x[.x]]] and assigns the
		'\\ extracted numbers to the appropriate elements of a VERINFO type variable.
		'
		'\\ Local Declarations
		Dim intOffset As Short
		Dim intAnchor As Short
		'
		On Error GoTo ErrCall
		'
		intOffset = InStr(strVersion, ".")
		If intOffset = 0 Then
			sVerInfo.nMSHi = Val(strVersion)
			GoTo PVIMSLo
		Else
			sVerInfo.nMSHi = Val(Left(strVersion, intOffset - 1))
			intAnchor = intOffset + 1
		End If
		'
		intOffset = InStr(intAnchor, strVersion, ".")
		If intOffset = 0 Then
			sVerInfo.nMSLo = Val(Mid(strVersion, intAnchor))
			GoTo PVILSHi
		Else
			sVerInfo.nMSLo = Val(Mid(strVersion, intAnchor, intOffset - intAnchor))
			intAnchor = intOffset + 1
		End If
		'
		intOffset = InStr(intAnchor, strVersion, ".")
		If intOffset = 0 Then
			sVerInfo.nLSHi = Val(Mid(strVersion, intAnchor))
			GoTo PVILSLo
		Else
			sVerInfo.nLSHi = Val(Mid(strVersion, intAnchor, intOffset - intAnchor))
			intAnchor = intOffset + 1
		End If
		'
		intOffset = InStr(intAnchor, strVersion, ".")
		If intOffset = 0 Then
			sVerInfo.nLSLo = Val(Mid(strVersion, intAnchor))
		Else
			sVerInfo.nLSLo = Val(Mid(strVersion, intAnchor, intOffset - intAnchor))
		End If
		'
		Exit Sub
		'
ErrCall: 
		sVerInfo.nMSHi = 0
PVIMSLo: 
		sVerInfo.nMSLo = 0
PVILSHi: 
		sVerInfo.nLSHi = 0
PVILSLo: 
		sVerInfo.nLSLo = 0
	End Sub
	Public Sub ClearSourceFiles()
		On Error GoTo ErrCall
		'
		'UPGRADE_NOTE: Object FilesSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FilesSource = Nothing
		'
		Exit Sub
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_ClearSourceFiles")
	End Sub
	
	
	Public Sub AddSourceFile(ByRef sFilename As String)
		On Error GoTo ErrCall
		'
		FilesSource.Add(sFilename)
		'
		Exit Sub
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_AddSourceFile")
	End Sub
	
	Private Function GetFileVerStruct(ByVal strFilename As String, ByRef sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile As Object = Nothing) As Boolean
		'\\ [fIsRemoteServerSupportFile] - True if file is a remote OLE automation server
		'\\ support file (.VBR)
		'
		'\\ Local Declarations
		Const strFIXEDFILEINFO As String = "\"
		Dim lVerSize As Integer
		Dim lVerHandle As Integer
		Dim lpBufPtr As Integer
		Dim byteVerData() As Byte
		'
		GetFileVerStruct = False
		'
		'\\ Initial Validation
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fIsRemoteServerSupportFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If IsNothing(fIsRemoteServerSupportFile) Then fIsRemoteServerSupportFile = False
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object fIsRemoteServerSupportFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If fIsRemoteServerSupportFile Then
			GetFileVerStruct = GetRemoteSupportFileVerStruct(strFilename, sVerInfo)
		Else
			'\\ Assertain file version info size, allocate buffer, and get version info.
			'\\ Query fixed file info portion, where the internal file version used by the
			'\\ VerInstallFile API is kept. Copy fixed file info into a VERINFO structure.
			lVerSize = GetFileVersionInfoSize(strFilename, lVerHandle)
			If lVerSize > 0 Then
				ReDim byteVerData(lVerSize)
				If GetFileVersionInfo(strFilename, lVerHandle, lVerSize, byteVerData(0)) <> 0 Then ' (Pass byteVerData array via reference to first element)
					If VerQueryValue(byteVerData(0), strFIXEDFILEINFO & "", lpBufPtr, lVerSize) <> 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object sVerInfo. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lmemcpy(sVerInfo, lpBufPtr, lVerSize)
						'CopyMemory sVerInfo, lpBufPtr, lVerSize
						GetFileVerStruct = True
					End If
				End If
			End If
		End If
	End Function
	Private Function GetRemoteSupportFileVerStruct(ByVal strFilename As String, ByRef sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile As Object = Nothing) As Boolean
		'\\ Retrieves file version information of a remote OLE server support file into a
		'\\ VERINFO data structure. These files do not possess a Windows version stamp, but
		'\\ they do have an internal version stamp that can be read.
		'
		'\\ Local Declarations
		Const strVersionKey As String = "Version="
		Dim cchVersionKey As Short
		Dim iFile As Short
		'
		On Error GoTo ErrCall
		'
		cchVersionKey = Len(strVersionKey)
		sVerInfo.nMSHi = gintNOVERINFO
		'
		iFile = FreeFile
		FileOpen(iFile, strFilename, OpenMode.Input, OpenAccess.Read, OpenShare.LockReadWrite)
		'
		Dim strLine As String
		Dim strVersion As String
		While (Not EOF(iFile))
			'\\ Local Declarations
			strLine = LineInput(iFile)
			If Left(strLine, cchVersionKey) = strVersionKey Then
				'\\ Version Key found.  Retrieve everything after equals sign.
				'\\ Local Declarations
				strVersion = Mid(strLine, cchVersionKey + 1)
				'\\ Parse Version Info.
				PackVerInfo(strVersion, sVerInfo)
				'\\ Convert the format 1.2.3 from the .VBR into '1.2.0.3
				sVerInfo.nLSLo = sVerInfo.nLSHi
				sVerInfo.nLSHi = 0
				'
				GetRemoteSupportFileVerStruct = True
				FileClose(iFile)
				Exit Function
			End If
		End While
		'
		FileClose(iFile)
		'
		Exit Function
		'
ErrCall: 
		GetRemoteSupportFileVerStruct = False
	End Function
	Public Function Exists(ByVal sFileAddr As String) As Boolean
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim WFD As WIN32_FIND_DATA
		'
		If Right(sFileAddr, 1) = "\" Then sFileAddr = Left(sFileAddr, Len(sFileAddr) - 1)
		'
		lAPIRes = FindFirstFile(sFileAddr, WFD) '\\ lAPIRes = file handle
		Exists = lAPIRes <> INVALID_FILE_HANDLE
		'
		FindClose(lAPIRes)
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_FileExists")
	End Function
	Public Sub ClearFilesDest()
		On Error GoTo ErrCall
		'
		'UPGRADE_NOTE: Object FilesDest may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FilesDest = Nothing
		'
		Exit Sub
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_ClearFilesDest")
	End Sub
	
	Public Sub AddDestFile(ByRef sFilename As String)
		On Error GoTo ErrCall
		'
		FilesDest.Add(sFilename)
		'
		Exit Sub
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_AddDestFile")
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		On Error GoTo ErrCall
		'
		blnUndo = True
		blnConfirmMakeDir = False
		blnConfirmOp = False
		strCustomText = csZeroLen
		blnIncDirs = False
		lngParentWnd = 0
		blnRenameCollision = False
		blnSilent = True
		'
		Exit Sub
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_evtInitialize")
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Function DoOperation(ByRef wFunc As Short) As Boolean
		On Error GoTo ErrCall
		'
		Dim n, ptrTmp As Integer
		Dim sfoClass As SHFILEOPSTRUCT
		Dim aryByte() As Byte
		Dim Bfr1() As Byte
		Dim Bfr2() As Byte
		Dim Bfr3() As Byte
		'
		With sfoClass
			'
			'\\ Parent Window Of Dialog Box
			.hWnd = lngParentWnd
			'\\ Operation To Perform
			.wFunc = wFunc
			'\\ Operation Flags
			.fFlags = 0
			'
			'\\ Build Flag Data
			If blnUndo Then .fFlags = .fFlags Or FOF_ALLOWUNDO
			If blnSilent Then .fFlags = .fFlags Or FOF_SILENT
			If blnRenameCollision Then .fFlags = .fFlags Or FOF_RENAMEONCOLLISION
			If Not blnConfirmOp Then .fFlags = .fFlags Or FOF_NOCONFIRMATION
			If Not blnConfirmMakeDir Then .fFlags = .fFlags Or FOF_NOCONFIRMMKDIR
			If Not blnIncDirs Then .fFlags = .fFlags Or FOF_FILESONLY
			'
			If Len(strCustomText) > 0 Then
				.lpszProgressTitle = strCustomText
				.fFlags = .fFlags Or FOF_SIMPLEPROGRESS
			End If
			'
			'\\ Build "From" String
			If FilesSource.Count() = 0 Then Err.Raise(vbObjectError + 1000,  , "No source files specified file operation")
			For n = 1 To FilesSource.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object FilesSource(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.pFrom = .pFrom & FilesSource.Item(n) & Chr(0)
			Next n
			'
			'\\ Build "To" String
			For n = 1 To FilesDest.Count()
				'UPGRADE_WARNING: Couldn't resolve default property of object FilesDest(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.pTo = .pTo & FilesDest.Item(n) & Chr(0)
			Next n
			If FilesDest.Count() > 1 Then .fFlags = .fFlags Or FOF_MULTIFilesDest
			'
			'\\ Note: Windows packs the SHFILEOPSTRUCT structure but VB
			'\\   does not. Therefore, all members following the two-byte
			'\\   fFlags member are offset by 2 bytes. To compensate,
			'\\   structure members are copied to a byte array with the
			'\\   proper alignment which is passed to SHFileOperation.
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			ReDim aryByte(LenB(sfoClass) - 2)
			CopyMemory(aryByte(0), .hWnd, Len(.hWnd))
			CopyMemory(aryByte(4), .wFunc, Len(.wFunc))
			'\\ Variable-length Strings Require Additonal Processing
			'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
			Bfr1 = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(.pFrom & Chr(0), vbFromUnicode))
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			ptrTmp = VarPtr(Bfr1(0))
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			CopyMemory(aryByte(8), ptrTmp, LenB(ptrTmp))
			'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
			Bfr2 = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(.pTo & Chr(0), vbFromUnicode))
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			ptrTmp = VarPtr(Bfr2(0))
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			CopyMemory(aryByte(12), ptrTmp, LenB(ptrTmp))
			CopyMemory(aryByte(16), .fFlags, Len(.fFlags))
			CopyMemory(aryByte(18), .fAnyOperationsAborted, Len(.fAnyOperationsAborted))
			CopyMemory(aryByte(22), .hNameMappings, Len(.hNameMappings))
			'UPGRADE_ISSUE: Constant vbFromUnicode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
			Bfr3 = System.Text.UnicodeEncoding.Unicode.GetBytes(StrConv(.lpszProgressTitle & Chr(0), vbFromUnicode))
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			ptrTmp = VarPtr(Bfr3(0))
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			CopyMemory(aryByte(26), ptrTmp, LenB(ptrTmp))
			'
			'\\ Call SHFileOperation
			n = SHFileOperation(aryByte(0))
			'\\ Retrieve fAnyOperationsAborted Flag
			CopyMemory(.fAnyOperationsAborted, aryByte(18), Len(.fAnyOperationsAborted))
			'\\ Return True If SHFileOperation Succeeded And No Operations Were Aborted
			DoOperation = Not CBool(n Or .fAnyOperationsAborted)
			'
		End With
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_DoOperation")
	End Function
	
	Public Function ConvertPath(ByRef sDirSrc As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim sDirBranch, sDirDest, sBranchExt As String
		Dim iSrcLen, iBranchLen As Short
		Dim iLocNext, iLocCur, iLocBranch As Short
		'
		iSrcLen = Len(sDirSrc)
		iLocCur = InStr(1, sDirSrc, "\", 1)
		sDirDest = Left(sDirSrc, iLocCur - 1)
		iLocCur = iLocCur + 1
		Do Until iLocCur > iSrcLen
			iLocNext = InStr(iLocCur, sDirSrc, "\", 1)
			iLocNext = IIf(iLocNext > iLocCur, iLocNext, iSrcLen + 1)
			sDirBranch = Mid(sDirSrc, iLocCur, iLocNext - iLocCur)
			iBranchLen = Len(sDirBranch)
			If iBranchLen > 8 Then
				If iLocNext = iSrcLen + 1 Then
					iLocBranch = IIf(InStr(1, sDirBranch, ".", 1) > 0, InStr(1, sDirBranch, ".", 1) + 1, iBranchLen)
					sBranchExt = Right(sDirBranch, iBranchLen - iLocBranch)
					sDirBranch = Left(sDirBranch, iBranchLen - Len(sBranchExt))
				End If
				If iBranchLen > 8 Then sDirBranch = Left(sDirBranch, 6) & "~1"
				If iLocNext = iSrcLen + 1 Then sDirBranch = sDirBranch & sBranchExt
			End If
			sDirDest = sDirDest & "\" & sDirBranch
			iLocCur = iLocNext + 1
		Loop 
		'
		ConvertPath = sDirDest
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_ConvertPath")
	End Function
	Public Function Execute(ByRef frmCur As System.Windows.Forms.Form, ByRef sTarget As String, ByRef sParam As String, ByRef sPath As String, ByRef lWinState As Integer) As Integer
		On Error GoTo ErrCall
		'
		lAPIRes = ShellExecute(frmCur.Handle.ToInt32, "Open", sTarget, sParam, sPath, lWinState)
		'
		Execute = lAPIRes
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_Execute")
	End Function
	
	Public Function GetFileVersion(ByVal strFilename As String, Optional ByVal fIsRemoteServerSupportFile As Object = Nothing) As String
		'\\ This function requires VB Setup Kit's Version.dll and "Declarations: Version.dll"
		'\\ declarations to be remmed in clsFileOps' (General) Declarations procedure.
		'\\ Returns the internal file version number for the specified file.  This can be
		'\\ different than the 'display' version number shown in the File Properties dialog.
		'\\ It is the same number as shown in the VB4 SetupWizard's File Details screen.  This
		'\\ is the number used by the Windows VerInstallFile API when comparing file versions.
		'\\ [fIsRemoteServerSupportFile] = whether or not this file is a remote OLE automation
		'\\ server support file (.VBR). If missing, False is assumed.
		'
		'\\ Local Declarations
		'UPGRADE_WARNING: Arrays in structure sVerInfo may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim sVerInfo As VERINFO
		Dim strVer As String
		'
		On Error GoTo ErrCall
		'
		'\\ Initial Validation
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object fIsRemoteServerSupportFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If IsNothing(fIsRemoteServerSupportFile) Then fIsRemoteServerSupportFile = False
		'
		If GetFileVerStruct(strFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
			strVer = VB6.Format(sVerInfo.nMSHi) & "." & VB6.Format(sVerInfo.nMSLo) & "."
			strVer = strVer & VB6.Format(sVerInfo.nLSHi) & "." & VB6.Format(sVerInfo.nLSLo)
			GetFileVersion = strVer
		Else
			GetFileVersion = csZeroLen
		End If
		'
		Exit Function
		'
ErrCall: 
		GetFileVersion = csZeroLen
		Err.Clear()
	End Function
	
	Public Sub WriteINIEntry(ByRef sSection As String, ByRef sKey As String, ByRef sValue As String, ByRef sFile As String)
		On Error GoTo ErrCall
		'
		WritePrivateProfileString(sSection, sKey, sValue, sFile)
		'
		Exit Sub
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_WriteINIEntry_Get")
	End Sub
	
	Public Sub CreateSubDirectory(ByVal sDirNew As String)
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim saMain As SECURITY_ATTRIBUTES
		Dim sDrive As String
		Dim sDirNewCreate As String
		Dim sDirSeg As String
		Dim sDirSegs() As String
		Dim iStrPos As Short
		Dim liN As Short
		'
		If Right(sDirNew, 1) <> "\" Then sDirNew = sDirNew & "\"
		'
		iStrPos = InStr(sDirNew, ":")
		'
		If iStrPos Then
			sDrive = fsDirParse(sDirNew, "\")
		Else
			sDrive = csZeroLen
		End If
		'
		Do Until sDirNew = ""
			sDirSeg = fsDirParse(sDirNew, "\")
			ReDim Preserve sDirSegs(liN)
			If liN = 0 Then sDirSeg = sDrive & sDirSeg
			sDirSegs(liN) = sDirSeg
			liN = liN + 1
		Loop 
		'
		liN = -1
		Do 
			liN = liN + 1
			sDirNewCreate = sDirNewCreate & sDirSegs(liN)
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			saMain.nLength = LenB(saMain)
			CreateDirectory(sDirNewCreate, saMain)
		Loop Until liN = UBound(sDirSegs)
		'
		Exit Sub
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_CustomText_Get")
	End Sub
	
	Public Function StripPath(ByRef sFileAddr As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim iPosCur, iPosFinal As Short
		'
		iPosCur = InStr(1, sFileAddr, "\", 1)
		Do Until iPosCur = 0
			iPosFinal = iPosCur
			iPosCur = InStr(iPosCur + 1, sFileAddr, "\", 1)
		Loop 
		'
		If iPosFinal = 0 Then
			StripPath = sFileAddr
		Else
			StripPath = Right(sFileAddr, Len(sFileAddr) - iPosFinal)
		End If
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_StripPath")
	End Function
	Public Function StripFile(ByRef sFileAddr As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim iPosCur, iPosFinal As Short
		'
		iPosCur = InStr(1, sFileAddr, "\", 1)
		Do Until iPosCur = 0
			iPosFinal = iPosCur
			iPosCur = InStr(iPosCur + 1, sFileAddr, "\", 1)
		Loop 
		'
		If iPosCur = 0 Then
			StripFile = sFileAddr
		Else
			StripFile = Left(sFileAddr, iPosFinal)
		End If
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_StripNameExtension")
	End Function
	
	Private Function fsLeadingZeroFormat(ByRef sSrc As String, Optional ByRef bZeros As Byte = 0) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim bN As Byte
		Dim sMask As String
		'
		If bZeros = 0 Then bZeros = 4
		'
		For bN = 1 To bZeros
			sMask = sMask & "0"
		Next 
		'
		fsLeadingZeroFormat = VB6.Format(sSrc, sMask)
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_fsLeadingZeroFormat")
	End Function
	Public Function EnsureUniqueFile(ByRef sFileAddr As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim sPath As String
		Dim sFile As String
		Dim sName As String
		Dim sExt As String
		'
		sPath = IsolatePath(sFileAddr)
		sFile = IsolateFile(sFileAddr)
		sName = IsolateName(sFile)
		sExt = IsolateExtension(sFile)
		'
		liN = 2
		Do Until Not Exists(sFileAddr)
			sFileAddr = sPath & sName & " (" & fsLeadingZeroFormat(CStr(liN)) & ")" & sExt
			liN = liN + 1
		Loop 
		'
		EnsureUniqueFile = sFileAddr
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_UniqueFileName")
	End Function
	Public Function StripExtension(ByRef psFile As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim iPosCur As Short
		Dim iPosFinal As Short
		'
		iPosCur = InStr(1, psFile, ".", 1)
		Do Until iPosCur = 0
			iPosFinal = iPosCur
			iPosCur = InStr(1 + iPosFinal, psFile, ".", 1)
		Loop 
		'
		If iPosFinal > 0 Then
			StripExtension = Left(psFile, iPosFinal - 1)
		Else
			StripExtension = psFile
		End If
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_StripExtension")
	End Function
	
	Public Function IsolateExtension(ByRef sFile As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim iPosCur, iPosFinal As Short
		'
		iPosCur = InStr(1, sFile, ".", 1)
		Do Until iPosCur = 0
			iPosFinal = iPosCur
			iPosCur = InStr(1 + iPosFinal, sFile, ".", 1)
		Loop 
		'
		If iPosFinal = 0 Then
			IsolateExtension = csZeroLen
		Else
			IsolateExtension = Right(sFile, Len(sFile) - (iPosFinal - 1))
		End If
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_IsolateExtension")
	End Function
	
	Public Function EnsureExtension(ByRef sFileOrFileAddr As String, ByRef sExt As String) As Object
		On Error GoTo ErrCall
		'
		If Left(sExt, 1) <> "." Then sExt = "." & sExt
		If Right(sFileOrFileAddr, 4) <> sExt Then sFileOrFileAddr = sFileOrFileAddr & sExt
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object EnsureExtension. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		EnsureExtension = sFileOrFileAddr
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: clsFileOps_EnsureExtension")
	End Function
	
	Public Function IsolatePath(ByRef sFileAddr As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim iPosCur, iPosFinal As Short
		'
		iPosCur = InStr(1, sFileAddr, "\", 1)
		Do Until iPosCur = 0
			iPosFinal = iPosCur
			iPosCur = InStr(iPosCur + 1, sFileAddr, "\", 1)
		Loop 
		'
		If iPosFinal = 0 Then
			IsolatePath = sFileAddr
		Else
			IsolatePath = Left(sFileAddr, iPosFinal)
		End If
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_IsolatePath")
	End Function
	
	Public Function IsolateName(ByVal sFile As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim iPosCur, iPosFinal As Short
		'
		sFile = IsolateFile(sFile)
		'
		iPosCur = InStr(1, sFile, ".", 1)
		Do Until iPosCur = 0
			iPosFinal = iPosCur
			iPosCur = InStr(1 + iPosFinal, sFile, ".", 1)
		Loop 
		'
		If iPosFinal = 0 Then
			IsolateName = sFile
		Else
			IsolateName = Left(sFile, iPosFinal - 1)
		End If
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_IsolateName")
	End Function
	
	Public Function IsolateFile(ByRef sFileAddr As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim iPosCur, iPosFinal As Short
		'
		iPosCur = InStr(1, sFileAddr, "\", 1)
		Do Until iPosCur = 0
			iPosFinal = iPosCur
			iPosCur = InStr(iPosCur + 1, sFileAddr, "\", 1)
		Loop 
		'
		If iPosFinal = 0 Then
			IsolateFile = sFileAddr
		Else
			IsolateFile = Right(sFileAddr, Len(sFileAddr) - iPosFinal)
		End If
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_IsolateFile")
	End Function
	
	Public Function StripName(ByRef sFile As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim iPosCur, iPosFinal As Short
		'
		iPosCur = InStr(1, sFile, ".", 1)
		Do Until iPosCur = 0
			iPosFinal = iPosCur
			iPosCur = InStr(1 + iPosFinal, sFile, ".", 1)
		Loop 
		'
		If iPosFinal = 0 Then
			StripName = csZeroLen
		Else
			StripName = Right(sFile, Len(sFile) - (iPosFinal - 1))
		End If
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_StripName")
	End Function
	
	
	
	Public Function IsFile(ByRef sFileAddr As String) As Boolean
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		'Dim liN As Integer
		'Dim iPosCur As Integer, iPosFinal As Integer
		''
		'sFileAddr = IsolateFile(sFileAddr)
		''
		'iPosCur = InStr(1, sFile, ".", 1)
		'Do Until iPosCur = 0
		'  iPosFinal = iPosCur
		'  iPosCur = InStr(1 + iPosFinal, sFile, ".", 1)
		'Loop
		''
		'IsolateName = IIf(iPosFinal = 0, sFile, Left$(sFile, iPosFinal - 1))
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_IsFile")
	End Function
	
	Public Function CreateFileID(Optional ByRef bYrDgts As Byte = 0, Optional ByRef sDlmtr As String = "") As String
		On Error GoTo ErrCall
		'
		If bYrDgts = 0 Then bYrDgts = 4
		If sDlmtr = csZeroLen Then sDlmtr = "."
		'
		If bYrDgts = 4 Then
			CreateFileID = VB6.Format(Now, "yyyymmdd" & sDlmtr & "hhnnss")
		Else
			CreateFileID = VB6.Format(Now, "yymmdd" & sDlmtr & "hhnnss")
		End If
		'
		Exit Function
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_CreateFileID")
	End Function
	
	Private Function fsDirParse(ByRef sSource As String, ByRef sDelimiter As String) As String
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim liN As Short
		Dim sSeg As String
		'
		liN = 1
		'
		Do 
			If Mid(sSource, liN, 1) = sDelimiter Then
				sSeg = Mid(sSource, 1, liN)
				sSource = Mid(sSource, liN + 1, Len(sSource))
				fsDirParse = sSeg
				Exit Function
			End If
			liN = liN + 1
		Loop 
		'
		Exit Function
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_fsDirParse")
	End Function
	
	Public Sub SplitPathFile(ByRef psFilename As String, ByRef psPath As String, ByRef psFile As String)
		On Error GoTo ErrCall
		'
		Dim iPos As Short
		'
		iPos = Len(psFilename)
		'
		Do While iPos > 0
			If InStr(iPos, psFilename, "\") Then
				Exit Do
			Else
				iPos = iPos - 1
			End If
		Loop 
		'
		psFile = Mid(psFilename, iPos + 1)
		psPath = Mid(psFilename, 1, Len(psFilename) - Len(psFile))
		'
		Exit Sub
		'
ErrCall: 
		iDlgRsp = MsgBox(fstrErrMsg(Err.Number), CDbl(MsgBoxStyle.OKOnly & MsgBoxStyle.Critical), "ERROR: clsFileOps_SplitPathFile")
	End Sub
	Private Function fstrErrMsg(ByRef errno As Integer) As String
		fstrErrMsg = ErrorToString(errno)
	End Function
	
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
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFileOps.FullPath", MsgBoxStyle.Critical, "Error")
	End Function
	
	Public Function MkDirAll(ByRef psPath As String) As Boolean
		On Error GoTo ErrCall
		'
		Dim iPos As Short
		Dim sPath As String
		Dim sOriginalPath As String
		'
		sPath = Me.IsolatePath(psPath)
		'
		If Len(sPath) > 0 Then
			On Error Resume Next
			MkDir(sPath)
			If Err.Number Then
				If Err.Number = 75 Then 'path exists
					MkDirAll = True
				Else
					sOriginalPath = sPath
					sPath = Left(sPath, Len(sPath) - 1)
					sPath = Me.IsolatePath(sPath)
					'
					If Len(sPath) > 2 Then
						If Me.MkDirAll(sPath) Then
							MkDir(sOriginalPath)
						Else
							MkDirAll = False
						End If
					Else
						MkDirAll = False
					End If
				End If
			Else
				MkDirAll = True
			End If
		End If
		'
		' CSErrorHandler begin - please do not modify or remove this line
		Exit Function
ErrCall: 
		MkDirAll = False
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFileOps.MkDir.", MsgBoxStyle.Critical, "Error")
	End Function
End Class