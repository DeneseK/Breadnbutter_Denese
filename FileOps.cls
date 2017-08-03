VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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
Private iDlgRsp             As Integer
Private lAPIRes             As Long
Private Const csZeroLen     As String = ""
Private blnUndo             As Boolean
Private blnConfirmMakeDir   As Boolean
Private blnConfirmOp        As Boolean
Private strCustomText       As String
Private blnIncDirs          As Boolean
Private lngParentWnd        As Long
Private blnRenameCollision  As Boolean
Private blnSilent           As Boolean
Private FilesSource         As New Collection
Private FilesDest           As New Collection
'
'\\ DLL: VB5STKIT.dll
Private Declare Sub lmemcpy Lib "VB5STKIT.DLL" (StrDest As Any, ByVal StrSrc As Any, ByVal lBytes As Long)
'
'\\ DLL: Version.dll
Private Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" (ByVal FLAGS&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Private Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFilename As String, lVerHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFilename As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
Private Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
Private Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type VERINFO 'Version FIXEDFILEINFO
  strPad1 As Long '\\ Pad out struct version
  strPad2 As Long '\\ Pad out struct signature
  nMSLo As Integer '\\ Low word of ver # MS DWord
  nMSHi As Integer '\\ High word of ver # MS DWord
  nLSLo As Integer '\\ Low word of ver # LS DWord
  nLSHi As Integer '\\ High word of ver # LS DWord
  strPad3(1 To 36) As Byte '\\ Pad out rest of VERINFO struct (36 bytes)
End Type
'
'\\ Win32API: General
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, pFrom As Any, ByVal lCount As Long)
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Const INVALID_FILE_HANDLE = -1
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Const gintNOVERINFO As Integer = 32767
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type
'
'\\ Win32API: SHFileOperation
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
Private Const FO_MOVE = 1
Private Const FO_COPY = 2
Private Const FO_DELETE = 3
Private Const FO_RENAME = 4
Private Const FOF_MULTIFilesDest = &H1
Private Const FOF_SILENT = &H4
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_WANTMAPPINGHANDLE = &H20
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Type SHFILEOPSTRUCT
  hWnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type
Public Function EnsureTrailingBackslash(sPath As String) As String
  On Local Error GoTo ErrCall:
  '
  sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", csZeroLen)
  '
  EnsureTrailingBackslash = sPath
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_EnsureTrailingBackslash")
End Function
Public Property Let AllowUndo(bAllowUndo As Boolean)
  On Local Error GoTo ErrCall:
  '
  blnUndo = bAllowUndo
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_AllowUndo_Let")
End Property

Public Function CreateBackupDir(dtTmp As Date) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim bN As Byte
  Dim sChr As String, sDT As String
  '
  sDT = dtTmp
  For bN = 1 To Len(sDT) - 1
    sChr = Mid$(sDT, bN, 1)
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
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_CreateBackupDir")
End Function
Public Property Get AllowUndo() As Boolean
  On Local Error GoTo ErrCall:
  '
  AllowUndo = blnUndo
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_AllowUndo_GLet")
End Property

Public Property Let ConfirmMakeDir(bConfirmMakeDir As Boolean)
  On Local Error GoTo ErrCall:
  '
  blnConfirmMakeDir = bConfirmMakeDir
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_ConfirmMakeDir_Let")
End Property

Public Property Get ConfirmMakeDir() As Boolean
  On Local Error GoTo ErrCall:
  '
  ConfirmMakeDir = blnConfirmMakeDir
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_ConfirmMakeDir_Get")
End Property

Public Property Let ConfirmOperation(bConfirmOperation As Boolean)
  On Local Error GoTo ErrCall:
  '
  blnConfirmOp = bConfirmOperation
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_ConfirmOperation_Let")
End Property

Public Property Get ConfirmOperation() As Boolean
  On Local Error GoTo ErrCall:
  '
  ConfirmOperation = blnConfirmOp
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_ConfirmOperation_Get")
End Property


Public Property Let CustomText(sCustomText As String)
  On Local Error GoTo ErrCall:
  '
  '\\ NOTE: If sCustomText = clsStrNull then dialog box displays
  '\\   the name of each file as it is processed.
  '
  strCustomText = sCustomText
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_CustomText_Let")
End Property

Public Property Get CustomText() As String
  On Local Error GoTo ErrCall:
  '
  CustomText = strCustomText
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_CustomText_Get")
End Property

Public Property Let IncludeDirectories(bIncludeDirectories As Boolean)
  On Local Error GoTo ErrCall:
  '
  '\\ Determines if operation affects directories when wildcards are specified
  '
  blnIncDirs = bIncludeDirectories
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_IncludeDirectories_Let")
End Property

Public Property Get IncludeDirectories() As Boolean
  On Local Error GoTo ErrCall:
  '
  '\\ Determines if operation affects directories when wildcards are specified
  '
  IncludeDirectories = blnIncDirs
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_IncludeDirectories_Get")
End Property

Public Property Let ParentWnd(hParentWnd As Long)
  On Local Error GoTo ErrCall:
  '
  lngParentWnd = hParentWnd
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_ParentWnd_Let")
End Property

Public Property Get ParentWnd() As Long
  On Local Error GoTo ErrCall:
  '
  ParentWnd = lngParentWnd
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_ParentWnd_Get")
End Property
Public Property Let RenameOnCollision(bRenameOnCollision As Boolean)
  On Local Error GoTo ErrCall:
  '
  blnRenameCollision = bRenameOnCollision
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_RenameOnCollision_Let")
End Property

Public Property Get RenameOnCollision() As Boolean
  On Local Error GoTo ErrCall:
  '
  RenameOnCollision = blnRenameCollision
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_RenameOnCollision_Get")
End Property


Public Property Let SilentMode(bSilentMode As Boolean)
  On Local Error GoTo ErrCall:
  '
  blnSilent = bSilentMode
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_SilentMode_Let")
End Property

Public Property Get SilentMode() As Boolean
  On Local Error GoTo ErrCall:
  '
  SilentMode = blnSilent
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_SilentMode_Get")
End Property


Public Function MoveFiles() As Boolean
  On Local Error GoTo ErrCall:
  '
  MoveFiles = DoOperation(FO_MOVE)
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_MoveFiles")
End Function

Public Function CopyFiles() As Boolean
  On Local Error GoTo ErrCall:
  '
  CopyFiles = DoOperation(FO_COPY)
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_CopyFiles")
End Function

Public Function DeleteFiles() As Boolean
  On Local Error GoTo ErrCall:
  '
  DeleteFiles = DoOperation(FO_DELETE)
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_DeleteFiles")
End Function

Public Function RenameFiles() As Boolean
  On Local Error GoTo ErrCall:
  '
  RenameFiles = DoOperation(FO_RENAME)
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_RenameFiles")
End Function

Private Sub PackVerInfo(ByVal strVersion As String, sVerInfo As VERINFO)
  '\\ Parses file version number string of the form x[.x[.x[.x]]] and assigns the
  '\\ extracted numbers to the appropriate elements of a VERINFO type variable.
  '
  '\\ Local Declarations
  Dim intOffset As Integer
  Dim intAnchor As Integer
  '
  On Local Error GoTo ErrCall
  '
  intOffset = InStr(strVersion, ".")
  If intOffset = 0 Then
    sVerInfo.nMSHi = Val(strVersion)
    GoTo PVIMSLo
  Else
    sVerInfo.nMSHi = Val(Left$(strVersion, intOffset - 1))
    intAnchor = intOffset + 1
  End If
  '
  intOffset = InStr(intAnchor, strVersion, ".")
  If intOffset = 0 Then
    sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor))
    GoTo PVILSHi
  Else
    sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
    intAnchor = intOffset + 1
  End If
  '
  intOffset = InStr(intAnchor, strVersion, ".")
  If intOffset = 0 Then
    sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor))
    GoTo PVILSLo
  Else
    sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
    intAnchor = intOffset + 1
  End If
  '
  intOffset = InStr(intAnchor, strVersion, ".")
  If intOffset = 0 Then
    sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor))
  Else
    sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
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
  On Local Error GoTo ErrCall:
  '
  Set FilesSource = Nothing
  '
  Exit Sub
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_ClearSourceFiles")
End Sub


Public Sub AddSourceFile(sFilename As String)
  On Local Error GoTo ErrCall:
  '
  FilesSource.Add sFilename
  '
  Exit Sub
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_AddSourceFile")
End Sub

Private Function GetFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
  '\\ [fIsRemoteServerSupportFile] - True if file is a remote OLE automation server
  '\\ support file (.VBR)
  '
  '\\ Local Declarations
  Const strFIXEDFILEINFO$ = "\"
  Dim lVerSize As Long
  Dim lVerHandle As Long
  Dim lpBufPtr As Long
  Dim byteVerData() As Byte
  '
  GetFileVerStruct = False
  '
  '\\ Initial Validation
  If IsMissing(fIsRemoteServerSupportFile) Then fIsRemoteServerSupportFile = False
  '
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
           lmemcpy sVerInfo, lpBufPtr, lVerSize
           'CopyMemory sVerInfo, lpBufPtr, lVerSize
           GetFileVerStruct = True
         End If
      End If
    End If
  End If
End Function
Private Function GetRemoteSupportFileVerStruct(ByVal strFilename As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
  '\\ Retrieves file version information of a remote OLE server support file into a
  '\\ VERINFO data structure. These files do not possess a Windows version stamp, but
  '\\ they do have an internal version stamp that can be read.
  '
  '\\ Local Declarations
  Const strVersionKey = "Version="
  Dim cchVersionKey As Integer
  Dim iFile As Integer
  '
  On Local Error GoTo ErrCall:
  '
  cchVersionKey = Len(strVersionKey)
  sVerInfo.nMSHi = gintNOVERINFO
  '
  iFile = FreeFile
  Open strFilename For Input Access Read Lock Read Write As #iFile
  '
  While (Not EOF(iFile))
    '\\ Local Declarations
    Dim strLine As String
    Line Input #iFile, strLine
    If Left$(strLine, cchVersionKey) = strVersionKey Then
      '\\ Version Key found.  Retrieve everything after equals sign.
      '\\ Local Declarations
      Dim strVersion As String
      strVersion = Mid$(strLine, cchVersionKey + 1)
      '\\ Parse Version Info.
      PackVerInfo strVersion, sVerInfo
      '\\ Convert the format 1.2.3 from the .VBR into '1.2.0.3
      sVerInfo.nLSLo = sVerInfo.nLSHi
      sVerInfo.nLSHi = 0
      '
      GetRemoteSupportFileVerStruct = True
      Close iFile
      Exit Function
    End If
  Wend
  '
  Close iFile
  '
  Exit Function
  '
ErrCall:
  GetRemoteSupportFileVerStruct = False
End Function
Public Function Exists(ByVal sFileAddr As String) As Boolean
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim WFD As WIN32_FIND_DATA
  '
  If Right$(sFileAddr, 1) = "\" Then sFileAddr = Left$(sFileAddr, Len(sFileAddr) - 1)
  '
  lAPIRes = FindFirstFile(sFileAddr, WFD) '\\ lAPIRes = file handle
  Exists = lAPIRes <> INVALID_FILE_HANDLE
  '
  FindClose lAPIRes
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_FileExists")
End Function
Public Sub ClearFilesDest()
  On Local Error GoTo ErrCall:
  '
  Set FilesDest = Nothing
  '
  Exit Sub
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_ClearFilesDest")
End Sub

Public Sub AddDestFile(sFilename As String)
  On Local Error GoTo ErrCall:
  '
  FilesDest.Add sFilename
  '
  Exit Sub
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_AddDestFile")
End Sub


Private Sub Class_Initialize()
  On Local Error GoTo ErrCall:
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
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_evtInitialize")
End Sub

Private Function DoOperation(wFunc As Integer) As Boolean
  On Local Error GoTo ErrCall:
  '
  Dim n As Long, ptrTmp As Long
  Dim sfoClass As SHFILEOPSTRUCT
  Dim aryByte() As Byte
  Dim Bfr1() As Byte, Bfr2() As Byte, Bfr3() As Byte
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
    If FilesSource.Count = 0 Then Err.Raise vbObjectError + 1000, , "No source files specified file operation"
    For n = 1 To FilesSource.Count
      .pFrom = .pFrom & FilesSource(n) & Chr$(0)
    Next n
    '
    '\\ Build "To" String
    For n = 1 To FilesDest.Count
      .pTo = .pTo & FilesDest(n) & Chr$(0)
    Next n
    If FilesDest.Count > 1 Then .fFlags = .fFlags Or FOF_MULTIFilesDest
    '
    '\\ Note: Windows packs the SHFILEOPSTRUCT structure but VB
    '\\   does not. Therefore, all members following the two-byte
    '\\   fFlags member are offset by 2 bytes. To compensate,
    '\\   structure members are copied to a byte array with the
    '\\   proper alignment which is passed to SHFileOperation.
    ReDim aryByte(LenB(sfoClass) - 2)
    CopyMemory aryByte(0), .hWnd, Len(.hWnd)
    CopyMemory aryByte(4), .wFunc, Len(.wFunc)
    '\\ Variable-length Strings Require Additonal Processing
    Bfr1 = StrConv(.pFrom & Chr$(0), vbFromUnicode)
    ptrTmp = VarPtr(Bfr1(0))
    CopyMemory aryByte(8), ptrTmp, LenB(ptrTmp)
    Bfr2 = StrConv(.pTo & Chr$(0), vbFromUnicode)
    ptrTmp = VarPtr(Bfr2(0))
    CopyMemory aryByte(12), ptrTmp, LenB(ptrTmp)
    CopyMemory aryByte(16), .fFlags, Len(.fFlags)
    CopyMemory aryByte(18), .fAnyOperationsAborted, Len(.fAnyOperationsAborted)
    CopyMemory aryByte(22), .hNameMappings, Len(.hNameMappings)
    Bfr3 = StrConv(.lpszProgressTitle & Chr$(0), vbFromUnicode)
    ptrTmp = VarPtr(Bfr3(0))
    CopyMemory aryByte(26), ptrTmp, LenB(ptrTmp)
    '
    '\\ Call SHFileOperation
    n = SHFileOperation(aryByte(0))
    '\\ Retrieve fAnyOperationsAborted Flag
    CopyMemory .fAnyOperationsAborted, aryByte(18), Len(.fAnyOperationsAborted)
    '\\ Return True If SHFileOperation Succeeded And No Operations Were Aborted
    DoOperation = Not CBool(n Or .fAnyOperationsAborted)
    '
  End With
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_DoOperation")
End Function

Public Property Get WindowsPath() As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim sPathTmp As String * 512
  '
  lAPIRes = GetWindowsDirectory(sPathTmp, 512)
  WindowsPath = Left$(sPathTmp, lAPIRes)
  WindowsPath = WindowsPath & IIf(Right$(WindowsPath, 1) = "\", csZeroLen, "\")
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
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_WinDir_Get")
End Property
Public Property Get SystemPath() As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim sPathTmp As String * 512
  '
  lAPIRes = GetSystemDirectory(sPathTmp, 512)
  SystemPath = Left$(sPathTmp, lAPIRes)
  SystemPath = SystemPath & IIf(Right$(SystemPath, 1) = "\", csZeroLen, "\")
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
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_WinSysDir_Get")
End Property

Public Function ConvertPath(sDirSrc As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim sDirDest As String, sDirBranch As String, sBranchExt As String
  Dim iSrcLen As Integer, iBranchLen As Integer
  Dim iLocCur As Integer, iLocNext As Integer, iLocBranch As Integer
  '
  iSrcLen = Len(sDirSrc)
  iLocCur = InStr(1, sDirSrc, "\", 1)
  sDirDest = Left$(sDirSrc, iLocCur - 1)
  iLocCur = iLocCur + 1
  Do Until iLocCur > iSrcLen
    iLocNext = InStr(iLocCur, sDirSrc, "\", 1)
    iLocNext = IIf(iLocNext > iLocCur, iLocNext, iSrcLen + 1)
    sDirBranch = Mid$(sDirSrc, iLocCur, iLocNext - iLocCur)
    iBranchLen = Len(sDirBranch)
    If iBranchLen > 8 Then
      If iLocNext = iSrcLen + 1 Then
        iLocBranch = IIf(InStr(1, sDirBranch, ".", 1) > 0, InStr(1, sDirBranch, ".", 1) + 1, iBranchLen)
        sBranchExt = Right$(sDirBranch, iBranchLen - iLocBranch)
        sDirBranch = Left$(sDirBranch, iBranchLen - Len(sBranchExt))
      End If
      If iBranchLen > 8 Then sDirBranch = Left$(sDirBranch, 6) & "~1"
      If iLocNext = iSrcLen + 1 Then sDirBranch = sDirBranch + sBranchExt
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
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_ConvertPath")
End Function
Public Function Execute(frmCur As Form, sTarget As String, sParam As String, sPath As String, lWinState As Long) As Long
  On Local Error GoTo ErrCall:
  '
  lAPIRes = ShellExecute(frmCur.hWnd, "Open", sTarget, sParam, sPath, lWinState)
  '
  Execute = lAPIRes
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_Execute")
End Function

Public Function GetFileVersion(ByVal strFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
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
  Dim sVerInfo As VERINFO
  Dim strVer As String
  '
  On Local Error GoTo ErrCall:
  '
  '\\ Initial Validation
  If IsMissing(fIsRemoteServerSupportFile) Then fIsRemoteServerSupportFile = False
  '
  If GetFileVerStruct(strFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
    strVer = Format$(sVerInfo.nMSHi) & "." & Format$(sVerInfo.nMSLo) & "."
    strVer = strVer & Format$(sVerInfo.nLSHi) & "." & Format$(sVerInfo.nLSLo)
    GetFileVersion = strVer
  Else
    GetFileVersion = csZeroLen
  End If
  '
  Exit Function
  '
ErrCall:
  GetFileVersion = csZeroLen
  Err = 0
End Function

Public Sub WriteINIEntry(sSection As String, sKey As String, sValue As String, sFile As String)
  On Local Error GoTo ErrCall:
  '
  WritePrivateProfileString sSection, sKey, sValue, sFile
  '
  Exit Sub
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_WriteINIEntry_Get")
End Sub

Public Sub CreateSubDirectory(ByVal sDirNew As String)
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim saMain As SECURITY_ATTRIBUTES
  Dim sDrive As String
  Dim sDirNewCreate  As String
  Dim sDirSeg As String
  Dim sDirSegs() As String
  Dim iStrPos As Integer
  Dim liN As Integer
  '
  If Right$(sDirNew, 1) <> "\" Then sDirNew = sDirNew & "\"
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
    ReDim Preserve sDirSegs(0 To liN) As String
    If liN = 0 Then sDirSeg = sDrive & sDirSeg
    sDirSegs(liN) = sDirSeg
    liN = liN + 1
  Loop
  '
  liN = -1
  Do
    liN = liN + 1
    sDirNewCreate = sDirNewCreate & sDirSegs(liN)
    saMain.nLength = LenB(saMain)
    CreateDirectory sDirNewCreate, saMain
  Loop Until liN = UBound(sDirSegs)
  '
  Exit Sub
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_CustomText_Get")
End Sub

Public Function StripPath(sFileAddr As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer
  Dim iPosCur As Integer, iPosFinal As Integer
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
    StripPath = Right$(sFileAddr, Len(sFileAddr) - iPosFinal)
  End If
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_StripPath")
End Function
Public Function StripFile(sFileAddr As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer
  Dim iPosCur As Integer, iPosFinal As Integer
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
    StripFile = Left$(sFileAddr, iPosFinal)
  End If
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_StripNameExtension")
End Function

Private Function fsLeadingZeroFormat(sSrc As String, Optional bZeros As Byte) As String
  On Local Error GoTo ErrCall:
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
  fsLeadingZeroFormat = Format(sSrc, sMask)
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_fsLeadingZeroFormat")
End Function
Public Function EnsureUniqueFile(sFileAddr As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN   As Integer
  Dim sPath As String
  Dim sFile As String
  Dim sName As String
  Dim sExt  As String
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
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_UniqueFileName")
End Function
Public Function StripExtension(psFile As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN       As Integer
  Dim iPosCur   As Integer
  Dim iPosFinal As Integer
  '
  iPosCur = InStr(1, psFile, ".", 1)
  Do Until iPosCur = 0
    iPosFinal = iPosCur
    iPosCur = InStr(1 + iPosFinal, psFile, ".", 1)
  Loop
  '
  If iPosFinal > 0 Then
    StripExtension = Left$(psFile, iPosFinal - 1)
  Else
    StripExtension = psFile
  End If
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_StripExtension")
End Function

Public Function IsolateExtension(sFile As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer
  Dim iPosCur As Integer, iPosFinal As Integer
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
    IsolateExtension = Right$(sFile, Len(sFile) - (iPosFinal - 1))
  End If
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_IsolateExtension")
End Function

Public Function EnsureExtension(sFileOrFileAddr As String, sExt As String)
  On Local Error GoTo ErrCall:
  '
  If Left$(sExt, 1) <> "." Then sExt = "." & sExt
  If Right$(sFileOrFileAddr, 4) <> sExt Then sFileOrFileAddr = sFileOrFileAddr & sExt
  '
  EnsureExtension = sFileOrFileAddr
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_EnsureExtension")
End Function

Public Function IsolatePath(sFileAddr As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer
  Dim iPosCur As Integer, iPosFinal As Integer
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
    IsolatePath = Left$(sFileAddr, iPosFinal)
  End If
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_IsolatePath")
End Function

Public Function IsolateName(ByVal sFile As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer
  Dim iPosCur As Integer, iPosFinal As Integer
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
    IsolateName = Left$(sFile, iPosFinal - 1)
  End If
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_IsolateName")
End Function

Public Function IsolateFile(sFileAddr As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer
  Dim iPosCur As Integer, iPosFinal As Integer
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
    IsolateFile = Right$(sFileAddr, Len(sFileAddr) - iPosFinal)
  End If
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_IsolateFile")
End Function

Public Function StripName(sFile As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer
  Dim iPosCur As Integer, iPosFinal As Integer
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
    StripName = Right$(sFile, Len(sFile) - (iPosFinal - 1))
  End If
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_StripName")
End Function



Public Function IsFile(sFileAddr As String) As Boolean
  On Local Error GoTo ErrCall:
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
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_IsFile")
End Function

Public Function CreateFileID(Optional bYrDgts As Byte, Optional sDlmtr As String) As String
  On Error GoTo ErrCall
  '
  If bYrDgts = 0 Then bYrDgts = 4
  If sDlmtr = csZeroLen Then sDlmtr = "."
  '
  If bYrDgts = 4 Then
    CreateFileID = Format(Now, "yyyymmdd" & sDlmtr & "hhnnss")
  Else
    CreateFileID = Format(Now, "yymmdd" & sDlmtr & "hhnnss")
  End If
  '
  Exit Function
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_CreateFileID")
End Function

Private Function fsDirParse(sSource As String, sDelimiter As String) As String
  On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim liN As Integer
  Dim sSeg As String
  '
  liN = 1
  '
  Do
    If Mid$(sSource, liN, 1) = sDelimiter Then
      sSeg = Mid$(sSource, 1, liN)
      sSource = Mid$(sSource, liN + 1, Len(sSource))
      fsDirParse = sSeg
      Exit Function
    End If
    liN = liN + 1
  Loop
  '
  Exit Function
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_fsDirParse")
End Function

Public Sub SplitPathFile(psFilename As String, psPath As String, psFile As String)
  On Error GoTo ErrCall
  '
  Dim iPos As Integer
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
  psFile = Mid$(psFilename, iPos + 1)
  psPath = Mid$(psFilename, 1, Len(psFilename) - Len(psFile))
  '
  Exit Sub
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbOKOnly & vbCritical, "ERROR: clsFileOps_SplitPathFile")
End Sub
Private Function fstrErrMsg(errno As Long) As String
  fstrErrMsg = Error(errno)
End Function

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
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFileOps.FullPath", vbCritical, "Error"
End Function

Public Function MkDirAll(psPath As String) As Boolean
  On Error GoTo ErrCall
  '
  Dim iPos As Integer
  Dim sPath As String
  Dim sOriginalPath As String
  '
  sPath = Me.IsolatePath(psPath)
  '
  If Len(sPath) > 0 Then
    On Error Resume Next
    MkDir sPath
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
            MkDir sOriginalPath
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
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFileOps.MkDir.", vbCritical, "Error"
End Function

Public Property Get DesktopPath() As String
  DesktopPath = WindowsPath & "Desktop\"
End Property

Public Property Get TimeDateStamp() As String
  TimeDateStamp = Year(Date) & "."
  TimeDateStamp = TimeDateStamp & IIf(Len(Month(Date)) < 2, "0" & Month(Date), Month(Date)) & "."
  TimeDateStamp = TimeDateStamp & IIf(Len(Day(Date)) < 2, "0" & Day(Date), Day(Date)) & "."
  TimeDateStamp = TimeDateStamp & IIf(Len(Hour(Time)) < 2, "0" & Hour(Time), Hour(Time)) & "."
  TimeDateStamp = TimeDateStamp & IIf(Len(Minute(Time)) < 2, "0" & Minute(Time), Minute(Time)) & "."
  TimeDateStamp = TimeDateStamp & IIf(Len(Second(Time)) < 2, "0" & Second(Time), Second(Time))
End Property



Public Property Get TemporaryPath() As String
On Local Error GoTo ErrCall:
  '
  '\\ Local Declarations
  Dim sPathTmp As String * 512
  '
  lAPIRes = GetTempPath(512, sPathTmp)
  TemporaryPath = Left$(sPathTmp, lAPIRes)
  TemporaryPath = EnsureTrailingBackslash(TemporaryPath)
  '
  Exit Property
  '
ErrCall:
  iDlgRsp = MsgBox(fstrErrMsg((Err.Number)), vbCritical + vbOKOnly, "ERROR: clsFileOps_WinDir_Get")
End Property
