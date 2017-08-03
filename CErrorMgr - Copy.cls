VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CErrorMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Raise(psModule As String, psProcedure As String, piErrID As Integer, psErrDescript As String, Optional plErrHelpID As Long, Optional piButtons As Integer, Optional psModuleEasy As String, Optional psProcedureEasy As String) As String
  On Error GoTo ErrCall:
  '
  Dim iFileTmp As Integer
  Dim sFeedback As String
  '
  iFileTmp = FreeFile
  '
  If psModuleEasy = "" Then psModuleEasy = psModule
  If psProcedureEasy = "" Then psProcedureEasy = psProcedure
  '
  sFeedback = InputBox("Error in " & psModuleEasy & " " & psProcedureEasy & ":" & vbCrLf & vbCrLf & "Error Number " & piErrID & vbCrLf & psErrDescript & vbCrLf & vbCrLf & "Please describe what happened when this error occurred.", "Error")
  '
  Open FileOps.FullPath(App.Path) & App.EXEName & " Error Log.txt" For Append As iFileTmp
    Print #iFileTmp, "[" & Format(Now, "Long Date") & " " & Format(Now, "Medium Time") & "]"
    Print #iFileTmp, psModule & " " & psProcedure
    Print #iFileTmp, "Error #:" & piErrID & "  " & psErrDescript
    Print #iFileTmp, sFeedback
    Print #iFileTmp, ""
  Close iFileTmp
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsErrMgr.Raise.", vbCritical, "Error"
End Function

Public Sub LogClear()
  '\\ Local Declarations
  Dim iFileTmp As Integer
  '
  On Error Resume Next
  Kill FileOps.FullPath(App.Path) & App.EXEName & " Error Log.txt"
  '
  On Error GoTo ErrCall
  iFileTmp = FreeFile
  Open FileOps.FullPath(App.Path) & App.EXEName & " Error Log.txt" For Append As iFileTmp
  Close iFileTmp
  '
  ' CSErrorHandler begin - please do not modify or remove this line
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsErrMgr.LogClear.", vbCritical, "Error"
End Sub

Public Property Get LogEntries() As Long
  On Error GoTo ErrCall
  '
  '\\ Local Declarations
  Dim iFileTmp As Integer
  Dim sTmp As String
  '
  iFileTmp = FreeFile
  Open FileOps.FullPath(App.Path) & App.EXEName & " Error Log.txt" For Input As iFileTmp
    Do Until EOF(iFileTmp)
      Input #iFileTmp, sTmp
      If InStr(1, sTmp, "[", 1) > 0 Then LogEntries = LogEntries + 1
    Loop
  Close iFileTmp
  '
  ' CSErrorHandler begin - please do not modify or remove this line
  Exit Property
ErrCall:
  If Err.Number = 53 Then
    LogEntries = 0
  Else
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsErrMgr.LogEntries.", vbCritical, "Error"
  End If
End Property

Public Sub LogComplete(sAct)
  On Error GoTo ErrCall
  '
  '\\ Local Declarations
  Dim iFileTmp As Integer
  '
  iFileTmp = FreeFile
  Open FileOps.FullPath(App.Path) & App.EXEName & " Error Log.txt" For Append As iFileTmp
    Print #iFileTmp, IIf(sAct <> vbNullString, sAct, "N/A")
    Print #iFileTmp, GetSetting(App.Title, "Miscellaneous", "ErrRes", "c") & vbCrLf
  Close iFileTmp
  '
  ' CSErrorHandler begin - please do not modify or remove this line
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsErrMgr.LogComplete.", vbCritical, "Error"
End Sub
