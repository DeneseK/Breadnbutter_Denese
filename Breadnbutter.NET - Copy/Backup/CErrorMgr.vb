Option Strict Off
Option Explicit On
Friend Class CErrorMgr
	
	Public Function Raise(ByRef psModule As String, ByRef psProcedure As String, ByRef piErrID As Short, ByRef psErrDescript As String, Optional ByRef plErrHelpID As Integer = 0, Optional ByRef piButtons As Short = 0, Optional ByRef psModuleEasy As String = "", Optional ByRef psProcedureEasy As String = "") As String
		On Error GoTo ErrCall
		'
		Dim iFileTmp As Short
		Dim sFeedback As String
		'
		iFileTmp = FreeFile
		'
		If psModuleEasy = "" Then psModuleEasy = psModule
		If psProcedureEasy = "" Then psProcedureEasy = psProcedure
		'
		sFeedback = InputBox("Error in " & psModuleEasy & " " & psProcedureEasy & ":" & vbCrLf & vbCrLf & "Error Number " & piErrID & vbCrLf & psErrDescript & vbCrLf & vbCrLf & "Please describe what happened when this error occurred.", "Error")
		'
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		FileOpen(iFileTmp, FileOps.FullPath(My.Application.Info.DirectoryPath) & My.Application.Info.AssemblyName & " Error Log.txt", OpenMode.Append)
		PrintLine(iFileTmp, "[" & VB6.Format(Now, "Long Date") & " " & VB6.Format(Now, "Medium Time") & "]")
		PrintLine(iFileTmp, psModule & " " & psProcedure)
		PrintLine(iFileTmp, "Error #:" & piErrID & "  " & psErrDescript)
		PrintLine(iFileTmp, sFeedback)
		PrintLine(iFileTmp, "")
		FileClose(iFileTmp)
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsErrMgr.Raise.", MsgBoxStyle.Critical, "Error")
	End Function
	
	Public Sub LogClear()
		'\\ Local Declarations
		Dim iFileTmp As Short
		'
		On Error Resume Next
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		Kill(FileOps.FullPath(My.Application.Info.DirectoryPath) & My.Application.Info.AssemblyName & " Error Log.txt")
		'
		On Error GoTo ErrCall
		iFileTmp = FreeFile
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		FileOpen(iFileTmp, FileOps.FullPath(My.Application.Info.DirectoryPath) & My.Application.Info.AssemblyName & " Error Log.txt", OpenMode.Append)
		FileClose(iFileTmp)
		'
		' CSErrorHandler begin - please do not modify or remove this line
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsErrMgr.LogClear.", MsgBoxStyle.Critical, "Error")
	End Sub
	
	Public ReadOnly Property LogEntries() As Integer
		Get
			On Error GoTo ErrCall
			'
			'\\ Local Declarations
			Dim iFileTmp As Short
			Dim sTmp As String
			'
			iFileTmp = FreeFile
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			FileOpen(iFileTmp, FileOps.FullPath(My.Application.Info.DirectoryPath) & My.Application.Info.AssemblyName & " Error Log.txt", OpenMode.Input)
			Do Until EOF(iFileTmp)
				Input(iFileTmp, sTmp)
				If InStr(1, sTmp, "[", 1) > 0 Then LogEntries = LogEntries + 1
			Loop 
			FileClose(iFileTmp)
			'
			' CSErrorHandler begin - please do not modify or remove this line
			Exit Property
ErrCall: 
			If Err.Number = 53 Then
				LogEntries = 0
			Else
				MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsErrMgr.LogEntries.", MsgBoxStyle.Critical, "Error")
			End If
		End Get
	End Property
	
	Public Sub LogComplete(ByRef sAct As Object)
		On Error GoTo ErrCall
		'
		'\\ Local Declarations
		Dim iFileTmp As Short
		'
		iFileTmp = FreeFile
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		FileOpen(iFileTmp, FileOps.FullPath(My.Application.Info.DirectoryPath) & My.Application.Info.AssemblyName & " Error Log.txt", OpenMode.Append)
		'UPGRADE_WARNING: Couldn't resolve default property of object sAct. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PrintLine(iFileTmp, IIf(sAct <> vbNullString, sAct, "N/A"))
		PrintLine(iFileTmp, GetSetting(My.Application.Info.Title, "Miscellaneous", "ErrRes", "c") & vbCrLf)
		FileClose(iFileTmp)
		'
		' CSErrorHandler begin - please do not modify or remove this line
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsErrMgr.LogComplete.", MsgBoxStyle.Critical, "Error")
	End Sub
End Class