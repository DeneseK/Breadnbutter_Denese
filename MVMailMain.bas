Attribute VB_Name = "MvMailMain"
Option Explicit
Public cn As New ADODB.Connection
Public StrUser As String
Public sFromAddress As String
Public sEmailAddress As String
Public sSubject As String
Public sBody As String
Public sCaller As String
Public sReceived As String
Public sMessageName As String
Public StrGroups As String
Public sGroupNumber As Integer
Public RefreshSpeed As Integer
Dim strDatapath As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
String, ByVal lpszFile As String, ByVal lpszParams As String, _
ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const SW_SHOWNORMAL = 1

Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&

Public Const ALLCALLS = 1
Public Const NEWCALLS = 2
Public Const OLDCALLS = 3
Public OldRecordCount As Integer
Public NewRecordCount As Integer
Public FromTimer As Boolean
Public RefreshList As Boolean
Public SavedIndex As String
Public StrKey As String
Public DeleteDays As Integer
      
Public Sub Main()
'
  Dim TSSettings As TextStream
  Dim fso As New FileSystemObject
  
  '
  Set TSSettings = fso.OpenTextFile(App.Path & "\Settings.txt", ForReading, True, TristateUseDefault)
  strDatapath = TSSettings.ReadLine
  RefreshSpeed = Val(TSSettings.ReadLine)
  DeleteDays = Val(TSSettings.ReadLine)
  TSSettings.Close
  '
'  cn.CursorLocation = adUseClient
'  cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDatapath & "voice.mdb"
  cn.Provider = "SQLOLEDB"
  cn.Properties("Data Source").Value = "HR_SERVER"
  cn.Properties("Initial Catalog").Value = "BNB_DATA"
  cn.Properties("User ID").Value = "JASONSEAY"
  cn.Properties("Password").Value = "JASONSEAY"
  cn.Properties("Persist Security Info") = False
  cn.CursorLocation = adUseClient
  
  
  cn.Open
  Load FUser
  FUser.Show
'
End Sub
Public Sub Shutdown()
'
  cn.Close
'
End Sub

Public Function GetLastUpdate() As String
  Dim rs As New ADODB.Recordset
  '
  rs.Open "SELECT * FROM TVMailSettings", cn, adOpenKeyset, adLockBatchOptimistic
  '
  GetLastUpdate = rs!LastUpdateTime & " " & rs!LastUpdateDate
  '
  rs.Close
  '
  Set rs = Nothing
End Function

Public Function GetRS(ListType As Integer) As ADODB.Recordset
  Dim rs As New ADODB.Recordset
  Dim qsource As String
  '
  qsource = "SELECT [MessageID], [Group], [MessageName], [PhoneNumber], " & _
              "[From], [Subject], " & _
              "[DateReceived], [TimeReceived], " & _
              "[MessageSize], [Completed], " & _
              "[User], [Caller], " & _
              "[Comments], [DateCompleted], " & _
              "[TimeCompleted], " & _
              "[FromAddress], " & _
              "[Body] " & _
            "From TVMailMessages "
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
  
  
  qsource = qsource + " ORDER BY TVMailMessages.DateReceived DESC , TVMailMessages.TimeReceived DESC;"
  
  rs.Open qsource, cn, adOpenKeyset, adLockBatchOptimistic
  '
  NewRecordCount = rs.RecordCount
  If NewRecordCount = OldRecordCount Then
    RefreshList = False
  Else
    RefreshList = True
    OldRecordCount = NewRecordCount
  End If
  Set GetRS = rs
 ' rs.Close
  'Set rs = Nothing
End Function

Public Sub FillListOLD(rs As Recordset, list As ListView)
  On Error GoTo EH
  '
  
  Dim LineCount As Integer
  Dim FieldPos As Integer
  Dim TotalCharacters As Integer
  Dim color As Variant
  '
  list.ListItems.Clear
  list.ColumnHeaders.Clear
  '
  TotalCharacters = 0
  LineCount = 0
  FieldPos = 0
  With rs
    If .RecordCount > 0 Then
      Do
        TotalCharacters = 0
        .MoveFirst
        Do
          TotalCharacters = Len(CStr(.Fields(FieldPos) & vbNullString)) + TotalCharacters
          .MoveNext
        Loop Until .EOF
        list.ColumnHeaders.Add , "w1" & FieldPos, .Fields(FieldPos).Name, 400 + ((TotalCharacters / .RecordCount) * 100)
        FieldPos = FieldPos + 1
      Loop Until FieldPos = .Fields.Count
      .MoveFirst
      FieldPos = 0
      LineCount = 0
      Do Until .EOF
          If !Completed = True Then
            color = vbBlack
          Else
            color = &H80&      'vbRed
          End If
          StrKey = "r" & .Fields(FieldPos)
          list.ListItems.Add , StrKey, Trim(.Fields(FieldPos) & vbNullString)
          list.ListItems.Item(StrKey).ForeColor = color
          FieldPos = FieldPos + 1
          Do
            list.ListItems.Item(StrKey).ListSubItems.Add(, , .Fields(FieldPos) & vbNullString).ForeColor = color
            FieldPos = FieldPos + 1
          Loop Until FieldPos = .Fields.Count
          FieldPos = 0
          .MoveNext
          LineCount = LineCount + 1
      Loop
    End If
  End With
Exit Sub
EH:
 MsgBox Err.Description & " in FillList."
End Sub

Public Sub PlayTextFile(StrFileName As String)
  Dim r As Long
Dim msg As String
'Dim StrFileName As String

          r = StartDoc(App.Path & "\" & StrFileName)
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
              MsgBox msg
          End If
End Sub

Public Sub PlaySound(StrFileName As String)
Dim r As Long
Dim msg As String
'Dim StrFileName As String

          r = StartDoc(strDatapath & "messages\" & StrFileName)
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
              MsgBox msg
          End If
End Sub


Public Function FillList(rs As Recordset, list As ListView) As Long
  On Error GoTo EH
  '
  Dim iCount As Integer
  Dim StrKey As String
  Dim LineCount As Integer
  Dim FieldPos As Integer
  Dim TotalCharacters As Integer
  Dim color As Variant
  Dim pos As Variant
  Dim sTemp As String
  Dim sGroupChecker1 As String
  Dim sGroupChecker2 As String
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
  list.ListItems.Clear
  list.ColumnHeaders.Clear
  '
  TotalCharacters = 0
  LineCount = 0
  FieldPos = 1
  With rs
    If .RecordCount > 0 Then
      Do
        TotalCharacters = 0
        .MoveFirst
        Do
           If FieldPos = 3 Or FieldPos = 4 And .Fields(FieldPos) & vbNullString = "" Then
                sTemp = "QQQQQQQ"
              Else
                sTemp = .Fields(FieldPos) & vbNullString
              End If
          TotalCharacters = Len(CStr(sTemp)) + TotalCharacters
          .MoveNext
        Loop Until .EOF
        list.ColumnHeaders.Add , "w1" & FieldPos, .Fields(FieldPos).Name, 400 + ((TotalCharacters / .RecordCount) * 100)
        FieldPos = FieldPos + 1
      Loop Until FieldPos = .Fields.Count - 1
      .MoveFirst
      FieldPos = 1
      LineCount = 0
      '
      
        Select Case StrGroups
          Case 1
            sGroupChecker1 = "Authorizations"
            sGroupChecker2 = "Authorizations"
          Case 2
            sGroupChecker1 = "Sales"
            sGroupChecker2 = "Sales"
          Case 3
            sGroupChecker1 = "Support"
            sGroupChecker2 = "Support"
          Case 5
            sGroupChecker1 = "Authorizations"
            sGroupChecker2 = "Sales"
          Case 6
            sGroupChecker1 = "Authorizations"
            sGroupChecker2 = "Support"
          Case 7
            sGroupChecker1 = "Sales"
            sGroupChecker2 = "Support"
          Case Else
            sGroupChecker1 = "Authorizations"
            sGroupChecker2 = "Authorizations"
        End Select
      
      'For iCount = 1 To rs.RecordCount
      Do Until .EOF
        
        If (!Group = sGroupChecker1) Or (!Group = sGroupChecker2) Or (StrGroups = 4) Then
        ' If Right$(.Fields(0), 3) = "WAV" Then
            If !Completed = True Then
              color = vbBlack
            Else
              If Right$(.Fields(2) & "", 3) = "WAV" Or Right$(.Fields(2) & "", 3) = "wav" Then
                color = &H80&      'vbRed
              Else
                color = &H8000000D
              End If
                
            End If
            StrKey = "r" & .Fields(0)
            list.ListItems.Add , StrKey, .Fields(FieldPos) & vbNullString
            list.ListItems.Item(StrKey).ForeColor = color
            FieldPos = FieldPos + 1
            Do
              list.ListItems.Item(StrKey).ListSubItems.Add(, , .Fields(FieldPos) & vbNullString).ForeColor = color
              FieldPos = FieldPos + 1
            Loop Until FieldPos = .Fields.Count - 1
            FieldPos = 1
            '.MoveNext
            LineCount = LineCount + 1
        ' Else
           ' .MoveNext
        'End If
          
      'Next
        End If
        .MoveNext
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
list.Visible = True
'
Exit Function
EH:
 MsgBox Err.Description & " in FillList."
End Function

Function StartDoc(DocName As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, _
    "", "C:\", SW_SHOWNORMAL)
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

Public Function CheckForKey(pKey As String, pList As ListView) As Boolean
Dim i As Integer
CheckForKey = False
For i = 1 To pList.ListItems.Count
  If pKey = pList.ListItems.Item(i).Key Then
    CheckForKey = True
  End If
 ' pList.ListItems.Item(i).Key
Next i
End Function

