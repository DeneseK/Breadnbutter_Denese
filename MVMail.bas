Attribute VB_Name = "MVMail"
Option Explicit
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public CN As New ADODB.Connection
Public iFlasher As Integer
Public choice As Integer
Public StrUser As String
Public sFromAddress As String
Public sEmailAddress As String
Public sSubject As String
Public sBody As String
Public sPhone As String
Public sCaller As String
Public sReceived As String
Public sMessageID As Long
Public sMessageName As String
Public StrGroups As String
Public iGroupNumber As Integer
Public RefreshSpeed As Integer
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
Public iContact As Long
Public iCompany As Long
Public sContact As String
'
'
'
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As _
String, ByVal lpszFile As String, ByVal lpszParams As String, _
ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Public Declare Function GetDesktopWindow Lib "USER32" () As Long

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
Public strKey As String
Public DeleteDays As Integer
      
Public Sub InitializeVmail()
'
  Dim TSSettings As TextStream
  Dim fso As New FileSystemObject
  Dim rs As New ADODB.Recordset
  Dim i As Integer
  '
  'RefreshSpeed = 30000

  Dim rsUser As New ADODB.Recordset
  '
    rsUser.Open "select * from tblEmployees", cnMain, adOpenKeyset, adLockBatchOptimistic
    '
    With rsUser
      Do While Not .eof
        If LCase(StrUser) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
          If iGroupNumber = Null Then
            !Groups = 15
            .UpdateBatch
            iGroupNumber = 15
          Else
            iGroupNumber = !Groups
          End If
        End If
        .MoveNext
      Loop
      .Close
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
  rs.Open "SELECT * FROM TVMailSettings", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  GetLastUpdate = rs!LastUpdateTime & " " & rs!LastUpdateDate
  '
  rs.Close
  '
  Set rs = Nothing
End Function

Public Function GetMessageRecord(ByVal plMessageID As Long) As ADODB.Recordset
  Dim rs As New ADODB.Recordset
  '
  Dim sQuery As String
  '
  sQuery = "SELECT [MessageID], [Group], [MessageName], [PhoneNumber], " & _
              "[From], [Subject], " & _
              "[DateReceived], [TimeReceived], " & _
              "[MessageSize], [Completed], " & _
              "[User], [Caller], " & _
              "[Comments], [DateCompleted], " & _
              "[TimeCompleted], " & _
              "[FromAddress], " & _
              "[Body], [Checked] " & _
            "From TVMailMessages WHERE MessageID = " & plMessageID
  '
  rs.Open sQuery, cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  Set GetMessageRecord = rs
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
              "[Checked] " & _
            "From TVMailMessages "
'    qsource = "SELECT [MessageID], [Group], [MessageName], [PhoneNumber], " & _
'              "[From], [Subject], " & _
'              "[DateReceived], [TimeReceived], " & _
'              "[MessageSize], [Completed], " & _
'              "[User], [Caller], " & _
'              "[Comments], [DateCompleted], " & _
'              "[TimeCompleted], " & _
'              "[FromAddress], " & _
'              "[Body], [Checked] " & _
'            "From TVMailMessages "
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
  
  rs.Open qsource, cnMain, adOpenKeyset, adLockBatchOptimistic
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
  Dim Color As Variant
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
        Loop Until .eof
        list.ColumnHeaders.Add , "w1" & FieldPos, .Fields(FieldPos).Name, 400 + ((TotalCharacters / .RecordCount) * 100)
        FieldPos = FieldPos + 1
      Loop Until FieldPos = .Fields.Count
      .MoveFirst
      FieldPos = 0
      LineCount = 0
      Do Until .eof
          If !Completed = True Then
            Color = vbBlack
          Else
            Color = &H80&      'vbRed
          End If
          strKey = "r" & .Fields(FieldPos)
          list.ListItems.Add , strKey, Trim(.Fields(FieldPos) & vbNullString)
          list.ListItems.Item(strKey).ForeColor = Color
          FieldPos = FieldPos + 1
          Do
            list.ListItems.Item(strKey).ListSubItems.Add(, , .Fields(FieldPos) & vbNullString).ForeColor = Color
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

Public Sub PlayTextFile(strFilename As String)
  Dim r As Long
Dim msg As String
'Dim StrFileName As String

          r = StartDoc(App.Path & "\" & strFilename)
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
Public Sub ClosePlayer(strFilename As String)
    Dim r As Long
    Dim msg As String
    '
    r = CloseDoc(App.Path & "\" & strFilename)
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
Public Sub PlaySound(strFilename As String)
Dim r As Long
Dim msg As String
'Dim StrFileName As String

          r = StartDoc(App.Path & "\Temp\" & strFilename) '(strDatapath & "messages\" & strFilename)
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
  'On Error GoTo EH
  '
  Dim iCount As Long
  Dim strKey As String
  Dim LineCount As Long
  Dim FieldPos As Long
  Dim TotalCharacters As Long
  Dim Color As Variant
  Dim pos As Variant
  Dim sTemp As String
  Dim sGroupChecker1 As String
  Dim sGroupChecker2 As String
  Dim sGroupChecker3 As String
  Dim sGroupChecker4 As String
  Dim iTempGroup As Integer
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
        Loop Until .eof
        '
        'If (400 + ((TotalCharacters / .RecordCount) * 100)) > 491504 Then TotalCharacters = ((491504 * .RecordCount) / 100) - 400
        '
        list.ColumnHeaders.Add , "w1" & FieldPos, .Fields(FieldPos).Name, 400 + ((TotalCharacters / .RecordCount) * 100)
        FieldPos = FieldPos + 1
      Loop Until FieldPos = .Fields.Count - 2

      .MoveFirst
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
        
        If (!Group = sGroupChecker1) Or (!Group = sGroupChecker2) Or (!Group = sGroupChecker3) Or (!Group = sGroupChecker4) Or (iGroupNumber = 15) Then
        ' If Right$(.Fields(0), 3) = "WAV" Then
            If !Completed = True Then
              Color = vbBlack
            Else
              If Right$(.Fields(2) & "", 3) = "WAV" Or Right$(.Fields(2) & "", 3) = "wav" Then
                Color = &H80&      'vbRed
              Else
                Color = &H8000000D
              End If
                
            End If
            strKey = "r" & .Fields(0)
            list.ListItems.Add , strKey, .Fields(FieldPos) & vbNullString
            list.ListItems.Item(strKey).ForeColor = Color
            FieldPos = FieldPos + 1
            Do
              list.ListItems.Item(strKey).ListSubItems.Add(, , .Fields(FieldPos) & vbNullString).ForeColor = Color
              FieldPos = FieldPos + 1
            Loop Until FieldPos = .Fields.Count - 2
        '------------------
        '  Debug.Print !Checked
            If IsNull(!Checked) Then
               ' list.ListItems.Item.Checked = False
                list.ListItems.Item(strKey).Checked = False
            Else
                If !Checked = True Then
              '  Debug.Print !MessageID
                   
                   ' list.ListItems.Item.Checked = True
                    list.ListItems.Item(strKey).Checked = True
                Else
                    If !Checked = False Then
                    'Debug.Print !MessageID
                       ' list.ListItems.Item.Checked = False
                        list.ListItems.Item(strKey).Checked = False
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

'
If list.ListItems.Count = 0 Then
  With FVMail
    .txtBody = ""
    .txtPhone = ""
    .txtsubject = ""
    .chkComp.Value = 0
    .cmbCaller.Clear
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

Function StartDoc(DocName As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, _
    "", "C:\", SW_SHOWNORMAL)
End Function

Function CloseDoc(DocName As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    CloseDoc = ShellExecute(Scr_hDC, "Close", DocName, _
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
  GetFirstName
  rsContact.Open "Select [FirstName], [LastName], [ID],[ComPanyID] from TContact Where [FirstName] = '" & sFirstName & "'", cnMain, adOpenDynamic, adLockBatchOptimistic
  With rsContact
    While Not .eof
      If sContact = !FirstName & " " & !LastName Then
        iContact = !ID
        iCompany = !CompanyID
      End If
      .MoveNext
    Wend
  End With
  '
End Sub

Public Sub GetFirstName()
Dim i As Integer
Dim sLetter As String
  '
  sFirstName = ""
  i = 1
  sLetter = Mid(sContact, i, 1)
  While Not sLetter = " "
    sFirstName = sFirstName & sLetter
    i = i + 1
    sLetter = Mid(sContact, i, 1)
  Wend
  If sFirstName = "Dr." Then
    sFirstName = sFirstName & " "
    GetRestOfName
  End If
  '
End Sub

Public Sub GetRestOfName()
Dim i As Integer
Dim sLetter As String
  '
  i = 5
  sLetter = Mid(sContact, i, 1)
  While Not sLetter = " "
    sFirstName = sFirstName & sLetter
    i = i + 1
    sLetter = Mid(sContact, i, 1)
  Wend
  '
End Sub
