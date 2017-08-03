VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContactList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function FillList(ByRef pContactList As ListView, rs As Recordset) As Long
  On Error GoTo EH
  '
  'list = pContactList
  Dim iStartField As Integer
  Dim strKey As String
  Dim iLineCount As Integer
  Dim iFieldPos As Integer
  Dim iTotalCharacters As Long
  Dim iIcon As Integer
  Dim lColor As Long
  Dim sColumnName As String
  '
  Dim sPreviousText As String
  '
  iStartField = 3
  '
  If Not (pContactList.SelectedItem Is Nothing) Then
    sPreviousText = pContactList.SelectedItem.Key
  End If
  '
  pContactList.ListItems.Clear
  pContactList.ColumnHeaders.Clear
  '
  iTotalCharacters = 0
  iLineCount = 0
  iFieldPos = iStartField
  With rs
    If .RecordCount > 0 Then
      Do
        '
        iTotalCharacters = 0
        .MoveFirst
        Do
          iTotalCharacters = Len(CStr(.Fields(iFieldPos) & vbNullString)) + iTotalCharacters
          .MoveNext
        Loop Until .eof
        '
'        Select Case .Fields(iFieldPos).Name
'         Case "LastName"
'          sColumnName = "Last"
'         Case "FirstName"
'          sColumnName = "First"
'         Case "AuthRemaining"
'          sColumnName = "Days"
'        Case Else
          sColumnName = .Fields.Item(iFieldPos).Name
'        End Select
        '
        pContactList.ColumnHeaders.Add , "w1" & iFieldPos, sColumnName, 400 + ((iTotalCharacters / .RecordCount) * 100)
        iFieldPos = iFieldPos + 1
      '
      Loop Until iFieldPos = .Fields.Count
      '
      .MoveFirst
      '
      iFieldPos = iStartField
      '
      iLineCount = 0
      '
      Do Until .eof
         Select Case !ContactType
          Case 0  ' unknown
            iIcon = 1
          Case 1 'adjuster
            iIcon = 2
          Case 2 'admin
            iIcon = 4
          Case 3 'tech
            iIcon = 3
          Case 4 'sec
            iIcon = 5
          Case 5 'unknown
            iIcon = 1
          Case Else
           iIcon = 2
        End Select
        '
        Select Case !Status
          '
          Case "Customer"
            lColor = &HC00000
          Case "Prospect"
            lColor = &HC000&
          Case "Future Prospect"
            lColor = &HC000&
          Case "Inactive"
            lColor = &H404040
          Case "Contact"
            lColor = &HC000C0
          Case Else
            lColor = &H0
        End Select
        '
        strKey = "ID" & Format(!ID)
        pContactList.ListItems.Add , strKey, .Fields(iFieldPos) & vbNullString, , iIcon
        pContactList.ListItems.Item(strKey).ForeColor = lColor
        iFieldPos = iFieldPos + 1
        Do
          pContactList.ListItems.Item(strKey).ListSubItems.Add(, , .Fields(iFieldPos) & vbNullString).ForeColor = lColor
          iFieldPos = iFieldPos + 1
        Loop Until iFieldPos = .Fields.Count
        iFieldPos = iStartField
        .MoveNext
        iLineCount = iLineCount + 1
      Loop
    End If
  End With
  FillList = iLineCount
'
  If sPreviousText <> "" Then
   Dim itmFound As ListItem   ' FoundItem variable.
   'Debug.Print pContactList.ListItems(sPreviousText).
  If VerifyKeyInList(pContactList, sPreviousText) Then
   Set itmFound = pContactList.ListItems(sPreviousText)
'  '
    If Not (itmFound Is Nothing) Then
       itmFound.EnsureVisible
       itmFound.Selected = True
       
  ' Set pContactList.SelectedItem = pContactList.ListItems(sPreviousText)
    End If
  End If
  End If
  '
Exit Function
EH:
 MsgBox Err.Description & " in FillList."
End Function

Private Function VerifyKeyInList(ByRef pContactList As ListView, ByVal psKey As String) As Boolean
  Dim i As Integer
  '
  VerifyKeyInList = False
  '
  For i = 1 To pContactList.ListItems.Count
    If pContactList.ListItems(i).Key = psKey Then VerifyKeyInList = True
  Next i
End Function
