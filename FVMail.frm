VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FVMail 
   Caption         =   "VMailClient"
   ClientHeight    =   8595
   ClientLeft      =   2310
   ClientTop       =   1365
   ClientWidth     =   11400
   ControlBox      =   0   'False
   Icon            =   "FVMail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11400
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCompleted 
      Caption         =   "&Completed"
      Height          =   495
      Left            =   5040
      TabIndex        =   35
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   615
      Left            =   1920
      TabIndex        =   33
      Top             =   3480
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   -10000
      TabIndex        =   31
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details"
      Height          =   675
      Left            =   120
      TabIndex        =   16
      Top             =   3420
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      Caption         =   "Details"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   11535
      Begin SHDocVwCtl.WebBrowser webBody 
         Height          =   1455
         Left            =   4440
         TabIndex        =   36
         Top             =   600
         Width           =   6975
         ExtentX         =   12303
         ExtentY         =   2566
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Changes"
         Height          =   375
         Left            =   5640
         TabIndex        =   34
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtBody 
         Height          =   1455
         Left            =   4440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Text            =   "FVMail.frx":08CA
         Top             =   600
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.CommandButton cmdContactInfo 
         Caption         =   "Contact Info."
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   30
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox cmbCaller 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton cmdGetNames 
         Caption         =   "Get Names"
         Height          =   375
         Left            =   3000
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtPhone 
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1200
         Width           =   2760
      End
      Begin VB.TextBox txtsubject 
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   480
         Width           =   3525
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   "Forward Message"
         Height          =   345
         Left            =   9750
         TabIndex        =   25
         Top             =   195
         Width           =   1500
      End
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "View Body In Browser"
         Height          =   345
         Left            =   7830
         TabIndex        =   24
         Top             =   195
         Width           =   1815
      End
      Begin VB.ComboBox cmbComment 
         Height          =   315
         Left            =   8040
         TabIndex        =   23
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox chkComp 
         Caption         =   "Completed"
         Height          =   375
         Left            =   4440
         TabIndex        =   22
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Comment:"
         Height          =   255
         Left            =   7230
         TabIndex        =   21
         Top             =   2205
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Caller:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Phone Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Subject:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Body:"
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "D&elete"
      Height          =   735
      Left            =   9480
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbMessageGroup 
      Height          =   315
      Left            =   8820
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Timer Timer1 
      Left            =   6030
      Top             =   4440
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&All"
      Height          =   705
      Left            =   3810
      TabIndex        =   5
      Top             =   4140
      Width           =   1605
   End
   Begin VB.CommandButton cmdOld 
      Caption         =   "O&ld"
      Height          =   645
      Left            =   2580
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   555
      Left            =   1260
      TabIndex        =   3
      Top             =   4230
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   705
      Left            =   6810
      TabIndex        =   2
      Top             =   4050
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CommandButton cmdEditGroups 
      Caption         =   "Edit Groups"
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   645
      Left            =   2730
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   3450
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3285
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Group"
         Object.Width           =   866958
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   9600
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblGroups 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALL"
      Height          =   255
      Left            =   9480
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblMessageGroup 
      Caption         =   "Message Group:"
      Height          =   255
      Left            =   9540
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblLastClient 
      Caption         =   "Label1"
      Height          =   195
      Left            =   2580
      TabIndex        =   10
      Top             =   5130
      Width           =   3555
   End
   Begin VB.Label lblLastServer 
      Caption         =   "Label1"
      Height          =   225
      Left            =   7440
      TabIndex        =   9
      Top             =   5130
      Width           =   3525
   End
   Begin VB.Label lblcount 
      Caption         =   "Label1"
      Height          =   225
      Left            =   0
      TabIndex        =   8
      Top             =   5130
      Width           =   2565
   End
   Begin VB.Label lblShow 
      Caption         =   "Show:"
      Height          =   255
      Left            =   570
      TabIndex        =   7
      Top             =   4170
      Width           =   495
   End
End
Attribute VB_Name = "FVMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private rs As New ADODB.Recordset
'
Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
'
Private rsMessages As New ADODB.Recordset
'

Private Sub SaveCheckMarks(pRS As Recordset)
 Dim iListCount As Integer
 Dim i As Integer
    iListCount = ListView1.ListItems.Count
    '
    For i = 1 To iListCount
      If Not pRS.eof Then
        pRS.MoveFirst
        pRS.Find "messageID=" & CLng(Right$(ListView1.ListItems.Item(i).Key, Len(ListView1.ListItems.Item(i).Key) - 1))
        '
        If Not pRS.eof Then
          pRS!Checked = ListView1.ListItems.Item(i).Checked
        End If
      Else
        Exit For
      End If
    Next
    pRS.UpdateBatch
End Sub

Public Sub RefreshMessages()
  Dim iPosition As Integer
  Dim rs As New ADODB.Recordset
  'On Error GoTo ErrorHandler
'  If rs.State <> 0 Then
'    rs.Close
'  End If
  '
  Set rs = GetRS(choice)
  '
  SaveCheckMarks rs
  '
  If Not ListView1.SelectedItem Is Nothing Then
    SavedIndex = ListView1.SelectedItem.Key
    iPosition = ListView1.SelectedItem.Index
  End If
  '
  lblcount.Caption = "Messages Shown: " & FillList(rs, ListView1)
  '
  On Error Resume Next
  Set ListView1.SelectedItem = ListView1.ListItems(SavedIndex)
  ListView1.ListItems(iPosition).EnsureVisible
  
  'On Error GoTo ErrorHandler
  '
  If DateDiff("n", (Mid(GetLastUpdate, 1, 11)), Time) > 30 Or DateDiff("d", (Mid(GetLastUpdate, 12, 11)), Date) > 0 Then
    lblLastServer.BackColor = vbRed
    lblLastServer = "Last Server Update: " & GetLastUpdate
    lblLastServer = lblLastServer & " Server May Be Down!!! Tell Supervisor"
  Else
    lblLastServer.BackColor = &H8000000F
    lblLastServer = "Last Server Update: " & GetLastUpdate
  End If
  '
  lblLastClient = "Last Client Update: " & Time & " " & Date
  '
  rs.Close
  Set rs = Nothing
  
  Exit Sub
'ErrorHandler:
 ' MsgBox "Error filling list"
End Sub

Private Sub cmbCaller_Change()
  cmdContactInfo.Enabled = True
End Sub

Private Sub cmdAll_Click()
  cmdNew.FontBold = False
  cmdAll.FontBold = True
  cmdOld.FontBold = False
  choice = ALLCALLS
  RefreshMessages
  ListView1_Click
End Sub

Private Sub LoadEmailBodyBody()
  Dim rsBody As New ADODB.Recordset
  '
  rsBody.Open "SELECT [Body] FROM TVMailMessages WHERE MessageID = '" & CStr(sMessageID) & "'", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  With rsBody
    If Not .eof Then
      If Not !Body = vbNullString Then
        txtBody.Text = !Body & vbNullString
        sBody = !Body & vbNullString
      Else
        txtBody.Text = vbNullString
        sBody = vbNullString
      End If
    Else
      txtBody.Text = vbNullString
      sBody = vbNullString
    End If
   End With
   '
   Set rsBody = Nothing
End Sub

Private Sub DisplayEmailBody()
  Dim TSTemp As TextStream
  Dim fso As New FileSystemObject
  '
  Set TSTemp = fso.OpenTextFile(App.Path & "\temp.html", ForWriting, True, TristateUseDefault)
  TSTemp.Write txtBody.Text
  TSTemp.Close
  webBody.Navigate2 App.Path & "\temp.html"
End Sub

Private Sub cmdBrowser_Click()
  Dim TSTemp As TextStream
  Dim fso As New FileSystemObject
  '
  Set TSTemp = fso.OpenTextFile(App.Path & "\temp.html", ForWriting, True, TristateUseDefault)
  TSTemp.Write txtBody.Text
  TSTemp.Close
  PlayTextFile "temp.html"
  'webBody.Navigate2 App.Path & "\temp.html"
End Sub

'Private Sub SendToOld(oldRS As Recordset)
'  With oldRS
'      If Not !Completed Then
'        !DateCompleted = Date
'        !TimeCompleted = Time
'        !User = StrUser
'      End If
'      '
'      !Completed = True
'      On Error GoTo ErrorHandler
'    End With
'  Exit Sub
'ErrorHandler:
'    MsgBox ("Error. It's possible somebody else has changed this record since this window was opened")
'End Sub

Private Sub cmdCompleted_Click()
  Dim i As Integer
  Dim oldRS As New Recordset
  Dim iListCount As Integer
  Dim lMessageID As Long
  Dim sKey As String
  'adOpenKeyset
  oldRS.Open "SELECT messageID, Checked, Completed, TimeCompleted, DateCompleted, [User] FROM TVMailMessages WHERE Completed = " & "'False'", cnMain, adOpenDynamic, adLockBatchOptimistic
  iListCount = ListView1.ListItems.Count
  '
  For i = 1 To iListCount
   If ListView1.ListItems.Item(i).Checked Then
      sKey = Trim(ListView1.ListItems.Item(i).Key)
      lMessageID = CLng(Right$(sKey, Len(sKey) - 1))
      oldRS.MoveFirst
      oldRS.Find "messageID=" & lMessageID
        If Not oldRS.eof Then
          oldRS!Checked = False
          'SendToOld oldRS
        If Not oldRS!Completed Then
          oldRS!DateCompleted = Date
          oldRS!TimeCompleted = Time
          oldRS!User = StrUser
        End If
      '
      oldRS!Completed = True
        End If
   End If
  Next
  '
  oldRS.UpdateBatch
  oldRS.Close
  '
  RefreshMessages
  ListView1_Click
  FMain.tmrMessages_Timer
End Sub

Private Sub cmdContactInfo_Click()
'Set FormMgr = New CFormMgr
'FormMgr.Setup FMain
'  sPhone = txtPhone.Text
'  bVMail = True
'  'fmain.tbMain.ToolBars(1).Tools.item..ID = "ID_Lookup"
'  FContact.Form_Load
'  FContact.Show
'  FContact.WindowState = 2
'  FContact.txtSearch = sPhone
'
'  FMain.tbMain_Go (FMain.tbMain.ToolBars(3).Tools.Item(4).ID)
'  Unload FVMail
 'Debug.Print "tools " & FMain.tbMain.ToolBars(3).Tools.Item(4).ID
  
  sContact = cmbCaller.Text
  sContact = Trim(sContact)
  If InStr(sContact, " ") <> 0 Then
    GetContactInfo
    'Company.Fetch iCompany
    'Company.Contact.Fetch iContact
    cmbCaller.Text = ""
    FContact.LoadContact iContact, True
    FormMgr.ShowForm Me, FContact
    
  Else
    MsgBox "Please enter First and Last Name", vbInformation, "Bread 'n' Butter"
  End If
End Sub

Private Sub cmdDelete_Click()
  MsgBox "Are You Sure You want to Delete Selected record?", vbOKCancel
End Sub

Private Sub cmdDetails_Click()
  '
  ListView1_Click
End Sub


Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdForward_Click()
Dim TSTemp As TextStream
Dim fso As New FileSystemObject
  '
  Set TSTemp = fso.OpenTextFile(App.Path & "\tempMessage.txt", ForWriting, True, TristateUseDefault)
  TSTemp.WriteLine ("Received: " & sReceived)
  TSTemp.WriteLine ("From: " & sCaller)
  TSTemp.WriteLine (sBody)
  TSTemp.Close
  
  FSendTo.Show vbModal
End Sub


Private Sub cmdGetNames_Click()
  FindNames
  If Trim(cmbCaller.Text) = "" Then
    cmdContactInfo.Enabled = False
  Else
    cmdContactInfo.Enabled = True
  End If
End Sub

Private Sub cmdNew_Click()
  cmdNew.FontBold = True
  cmdAll.FontBold = False
  cmdOld.FontBold = False
  choice = NEWCALLS
  RefreshMessages
  ListView1_Click
End Sub

Private Sub cmdOld_Click()
  cmdNew.FontBold = False
  cmdAll.FontBold = False
  cmdOld.FontBold = True
  choice = OLDCALLS
  RefreshMessages
  ListView1_Click
End Sub

Private Sub CmdPlay_Click()
  Dim strStream As ADODB.Stream
  Dim FileSys As New FileSystemObject
  Dim myStr As String
  Dim rsAttach As New ADODB.Recordset
  '
  If Not FileSys.FolderExists(App.Path & "\Temp") Then
     FileSys.CreateFolder (App.Path & "\Temp")
  End If
  '
  myStr = ListView1.SelectedItem.SubItems(1) & vbNullString
  rsAttach.Open "SELECT * FROM TVMailMessages WHERE MessageID like '" & CStr(sMessageID) & "'", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  With rsAttach
    .MoveFirst
    .Find "MessageName = " & "'" & myStr & "'"
    If Not .eof Then
      If Not !MessageName = vbNullString Then
        If Right$(!MessageName, 3) = "WAV" Or Right$(!MessageName, 3) = "wav" Then
           Set strStream = New ADODB.Stream
           strStream.Type = adTypeBinary
           strStream.Open
           strStream.Write !Attachment
           If Not FileSys.FileExists(App.Path & "\Temp\" & !MessageName) Then
              strStream.SaveToFile App.Path & "\Temp\" & !MessageName, adSaveCreateOverWrite
           End If
           strStream.Close
           Set strStream = Nothing
          '
           PlaySound (!MessageName)
        Else
          MsgBox "Invalid File Format", vbExclamation, "Warning"
        End If
      Else
        MsgBox "There is no file attached to play!", vbInformation
      End If
    End If
   End With
   '
   Set rsAttach = Nothing
   Set strStream = Nothing
   Set FileSys = Nothing
End Sub

Private Sub cmdEditGroups_Click()
  FUserGroups.Show vbModal
  RefreshMessages
  FMain.tmrMessages_Timer
End Sub

Private Sub cmdRefresh_Click()
  RefreshMessages
  ListView1_Click
End Sub

Private Sub cmdSave_Click()
  With rsMessages
    If Not !Completed Then
      !DateCompleted = Date
      !TimeCompleted = Time
      !User = StrUser
    End If
    '
    If cmbComment.Text <> vbNullString Then
      !Comments = cmbComment.Text
    End If
    If cmbCaller.Text <> vbNullString Then
      !Caller = cmbCaller.Text
    End If
    If chkComp.Value = 1 Then
      !Completed = True
    Else
      !Completed = False
    End If
    On Error GoTo ErrorHandler
    .UpdateBatch
  End With
  '
  RefreshMessages
  ListView1_Click
  FMain.tmrMessages_Timer
Exit Sub
ErrorHandler:
  MsgBox ("Error. It's possible somebody else has changed this record since this window was opened")
  '
End Sub

Private Sub Form_Load()
  Dim rsComments As New ADODB.Recordset
  '
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.MinHeight = 5475
  FormControl.MinWidth = 10590
  FormControl.DataForm = True
  '
  choice = NEWCALLS
  cmdNew.FontBold = True
  cmdAll.FontBold = False
  cmdOld.FontBold = False
  '
  cmbComment.Clear
  rsComments.Open "select * from TVMailComment", cnMain, adOpenKeyset, adLockBatchOptimistic
  With rsComments
    Do While Not .eof
      cmbComment.AddItem !Comment
    .MoveNext
    Loop
    .Close
  End With
  '
  Set rsComments = Nothing
  '
  cmbMessageGroup.AddItem "All"
  cmbMessageGroup.AddItem "Authorizations"
  cmbMessageGroup.AddItem "Sales"
  cmbMessageGroup.AddItem "Support"
  GetUserGroups
  '
  cmbMessageGroup.ListIndex = 0
  '
  Me.Caption = ""
  '
  InitializeVmail
  '
  Timer1.Interval = 30000
  '
  GetColumnWidths
  '
  ListView1.Checkboxes = True
  '
  RefreshMessages
  '
  ListView1_Click
  '
End Sub

Private Sub Form_Resize()
  If Me.Height > 5000 And Me.Width > 1000 Then
    ListView1.Move 0, 400, Me.Width - 100, Me.Height - 4000
    fraDetails.Move 150, Me.Height - 3450, 11535, 2700
    Shape1.Move 100, Me.Height - 3500, Me.Width - 300, 2800
    cmdDetails.Move 0, 0, 1000, 400
    cmdPlay.Move 0, 0, 1000, 400
    cmdCompleted.Move 1010, 0, 1000, 400
    cmdRefresh.Move 2020, 0, 1000, 400
    cmdNew.Move 5000, 0, 600, 400
    cmdOld.Move 5610, 0, 600, 400
    cmdAll.Move 6220, 0, 600, 400
    lblMessageGroup.Move 7000, 100
    lblGroups.Move 8300, 100, 1750
    cmbMessageGroup.Move 8250, 50
    cmdEditGroups.Move 10100, 0, 1000, 400
    lblShow.Move 4300, 100, 500, 400
    lblcount.Move 0, Me.Height - 600
    lblLastClient.Move 2000, Me.Height - 600
    lblLastServer.Move 5750, Me.Height - 600, 8000
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim FileSys As New FileSystemObject

  Dim fileCount As Integer
  Dim Index As Integer
  '
  Timer1_Timer
  If FileSys.FolderExists(App.Path & "\Temp") Then
    File1.Path = App.Path & "\Temp"
    File1.Refresh
    fileCount = File1.ListCount
    If fileCount > 0 Then
    For Index = 0 To fileCount - 1
    
     On Error Resume Next
       FileSys.DeleteFile (App.Path & "\Temp\" & File1.list(Index))
    Next
    End If
  End If
  SaveColumnWidths
  'rs.Close
End Sub

Private Sub ListView1_DblClick()
  ListView1_Click
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next

  RefreshMessages
  '


  'ListView1.SetFocus
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  On Error GoTo ErrorHandler
  '
  SortListView ListView1, ColumnHeader
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.lvwLog.ColumnClick"
End Sub

Public Sub SortListView(ByVal lvwCur As MSComctlLib.ListView, ByVal colHdr As MSComctlLib.ColumnHeader, Optional ByVal sSortOrder As String)
  On Error GoTo ErrorHandler
  '
  With lvwCur
    '
    If .SortKey > -1 Then .ColumnHeaders.Item(.SortKey + 1).Icon = 0
    '
    .SortKey = colHdr.Index - 1
    '
    If sSortOrder <> vbNullString Then
      .SortOrder = IIf(sSortOrder = "Ascending", lvwAscending, lvwDescending)
    Else
      .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    End If
    '
    .Sorted = True
    '
    '.ColumnHeaders.Item(colHdr.Index).Icon = IIf(.SortOrder = lvwAscending, "imgAscending", "imgDescending")
    '
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox Error((Err.Number)), vbCritical + vbOKOnly, "Error: FPrimary.General.SortListView"
End Sub

Public Sub GetUserGroups()
  Dim iTemp As Integer
  Dim sTempMsg As String
  '
  '
  If iGroupNumber = 15 Then
    lblGroups.Caption = "All"
  Else
    iTemp = iGroupNumber
    If iTemp >= 8 Then
      sTempMsg = "Authorizations"
      iTemp = iTemp - 8
    End If
    '
    If iTemp >= 4 Then
      If Len(sTempMsg) > 0 Then
        sTempMsg = sTempMsg & ", Sales"
      Else
        sTempMsg = "Sales"
      End If
      iTemp = iTemp - 4
    End If
    '
    If iTemp >= 2 Then
      If Len(sTempMsg) > 0 Then
        sTempMsg = sTempMsg & ", Support"
      Else
        sTempMsg = "Support"
      End If
      iTemp = iTemp - 2
    End If
    '
    If iTemp >= 1 Then
      If Len(sTempMsg) > 0 Then
        sTempMsg = sTempMsg & ", Operator"
      Else
        sTempMsg = "Operator"
      End If
    End If
    lblGroups.Caption = sTempMsg
  End If
'  Select Case iGroupNumber
'    Case 1
'      lblGroups.Caption = "Authorizations"
'    Case 2
'      lblGroups.Caption = "Sales"
'    Case 3
'      lblGroups.Caption = "Support"
'    Case 4
'      lblGroups.Caption = "All"
'    Case 5
'      lblGroups.Caption = "Authorizations, Sales"
'    Case 6
'      lblGroups.Caption = "Authorizations, Support"
'    Case 7
'      lblGroups.Caption = "Sales, Support"
'    Case Else
'  End Select
  '
End Sub

Public Sub ListView1_Click()
  
  Dim lMessageID As Long
  '
  If FVMail.ListView1.ListItems.Count > 0 Then
    If rsMessages.State <> 0 Then
        rsMessages.Close
    End If
    '
    lMessageID = CLng(Right$(FVMail.ListView1.SelectedItem.Key, Len(FVMail.ListView1.SelectedItem.Key) - 1))
    '
    Set rsMessages = GetMessageRecord(lMessageID)
      '
    With rsMessages
      .MoveFirst
      '.Find "messageID = " & X
      If Not .eof Then
        sMessageID = lMessageID
        sMessageName = !MessageName & vbNullString
        '
        txtPhone.Text = !PhoneNumber & vbNullString
        '
        If !Caller & vbNullString <> "" Then
            cmbCaller.Text = !Caller & vbNullString
        Else
            cmbCaller.Text = ""
        End If
        '
        If txtPhone.Text = "" Then
          cmdGetNames.Enabled = False
        Else
          cmdGetNames.Enabled = True
        End If
        '
        sReceived = !TimeReceived & " " & !DateReceived & vbNullString
        '
        cmbComment.Text = !Comments & vbNullString
        '
        If Not !Caller = "" Then
          sCaller = !Caller & vbNullString
        Else
          sCaller = !From & vbNullString
        End If
        '
        txtsubject.Text = !Subject & vbNullString
        ' Body only used here!
        'txtBody.Text = !Body & vbNullString
        'sBody = !Body & vbNullString
        LoadEmailBodyBody
        '
        If txtBody.Text <> "" Then
          DisplayEmailBody
        Else
         webBody.Navigate2 "about:blank"
        End If
        '
        If !PhoneNumber <> "" Then
          sSubject = !PhoneNumber & vbNullString
        Else
          sSubject = !Subject & vbNullString
        End If
        '
        If !Completed = True Then
          chkComp.Value = 1
        Else
          chkComp.Value = 0
        End If
        '
       End If
    End With
    '
    If rsMessages!MessageName & vbNullString = "" Then
      cmdPlay.Enabled = False
    Else
      cmdPlay.Enabled = True
    End If
    '
    'If ListView1.Visible Then
    '  ListView1.SetFocus
   ' End If
  End If
End Sub

Public Sub FindNames()
  Dim rs As New ADODB.Recordset
  Dim i As Integer
  '
  cmbCaller.Clear
  '
  i = 0
  '
  rs.Open "SELECT [FirstName], [LastName],[Phone1], [Phone2],[fax], [ID]  FROM TContact", cnMain, adOpenKeyset, adLockBatchOptimistic
  
  With rs
    .MoveFirst
    While Not .eof
      If Trim(!Phone1) = txtPhone.Text Or Trim(!Phone2) = txtPhone.Text Or Trim(!Fax) = txtPhone.Text Then
        cmbCaller.AddItem !FirstName & " " & !LastName
        i = i + 1
      End If
      .MoveNext
    Wend
  End With
  '
  If Not i = 0 Then
    cmbCaller.ListIndex = 0
  Else
    MsgBox "No Results Found For Phone Number", vbInformation, "vMail"
  End If
  '
  rs.Close
  '
  Set rs = Nothing
End Sub

Public Sub GetColumnWidths()
Dim rsColumns As New ADODB.Recordset
'
  rsColumns.Open "Select * from TEmailAddresses where [Name] = '" & StrUser & "'", cnMain, adOpenDynamic, adLockBatchOptimistic
  '
  With rsColumns
    If Not .eof Then
      iLenGroup = nnNum(!LenGroup)
      iLenMessage = nnNum(!LenMessage)
      iLenPhone = nnNum(!LenPhone)
      iLenFrom = nnNum(!LenFrom)
      iLenSubject = nnNum(!LenSubject)
      iLenDateRec = nnNum(!LenDateRec)
      iLenTimeRec = nnNum(!LenTimeRec)
      iLenMessageNum = nnNum(!LenMessageNum)
      iLenUser = nnNum(!LenUser)
      iLenCaller = nnNum(!LenCaller)
      iLenComments = nnNum(!LenComments)
      iLenDateCom = nnNum(!LenDateCom)
      iLenTimeCom = nnNum(!LenTimeCom)
      iFromAddress = nnNum(!LenFromAddress)
    End If
  End With
  bLoad = True
  '
End Sub

Public Sub SaveColumnWidths()
Dim rsColumns As New ADODB.Recordset
'
  rsColumns.Open "Select * from TEmailAddresses where [Name] = '" & StrUser & "'", cnMain, adOpenDynamic, adLockBatchOptimistic
  '
  With rsColumns
    If Not .eof Then
      !LenGroup = iLenGroup
      !LenMessage = iLenMessage
      !LenPhone = iLenPhone
      !LenFrom = iLenFrom
      !LenSubject = iLenSubject
      !LenDateRec = iLenDateRec
      !LenTimeRec = iLenTimeRec
      !LenMessageNum = iLenMessageNum
      !LenUser = iLenUser
      !LenCaller = iLenCaller
      !LenComments = iLenComments
      !LenDateCom = iLenDateCom
      !LenTimeCom = iLenTimeCom
      !LenFromAddress = iFromAddress
      .UpdateBatch
    End If
  End With
  '
End Sub
