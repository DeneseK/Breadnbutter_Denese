VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FSupportOpen 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Open Calls"
   ClientHeight    =   7365
   ClientLeft      =   2160
   ClientTop       =   2355
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      Begin VB.CommandButton Command1 
         Caption         =   "close selected calls"
         Height          =   315
         Left            =   9360
         TabIndex        =   13
         Top             =   6600
         Width           =   1995
      End
      Begin VB.Frame Group 
         BackColor       =   &H00FF8080&
         Height          =   1635
         Left            =   4680
         TabIndex        =   22
         Top             =   5280
         Width           =   2175
         Begin VB.ListBox lstGroup 
            Height          =   1185
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CheckBox chkDate 
         BackColor       =   &H00FF8080&
         Caption         =   "Date"
         Height          =   255
         Left            =   6960
         TabIndex        =   4
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Frame Date 
         BackColor       =   &H00FF8080&
         Height          =   1635
         Left            =   6960
         TabIndex        =   18
         Top             =   5280
         Width           =   2175
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   38209
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "To"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   195
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "From"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   120
            Width           =   345
         End
      End
      Begin VB.OptionButton optGroup 
         BackColor       =   &H00FF8080&
         Caption         =   "Group"
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optUser 
         BackColor       =   &H00FF8080&
         Caption         =   "User"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Frame User 
         BackColor       =   &H00FF8080&
         Height          =   1635
         Left            =   2400
         TabIndex        =   15
         Top             =   5280
         Width           =   2175
         Begin VB.ListBox lstUsers 
            Height          =   1185
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FF8080&
         Height          =   1635
         Left            =   120
         TabIndex        =   14
         Top             =   5280
         Width           =   2175
         Begin VB.ListBox lstCategory 
            Height          =   1185
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   240
            Width           =   1935
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdHistory 
         Bindings        =   "FSupportOpen.frx":0000
         Height          =   4695
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   11565
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   11
         UseGroups       =   -1  'True
         AllowColumnShrinking=   0   'False
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         ForeColorEven   =   0
         BackColorEven   =   -2147483633
         BackColorOdd    =   -2147483633
         Levels          =   3
         RowHeight       =   1270
         Groups(0).Width =   20399
         Groups(0).Caption=   $"FSupportOpen.frx":001B
         Groups(0).CaptionAlignment=   0
         Groups(0).Columns.Count=   11
         Groups(0).Columns(0).Width=   3200
         Groups(0).Columns(0).Visible=   0   'False
         Groups(0).Columns(0).Caption=   "RecID"
         Groups(0).Columns(0).Name=   "RecID"
         Groups(0).Columns(0).Alignment=   1
         Groups(0).Columns(0).CaptionAlignment=   1
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).DataType=   3
         Groups(0).Columns(0).FieldLen=   256
         Groups(0).Columns(0).Locked=   -1  'True
         Groups(0).Columns(1).Width=   3200
         Groups(0).Columns(1).Visible=   0   'False
         Groups(0).Columns(1).Caption=   "CustRecID"
         Groups(0).Columns(1).Name=   "CustRecID"
         Groups(0).Columns(1).Alignment=   1
         Groups(0).Columns(1).CaptionAlignment=   1
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).DataType=   3
         Groups(0).Columns(1).FieldLen=   256
         Groups(0).Columns(1).Locked=   -1  'True
         Groups(0).Columns(2).Width=   3016
         Groups(0).Columns(2).Caption=   "Date"
         Groups(0).Columns(2).Name=   "Date"
         Groups(0).Columns(2).CaptionAlignment=   1
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   7
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(2).Locked=   -1  'True
         Groups(0).Columns(2).HasBackColor=   -1  'True
         Groups(0).Columns(2).BackColor=   12632256
         Groups(0).Columns(3).Width=   3228
         Groups(0).Columns(3).Caption=   "Time"
         Groups(0).Columns(3).Name=   "Time"
         Groups(0).Columns(3).CaptionAlignment=   1
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   7
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(3).Locked=   -1  'True
         Groups(0).Columns(3).HasBackColor=   -1  'True
         Groups(0).Columns(3).BackColor=   12632256
         Groups(0).Columns(4).Width=   3757
         Groups(0).Columns(4).Caption=   "Type"
         Groups(0).Columns(4).Name=   "Type"
         Groups(0).Columns(4).CaptionAlignment=   0
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(4).Locked=   -1  'True
         Groups(0).Columns(4).HasBackColor=   -1  'True
         Groups(0).Columns(4).BackColor=   12632256
         Groups(0).Columns(5).Width=   3810
         Groups(0).Columns(5).Caption=   "User"
         Groups(0).Columns(5).Name=   "User"
         Groups(0).Columns(5).CaptionAlignment=   0
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   8
         Groups(0).Columns(5).FieldLen=   256
         Groups(0).Columns(5).Locked=   -1  'True
         Groups(0).Columns(5).HasBackColor=   -1  'True
         Groups(0).Columns(5).BackColor=   12632256
         Groups(0).Columns(6).Width=   6588
         Groups(0).Columns(6).Caption=   "Subject"
         Groups(0).Columns(6).Name=   "Subject"
         Groups(0).Columns(6).CaptionAlignment=   0
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   8
         Groups(0).Columns(6).FieldLen=   256
         Groups(0).Columns(6).Locked=   -1  'True
         Groups(0).Columns(6).HasBackColor=   -1  'True
         Groups(0).Columns(6).BackColor=   12632256
         Groups(0).Columns(7).Width=   9419
         Groups(0).Columns(7).Caption=   "Company"
         Groups(0).Columns(7).Name=   "Company"
         Groups(0).Columns(7).DataField=   "Column 7"
         Groups(0).Columns(7).DataType=   8
         Groups(0).Columns(7).Level=   1
         Groups(0).Columns(7).FieldLen=   256
         Groups(0).Columns(7).Locked=   -1  'True
         Groups(0).Columns(7).HasBackColor=   -1  'True
         Groups(0).Columns(7).BackColor=   16777215
         Groups(0).Columns(8).Width=   10980
         Groups(0).Columns(8).Caption=   "Name"
         Groups(0).Columns(8).Name=   "Name"
         Groups(0).Columns(8).DataField=   "Column 8"
         Groups(0).Columns(8).DataType=   8
         Groups(0).Columns(8).Level=   1
         Groups(0).Columns(8).FieldLen=   256
         Groups(0).Columns(8).Locked=   -1  'True
         Groups(0).Columns(8).HasBackColor=   -1  'True
         Groups(0).Columns(8).BackColor=   16777215
         Groups(0).Columns(9).Width=   18918
         Groups(0).Columns(9).Caption=   "Results"
         Groups(0).Columns(9).Name=   "Results"
         Groups(0).Columns(9).CaptionAlignment=   0
         Groups(0).Columns(9).DataField=   "Column 9"
         Groups(0).Columns(9).DataType=   8
         Groups(0).Columns(9).Level=   2
         Groups(0).Columns(9).FieldLen=   2000
         Groups(0).Columns(9).VertScrollBar=   -1  'True
         Groups(0).Columns(9).Locked=   -1  'True
         Groups(0).Columns(9).HasForeColor=   -1  'True
         Groups(0).Columns(9).HasBackColor=   -1  'True
         Groups(0).Columns(9).ForeColor=   -2147483640
         Groups(0).Columns(9).BackColor=   -2147483643
         Groups(0).Columns(10).Width=   1482
         Groups(0).Columns(10).Caption=   "Close"
         Groups(0).Columns(10).Name=   "Close"
         Groups(0).Columns(10).DataField=   "Column 10"
         Groups(0).Columns(10).DataType=   8
         Groups(0).Columns(10).Level=   2
         Groups(0).Columns(10).FieldLen=   256
         Groups(0).Columns(10).Style=   4
         _ExtentX        =   20399
         _ExtentY        =   8281
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Report"
         Height          =   315
         Left            =   9360
         TabIndex        =   11
         Top             =   5640
         Width           =   1995
      End
      Begin VB.CommandButton cmdShowResults 
         Caption         =   "Show Results"
         Height          =   315
         Left            =   9360
         TabIndex        =   10
         Top             =   5160
         Width           =   1995
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Results to Clipboad"
         Height          =   315
         Left            =   9360
         TabIndex        =   12
         Top             =   6120
         Width           =   1995
      End
      Begin VB.Label lblCategory 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Category"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Width           =   630
      End
      Begin VB.Label LblCount 
         BackColor       =   &H00FF8080&
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Notes Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "FSupportOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Report As New CReport
'
Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
Private sQueryString As String
Private sText As String

Private Sub chkDate_Click()
  If chkDate.Value = 1 Then
    Date.Enabled = True
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
  Else
    Date.Enabled = False
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
  End If
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  '
  On Error GoTo EH
  '
  Clipboard.SetText sText, vbCFText
  '
  Exit Sub
EH:
 MsgBox Err.Description & " in FSupportOpen.cmdShowResults_Click."
End Sub

Private Sub Command1_Click()
CloseAllCalls
End Sub

Private Sub Form_Activate()
  On Error GoTo EH
  '
  Frame1.Move (Width - Frame1.Width) / 2
  Exit Sub
EH:
 MsgBox Err.Description & " in FSupportOpen.Form_Activate."
End Sub

Private Sub Form_Resize()
  On Error GoTo EH
  '
  Frame1.Move (Width - Frame1.Width) / 2
  Exit Sub
EH:
 MsgBox Err.Description & " in FSupportOpen.Form_Resize."
End Sub

Private Sub Form_Load()
  On Error GoTo EH
  Dim rs As New ADODB.Recordset
  '
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.Setup Me, False, , , "Open Support Calls"
  '
  rs.Open "SELECT * FROM tblactivities ORDER BY Activity", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rs.eof
    lstCategory.AddItem rs!Activity
    rs.MoveNext
  Loop
  '
  rs.Close
  '
  rs.Open "SELECT * FROM tblEmployees ORDER BY EmployeeLast", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rs.eof
    lstUsers.AddItem rs!EmployeeFirst & " " & rs!EmployeeLast
    rs.MoveNext
  Loop
  '
  rs.Close
  '
  lstGroup.AddItem "Management"
  lstGroup.AddItem "Sales"
  lstGroup.AddItem "Support"
  lstGroup.AddItem "Development"
  '
  optUser.Value = True
  chkDate.Value = 0
  Date.Enabled = False
  DTPicker1.Enabled = False
  DTPicker2.Enabled = False
  DTPicker1.Value = Now - 1
  DTPicker2.Value = Now
  '
  Exit Sub
EH:
 MsgBox Err.Description & " in FSupportOpen.Form_Load."
End Sub


Private Sub grdHistory_BtnClick()
  If MsgBox("Are you sure you want to close this call?", vbYesNo, "Close Call") = vbYes Then
    CloseCall grdHistory.Columns(0).Value
    LoadHistory
  End If
End Sub

Private Sub grdHistory_DblClick()
  On Error GoTo EH
    'grdHistory.Redraw = False
    'FEditDetail.ShowRecord grdHistory.Columns(0).Value, True
    'grdHistory.Redraw = True
    'LoadHistory
    '
  'Load FResult
  'FResult.TextResult.Text = grdHistory.Columns(9).Value
  'FResult.Show vbModal
  Exit Sub
EH:
 MsgBox Err.Description & " in FSupportOpen.grdHistory_DblClick."
End Sub

Private Sub cmdPrint_Click()
  'On Error GoTo EH
      RHistory.Company.Text = "Open Calls"
      RHistory.adc.Connection = cnMain
      RHistory.adc.Source = sQueryString
      RHistory.Show
'EH:
' MsgBox Err.Description & " in FSupportOpen.cmdPrint_Click."
End Sub

Private Sub cmdShowResults_Click()
LoadHistory
'  '
'  Dim Count As Long
'  On Error GoTo EH
'  SendData
'  Me.grdHistory.RemoveAll
'  With Report.rsReport
'      Do While Not .eof
'         Count = Count + 1
'         grdHistory.AddItem !RecID & vbTab & !CustRecID & vbTab & !Date & vbTab _
'         & !Time & vbTab & !Type & vbTab & !User & vbTab & !Subject & vbTab _
'         & "Company: " & !Company & vbTab & "Contact: " & !FirstName & " " & !LastName _
'         & vbTab & !Results
'         .MoveNext
'      Loop
'  End With
'  '
'  LblCount.Caption = Count
'  Exit Sub
'EH:
' MsgBox Err.Description & " in FSupportOpen.cmdShowResults_Click."
End Sub

Private Sub CloseAllCalls()
  '
  Dim iWorkGroup As Integer
  Dim sCategory As String
  Dim sUsers As String
  Dim sDate As String
  Dim iIndexNum As Integer
  Dim iConjunction As Integer
  Dim Employee As New CEmployee
  '
  Me.MousePointer = vbHourglass
  '
  sText = ""
  '
  grdHistory.Redraw = False
  '
  Me.grdHistory.RemoveAll
  '
  'If lCustomerID <> -1 Then
    Dim rsHistory As ADODB.Recordset
    '
    Set rsHistory = New ADODB.Recordset
    '
    iConjunction = 0
    For iIndexNum = 0 To lstCategory.ListCount - 1
      If lstCategory.Selected(iIndexNum) = True Then
        If iConjunction = 0 Then
          sCategory = "AND ("
          iConjunction = 1
        Else
          sCategory = sCategory + "OR"
        End If
        sCategory = sCategory + " (TSupportActs.Type = '" & lstCategory.list(iIndexNum) & "') "
      End If
    Next
    If Len(sCategory) > 0 Then sCategory = sCategory + ")"
    '
    If optUser.Value = True Then
      iConjunction = 0
      For iIndexNum = 0 To lstUsers.ListCount - 1
        If lstUsers.Selected(iIndexNum) = True Then
          If iConjunction = 0 Then
            sUsers = "AND ("
            iConjunction = 1
          Else
            sUsers = sUsers + "OR"
          End If
          sUsers = sUsers + " (TSupportActs.[User] = '" & lstUsers.list(iIndexNum) & "') "
        End If
      Next
      If Len(sUsers) > 0 Then sUsers = sUsers + ")"
    Else '##################################################################################
      iConjunction = 0
      For iWorkGroup = 0 To lstGroup.ListCount - 1
        If lstGroup.Selected(iWorkGroup) = True Then
          For iIndexNum = 0 To lstUsers.ListCount - 1
            If Employee.InGroup(lstUsers.list(iIndexNum), lstGroup.list(iWorkGroup)) = True Then
              If iConjunction = 0 Then
                sUsers = "AND ("
                iConjunction = 1
              Else
                sUsers = sUsers + "OR"
              End If
              sUsers = sUsers + " (TSupportActs.[User] = '" & lstUsers.list(iIndexNum) & "') "
            End If
          Next
        End If
      Next
      If Len(sUsers) > 0 Then sUsers = sUsers + ")"
    End If '###############################################################################
        '
    If chkDate.Value = 1 Then
      If DTPicker1 < DTPicker2 Then
        sDate = "AND (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) "
      Else
        MsgBox "Incorrect Date Values", vbExclamation, "Date Error"
        LblCount.Caption = 0
        Exit Sub
      End If
    Else
      sDate = ""
    End If
    '
    sQueryString = "SELECT TSupportActs.*, TContact.FirstName, TContact.LastName, TCompany.Name AS CompanyName FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID RIGHT OUTER JOIN TSupportActs ON TContact.ID = TSupportActs.CustRecID Where (TSupportActs.OpenCall = 1) " & sCategory & sUsers & sDate & "ORDER BY TSupportActs.[Date] DESC, TSupportActs.[Time] DESC"
    
    rsHistory.Open sQueryString, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  With rsHistory
    If Not (.eof And .BOF) Then
      Do While Not .eof
        CloseCall !RecID
        .MoveNext
      Loop
    Else
      MsgBox "No Open Calls Found!", vbExclamation, "Close Open Calls"
      LblCount.Caption = 0
    End If
    End With 'rsHistory
    '
    rsHistory.Close
    Set rsHistory = Nothing
    Set Employee = Nothing
'  End If
  '
  grdHistory.Redraw = True
  '
  Me.MousePointer = vbNormal
End Sub

Private Sub LoadHistory()
'  On Error GoTo EH
  '
  Dim iWorkGroup As Integer
  Dim sCategory As String
  Dim sUsers As String
  Dim sDate As String
  Dim iIndexNum As Integer
  Dim iConjunction As Integer
  Dim Employee As New CEmployee
  '
  Me.MousePointer = vbHourglass
  '
  sText = ""
  '
  grdHistory.Redraw = False
  '
  Me.grdHistory.RemoveAll
  '
  'If lCustomerID <> -1 Then
    Dim rsHistory As ADODB.Recordset
    '
    Set rsHistory = New ADODB.Recordset
    '
    iConjunction = 0
    For iIndexNum = 0 To lstCategory.ListCount - 1
      If lstCategory.Selected(iIndexNum) = True Then
        If iConjunction = 0 Then
          sCategory = "AND ("
          iConjunction = 1
        Else
          sCategory = sCategory + "OR"
        End If
        sCategory = sCategory + " (TSupportActs.Type = '" & lstCategory.list(iIndexNum) & "') "
      End If
    Next
    If Len(sCategory) > 0 Then sCategory = sCategory + ")"
    '
    If optUser.Value = True Then
      iConjunction = 0
      For iIndexNum = 0 To lstUsers.ListCount - 1
        If lstUsers.Selected(iIndexNum) = True Then
          If iConjunction = 0 Then
            sUsers = "AND ("
            iConjunction = 1
          Else
            sUsers = sUsers + "OR"
          End If
          sUsers = sUsers + " (TSupportActs.[User] = '" & lstUsers.list(iIndexNum) & "') "
        End If
      Next
      If Len(sUsers) > 0 Then sUsers = sUsers + ")"
    Else '##################################################################################
      iConjunction = 0
      For iWorkGroup = 0 To lstGroup.ListCount - 1
        If lstGroup.Selected(iWorkGroup) = True Then
          For iIndexNum = 0 To lstUsers.ListCount - 1
            If Employee.InGroup(lstUsers.list(iIndexNum), lstGroup.list(iWorkGroup)) = True Then
              If iConjunction = 0 Then
                sUsers = "AND ("
                iConjunction = 1
              Else
                sUsers = sUsers + "OR"
              End If
              sUsers = sUsers + " (TSupportActs.[User] = '" & lstUsers.list(iIndexNum) & "') "
            End If
          Next
        End If
      Next
      If Len(sUsers) > 0 Then sUsers = sUsers + ")"
    End If '###############################################################################
        '
    If chkDate.Value = 1 Then
      If DTPicker1 < DTPicker2 Then
        sDate = "AND (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) "
      Else
        MsgBox "Incorrect Date Values", vbExclamation, "Date Error"
        LblCount.Caption = 0
        Exit Sub
      End If
    Else
      sDate = ""
    End If
    '
    sQueryString = "SELECT TSupportActs.*, TContact.FirstName, TContact.LastName, TCompany.Name AS CompanyName FROM TCompany RIGHT OUTER JOIN TContact ON TCompany.ID = TContact.CompanyID RIGHT OUTER JOIN TSupportActs ON TContact.ID = TSupportActs.CustRecID Where (TSupportActs.OpenCall = 1) " & sCategory & sUsers & sDate & "ORDER BY TSupportActs.[Date] DESC, TSupportActs.[Time] DESC"
    
    rsHistory.Open sQueryString, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
    '
    With rsHistory
    If Not (.eof And .BOF) Then
      LblCount.Caption = .RecordCount
      Do While Not .eof
      '###################################################################################
       sText = sText & !FirstName & " " & !LastName & ", "
       sText = sText & !CompanyName & vbCrLf
       sText = sText & !Date & ", "
       sText = sText & !Time & ", "
       sText = sText & !User & vbCrLf
       sText = sText & !Type & ", "
      
       If !Subject & vbNullString <> "" Then
          sText = sText & !Subject & ", "
       End If
       '
       sText = sText & !Results & vbCrLf & vbCrLf
       '###################################################################################
        grdHistory.AddItem !RecID & vbTab & !CustRecID & vbTab & !Date & vbTab _
           & !Time & vbTab & !Type & vbTab & !User & vbTab & !Subject & vbTab _
           & "Company: " & !CompanyName & vbTab & "Contact: " & !FirstName & " " & !LastName _
           & vbTab & !Results & vbTab & "Close"
        '
        .MoveNext
      Loop
    Else
      MsgBox "No Open Calls Found!", vbExclamation, "Open Calls"
      LblCount.Caption = 0
    End If
    End With 'rsHistory
    '
    rsHistory.Close
    Set rsHistory = Nothing
    Set Employee = Nothing
'  End If
  '
  grdHistory.Redraw = True
  '
  Me.MousePointer = vbNormal
  '
  Exit Sub
'EH:
'  grdHistory.Redraw = True
'  MsgBox Err.Description
End Sub

Private Sub CloseCall(lID As Long)
  Dim sSQL As String
  '
  sSQL = "UPDATE TSupportActs SET " & _
        "ClosedTime = '" & Replace(Format(Now, "m/d/yy h:mm AM/PM"), "'", "''") & "', OpenCall = '0' " & _
        "WHERE RecID = " & lID
  '
  cnMain.Execute sSQL
  '
End Sub

Private Sub optGroup_Click()
  If optGroup.Value = True Then
    User.Enabled = False
    lstUsers.Enabled = False
    Group.Enabled = True
    lstGroup.Enabled = True
  Else
    User.Enabled = True
    lstUsers.Enabled = True
    Group.Enabled = False
    lstGroup.Enabled = False
  End If
End Sub

Private Sub optUser_Click()
  If optUser.Value = True Then
    User.Enabled = True
    lstUsers.Enabled = True
    Group.Enabled = False
    lstGroup.Enabled = False
  Else
    User.Enabled = False
    lstUsers.Enabled = False
    Group.Enabled = True
    lstGroup.Enabled = True
  End If
End Sub

'Private Function InGroup(sName As String, sWorkGroup As String) As Boolean
'  InGroup = False
'  Dim rsList As New ADODB.Recordset
'  Dim EmployeeData As CEmployeeData
'  Dim iWorkGroupNum As Integer
'  '
'  rsList.Open "select * from tblEmployees", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
'  With rsList
'    Do While Not .eof
'      If LCase(sName) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
'        iWorkGroupNum = nnNum(!WorkGroups)
'      End If
'      .MoveNext
'    Loop
'  End With
'    Select Case sWorkGroup
'      Case "Management"
'        If iWorkGroupNum > 7 Then InGroup = True
'      Case "Sales"
'        Select Case iWorkGroupNum
'          Case 4, 5, 6, 7, 12, 13, 14, 15
'            InGroup = True
'        End Select
'      Case "Support"
'        Select Case iWorkGroupNum
'          Case 2, 3, 6, 7, 10, 11, 14, 15
'            InGroup = True
'        End Select
'      Case "Development"
'        Select Case iWorkGroupNum
'          Case 1, 3, 5, 7, 9, 11, 13, 15
'            InGroup = True
'        End Select
'    End Select
'    '
'    rsList.Close
'  Set rsList = Nothing
'  Set EmployeeData = Nothing
'  '
'End Function
