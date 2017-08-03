VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FAuthLog 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   1020
   ClientTop       =   1665
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Options"
      Height          =   885
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5325
      Begin SSDataWidgets_B.SSDBCombo cboFilter 
         Height          =   285
         Left            =   1290
         TabIndex        =   1
         Top             =   330
         Width           =   2175
         DataFieldList   =   "Column 0"
         MaxDropDownItems=   10
         AllowInput      =   0   'False
         ListWidth       =   3836
         _Version        =   196617
         DataMode        =   2
         Cols            =   1
         ColumnHeaders   =   0   'False
         DefColWidth     =   3836
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   3836
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.Label Label19 
         Caption         =   "Log Filter:"
         Height          =   225
         Left            =   270
         TabIndex        =   3
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblActs 
         Caption         =   "(0000 of 0000)"
         Height          =   195
         Left            =   3660
         TabIndex        =   2
         Top             =   360
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView lvwLog 
      Height          =   5685
      Left            =   30
      TabIndex        =   4
      Top             =   1320
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   10028
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilLog"
      SmallIcons      =   "ilLog"
      ColHdrIcons     =   "ilLog"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date/Time"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Company"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "User"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Action"
         Object.Width           =   4108
      EndProperty
   End
   Begin MSComctlLib.ImageList ilLog 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAuthLog.frx":0000
            Key             =   "imgAscending"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAuthLog.frx":015A
            Key             =   "imgDescending"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAuthLog.frx":02B4
            Key             =   "imgAuthorize"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAuthLog.frx":040E
            Key             =   "imgDeauthorize"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAuthLog.frx":0568
            Key             =   "imgRestore"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "FAuthLog.frx":09BA
      Top             =   1020
      Width           =   240
   End
   Begin VB.Label Label18 
      Caption         =   "Activity Log"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   990
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   8520
      Picture         =   "FAuthLog.frx":0B04
      Stretch         =   -1  'True
      Top             =   960
      Width           =   300
   End
   Begin VB.Label Label12 
      Caption         =   "Double-click activity to view full details"
      Height          =   195
      Left            =   8880
      TabIndex        =   5
      Top             =   990
      Width           =   2775
   End
End
Attribute VB_Name = "FAuthLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
'
Public rsLog  As New Recordset

Private Sub Form_Initialize()
  On Error GoTo ErrCall
  '
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.MinHeight = 1965
  FormControl.MinWidth = Me.Width
  FormControl.DataForm = True
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & "in frmAuthLog.Form_Initialize.", vbCritical, "Error"
End Sub

Private Sub OpenLog()
  On Error GoTo ErrorHandler
  '
  '\\ Log
  Dim rsLog As New ADODB.Recordset
  '
  rsLog.LockType = adLockPessimistic
  rsLog.CursorType = adOpenDynamic
  rsLog.Open "SELECT * from tbllog", cnMain '
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: OpenLog"
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
    .ColumnHeaders.Item(colHdr.Index).Icon = IIf(.SortOrder = lvwAscending, "imgAscending", "imgDescending")
    '
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox Error((Err.number)), vbCritical + vbOKOnly, "Error: FPrimary.General.SortListView"
End Sub

Public Sub RefreshLogDisplay()
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iEntryNo    As Integer
  Dim iActionIcon As Integer
  Dim lActTtl     As Long
  Dim lActCur     As Long
  Dim sEmp()      As String
  '
  rsLog.LockType = adLockPessimistic
  rsLog.CursorType = adOpenDynamic
  rsLog.Open "SELECT * from tblLog", cnMain '
  lblActs.Caption = "(0 of 0)"
  '
  With rsLog
    If ((.BOF = True) And (.eof = True)) = False Then
      .MoveLast
      lActTtl = .RecordCount
      lblActs.Caption = "(" & CStr(.RecordCount) & " of " & CStr(.RecordCount) & ")"
      .MoveFirst
      rsLog.Close
      Select Case cboFilter
        Case "None"
          '
          rsLog.Open "SELECT * from tblLog", cnMain '
          '
        Case "Authorizations, All"
          '
          rsLog.Open "SELECT * FROM [tblLog] WHERE [ActionType] = 'Authorization'", cnMain
          '
        Case "Authorizations, New"
          '
          rsLog.Open "SELECT * FROM [tblLog] WHERE [ActionSubType] = 'New'", cnMain
          '
        Case "Authorizations, Extensions"
          '
          rsLog.Open "SELECT * FROM [tblLog] WHERE [ActionSubType] = 'Extension'", cnMain
          '
        Case "Deauthorizations"
          '
          rsLog.Open "SELECT * FROM [tblLog] WHERE [ActionType] = 'Deauthorization'", cnMain
          '
        Case "Restorations"
          '
          rsLog.Open "SELECT * FROM [tblLog] WHERE [ActionType] = 'Restoration'", cnMain
          '
      End Select
    Else
      Exit Sub
    End If
  End With
  '
  With rsLog
    If ((.BOF = True) And (.eof = True)) = False Then
      .MoveLast
      lblActs.Caption = "(" & CStr(.RecordCount) & " of " & CStr(lActTtl) & ")"
      .MoveFirst
      lvwLog.ListItems.Clear
      Do Until .eof
        sEmp = Split(.Fields("Employee").Value)
        Select Case .Fields("ActionType").Value
          Case "Authorization"
            iActionIcon = 3
          Case "Deauthorization"
            iActionIcon = 4
          Case "Restoration"
            iActionIcon = 5
        End Select
        lvwLog.ListItems.Add , "r" & CStr(.Fields("ID").Value), Format(.Fields("ActionDateTime").Value, "YYYY.Mm.Dd") & "  " & Format(.Fields("ActionDateTime").Value, "Hh:Nn:Ss"), , iActionIcon
        iEntryNo = lvwLog.ListItems.Count
        lvwLog.ListItems.Item(iEntryNo).SubItems(1) = sEmp(0)
        lvwLog.ListItems.Item(iEntryNo).SubItems(2) = .Fields("Company").Value
        lvwLog.ListItems.Item(iEntryNo).SubItems(3) = .Fields("User").Value
        lvwLog.ListItems.Item(iEntryNo).SubItems(4) = .Fields("ActionType").Value & ": " & IIf(.Fields("ActionType").Value <> "Restoration", Format(.Fields("SiteDays").Value, "0000") & " Days", Format(.Fields("SiteDateTime").Value, "Short Date"))
        .MoveNext
      Loop
    End If
    '
    .Close
    '
  End With
  '
  'SortListView lvwLog, lvwLog.ColumnHeaders(1), "Descending"
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.General.RefreshLogDisplay"
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
  '
  cboFilter.Text = GetSetting(App.Title, "General", "Filter", "None")
  '
  RefreshLogDisplay
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.lvwLog.ColumnClick"

End Sub

Private Sub Form_Resize()
  On Error GoTo ErrorHandler
  '
  If Width > 1000 And Height > 1000 Then
    lvwLog.Move 30, 1320, Width - 100, Height - 1275
  End If
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.lvwLog.ColumnClick"

End Sub

Private Sub lvwLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  On Error GoTo ErrorHandler
  '
  SortListView lvwLog, ColumnHeader
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.lvwLog.ColumnClick"
End Sub

Private Sub lvwLog_DblClick()
  On Error GoTo ErrorHandler
  '
  FActivity.Show vbModal
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.lvwLog.DoubleClick"
End Sub

Private Sub cboFilter_CloseUp()
  On Error GoTo ErrorHandler
  '
  SaveSetting App.Title, "General", "Filter", cboFilter.Text
  RefreshLogDisplay
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.cboFilter.CloseUp"
End Sub
Private Sub cboFilter_GotFocus()
  On Error GoTo ErrorHandler
  '
  SelectText cboFilter
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.cboFilter.GotFocus"
End Sub
Private Sub cboFilter_InitColumnProps()
  On Error GoTo ErrorHandler
  '
  With cboFilter
    .AddItem "None"
    .AddItem "Authorizations, All"
    .AddItem "Authorizations, New"
    .AddItem "Authorizations, Extensions"
    .AddItem "Deauthorizations"
    .AddItem "Restorations"
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.cboFilter.InitColumnProps"
End Sub

