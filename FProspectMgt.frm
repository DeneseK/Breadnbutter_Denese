VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FProspectMgt 
   BorderStyle     =   0  'None
   ClientHeight    =   7125
   ClientLeft      =   1290
   ClientTop       =   1605
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkTtls 
      Caption         =   "Show &Totals"
      Height          =   225
      Left            =   2280
      TabIndex        =   8
      Top             =   4590
      Width           =   1185
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort by &Label"
      Height          =   255
      Index           =   1
      Left            =   1770
      TabIndex        =   7
      Top             =   4920
      Width           =   1305
   End
   Begin VB.OptionButton optSort 
      Caption         =   "Sort by &Group"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "AM &Best"
      Height          =   225
      Index           =   1
      Left            =   1290
      TabIndex        =   5
      Top             =   4590
      Width           =   915
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "&Standard"
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   4590
      Width           =   975
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdProspectGroup 
      Bindings        =   "FProspectMgt.frx":0000
      Height          =   4725
      Left            =   3630
      TabIndex        =   3
      Top             =   450
      Width           =   7965
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   7
      AllowUpdate     =   0   'False
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "TContact.ID"
      Columns(0).Name =   "TContact.ID"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "TCompany.ID"
      Columns(1).Name =   "TCompany.ID"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   3678
      Columns(2).Caption=   "Full Name"
      Columns(2).Name =   "FullName"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4154
      Columns(3).Caption=   "Company"
      Columns(3).Name =   "Company"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1191
      Columns(4).Caption=   "State"
      Columns(4).Name =   "State"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2302
      Columns(5).Caption=   "Status"
      Columns(5).Name =   "Status"
      Columns(5).CaptionAlignment=   0
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1746
      Columns(6).Caption=   "AuthDate"
      Columns(6).Name =   "AuthDate"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   1
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   7
      Columns(6).FieldLen=   256
      _ExtentX        =   14049
      _ExtentY        =   8334
      _StockProps     =   79
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
   Begin VB.ListBox lstGroups 
      Height          =   3960
      Left            =   210
      TabIndex        =   0
      Top             =   450
      Width           =   3255
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdHistory 
      Bindings        =   "FProspectMgt.frx":001E
      Height          =   1725
      Left            =   210
      TabIndex        =   9
      Top             =   5310
      Width           =   11385
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      ColumnHeaders   =   0   'False
      Col.Count       =   8
      UseGroups       =   -1  'True
      AllowUpdate     =   0   'False
      AllowColumnShrinking=   0   'False
      ForeColorEven   =   0
      BackColorEven   =   -2147483633
      BackColorOdd    =   -2147483633
      Levels          =   2
      RowHeight       =   847
      Groups(0).Width =   19103
      Groups(0).Caption=   $"FProspectMgt.frx":0039
      Groups(0).CaptionAlignment=   0
      Groups(0).Columns.Count=   8
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
      Groups(0).Columns(2).Width=   2672
      Groups(0).Columns(2).Caption=   "Date"
      Groups(0).Columns(2).Name=   "Date"
      Groups(0).Columns(2).CaptionAlignment=   1
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   7
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(3).Width=   2910
      Groups(0).Columns(3).Caption=   "Time"
      Groups(0).Columns(3).Name=   "Time"
      Groups(0).Columns(3).CaptionAlignment=   1
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   7
      Groups(0).Columns(3).FieldLen=   256
      Groups(0).Columns(4).Width=   3598
      Groups(0).Columns(4).Caption=   "Type"
      Groups(0).Columns(4).Name=   "Type"
      Groups(0).Columns(4).CaptionAlignment=   0
      Groups(0).Columns(4).DataField=   "Column 4"
      Groups(0).Columns(4).DataType=   8
      Groups(0).Columns(4).FieldLen=   256
      Groups(0).Columns(5).Width=   3625
      Groups(0).Columns(5).Caption=   "User"
      Groups(0).Columns(5).Name=   "User"
      Groups(0).Columns(5).CaptionAlignment=   0
      Groups(0).Columns(5).DataField=   "Column 5"
      Groups(0).Columns(5).DataType=   8
      Groups(0).Columns(5).FieldLen=   256
      Groups(0).Columns(6).Width=   6297
      Groups(0).Columns(6).Caption=   "Subject"
      Groups(0).Columns(6).Name=   "Subject"
      Groups(0).Columns(6).CaptionAlignment=   0
      Groups(0).Columns(6).DataField=   "Column 6"
      Groups(0).Columns(6).DataType=   8
      Groups(0).Columns(6).FieldLen=   256
      Groups(0).Columns(7).Width=   19103
      Groups(0).Columns(7).Caption=   "Results"
      Groups(0).Columns(7).Name=   "Results"
      Groups(0).Columns(7).CaptionAlignment=   0
      Groups(0).Columns(7).DataField=   "Column 7"
      Groups(0).Columns(7).DataType=   8
      Groups(0).Columns(7).Level=   1
      Groups(0).Columns(7).FieldLen=   256
      Groups(0).Columns(7).HasForeColor=   -1  'True
      Groups(0).Columns(7).HasBackColor=   -1  'True
      Groups(0).Columns(7).ForeColor=   -2147483640
      Groups(0).Columns(7).BackColor=   -2147483643
      _ExtentX        =   20082
      _ExtentY        =   3043
      _StockProps     =   79
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
   Begin VB.Label Label1 
      Caption         =   "Current Group Members"
      Height          =   255
      Index           =   1
      Left            =   3660
      TabIndex        =   2
      Top             =   150
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "Prospecting Groups"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   150
      Width           =   2955
   End
End
Attribute VB_Name = "FProspectMgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData    As CFormData
Attribute FormData.VB_VarHelpID = -1

Private fShowTotals       As Boolean
Private fSettingPrefs     As Boolean

Private lProspectGroupID  As Long

Private rsGroupCategories   As ADODB.Recordset
Private rsProspectGroup     As ADODB.Recordset
Private rsHistory           As ADODB.Recordset

Private Enum eFilter
  FilterNone
  FilterStandard
  FilterAMBest
  FilterAll
End Enum

Private Enum eSortColumn
  SortByGroup
  SortByLabel
End Enum

Private Filter            As eFilter
Private SortColumn        As eSortColumn

Private Sub chkFilter_Click(Index As Integer)
  '*
  '
  On Error GoTo ErrHndlr
  '
  If fSettingPrefs = True Then Exit Sub
  '
  If chkFilter(0).Value = vbChecked And chkFilter(1).Value = vbChecked Then
    Filter = FilterAll
  ElseIf chkFilter(0).Value = vbChecked Then
    Filter = FilterStandard
  ElseIf chkFilter(1).Value = vbChecked Then
    Filter = FilterAMBest
  Else
    Filter = FilterNone
  End If
  '
  SaveSetting App.Title, "ProspectMgt", "FilterGroups", Filter
  '
  SetupGroups
  SelectGroup
  '
  Exit Sub
  '
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.chkFilter.Click", vbCritical, "Error"
End Sub

Private Sub chkTtls_Click()
  '*
  '
  On Error GoTo ErrHndlr
  '
  If fSettingPrefs Then Exit Sub
  '
  If chkTtls.Value = vbChecked Then
    MsgBox "WARNING: This option is fast as dirt. You've been warned.", _
           vbInformation, "I Wouldn't Do That If I Were You"
  End If
  '
  fShowTotals = IIf(chkTtls.Value = vbChecked, True, False)
  SaveSetting App.Title, "ProspectMgt", "ShowTotals", fShowTotals
  '
  SetupGroups
  SelectGroup
  '
  Exit Sub
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.chkTtls.Click", vbCritical, "Error"
End Sub
Private Sub Form_Activate()
  '*
  '
  On Error GoTo ErrHndlr
  '
  Dim lLastContactID As Long
  '
  ReadPreferences
  '
  SetupGroups
  SelectGroup
  '
  On Error Resume Next
  rsProspectGroup.MoveFirst
  lLastContactID = GetSetting(App.Title, "ProspectMgt", "CurrentRow", -1)
  '
  With Me.grdProspectGroup
    .Redraw = False
    '
    Dim iRow As Integer
    '
    For iRow = 0 To .Rows - 1
      .Bookmark = .AddItemBookmark(iRow)
      '
      If .Columns(0).Value = lLastContactID Then
        Exit For
      End If
    Next iRow
    '
    .Redraw = True
  End With
  '
  Exit Sub
  '
ErrHndlr:
  Me.grdProspectGroup.Redraw = True
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.Form_Activate.", vbCritical, "Error"
End Sub

Private Sub Form_Load()
  '*
  '
  On Error GoTo ErrHndlr
  '
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.MinHeight = Me.Height
  FormControl.MinWidth = Me.Width
  FormControl.DataForm = False
  '
  Set rsProspectGroup = New ADODB.Recordset
  Set rsHistory = New ADODB.Recordset
  '
  Exit Sub
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.Form_Load.", vbCritical, "Error"
End Sub

Public Sub SetupGroups()
  '*
  '
  On Error GoTo ErrHndlr
  '
  Dim rsGroups As ADODB.Recordset
  '
  Screen.MousePointer = vbHourglass
  lstGroups.Clear
  '
  Set rsGroupCategories = New ADODB.Recordset
  '
  SortColumn = GetSetting(App.Title, "ProspectMgt", "SortColumn", SortByGroup)
  '
  If Filter = FilterAll Then
    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] ORDER BY [" & SortColumnText(SortColumn) & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  ElseIf Filter = FilterStandard Then
    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] NOT LIKE 'AM Best%' ORDER BY [" & SortColumnText(SortColumn) & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  ElseIf Filter = FilterAMBest Then
    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] LIKE 'AM Best%' ORDER BY [" & SortColumnText(SortColumn) & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  ElseIf Filter = FilterNone Then
    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] = 'Dummy'", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  End If
  '
  Set rsGroups = New ADODB.Recordset
  '
  With rsGroupCategories
    Do While .eof = False
      If rsGroups.State = adStateOpen Then rsGroups.Close
      If fShowTotals Then
        If ConnType = Access Then
          rsGroups.Open "SELECT Count(*) as GroupCount FROM QProspectMgt WHERE " & _
            .Fields("Formula").Value, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
          '
        Else
          Dim sSQL As String
          '
          sSQL = "SELECT Count(*) as GroupCount " & _
            "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
            "WHERE " & ConvertFormula(.Fields("Formula").Value)
          '
          rsGroups.Open sSQL, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
        End If
        '
        lstGroups.AddItem .Fields("Label").Value & " (" & rsGroups.Fields("GroupCount").Value & ")"
      Else
        lstGroups.AddItem .Fields("Label").Value
      End If
      lstGroups.ItemData(lstGroups.NewIndex) = .Fields("RecID").Value
      .MoveNext
    Loop
  End With
  '
  Screen.MousePointer = vbDefault
  '
  Exit Sub
  '
ErrHndlr:
  Screen.MousePointer = vbDefault
  DoEvents
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.SetupGroups.", vbCritical, "Error"
End Sub

Private Sub Form_Resize()
  '*
  '
  On Error Resume Next
  grdHistory.Height = FProspectMgt.Height - grdHistory.Top - 160
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '*
  '
  On Error Resume Next
  SaveSetting App.Title, "ProspectMgt", "CurrentRow", grdProspectGroup.Columns(0).Value
End Sub

Private Sub grdProspectGroup_Click()
  '*
  '
  On Error GoTo ErrHndlr
  '
  LoadHistory Me.grdProspectGroup.Columns(0).Value
  '
  Exit Sub
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.grdProspectGroup_Click.", vbCritical, "Error"
End Sub

Private Sub grdProspectGroup_DblClick()
  '*
  '
  On Error GoTo ErrHndlr
  '
 ' Company.Fetch grdProspectGroup.Columns(1).Value
  'Company.Contact.Fetch grdProspectGroup.Columns(0).Value
  '
  FContact.LoadContact grdProspectGroup.Columns(0).Value, True
  FormMgr.ShowForm Me, FContact
  '
  Exit Sub
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.grdProspectGroup_DblClick.", vbCritical, "Error"
End Sub

Private Sub lstGroups_Click()
  '*
  '
  On Error GoTo ErrHndlr
  '
  Screen.MousePointer = vbHourglass
  DoEvents
  '
  rsGroupCategories.MoveFirst
  rsGroupCategories.Find "RecID = " & lstGroups.ItemData(lstGroups.ListIndex), , adSearchForward
  '
  If Not rsGroupCategories.eof Then
    Me.grdProspectGroup.Redraw = False
    Me.grdProspectGroup.RemoveAll
    '
    If rsProspectGroup.State = adStateOpen Then
      rsProspectGroup.Close
    End If
    '
    lProspectGroupID = lstGroups.ItemData(lstGroups.ListIndex)
    SaveSetting App.Title, "ProspectMgt", "CurrentGroupRecID", lProspectGroupID
    '
    If ConnType = Access Then
      rsProspectGroup.Open "SELECT * FROM QProspectMgt WHERE " & _
        rsGroupCategories!Formula, cnMain, adOpenKeyset, adLockReadOnly, adCmdText
    Else 'SQL
      Dim sSQL As String
      '
      sSQL = "SELECT  TCompany.ID AS CompanyID, TContact.ID AS ContactID, " & _
        "TCompany.Name AS Company, " & _
        "TContact.[FirstName] + ' ' + [LastName] AS FullName, " & _
        "TContact.State, TContact.Status, TContact.AuthDate, " & _
        "TContact.Status, TContact.ShipStatus, TContact.AuthStatus, " & _
        "TContact.AuthDate, TContact.AuthDays, TContact.ShipDate " & _
        "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
        "WHERE " & ConvertFormula(rsGroupCategories!Formula) & " ORDER BY TContact.State"
      '
      rsProspectGroup.Open sSQL, cnMain, adOpenKeyset, adLockReadOnly, adCmdText
    End If
    '
    With rsProspectGroup
      Do Until .eof
        If ConnType = Access Then
          Me.grdProspectGroup.AddItem .Fields("TContact.ID").Value & vbTab & _
            .Fields("TCompany.ID").Value & vbTab & _
            !FullName & vbTab & _
            !Company & vbTab & _
            !State & vbTab & _
            !Status & vbTab & _
            !AuthDate
        Else
          Me.grdProspectGroup.AddItem .Fields("ContactID").Value & vbTab & _
            .Fields("CompanyID").Value & vbTab & _
            !FullName & vbTab & _
            !Company & vbTab & _
            !State & vbTab & _
            !Status & vbTab & _
            !AuthDate
        End If
        '
        .MoveNext
      Loop
    End With
    '
    Me.grdProspectGroup.Redraw = True
  End If
  '
  Screen.MousePointer = vbDefault
  DoEvents
  '
  Exit Sub
  '
ErrHndlr:
  Me.grdProspectGroup.Redraw = True
  Screen.MousePointer = vbDefault
  DoEvents
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.lstGroups_Click.", vbCritical, "Error"
End Sub
Public Sub ReadPreferences()
  '*
  '
  On Error GoTo ErrHndlr
  '
  fSettingPrefs = True
  '
  '\\ Filter
  Filter = GetSetting(App.Title, "ProspectMgt", "FilterGroups", FilterAll)
  '
  If Filter = FilterAll Then
    chkFilter(0).Value = vbChecked
    chkFilter(1).Value = vbChecked
  Else
    chkFilter(0).Value = IIf(Filter = FilterStandard, vbChecked, vbUnchecked)
    chkFilter(1).Value = IIf(Filter = FilterAMBest, vbChecked, vbUnchecked)
  End If
  '
  '\\ Totals
  fShowTotals = GetSetting(App.Title, "ProspectMgt", "ShowTotals", "False")
  chkTtls.Value = IIf(fShowTotals, vbChecked, vbUnchecked)
  '
  '\\ Sort
  SortColumn = GetSetting(App.Title, "ProspectMgt", "SortColumn", SortByGroup)
  '
  If SortColumn = SortByGroup Then
    optSort(0).Value = True
  Else
    optSort(1).Value = True
  End If
  '
  fSettingPrefs = False
  '
  Exit Sub
  '
ErrHndlr:
  fSettingPrefs = False
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.General.ReadPreferences", vbCritical, "Error"
End Sub

Private Sub optSort_Click(Index As Integer)
  '*
  '
  On Error GoTo ErrHndlr
  '
  If fSettingPrefs = True Then Exit Sub
  '
  SortColumn = Index
  SaveSetting App.Title, "ProspectMgt", "SortColumn", Index
  '
  SetupGroups
  SelectGroup
  '
  Exit Sub
  '
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.optSort.Click", vbCritical, "Error"
End Sub

Public Sub SelectGroup()
  '*
  '
  On Error GoTo ErrHndlr
  '
  '\\ Local Declarations
  Dim iCur      As Integer
  Dim iCt       As Integer
  '
  lProspectGroupID = GetSetting(App.Title, "ProspectMgt", "CurrentGroupRecID", 0)
  '
  With lstGroups
    iCt = lstGroups.ListCount - 1
    If iCt < 0 Then Exit Sub
    '
    For iCur = 0 To iCt
      If .ItemData(iCur) = lProspectGroupID Then
        .ListIndex = iCur
        Exit Sub
      End If
    Next
    '
    .ListIndex = 0
  End With
  '
  Exit Sub
  '
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.General.SelectGroup", vbCritical, "Error"
End Sub

Private Sub LoadHistory(ByVal plCustomerID As Long)
  '*
  '
  On Error GoTo EH
  '
  grdHistory.Redraw = False
  '
  Me.grdHistory.RemoveAll
  '
  If plCustomerID <> -1 Then
    Dim rsHistory As ADODB.Recordset
    '
    Set rsHistory = New ADODB.Recordset
    '
    If ConnType = Access Then
      rsHistory.Open "SELECT * FROM tblSupportActs WHERE CustRecID = " & plCustomerID & " ORDER BY Date DESC, Time DESC", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
      Dim cmdHistory As New ADODB.Command
      '
      With cmdHistory
        Set .ActiveConnection = cnMain
        .CommandText = "dbo.UpParmSelSupportActs"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("CustomerID", adInteger, adParamInput, , plCustomerID)
        Set rsHistory = .Execute
      End With 'cmdSupportAct
      '
      Set cmdHistory = Nothing
    End If
    '
    With rsHistory
    Do While Not .eof
      grdHistory.AddItem !RecID & vbTab & !CustRecID & vbTab & !Date & vbTab & !Time & vbTab & !Type & vbTab & !User & vbTab & !Subject & vbTab & !Results
      .MoveNext
    Loop
    End With 'rsHistory
    '
    rsHistory.Close
    Set rsHistory = Nothing
  End If
  '
  grdHistory.Redraw = True
  '
  Exit Sub
EH:
  grdHistory.Redraw = True
  MsgBox Err.Description
End Sub

Private Function SortColumnText(ByVal SortType As eSortColumn) As String
  '*
  '
  If SortType = SortByGroup Then
    SortColumnText = "Priority"
  ElseIf SortType = SortByLabel Then
    SortColumnText = "Label"
  End If
End Function

Private Function ConvertFormula(ByVal psFormula As String) As String
  On Error GoTo EH
  '
  Dim sFormula As String
  '
  sFormula = psFormula
  '
  sFormula = Replace(sFormula, "ShipDays", "DateDiff(Day,[ShipDate],GETDATE())")
  sFormula = Replace(sFormula, "AuthDaysRemaining", "([AuthDays] - DateDiff(Day, [AuthDate], GETDATE()))")
  sFormula = Replace(sFormula, "isnull(shipdate)", "(ShipDate = Null)")
  '
  ConvertFormula = sFormula
  '
  Exit Function
EH:
  MsgBox Err.Description & " in Convert Formula.)"
End Function
