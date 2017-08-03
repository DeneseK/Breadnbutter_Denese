VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FSupportLog 
   BorderStyle     =   0  'None
   Caption         =   "Support Log"
   ClientHeight    =   5400
   ClientLeft      =   3810
   ClientTop       =   3225
   ClientWidth     =   9510
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFirst 
      Caption         =   "Go to First Record"
      Height          =   315
      Left            =   2790
      TabIndex        =   4
      Top             =   5010
      Width           =   1665
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Go to Last Record"
      Height          =   315
      Left            =   4500
      TabIndex        =   3
      Top             =   5010
      Width           =   1665
   End
   Begin VB.ComboBox cboShow 
      Height          =   315
      ItemData        =   "FSupportLog.frx":0000
      Left            =   720
      List            =   "FSupportLog.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5010
      Width           =   1995
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdSupportLog 
      Bindings        =   "FSupportLog.frx":002E
      Height          =   4965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   1
      AllowUpdate     =   0   'False
      SelectTypeRow   =   3
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      ExtraHeight     =   26
      Columns.Count   =   7
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "RecID"
      Columns(0).Name =   "RecID"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "CustRecID"
      Columns(1).Name =   "CustRecID"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Caption=   "CompanyContact"
      Columns(2).Name =   "CompanyContact"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2037
      Columns(3).Caption=   "Date"
      Columns(3).Name =   "Date"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   7
      Columns(3).FieldLen=   256
      Columns(4).Width=   2566
      Columns(4).Caption=   "Type"
      Columns(4).Name =   "Type"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2540
      Columns(5).Caption=   "User"
      Columns(5).Name =   "User"
      Columns(5).CaptionAlignment=   0
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   4815
      Columns(6).Caption=   "Results"
      Columns(6).Name =   "Results"
      Columns(6).CaptionAlignment=   0
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).VertScrollBar=   -1  'True
      _ExtentX        =   16113
      _ExtentY        =   8758
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
   Begin VB.Label lblShow 
      Caption         =   "Show:"
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   5040
      Width           =   645
   End
End
Attribute VB_Name = "FSupportLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1

Private rsSupportAct As ADODB.Recordset
Private lSupportActRecID As Long

Private Sub cboShow_Click()
  SetRecordset
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo EH
  '
  rsSupportAct.MoveFirst
  Me.grdSupportLog.Bookmark = rsSupportAct.Bookmark
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in Command First Click."
End Sub

Private Sub cmdLast_Click()
  On Error GoTo EH
  '
  rsSupportAct.MoveLast
  Me.grdSupportLog.Bookmark = rsSupportAct.Bookmark
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in Command Last Click."
End Sub

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
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSupportLog.Form_Initialize.", vbCritical, "Error"
End Sub

Private Sub Form_Load()
  On Error GoTo EH
  '
  Set rsSupportAct = New ADODB.Recordset
  '
  cboShow.ListIndex = 0
  SetRecordset
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in FSupportLog:Load"
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  grdSupportLog.Redraw = False
  grdSupportLog.Move 0, 0, Me.ScaleWidth, Me.Height - 435
  lblShow.Top = grdSupportLog.Height + 60
  cboShow.Top = grdSupportLog.Height + 45
  cmdLast.Top = cboShow.Top
  cmdFirst.Top = cboShow.Top
  '
  grdSupportLog.Columns(6).Width = grdSupportLog.Width - grdSupportLog.Columns(6).Left - 255
  grdSupportLog.Redraw = True
End Sub

Private Sub SetRecordset()
  On Error GoTo EH
  '
  Screen.MousePointer = vbHourglass
  DoEvents
  '
  If rsSupportAct.State = adStateOpen Then
    rsSupportAct.Close
  End If
  '
  Dim sSQL As String
  '
  If ConnType = Access Then
    Select Case cboShow.ListIndex
    Case 0 'Today
      sSQL = "SELECT * FROM QSupportActs WHERE Date = #" & Format(Now, "Short Date") & "# ORDER BY Date DESC, Time DESC"
    Case 1 'Previous 7 Days
      sSQL = "SELECT * FROM QSupportActs where Date > #" & Format(DateAdd("d", -7, Now), "Short Date") & "# ORDER BY Date DESC, Time DESC"
    Case Else 'All
      sSQL = "SELECT * FROM QSupportActs ORDER BY Date DESC, Time DESC"
    End Select
    '
    rsSupportAct.Open sSQL, cnMain, adOpenStatic, adLockReadOnly, adCmdText
  Else
    Dim cmdRecords As New ADODB.Command
    '
    With cmdRecords
      Set .ActiveConnection = cnMain
      .CommandText = "dbo.UpParmSelSupportActsByDateRange"
      .CommandType = adCmdStoredProc
      '
      Dim dtOldestDate As Date
      '
      Select Case cboShow.ListIndex
      Case 0 'Today
        dtOldestDate = DateAdd("d", -1, Now)
      Case 1 'Previous 7 Days
        dtOldestDate = DateAdd("d", -8, Now)
      Case Else 'All
        dtOldestDate = CDate("01/01/1980")
      End Select
      '
      .Parameters.Append .CreateParameter("Date", adDate, adParamInput, , dtOldestDate)
      Set rsSupportAct = .Execute
    End With
    '
    Set cmdRecords = Nothing
  End If
  '
  Me.grdSupportLog.ReBind
  '
  Screen.MousePointer = vbDefault
  DoEvents
  '
  Exit Sub
EH:
  Screen.MousePointer = vbDefault
  DoEvents
  MsgBox Err.Description & " in FSupportLog: Set Recordset."
End Sub

Private Sub grdSupportLog_UnboundReadData(ByVal RowBuf As SSDataWidgets_B_OLEDB.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  On Error GoTo EH
  '
  Dim iRBRow As Integer
  Dim iGridRows As Integer
  Dim iFld As Integer
  '
  iGridRows = 0
  '
  'This code initializes the procedure by declaring the variables
  'that will be used to move data into the Grid.
  'iRBRow will be used to count the number of rows of data
  'requested by the ssRowBuffer object (RowBuf) that is passed to the event.
  'iGridRows will count how many rows of data should be supplied to the Grid
  'from the data source.
  'iFld will be used as a generic counter when pulling data from the recordset.
  'Setting iGridRows to 0 indicates that, at the start of the event,
  'no rows have been read from the recordset.
  '
  With rsSupportAct
    If Not (.BOF And .eof) Then
      If IsNull(StartLocation) Then 'If the Grid is empty
        If ReadPriorRows Then     'If ReadPriorRows is True
          .MoveLast               'then the Grid is being
        Else                      'scrolled up towards the top
          .MoveFirst
        End If
      Else                         'If Grid contains data
        .Bookmark = StartLocation
        If ReadPriorRows Then
          .MovePrevious
        Else
          .MoveNext
        End If
      End If
      '
      For iRBRow = 0 To RowBuf.RowCount - 1
        If .BOF Or .eof Then Exit For
        '
        Select Case RowBuf.ReadType
          Case ssReadTypeAllData      'All data must be read
            RowBuf.Bookmark(iRBRow) = .Bookmark
            '
            RowBuf.value(iRBRow, 0) = !RecID
            RowBuf.value(iRBRow, 1) = !CustRecID
            RowBuf.value(iRBRow, 2) = !CompanyContact
            RowBuf.value(iRBRow, 3) = !Date
            RowBuf.value(iRBRow, 4) = !Type
            RowBuf.value(iRBRow, 5) = !User
            RowBuf.value(iRBRow, 6) = !Results
          Case ssReadTypeBookmarkOnly 'Only bookmarks must be read
            RowBuf.Bookmark(iRBRow) = .Bookmark
        End Select    'Cases 2 and 3 are not used by DBGrid
        '
        If ReadPriorRows Then
          .MovePrevious
        Else
          .MoveNext
        End If
        '
        iGridRows = iGridRows + 1
      Next iRBRow
      '
      RowBuf.RowCount = iGridRows
    End If
  End With 'rsSupportAct
  '
  Exit Sub
EH:
  MsgBox Err.Description
End Sub
