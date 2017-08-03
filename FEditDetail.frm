VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FEditDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Call"
   ClientHeight    =   2595
   ClientLeft      =   2715
   ClientTop       =   4110
   ClientWidth     =   9360
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOpenCall 
      Caption         =   "Open Call"
      Enabled         =   0   'False
      Height          =   255
      Left            =   150
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboCase 
      Enabled         =   0   'False
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboType 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1995
      DataFieldList   =   "Column 0"
      _Version        =   196617
      DataMode        =   2
      ColumnHeaders   =   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).Caption=   "Type"
      Columns(0).Name =   "Type"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   3519
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      Enabled         =   0   'False
      DataFieldToDisplay=   "Column 0"
   End
   Begin TDBTime6Ctl.TDBTime ttmTime 
      Height          =   315
      Left            =   720
      TabIndex        =   14
      Top             =   840
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "FEditDetail.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FEditDetail.frx":006C
      Spin            =   "FEditDetail.frx":00BC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn AMPM"
      EditMode        =   0
      Enabled         =   0
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "14:01"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   0.584363425925926
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   315
      Left            =   6360
      TabIndex        =   0
      Top             =   2130
      Width           =   1035
   End
   Begin VB.TextBox txtResults 
      DataField       =   "Results"
      Enabled         =   0   'False
      Height          =   1605
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "1"
      Top             =   450
      Width           =   3765
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   1155
      Left            =   7800
      Picture         =   "FEditDetail.frx":00E4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   1155
      Left            =   7800
      Picture         =   "FEditDetail.frx":0526
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1155
   End
   Begin VB.TextBox txtSubject 
      DataField       =   "Subject"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Tag             =   "1"
      Top             =   90
      Width           =   3765
   End
   Begin GTMaskDate.GTMaskDate mskDate 
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Tag             =   "1"
      Top             =   480
      Width           =   1275
      _Version        =   65537
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Enabled         =   0   'False
      AllowNull       =   0   'False
      BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MaskCentury     =   2
      DataField       =   "Date"
      BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Caption         =   "Case:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   780
      TabIndex        =   13
      Top             =   1710
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "Results:"
      Height          =   225
      Index           =   1
      Left            =   2970
      TabIndex        =   12
      Top             =   450
      Width           =   645
   End
   Begin VB.Label Label5 
      Caption         =   "Type:"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "User:"
      Height          =   285
      Index           =   5
      Left            =   150
      TabIndex        =   10
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Date:"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   510
      Width           =   465
   End
   Begin VB.Label Label5 
      Caption         =   "Time:"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Subject:"
      Height          =   225
      Index           =   3
      Left            =   2940
      TabIndex        =   7
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "FEditDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CboEvents As New CComboSearch
Private lID As Long
Private OpenEdit As Boolean
Private ClosedDate As Date

Private Sub cboType_GotFocus()
  CboEvents.Setup cboType
End Sub

Private Sub cboType_InitColumnProps()
  On Error GoTo EH
  '
  Dim rsType As ADODB.Recordset
  '
  Set rsType = New ADODB.Recordset
  rsType.Open "SELECT * FROM tblActivities ORDER BY Activity", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  cboType.Redraw = False
  Do While Not rsType.eof
    cboType.AddItem rsType!Activity
    rsType.MoveNext
  Loop
  cboType.Redraw = True
  '
  DBOps.ZapRS rsType
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in FEditDetail: cboType_InitColumnProps."
End Sub

Private Sub chkOpenCall_Click()
  'ClosedDate = Format(Now, "m/d/yy h:mm AM/PM")
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdEdit_Click()
  On Error Resume Next
  cboType.Enabled = True
  cboType.SetFocus
  txtSubject.Enabled = True
  txtResults.Enabled = True
  cboCase.Enabled = True
  If OpenEdit = True Then
    chkOpenCall.Enabled = True
  End If
End Sub

Private Sub cmdSave_Click()
  On Error GoTo EH
  '
  Dim sSQL As String
  Dim rsCase As New ADODB.Recordset
  Dim rsCaseLink As New ADODB.Recordset
  Dim iCaseNum As Integer
  '
  If chkOpenCall = False Then
    If ConnType = Access Then
      sSQL = "UPDATE tblSupportActs SET Type = '" & cboType.Text & "', " & _
            "Results = '" & Replace(txtResults.Text, "'", "''") & "', Subject = '" & Replace(txtSubject.Text, "'", "''") & "', ClosedTime = '" & Replace(Format(Now, "m/d/yy h:mm AM/PM"), "'", "''") & "', OpenCall = '" & Replace(chkOpenCall.Value, "'", "''") & "' " & _
            "WHERE RecID = " & lID
    Else
      sSQL = "UPDATE TSupportActs SET Type = '" & cboType.Text & "', " & _
            "Results = '" & Replace(txtResults.Text, "'", "''") & "', Subject = '" & Replace(txtSubject.Text, "'", "''") & "', ClosedTime = '" & Replace(Format(Now, "m/d/yy h:mm AM/PM"), "'", "''") & "', OpenCall = '" & Replace(chkOpenCall.Value, "'", "''") & "' " & _
            "WHERE RecID = " & lID
    End If
  Else
    If ConnType = Access Then
      sSQL = "UPDATE tblSupportActs SET Type = '" & cboType.Text & "', " & _
            "Results = '" & Replace(txtResults.Text, "'", "''") & "', Subject = '" & Replace(txtSubject.Text, "'", "''") & "', OpenCall = '" & Replace(chkOpenCall.Value, "'", "''") & "' " & _
            "WHERE RecID = " & lID
    Else
      sSQL = "UPDATE TSupportActs SET Type = '" & cboType.Text & "', " & _
            "Results = '" & Replace(txtResults.Text, "'", "''") & "', Subject = '" & Replace(txtSubject.Text, "'", "''") & "', OpenCall = '" & Replace(chkOpenCall.Value, "'", "''") & "' " & _
            "WHERE RecID = " & lID
    End If
  End If
  '
  cnMain.Execute sSQL
  '
  If Not cboCase.ListIndex = -1 Then
    rsCase.Open "Select [CaseName], [CaseID] from TCase Where [CaseName] = '" & cboCase.Text & "'", cnMain, adOpenDynamic, adLockBatchOptimistic
      With rsCase
        If Not .eof Then
          iCaseNum = !CaseID
        End If
      End With
      '
    rsCaseLink.Open "Select * from TCaseSupportActLink where CaseID = '" & iCaseNum & "'", cnMain, adOpenDynamic, adLockBatchOptimistic
      With rsCaseLink
        If Not .eof Then
          While Not .eof
            If !SupportActID = lID Then
              cboCase.Visible = False
              Label1.Visible = False
              Unload Me
              Exit Sub
            End If
            .MoveNext
          Wend
        End If
        .Close
      End With
      '
    rsCaseLink.Open "Select * from TCaseSupportActLink", cnMain, adOpenDynamic, adLockBatchOptimistic
      With rsCaseLink
        .AddNew
        !CaseID = iCaseNum
        !SupportActID = lID
        .UpdateBatch
      End With
    rsCase.Close
    rsCaseLink.Close
  End If
  '
  Unload Me
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in FEditDetail: cmdSave_Click."
End Sub

Public Sub ShowRecord(plID As Long, Optional pOpenEdit As Boolean)
  'On Error GoTo EH
  '
  Dim rsDetail As ADODB.Recordset
  Dim rsListCases As New ADODB.Recordset
  Dim rsCaseLink As New ADODB.Recordset
  Dim i As Integer
  Dim sCaseName As String
  Dim iCaseID As Integer
  '
  lID = plID
  OpenEdit = pOpenEdit
  If OpenEdit = True Then
    chkOpenCall.Visible = True
  Else
    chkOpenCall.Visible = False
  End If
  '
  Set rsDetail = New ADODB.Recordset
  '
  If ConnType = Access Then
    rsDetail.Open "SELECT * FROM tblSupportActs WHERE RecID = " & lID, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  Else
    rsDetail.Open "SELECT * FROM TSupportActs WHERE RecID = " & lID, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  End If
  '
  With rsDetail
  If Not .eof Then
    Me.cboType.Text = !Type & vbNullString
    Me.mskDate.DateValue = nnNum(!Date)
    Me.ttmTime.ValidateMode = nnNum(!Time)
    Me.txtSubject = !Subject & vbNullString
    Me.txtResults = !Results & vbNullString
    Me.lblUser = !User & vbNullString
    If !OpenCall = True Then
      Me.chkOpenCall.Value = 1
    Else
      Me.chkOpenCall.Value = 0
    End If
  Else
    MsgBox "Record not found.", vbInformation, "Edit Detail"
  End If
  End With
  '
  'Loads combo box with case names
  If Not bCases Then
    cboCase.Clear
    cboCase.Visible = True
    Label1.Visible = True
    rsListCases.Open "Select [CaseName] from TCase", cnMain, adOpenDynamic, adLockBatchOptimistic
    With rsListCases
      If Not .eof Then
        While Not .eof
          cboCase.AddItem !CaseName
          .MoveNext
        Wend
      End If
      .Close
    End With
  End If
  '
  'Checks to see if Message is connected to a case and displays it in combo box
  rsCaseLink.Open "Select * from TCaseSupportActLink where [SupportActID] = '" & lID & "'", cnMain, adOpenDynamic, adLockBatchOptimistic
    With rsCaseLink
      If Not .eof Then
        iCaseID = !CaseID
      End If
    End With
    '
  If iCaseID > 0 Then
    rsListCases.Open "Select * from TCase where [CaseID] = '" & iCaseID & "'", cnMain, adOpenDynamic, adLockBatchOptimistic
      With rsListCases
        If Not .eof Then
          sCaseName = !CaseName
        End If
      End With
      '
    For i = 0 To cboCase.ListCount
      If sCaseName = cboCase.list(i) Then
        cboCase.ListIndex = i
      End If
    Next i
  End If
  '
  Me.Show vbModal, FMain
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in FEditDetail: ShowRecord."
End Sub

