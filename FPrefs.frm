VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FPrefs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2640
      TabIndex        =   3
      Top             =   1650
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1050
      TabIndex        =   2
      Top             =   1650
      Width           =   1395
   End
   Begin SSDataWidgets_B.SSDBCombo cboStatus 
      DataField       =   "Status"
      Height          =   315
      Left            =   2610
      TabIndex        =   0
      Tag             =   "1"
      Top             =   270
      Width           =   2445
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   4313
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   16777215
   End
   Begin SSDataWidgets_B.SSDBCombo cboAuthStatus 
      DataField       =   "AuthStatus"
      Height          =   315
      Left            =   2610
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1050
      Width           =   2445
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   4313
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo cboShipStatus 
      DataField       =   "AuthStatus"
      Height          =   315
      Left            =   2610
      TabIndex        =   6
      Tag             =   "1"
      Top             =   660
      Width           =   2445
      DataFieldList   =   "Column 0"
      ListAutoValidate=   0   'False
      _Version        =   196617
      DataMode        =   2
      Cols            =   1
      ColumnHeaders   =   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   4313
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label Label2 
      Caption         =   "Initial Shipping Status:"
      Height          =   285
      Index           =   1
      Left            =   420
      TabIndex        =   7
      Top             =   690
      Width           =   2025
   End
   Begin VB.Label Label2 
      Caption         =   "Initial Authorization Status:"
      Height          =   285
      Index           =   0
      Left            =   420
      TabIndex        =   5
      Top             =   1080
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "Initial Customer Status:"
      Height          =   285
      Left            =   420
      TabIndex        =   1
      Top             =   300
      Width           =   2085
   End
End
Attribute VB_Name = "FPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAuthStatus_InitColumnProps()
  On Error GoTo ErrCall
  '
  Dim rs As ADODB.Recordset
  '
  'Set rs = dbMain.OpenRecordset("SELECT * FROM tblAuthStatus ORDER BY RecID", dbOpenForwardOnly)
  Set rs = New ADODB.Recordset
  rs.Open "SELECT * FROM tblAuthStatus ORDER BY RecID", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rs.EOF
    cboAuthStatus.AddItem rs!Status
    rs.MoveNext
  Loop
  '
  DBOps.ZapRS rs
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmPrefs.cboAuthStatus_InitColumnProps.", vbCritical, "Error"
End Sub

Private Sub cboStatus_InitColumnProps()
  On Error GoTo ErrCall
  '
  Dim rsStatus As ADODB.Recordset
  '
  'Set rsStatus = dbMain.OpenRecordset("SELECT * FROM tblStatus", dbOpenForwardOnly)
  Set rsStatus = New ADODB.Recordset
  rsStatus.Open "SELECT * FROM tblStatus", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rsStatus.EOF
    cboStatus.AddItem "" & rsStatus!Status
    rsStatus.MoveNext
  Loop
  '
  DBOps.ZapRS rsStatus
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmPrefs.cboStatus_InitColumnProps.", vbCritical, "Error"
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  SaveSetting App.Title, "Preferences", "InitStatus", cboStatus.Text
  SaveSetting App.Title, "Preferences", "InitShipStatus", cboShipStatus.Text
  SaveSetting App.Title, "Preferences", "InitAuthStatus", cboAuthStatus.Text
  Unload Me
End Sub

Private Sub Form_Load()
  cboStatus.Text = GetSetting(App.Title, "Preferences", "InitStatus", "Prospect")
  cboShipStatus.Text = GetSetting(App.Title, "Preferences", "InitShipStatus", "Not Shipped")
  cboAuthStatus.Text = GetSetting(App.Title, "Preferences", "InitAuthStatus", "Not Authorized")
End Sub

Private Sub cboShipStatus_InitColumnProps()
  On Error GoTo ErrCall
  '
  Dim rsStatus As ADODB.Recordset
  '
  'Set rsStatus = dbMain.OpenRecordset("SELECT * FROM tblShipStatus", dbOpenForwardOnly)
  Set rsStatus = New ADODB.Recordset
  rsStatus.Open "SELECT * FROM tblShipStatus", cnMain, adOpenForwardOnly, adLockReadOnly
  '
  Do While Not rsStatus.EOF
    cboShipStatus.AddItem "" & rsStatus!Status
    rsStatus.MoveNext
  Loop
  '
  DBOps.ZapRS rsStatus
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmPrefs.SSDBCombo1_InitColumnProps.", vbCritical, "Error"
End Sub
