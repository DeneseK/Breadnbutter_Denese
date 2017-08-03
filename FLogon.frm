VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Begin VB.Form FLogon 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4215
   ClientLeft      =   7680
   ClientTop       =   2595
   ClientWidth     =   4905
   Icon            =   "FLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboServer 
      Height          =   315
      ItemData        =   "FLogon.frx":000C
      Left            =   1500
      List            =   "FLogon.frx":0013
      TabIndex        =   6
      Text            =   "cboServer"
      Top             =   840
      Width           =   2985
   End
   Begin VB.ComboBox cboDatabase 
      Height          =   315
      ItemData        =   "FLogon.frx":0023
      Left            =   1500
      List            =   "FLogon.frx":002A
      TabIndex        =   7
      Text            =   "cboDatabase"
      Top             =   1230
      Width           =   2985
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   1005
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1155
   End
   Begin VB.TextBox txtName 
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1500
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1710
      Width           =   2985
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1500
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2100
      Width           =   2985
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   1005
      Left            =   3270
      Picture         =   "FLogon.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log"
      Default         =   -1  'True
      Height          =   1005
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1185
   End
   Begin VB.Timer tmrClock 
      Interval        =   60000
      Left            =   5370
      Top             =   1650
   End
   Begin TDBTime6Ctl.TDBTime ttmTime 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   2520
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   556
      Caption         =   "FLogon.frx":0342
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FLogon.frx":03AE
      Spin            =   "FLogon.frx":03FE
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
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn AMPM"
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
      Text            =   "__:__ "
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   0.6490625
   End
   Begin TDBDate6Ctl.TDBDate tdtDate 
      Height          =   315
      Left            =   1500
      TabIndex        =   4
      Top             =   2520
      Width           =   1425
      _Version        =   65536
      _ExtentX        =   2514
      _ExtentY        =   556
      Calendar        =   "FLogon.frx":0426
      Caption         =   "FLogon.frx":053E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FLogon.frx":05AA
      Keys            =   "FLogon.frx":05C8
      Spin            =   "FLogon.frx":0626
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   3.62586437654789E-316
      CenturyMode     =   0
   End
   Begin VB.Label Label5 
      Caption         =   "Server:"
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   870
      Width           =   645
   End
   Begin VB.Label Label4 
      Caption         =   "Database:"
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   390
      TabIndex        =   12
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Date/Time:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   2580
      Width           =   945
   End
   Begin VB.Label lblWelcomeMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   390
      TabIndex        =   9
      Top             =   0
      Width           =   4155
   End
End
Attribute VB_Name = "FLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Public Enum enMode
  enLogin
  enLogout
End Enum

Private iMode As enMode

Private Sub cmdCancel_Click()
  On Error GoTo ErrCall
  '
  txtPassword = ""
  User.LogResults = False
  Unload Me
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FLogon.cmdCancel_Click.", vbCritical, "Error"
End Sub

Private Function Submit(InOut%, bLog As Boolean) As Boolean
  'On Error GoTo ErrCall
  '
  Dim i As Integer
  Dim j As Integer
  Dim strMsg As String
  Dim strDate As String
  Dim resp%
  Dim rsEmployee As ADODB.Recordset
  Dim rsLog As ADODB.Recordset
  Dim bEdit As Boolean
  '####################################################################
  If SQLLogon = False Then
    MsgBox "SQL Server Logon Failed!"
    Exit Function
  End If
  If cnMain.State = adStateOpen Then
  '####################################################################
  Set rsEmployee = New ADODB.Recordset
  '
  If ConnType = Access Then
    rsEmployee.Open "SELECT *, EmployeeFirst & ' ' & EmployeeLast AS Name FROM tblEmployees", cnMain, adOpenDynamic, adLockOptimistic, adCmdText
  Else 'SQL Server
    rsEmployee.Open "SELECT *, EmployeeFirst + ' ' + EmployeeLast AS Name FROM tblEmployees", cnMain, adOpenDynamic, adLockOptimistic, adCmdText
  End If
  '
  Set rsLog = New ADODB.Recordset
  '
  If ConnType = Access Then
    rsLog.Open "SELECT * FROM tblHours WHERE LogDate = #" & Format(tdtDate.Text, "mm/dd/yy") & "# ORDER BY RecID", cnMain, adOpenKeyset, adLockOptimistic, adCmdText
  Else 'SQL Server
    rsLog.Open "SELECT * FROM tblHours WHERE LogDate = '" & Format(tdtDate.Text, "mm/dd/yy") & "' ORDER BY RecID", cnMain, adOpenKeyset, adLockOptimistic, adCmdText
  End If
  '
  rsEmployee.Find "Name = '" & txtName.Text & "'"
  '
  If Not rsEmployee.eof Then
    If DecryptStr(rsEmployee!Password & "") = txtPassword.Text Then
      If bLog Then
        If InOut% = 1 Then '(In)
          If Not rsLog.eof Then rsLog.MoveLast
          rsLog.Find "Employee = '" & txtName.Text & "'", , adSearchBackward
          If rsLog.BOF Then
            rsLog.AddNew
          Else
            If IsNull(rsLog!actualout) Then
              MsgBox "You have already logged in but not logged out for this date. Please log out before logging back in."
            Else
              rsLog.AddNew
            End If
          End If
        Else 'InOut% = 0 (Out)
          rsLog.MoveLast
          rsLog.Find "Employee = '" & txtName & "'", , adSearchBackward
          If rsLog.BOF Then
            MsgBox "You have not logged in on this date. Please log in before logging out."
          Else
            If IsNull(rsLog!actualout) Then
              bEdit = True
            Else
              MsgBox "You have logged in and logged out for this date. Please log in before logging out."
            End If
          End If
        End If
        '
        If bEdit Or rsLog.EditMode = adEditAdd Then
          strMsg = "The following will be submitted: " & vbCrLf & _
            vbCrLf & _
            IIf(InOut%, "LOG IN", "LOG OUT") & vbCrLf & vbCrLf & _
            "Name: " & vbTab & vbTab & txtName & vbCrLf & _
            "Time: " & vbTab & vbTab & tdtDate.Text & " " & ttmTime.Text
          resp% = MsgBox(strMsg, vbOKCancel)
          '
          If resp% = vbOK Then
            'If bEdit Then rsLog.Edit
            '
            If InOut% = 1 Then
              rsLog!actualin = Now
              rsLog!logdate = CDate(Format(tdtDate.Text, "mm/dd/yy"))
              rsLog!intime = ttmTime
              User.Name = txtName
            Else
              rsLog!actualout = Now
              rsLog!outtime = ttmTime
              sUserName = ""
            End If
            '
            rsLog!Employee = txtName
            rsLog.Update
            '
            SaveSetting App.Title, "Settings", "User", txtName
            txtPassword = ""
            txtPassword.SetFocus
            '
            Submit = True
          Else
            Submit = False
          End If
        Else
          Submit = False
        End If
      Else
        User.Name = txtName.Text
        Submit = True
      End If
    Else
      MsgBox "Password incorrect. Please try again."
      txtPassword.SetFocus
      Submit = False
    End If
  Else
    MsgBox "Employee name not found. Please try again."
    txtName.SetFocus
    Submit = False
  End If
   '
  DBOps.ZapRS rsEmployee
  DBOps.ZapRS rsLog
   '
  End If
  Exit Function
 'ErrCall:
 '  Submit = False
 '  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FLogon.Submit", vbCritical, "Error"
End Function

Private Sub cmdContinue_Click()
'  On Error GoTo ErrCall
  '
  Select Case iMode
  Case enLogin
    If Submit(1, False) Then
      User.LogResults = True
      SaveSetting App.Title, "Settings", "User", txtName.Text
      Unload Me
    Else
      User.LogResults = False
    End If
  Case enLogout
    User.LogResults = True
    SaveSetting App.Title, "Settings", "User", txtName.Text
    Unload Me
  End Select
  StrUser = txtName.Text
  '
  Exit Sub
'ErrCall:
'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FLogon.cmdContinue_Click.", vbCritical, "Error"
End Sub

Private Sub cmdLog_Click()
'  On Error GoTo ErrCall
  '
  Select Case iMode
  Case enLogin
    If Submit(1, True) Then
      User.LogResults = True
      SaveSetting App.Title, "Settings", "User", txtName.Text
      Unload Me
    Else
      User.LogResults = False
    End If
  Case enLogout
    If Submit(0, True) Then
      User.LogResults = True
      SaveSetting App.Title, "Settings", "User", txtName.Text
      Unload Me
    Else
      User.LogResults = False
    End If
  End Select
  StrUser = txtName.Text
  '
  Exit Sub
'ErrCall:
'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FLogon.cmdLog_Click.", vbCritical, "Error"
End Sub

Private Sub Form_Activate()
  On Error GoTo ErrCall
  '
  Dim sTimeofDay As String
  '
  Select Case CInt(Format(Now, "hh"))
  Case Is < 12
    If iMode = enLogout Then
      sTimeofDay = "day"
    Else
      sTimeofDay = "morning"
    End If
  Case Is < 17
    sTimeofDay = "afternoon"
  Case Else
    sTimeofDay = "evening"
  End Select
  '
  Select Case iMode
  Case enLogin
    lblWelcomeMessage = "Good " & sTimeofDay & ". Welcome to " & App.Title & ". Please provide your name and password below:"
    cmdContinue.Caption = "Continue"
    cmdLog.Caption = "Continue and Log In"
  Case enLogout
    lblWelcomeMessage = "Thanks for using " & App.Title & ". Please supply your name and password if you will be logging out. Have a good " & sTimeofDay & "."
    cmdContinue.Caption = "Exit"
    cmdLog.Caption = "Exit and Log Out"
  End Select
  '
  ttmTime = Now
  tdtDate = Now
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FLogon.Form_Activate", vbCritical, "Error"
End Sub

Private Sub Form_Load()
  On Error GoTo ErrCall
  '
  ' Get default username
  '
  txtName = GetSetting(App.Title, "Settings", "User", "")
  txtName.SelStart = 0
  txtName.SelLength = Len(txtName)
  Me.cboServer.Text = GetSetting(App.Title, "Database", "Server2", "HRI-SVR-03")
  'Me.cboServer.Text = GetSetting(App.Title, "Database", "Server", "HAWKINS-MAIN")
  Me.cboDatabase.Text = GetSetting(App.Title, "Database", "SQLDB", "BNB_DATA")
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FLogon.Form_Load", vbCritical, "Error"
End Sub

Private Sub tmrClock_Timer()
  On Error GoTo ErrCall
  '
  Static Minutes As Integer
  '
  If Minutes < 3 Then
    Minutes = Minutes + 1
  Else
    ttmTime = Now
    tdtDate = Now
    Minutes = 0
  End If
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FLogon.tmrClock_Timer", vbCritical, "Error"
End Sub

Public Property Get Mode() As enMode
  Mode = iMode
End Property

Public Property Let Mode(ByVal piMode As enMode)
  iMode = piMode
End Property

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 39 Then KeyAscii = 180
End Sub

Private Function SQLLogon() As Boolean
  On Error GoTo ErrCall
  '
  If cnMain.State = adStateClosed Then
    cnMain.CursorLocation = adUseClient
    cnMain.Provider = "SQLOLEDB"
    cnMain.Properties("Data Source").Value = cboServer
    cnMain.Properties("Initial Catalog").Value = cboDatabase
    cnMain.Properties("User ID").Value = Replace(txtName, " ", "")
    cnMain.Properties("Password").Value = Rot39(txtPassword)
    cnMain.Properties("Persist Security Info") = False
    cnMain.CursorLocation = adUseClient
    cnMain.Open
  End If
  SQLLogon = True
  SaveSetting App.Title, "Database", "Server2", Me.cboServer.Text
  SaveSetting App.Title, "Database", "SQLDB", Me.cboDatabase.Text
  Exit Function
  '
ErrCall:
  SQLLogon = False
End Function
