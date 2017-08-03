VERSION 5.00
Begin VB.Form FSelectDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Database"
   ClientHeight    =   3705
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6255
   Icon            =   "FSelectDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   300
      TabIndex        =   15
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton cmdSelectDB 
      Height          =   315
      Left            =   4290
      Picture         =   "FSelectDB.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3150
      Width           =   315
   End
   Begin VB.TextBox txtDatabase 
      Height          =   315
      Left            =   1470
      TabIndex        =   8
      Top             =   3120
      Width           =   2775
   End
   Begin VB.ComboBox cboDatabase 
      Height          =   315
      ItemData        =   "FSelectDB.frx":06D4
      Left            =   1440
      List            =   "FSelectDB.frx":06DB
      TabIndex        =   6
      Text            =   "cboDatabase"
      Top             =   960
      Width           =   3165
   End
   Begin VB.ComboBox cboServer 
      Height          =   315
      ItemData        =   "FSelectDB.frx":06E9
      Left            =   1440
      List            =   "FSelectDB.frx":06F0
      TabIndex        =   4
      Text            =   "cboServer"
      Top             =   570
      Width           =   3165
   End
   Begin VB.OptionButton optDBType 
      Caption         =   "Access"
      Height          =   315
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   2640
      Width           =   3555
   End
   Begin VB.OptionButton optDBType 
      Caption         =   "SQL Server"
      Height          =   315
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   150
      Width           =   3555
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4860
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4860
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label lblUserID 
      Caption         =   "User ID"
      Height          =   285
      Left            =   600
      TabIndex        =   11
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Database:"
      Height          =   285
      Left            =   660
      TabIndex        =   9
      Top             =   3150
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Database:"
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   990
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   600
      Width           =   645
   End
End
Attribute VB_Name = "FSelectDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Cancelled As Boolean

Private Sub CancelButton_Click()
  Cancelled = True
  Me.Hide
End Sub

Private Sub cmdSelectDB_Click()
  
  Dim sDBPath As String
  Dim sDBName As String
  '
  DBOps.GetPathFile sDBPath, sDBName, "Bread 'n' Butter Data"
  '
  Me.txtDatabase.Text = sDBPath & sDBName
  
End Sub

Private Sub Form_Load()
  
  Dim ConnectionType As ConnectionTypeEnum
  '
  ConnectionType = GetSetting(App.Title, "Database", "Type", SQL)
  '
  If ConnectionType = SQL Then
    Me.optDBType(0).Value = True
  Else
    Me.optDBType(1).Value = True
  End If
  '
  Me.cboServer.Text = GetSetting(App.Title, "Database", "Server", "HAWKINS-MAIN")
  Me.cboDatabase.Text = GetSetting(App.Title, "Database", "SQLDB", "BNB_DATA")
  Me.txtDatabase.Text = GetSetting(App.Title, "Database", "AccessDB", vbNullString)
  Me.chkLogin.Value = GetSetting(App.Title, "Database", "Login", "0")
  Me.txtPassword.Text = GetSetting(App.Title, "Database", "Password", "")
  GetLoginSettings
  '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If UnloadMode = vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
  
End Sub

Private Sub OKButton_Click()
  
 sLogin = chkLogin.Value
 sLoginName = txtUserID.Text
 sPassword = txtPassword.Text
  
  If optDBType(0).Value = True Then
    ConnType = SQL
  Else
    ConnType = Access
  End If
  '
  SaveSetting App.Title, "Database", "Type", ConnType
  SaveSetting App.Title, "Database", "Server", Me.cboServer.Text
  SaveSetting App.Title, "Database", "SQLDB", Me.cboDatabase.Text
  SaveSetting App.Title, "Database", "AccessDB", Me.txtDatabase.Text
  SaveSetting App.Title, "Database", "Login", Me.chkLogin.Value
  SaveSetting App.Title, "Database", "Password", Me.txtPassword.Text
  '
  Me.Hide
  
End Sub

Private Sub optDBType_Click(Index As Integer)
  
  If Me.optDBType(0).Value = True Then
    Me.cboServer.Enabled = True
    Me.cboDatabase.Enabled = True
    Me.txtDatabase.Enabled = False
  Else
    Me.cboServer.Enabled = False
    Me.cboDatabase.Enabled = False
    Me.txtDatabase.Enabled = True
  End If
    
    
End Sub

Private Sub GetLoginSettings()
  
  If chkLogin.Value = 1 Then
    txtUserID.Text = UCase(GetSetting(App.Title, "Settings", "User", ""))
    EditTextBox
    txtUserID.Visible = True
    txtPassword.Visible = True
    lblUserID.Visible = True
    lblPassword.Visible = True
    txtDatabase.Move 1470, 3120, 2775, 315
    Label3.Move 660, 3150, 795, 285
    optDBType(1).Move 300, 2640, 3555, 315
    cmdSelectDB.Move 4290, 3150, 315, 315
    FSelectDB.Move 2715, 3420, 6345, 4080
  Else
    txtUserID.Visible = False
    txtPassword.Visible = False
    lblUserID.Visible = False
    lblPassword.Visible = False
    txtDatabase.Move 1470, 2160, 2775, 315
    Label3.Move 660, 2190, 795, 285
    optDBType(1).Move 300, 1800, 3555, 315
    cmdSelectDB.Move 4290, 2190, 315, 315
    FSelectDB.Move 2715, 3420, 6345, 3030
  End If
  '
End Sub

Private Sub chkLogin_Click()
  If chkLogin.Value = 1 Then
    txtUserID.Text = UCase(GetSetting(App.Title, "Settings", "User", ""))
    EditTextBox
    txtUserID.Visible = True
    txtPassword.Visible = True
    lblUserID.Visible = True
    lblPassword.Visible = True
    txtDatabase.Move 1470, 3120, 2775, 315
    Label3.Move 660, 3150, 795, 285
    optDBType(1).Move 300, 2640, 3555, 315
    cmdSelectDB.Move 4290, 3150, 315, 315
    FSelectDB.Move 2715, 3420, 6345, 4080
  Else
    txtUserID.Visible = False
    txtPassword.Visible = False
    lblUserID.Visible = False
    lblPassword.Visible = False
    txtDatabase.Move 1470, 2160, 2775, 315
    Label3.Move 660, 2190, 795, 285
    optDBType(1).Move 300, 1800, 3555, 315
    cmdSelectDB.Move 4290, 2190, 315, 315
    FSelectDB.Move 2715, 3420, 6345, 3030
  End If
End Sub
Private Sub EditTextBox()
Dim i As Integer
Dim sLetter As String
Dim sName As String
  '
  For i = 1 To Len(txtUserID.Text)
    sLetter = Mid(txtUserID.Text, i, 1)
    If Not sLetter = " " Then
      sName = sName & sLetter
    End If
  Next i
  txtUserID.Text = UCase(sName)
  '
End Sub

