VERSION 5.00
Begin VB.Form FSetPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Password"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1980
      TabIndex        =   5
      Top             =   1170
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   420
      TabIndex        =   4
      Top             =   1170
      Width           =   1395
   End
   Begin VB.TextBox txtVerifyPwd 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1875
   End
   Begin VB.TextBox txtNewPwd 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   210
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Verify Password:"
      Height          =   285
      Index           =   2
      Left            =   300
      TabIndex        =   3
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "New Password:"
      Height          =   285
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FSetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iMinPwdLen As Integer
Private iMaxPwdLen As Integer

Public NewPwd As String
Public Cancelled As Boolean
Public PwdOK As Boolean

Private Sub cmdOK_Click()
  On Error GoTo ErrCall
  '
  Dim bPwdOK As Boolean
  bPwdOK = True
  '
  bPwdOK = (txtNewPwd.Text = txtVerifyPwd.Text)
  '
  If iMaxPwdLen > 0 Then If bPwdOK Then bPwdOK = Len(txtNewPwd.Text) <= iMaxPwdLen
  '
  If iMinPwdLen > 0 Then If bPwdOK Then bPwdOK = Len(txtNewPwd.Text) >= iMinPwdLen
  '
  If bPwdOK Then
    NewPwd = txtNewPwd.Text
    PwdOK = True
    Cancelled = False
    Me.Hide
  Else
    txtNewPwd.SetFocus
    txtNewPwd.SelStart = 0
    txtNewPwd.SelLength = Len(txtNewPwd.Text)
    '
    MsgBox "Valid password not supplied.", vbInformation, "Invalid Password"
    PwdOK = False
  End If
  '
  Exit Sub
ErrCall:
  PwdOK = False
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSetPassword.cmdOK_Click", vbCritical, "Error"
End Sub

Private Sub Command1_Click()
  On Error GoTo ErrCall
  '
  Cancelled = True
  Me.Hide
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSetPassword.Command1_Click", vbCritical, "Error"
End Sub

Private Sub Form_Activate()
  On Error GoTo ErrCall
  '
  txtNewPwd.SetFocus
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSetPassword.Form_Activate", vbCritical, "Error"
End Sub

Public Sub Setup(psOldPwd As String, piMinPwdLen As Integer, piMaxPwdLen As Integer, pbHideOldPwd As Boolean)
  On Error GoTo ErrCall
  '
  iMinPwdLen = piMinPwdLen
  iMaxPwdLen = piMaxPwdLen
  '
  txtNewPwd.MaxLength = piMaxPwdLen
  txtVerifyPwd.MaxLength = piMaxPwdLen
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSetPassword.Setup", vbCritical, "Error"
End Sub

Private Sub txtNewPwd_KeyPress(KeyAscii As Integer)
  If KeyAscii = 39 Then KeyAscii = 180
End Sub

Private Sub txtOldPwd_KeyPress(KeyAscii As Integer)
  If KeyAscii = 39 Then KeyAscii = 180
End Sub

Private Sub txtVerifyPwd_KeyPress(KeyAscii As Integer)
  If KeyAscii = 39 Then KeyAscii = 180
End Sub
