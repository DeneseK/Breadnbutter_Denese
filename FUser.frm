VERSION 5.00
Begin VB.Form FUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Select"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "FUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2430
      TabIndex        =   3
      Top             =   1380
      Width           =   2085
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1380
      Width           =   1935
   End
   Begin VB.ComboBox cmbUser 
      Height          =   315
      Left            =   810
      TabIndex        =   0
      Top             =   510
      Width           =   3165
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select User"
      Height          =   1095
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   0
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "FUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
Dim rsUser As New ADODB.Recordset
  '
   If cmbUser.Text <> "" Then
    StrUser = cmbUser
    rsUser.Open "select * from tblEmployees", cnMain, adOpenKeyset, adLockBatchOptimistic
    '
    With rsUser
      Do While Not .eof
        If LCase(StrUser) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
          iGroupNumber = !Groups
        End If
        .MoveNext
      Loop
      .Close
    End With
    '
    FVMail.Show
    Unload Me
  Else
    MsgBox ("Please enter user Name")
  End If
  '
  

 
End Sub

Private Sub Form_Load()
Dim rsUser As New ADODB.Recordset
Dim x As Long
'
rsUser.Open "select * from tblEmployees", cnMain, adOpenKeyset, adLockBatchOptimistic
With rsUser
  Do While Not .eof
    cmbUser.AddItem !EmployeeFirst & " " & !EmployeeLast
    .MoveNext
  Loop
  .Close
End With
End Sub
