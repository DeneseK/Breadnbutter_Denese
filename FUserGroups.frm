VERSION 5.00
Begin VB.Form FUserGroups 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select your groups"
   ClientHeight    =   2535
   ClientLeft      =   5775
   ClientTop       =   4365
   ClientWidth     =   2970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   2970
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Changes"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox chkOperator 
      Caption         =   "Operator"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CheckBox chkSupport 
      Caption         =   "Support"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.CheckBox chkSales 
      Caption         =   "Sales"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.CheckBox chkAuthorizations 
      Caption         =   "Authorizations"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FUserGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim iNum As Integer

Private Sub chkOperator_Click()
  '
  
End Sub

Private Sub chkAuthorizations_Click()
  '
  
End Sub

Private Sub chkSales_Click()
  '
  
End Sub

Private Sub chkSupport_Click()
  '
  
End Sub

Private Sub cmdCancel_Click()
  Form_Load
  Me.Hide
  FVMail.GetUserGroups
End Sub

Private Sub cmdSave_Click()
  '
  iGroupNumber = 0
  '
  If chkOperator.Value = 1 Then
    iGroupNumber = iGroupNumber + 1
  End If
  '
  If chkSupport.Value = 1 Then
    iGroupNumber = iGroupNumber + 2
  End If
  '
  If chkSales.Value = 1 Then
    iGroupNumber = iGroupNumber + 4
  End If
  '
  If chkAuthorizations.Value = 1 Then
    iGroupNumber = iGroupNumber + 8
  End If
  '
  If chkAuthorizations.Value = 0 And chkSales.Value = 0 And chkSupport.Value = 0 And chkOperator.Value = 0 Then
    MsgBox "You must check at least 1 Group.", vbInformation
  Else
    SaveData
    Form_Load
    FUserGroups.Hide
    FVMail.RefreshMessages
    FVMail.GetUserGroups
    FVMail.listview1_Click
  End If
End Sub

Private Sub Form_Load()
  Dim iTemp As Integer
  '
  FUserGroups.Caption = "User Groups For: " & StrUser
  '
  iTemp = iGroupNumber
  If iTemp >= 8 Then
    chkAuthorizations.Value = 1
    iTemp = iTemp - 8
  Else
    chkAuthorizations.Value = 0
  End If
  '
  If iTemp >= 4 Then
    chkSales.Value = 1
    iTemp = iTemp - 4
  Else
    chkSales.Value = 0
  End If
  '
  If iTemp >= 2 Then
    chkSupport.Value = 1
    iTemp = iTemp - 2
  Else
    chkSupport.Value = 0
  End If
  '
  If iTemp >= 1 Then
    chkOperator.Value = 1
  Else
    chkOperator.Value = 0
  End If
  '
End Sub

Private Sub SaveData()
Dim rsUser As New ADODB.Recordset
  '
  rsUser.Open "select * from tblEmployees", cnMain, adOpenKeyset, adLockBatchOptimistic
  With rsUser
    Do While Not .eof
      If LCase(StrUser) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
        !Groups = iGroupNumber
      End If
      .UpdateBatch
      .MoveNext
    Loop
    .Close
  End With
  'StrGroups = iGroupNumber
End Sub
