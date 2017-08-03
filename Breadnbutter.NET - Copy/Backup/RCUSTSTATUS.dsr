VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RCustStatus 
   Caption         =   "ActiveReport1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   17648
   _ExtentY        =   10398
   SectionData     =   "RCustStatus.dsx":0000
End
Attribute VB_Name = "RCustStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub DBName(psDBName As String)
  adc.ConnectionString = cnMain.ConnectionString
End Sub

Private Sub Detail_Format()
  On Error GoTo ErrCall
  '
  With adc.Recordset
  If Not (.BOF Or .EOF) Then
    txtName.Text = !FirstName & " " & !LastName
    '
    If Not IsNull(!AuthDate) Then
      txtRemaining.Text = DateDiff("d", Now, DateAdd("d", CDbl(nnNum(!AuthDays)), !AuthDate))
    Else
      txtRemaining.Text = "*"
    End If
  End If
  End With
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description & vbCrLf & "in rptCustStatus.Detail Format"
End Sub

Private Sub PageFooter_Format()
  txtPage = Me.pageNumber
End Sub
