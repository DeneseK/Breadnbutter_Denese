VERSION 5.00
Begin VB.Form FSendTo 
   Caption         =   "Forward Message To:"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   4890
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Send To:"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cboSendTo 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblEmailAddress 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   0
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "FSendTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private boolAttachment As Boolean
Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdSend_Click()
  SendMail
  cboSendTo.Text = vbNullString
  lblEmailAddress.Caption = vbNullString
  Me.Hide
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
  FSendTo.Caption = StrUser & " Forward Message To:"
  rs.Open "SELECT * FROM tblEmployees ORDER BY EmployeeLast", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  With rs
    While Not .eof
      cboSendTo.AddItem (!EmployeeFirst & " " & !EmployeeLast)
        If LCase(!EmployeeFirst & " " & !EmployeeLast) = LCase(StrUser) Then
'          sFromAddress = !EMailAddress & "@powerclaim.com"
        End If
      .MoveNext
    Wend
  End With
  '
  Set rs = Nothing
  '
End Sub

Private Sub cboSendTo_Click()
Dim rs As New ADODB.Recordset
  '
  rs.Open "SELECT * FROM tblEmployees", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  With rs
    .MoveFirst
    While Not .eof
      If (!EmployeeFirst & " " & !EmployeeLast) = cboSendTo.Text Then
        sEmailAddress = !EMailAddress '& "@powerclaim.com"
        lblEmailAddress.Caption = sEmailAddress
        Exit Sub
      Else
        sEmailAddress = cboSendTo.Text
      End If
      .MoveNext
    Wend
  End With
  lblEmailAddress.Caption = sEmailAddress
  '
End Sub

Private Sub cboSendTo_LostFocus()
Dim rs As New ADODB.Recordset
  '
  rs.Open "SELECT * FROM tblEmployees", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  With rs
    .MoveFirst
    While Not .eof
      If (!EmployeeFirst & " " & !EmployeeLast) = cboSendTo.Text Then
        sEmailAddress = !EMailAddress '& "@powerclaim.com"
        lblEmailAddress.Caption = sEmailAddress
        Exit Sub
      Else
        sEmailAddress = cboSendTo.Text
      End If
      .MoveNext
    Wend
  End With
  lblEmailAddress.Caption = sEmailAddress
  '
End Sub
Private Sub getAttachment()
Dim FileSys As New FileSystemObject
Dim myStr As String
Dim rsAttach As New ADODB.Recordset
Dim sMessageName As String

'
 If Not FileSys.FolderExists(App.Path & "\Attachment") Then
    FileSys.CreateFolder (App.Path & "\Attachment")
 End If
'rsAttach.Open "SELECT * FROM TVMailMessages WHERE MessageName like '" & sMessageName & "'", cnMain, adOpenKeyset, adLockBatchOptimistic
rsAttach.Open "SELECT * FROM TVMailMessages WHERE MessageID like '" & CStr(sMessageID) & "'", cnMain, adOpenKeyset, adLockBatchOptimistic
'

'
    With rsAttach
 sMessageName = "" & !MessageName
        If sMessageName <> "" And UCase(Right(sMessageName, 4)) = ".WAV" Then

       ' If !Attachment <> Null Then
            Set strStream = New ADODB.Stream
            strStream.Type = adTypeBinary
            strStream.Open
            strStream.Write !Attachment
            strStream.SaveToFile App.Path & "\Attachment\" & !MessageName, adSaveCreateOverWrite
            strStream.Close
            Set strStream = Nothing
                 boolAttachment = True
        Else
       '     Debug.Print !MessageName
            boolAttachment = False
        End If
            
     End With


'
'getAttachment = boolAttachment
End Sub

Private Sub SendMail()
Dim SMTP As Object
Dim FileSys As New FileSystemObject
Dim X%

  Set SMTP = CreateObject("EasyMail.SMTP.6")
  SMTP.LicenseKey = "Hawkins Research (Single Developer)/00B0630C10151C00BC30"
  SMTP.MailServer = "HRI-svr-02"
  SMTP.FromAddr = sFromAddress
  SMTP.AddRecipient "", sEmailAddress, 1
  SMTP.Subject = sSubject
  SMTP.BodyText = sBody
  getAttachment
  If boolAttachment Then
        X = SMTP.AddAttachment(App.Path & "\Attachment\" & sMessageName, 0)
  End If
        X% = SMTP.Send
        If X = 0 Then
          MsgBox "Message sent successfully."
        Else
          MsgBox "There was an error sending your message.  Error: " & CStr(X%)
        End If
  
  If FileSys.FolderExists(App.Path & "\Attachment") Then
    FileSys.DeleteFolder (App.Path & "\Attachment")
  End If
  Set SMTP = Nothing
End Sub


Private Function sBody()
Dim TSTemp As TextStream
Dim fso As New FileSystemObject
  '
  Set TSTemp = fso.OpenTextFile(App.Path & "\tempMessage.txt", ForReading, True, TristateUseDefault)
  sBody = TSTemp.ReadAll
  '
  
End Function

Private Function sFromAddress()
Dim rsUser As New ADODB.Recordset
  '
  rsUser.Open "select * from tblEmployees", cnMain, adOpenKeyset, adLockBatchOptimistic
    
  With rsUser
    .MoveFirst
    While Not .eof
      If LCase(StrUser) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
        sFromAddress = !EMailAddress '& "@powerclaim.com"
      End If
      .MoveNext
    Wend
  End With
  '
    
End Function
