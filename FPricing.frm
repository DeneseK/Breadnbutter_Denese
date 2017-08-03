VERSION 5.00
Begin VB.Form FPricing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PowerClaim Pricing"
   ClientHeight    =   5052
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7944
   Icon            =   "FPricing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   7944
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   4200
      Picture         =   "FPricing.frx":030A
      ScaleHeight     =   4764
      ScaleWidth      =   3564
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblYearlyCost 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$599/yr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   2310
      TabIndex        =   10
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label lblCopies 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "1000+ Copies ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   1965
   End
   Begin VB.Label lblCopies 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "1-24 Copies ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1485
      Width           =   1965
   End
   Begin VB.Label lblCopies 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "25-99 Copies ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   255
      TabIndex        =   7
      Top             =   1800
      Width           =   1965
   End
   Begin VB.Label lblCopies 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "100+ Copies ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   255
      TabIndex        =   6
      Top             =   2160
      Width           =   1965
   End
   Begin VB.Label lblYearlyCost 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$999/yr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2325
      TabIndex        =   5
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label lblYearlyCost 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$899/yr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2325
      TabIndex        =   4
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label lblYearlyCost 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$799/yr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   2325
      TabIndex        =   3
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Option: $120.00 per month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Quarterly Option: $300.00 per quarter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   6135
   End
End
Attribute VB_Name = "FPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function FixPhoneNumber(psNumber As String) As String
  Dim sTemp As String
  Dim l As Long
  Dim sChar As String
  '
  For l = 1 To Len(psNumber)
    sChar = Mid(psNumber, l, 1)
    If (Asc(sChar) >= 48) And (Asc(sChar) <= 57) Then
      sTemp = sTemp & sChar
    End If
  Next
  '
  FixPhoneNumber = sTemp
End Function

'Private Sub Command1_Click()
'  Dim ContactData As New CContactData
'  Dim contact As New CContact
'  Dim rs As New Recordset
'  '
'  rs.Open "SELECT * FROM Tcin", cnMain, adOpenForwardOnly, adLockReadOnly
'  '
'  While Not rs.eof
'    Set ContactData = New CContactData
'    '
'    With ContactData
'      .CompanyID = 5728
'      .FirstName = Trim(rs!FirstName & vbNullString)
'      .LastName = Trim(rs!LastName & vbNullString)
'      .Phone1 = FixPhoneNumber(rs!OfficePhone & vbNullString)
'      .Email = rs!Emailaddress & vbNullString
'      .MailState = rs!State & vbNullString
'      .MailZip = rs!Zip & vbNullString
'      .MailCity = rs!City & vbNullString
'      .MailAddress1 = rs!POBox & vbNullString
'      .PreferredAddress = 1
'      .AdjusterID = rs!AdjNo & vbNullString
'      .Status = "Customer"
'      .ContactType = 1
'    End With
'    '
'    contact.Save ContactData, True
'    '
'    rs.MoveNext
'  Wend
'End Sub
'
'Private Sub Command2_Click()
'Dim ContactData As New CContactData
'  Dim contact As New CContact
'  Dim rs As New Recordset
'  '
'  rs.Open "SELECT * FROM TContact WHERE CompanyID = 5728", cnMain, adOpenKeyset, adLockOptimistic
'  '
'
'  While Not rs.eof
'    contact.Load ContactData, rs!ID
'    ContactData.FirstName = Trim(ContactData.FirstName)
'    ContactData.LastName = Trim(ContactData.LastName)
'    contact.Save ContactData, False
'    rs.MoveNext
'  Wend
''    Set ContactData = New CContactData
''    '
''    With ContactData
''      .CompanyID = 5728
''      .FirstName = rs!FirstName & vbNullString
''      .LastName = rs!LastName & vbNullString
''      .Phone1 = FixPhoneNumber(rs!OfficePhone & vbNullString)
''      .Email = rs!Emailaddress & vbNullString
''      .MailZip = rs!Zip & vbNullString
''      .MailState = rs!State & vbNullString
''      .MailZip = rs!City & vbNullString
''      .MailAddress1 = rs!POBox & vbNullString
''      .PreferredAddress = 1
''      .AdjusterID = rs!AdjNo & vbNullString
''      .Status = "Customer"
''      .ContactType = 1
''    End With
''    '
''    contact.Save ContactData, True
''    '
''    rs.MoveNext
''  Wend
'End Sub
Private Sub Form_Load()

End Sub

