VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FCompany2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Company"
   ClientHeight    =   2790
   ClientLeft      =   6045
   ClientTop       =   5130
   ClientWidth     =   7005
   Icon            =   "FCompany2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7005
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   5280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   4560
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   4800
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   600
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2040
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCompany2.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCompany2.frx":0377
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox StarPicture 
      Height          =   255
      Index           =   0
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6015
      TabIndex        =   5
      Top             =   2355
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   2355
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      MaxLength       =   100
      TabIndex        =   0
      Top             =   150
      Width           =   3375
   End
   Begin VB.TextBox txtOffice 
      DataField       =   "DisplayName"
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CheckBox chkDoNotContact 
      Caption         =   "Do not contact individuals"
      Height          =   255
      Left            =   135
      TabIndex        =   2
      Top             =   555
      Width           =   2175
   End
   Begin VB.TextBox txtNote 
      Height          =   1095
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1155
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Interest Rank:"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblName2 
      Caption         =   "Division:"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   180
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblName1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   465
   End
   Begin VB.Label lblCoNotes 
      AutoSize        =   -1  'True
      Caption         =   "Company Notes:"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   915
      Width           =   1170
   End
End
Attribute VB_Name = "FCompany2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lResultID As Long
Dim bNew As Boolean
Dim Company As CCompany
Dim CompanyData As CCompanyData

Private Sub cmdCancel_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Unload Me
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FCompany2", "cmdCancel_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdSave_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   CompanyData.Name = Me.txtName
110   CompanyData.Note = Me.txtNote
115   CompanyData.InterestRank = GetRank
      '
120   CompanyData.DoNotContact = IIf((Me.chkDoNotContact = 1), True, False)
      '
130   If Company.Save(CompanyData, bNew) Then
140     lResultID = CompanyData.ID
150   Else
160     lResultID = 0
170   End If
      '
180   Unload Me
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FCompany2", "cmdSave_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Function NewCompany() As Long
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Set CompanyData = New CCompanyData
110   Set Company = New CCompany
      '
120   bNew = True
      '
130   Me.Caption = "Add New Company"
135   SetStar 0
140   Me.Show vbModal
      '
150   NewCompany = lResultID
      '
160   Set CompanyData = Nothing
170   Set Company = Nothing
      '<EhFooter>
      '
      Exit Function
      '
EH:
      ErrorMgr.Raise "FCompany2", "NewCompany", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Public Function LoadCompany(plCompanyID As Long) As Long
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Set CompanyData = New CCompanyData
110   Set Company = New CCompany
      '
120   bNew = False
      '
130   Company.Load CompanyData, plCompanyID
      '
140   Me.txtName.Text = CompanyData.Name
150   Me.txtNote.Text = CompanyData.Note
160   Me.chkDoNotContact.Value = IIf(CompanyData.DoNotContact, 1, 0)
165   SetStar CompanyData.InterestRank
      '
170   Me.Show vbModal
      '
180   LoadCompany = lResultID
      '
190   Set CompanyData = Nothing
200   Set Company = Nothing
      '<EhFooter>
      '
      Exit Function
      '
EH:
      ErrorMgr.Raise "FCompany2", "LoadCompany", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Private Sub SetStar(Index As Integer)
  Dim x As Integer
    For x = 0 To 5
      StarPicture(x).Picture = Me.ImageList2.ListImages(1).Picture
    Next x
    '
    For x = 1 To Index
      StarPicture(x).Picture = Me.ImageList2.ListImages(2).Picture
    Next x
End Sub

Private Function GetRank() As Integer
  Dim x As Integer
  GetRank = 0
    For x = 0 To 5
      If StarPicture(x).Picture = Me.ImageList2.ListImages(2).Picture Then GetRank = x
    Next x
    '
End Function

Private Sub StarPicture_Click(Index As Integer)
  SetStar Index
End Sub
