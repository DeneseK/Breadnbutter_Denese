VERSION 5.00
Begin VB.Form FBranch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Branches"
   ClientHeight    =   4785
   ClientLeft      =   3510
   ClientTop       =   3465
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3600
      TabIndex        =   14
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   12
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtFaxNumber 
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtPhoneNumber 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   10
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtAddress3 
      Height          =   285
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2760
      Width           =   5415
   End
   Begin VB.TextBox txtAddress2 
      Height          =   285
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   5
      Top             =   2400
      Width           =   5415
   End
   Begin VB.TextBox txtAddress1 
      Height          =   285
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   4
      Top             =   2040
      Width           =   5415
   End
   Begin VB.TextBox txtManagerLastName 
      Height          =   285
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtManagerFirstName 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail"
      Height          =   195
      Left            =   4560
      TabIndex        =   27
      Top             =   3600
      Width           =   435
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Fax Number"
      Height          =   195
      Left            =   2400
      TabIndex        =   26
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Phone Number"
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   3600
      Width           =   1065
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "City"
      Height          =   195
      Left            =   720
      TabIndex        =   24
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "State"
      Height          =   195
      Left            =   3240
      TabIndex        =   23
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Zip"
      Height          =   195
      Left            =   4800
      TabIndex        =   22
      Top             =   3120
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Address 3"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Address 2"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Address 1"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Manager's Last Name"
      Height          =   195
      Left            =   3480
      TabIndex        =   18
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Manager's First Name"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Number"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "FBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lResultID As Long
Private bAddNew As Boolean
Private BranchData As New CBranchData
Private Branch As New CBranch

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  '
  BranchData.Name = txtName.Text
  BranchData.Number = txtNumber.Text
  BranchData.ManagerFirstName = txtManagerFirstName.Text
  BranchData.ManagerLastName = txtManagerLastName.Text
  BranchData.Address1 = txtAddress1.Text
  BranchData.Address2 = txtAddress2.Text
  BranchData.Address3 = txtAddress3.Text
  BranchData.City = txtCity.Text
  BranchData.State = txtState.Text
  BranchData.Zip = txtZip.Text
  BranchData.PhoneNumber = txtPhoneNumber.Text
  BranchData.FaxNumber = txtFaxNumber.Text
  BranchData.Email = txtEmail.Text
  '
  If Branch.Save(BranchData, bAddNew) Then
    lResultID = BranchData.BranchID
  Else
    lResultID = 0
  End If
  '
  Unload Me
End Sub

Public Function NewBranch(pCompanyID As Integer) As Long
  Set Branch = New CBranch
  Set BranchData = New CBranchData
  '
  bAddNew = True
  '
  BranchData.CompanyID = pCompanyID
  '
  Me.Show vbModal
  '
  NewBranch = lResultID
  '
  Set BranchData = Nothing
  Set Branch = Nothing
End Function

Public Function EditBranch(pBranchID As Long) As Long
  Set Branch = New CBranch
  Set BranchData = New CBranchData
  '
  bAddNew = False
  '
  Branch.Load BranchData, pBranchID
  '
  'lCompanyID = BranchData.CompanyID
  txtName.Text = BranchData.Name
  txtNumber.Text = BranchData.Number
  txtManagerFirstName.Text = BranchData.ManagerFirstName
  txtManagerLastName.Text = BranchData.ManagerLastName
  txtAddress1.Text = BranchData.Address1
  txtAddress2.Text = BranchData.Address2
  txtAddress3.Text = BranchData.Address3
  txtCity.Text = BranchData.City
  txtState.Text = BranchData.State
  txtZip.Text = BranchData.Zip
  txtPhoneNumber.Text = BranchData.PhoneNumber
  txtFaxNumber.Text = BranchData.FaxNumber
  txtEmail.Text = BranchData.Email
  '
  Me.Show vbModal
  '
  EditBranch = lResultID
  '
  Set BranchData = Nothing
  Set Branch = Nothing
End Function


