VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FChangeCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Company"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "FChangeCompany.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwCompany 
      Height          =   5445
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Company"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "FChangeCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ContactData As CContactData
Private CompanyData As CCompanyData
Private Company As CCompany
Private Contact As CContact
Private Companys As New CCompanys

Private lContactID As Long
Private bChanged As Boolean

Private Sub cmdCancel_Click()
  bChanged = False
  '
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If Not lvwCompany.SelectedItem Is Nothing Then
    ContactData.CompanyID = GetIDFromKey(lvwCompany.SelectedItem.Key)
    '
    If Contact.Save(ContactData, False) Then
      bChanged = True
      '
      Unload Me
    Else
      MsgBox "Contact could not be saved"
    End If
  Else
    MsgBox "Invalid Selection"
  End If
End Sub

Private Sub Form_Load()
  Set ContactData = New CContactData
  Set CompanyData = New CCompanyData
  Set Contact = New CContact
  Set Company = New CCompany
  Set Companys = New CCompanys
  '
  Dim lPos As Long
  '
  Contact.Load ContactData, lContactID
  Company.Load CompanyData, ContactData.CompanyID
  '
  Company.LoadCollection Companys
  '
  lvwCompany.Visible = False
  '
  lvwCompany.ListItems.Clear
  '
  For lPos = 1 To Companys.Count
    lvwCompany.ListItems.Add , "A" & Companys.Item(lPos).ID, Companys.Item(lPos).Name
  Next
  '
  Set lvwCompany.SelectedItem = lvwCompany.ListItems("A" & CompanyData.ID)
  '
  lvwCompany.SelectedItem.EnsureVisible
  '
  lvwCompany.Visible = True
  '
  
End Sub

Public Function ChangeCompany(plContactID As Long) As Boolean
  lContactID = plContactID
  '
  Me.Show vbModal
  '
  ChangeCompany = bChanged
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set Companys = Nothing
  Set Contact = Nothing
  Set Company = Nothing
  Set CompanyData = Nothing
  Set ContactData = Nothing
End Sub
