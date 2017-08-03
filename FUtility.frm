VERSION 5.00
Begin VB.Form FUtility 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScanForLicenses 
      Caption         =   "Scan for Licenses"
      Height          =   315
      Left            =   690
      TabIndex        =   0
      Top             =   660
      Width           =   1755
   End
End
Attribute VB_Name = "FUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdScanForLicenses_Click()
  Dim rsContact As ADODB.Recordset
  Dim rsLicense As ADODB.Recordset
  '
  Set rsContact = New ADODB.Recordset
  Set rsLicense = New ADODB.Recordset
  '
  rsContact.Open "SELECT * FROM TContact WHERE Status = 'Customer'", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  rsLicense.Open "SELECT * FROM TLicense", cnMain, adOpenDynamic, adLockOptimistic, adCmdText
  '
  Do Until rsContact.EOF
    rsLicense.AddNew
    '
    rsLicense!ContactID = rsContact!ID
    rsLicense!LicenseDate = IIf(IsNull(rsContact!AuthDate), 0, rsContact!AuthDate)
    rsLicense!Days = rsContact!AuthDays
    rsLicense!Amount = rsContact!Rate
    '
    rsLicense.Update
    rsContact.MoveNext
  Loop
End Sub
