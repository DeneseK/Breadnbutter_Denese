VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RWeeklyEvalAuthorizations 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12615
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22251
   _ExtentY        =   13811
   SectionData     =   "RWeeklyEvalAuthorizations.dsx":0000
End
Attribute VB_Name = "RWeeklyEvalAuthorizations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dStartDate As Date
Private dEndDate As Date

Public Sub Setup(pCN As ADODB.Connection, pdStartDate As Date, pdEndDate As Date)
  On Error GoTo ErrCall
  '
  dStartDate = pdStartDate
  dEndDate = pdEndDate
  '
  Set adc.Connection = pCN
  '
  If ConnType = Access Then
    adc.Source = "SELECT * FROM QWeeklyEvalAuthorizations WHERE AuthDate Between #" & pdStartDate & "# AND #" & pdEndDate & "# ORDER BY AuthDate"
  Else
    adc.Source = "SELECT TCompany.Name AS Company, [FirstName] + ' ' + [LastName] AS Name, TContact.AuthDate, TContact.AuthDays, TContact.ShipStatus, TContact.ShipDate, TContact.Status, TContact.AuthStatus, TSupportActs.Type, TSupportActs.[User] " & _
      "FROM (TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID) LEFT JOIN TSupportActs ON TContact.ID = TSupportActs.CustRecID " & _
      "Where ((TContact.Status = 'Prospect') And (TContact.AuthStatus = 'Evaluation') And (TSupportActs.Type = 'Eval Authorized')) " & _
      "AND (AuthDate Between '" & pdStartDate & "' AND '" & pdEndDate & "') " & _
      "ORDER BY TContact.AuthDate"
  End If
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description & " in Weekly Eval Authorizations Report Setup"
End Sub

Private Sub GroupHeader1_Format()
  lblPeriod.Caption = "From " & dStartDate & " to " & dEndDate
End Sub
