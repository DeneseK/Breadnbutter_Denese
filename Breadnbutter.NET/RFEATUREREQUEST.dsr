VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RFeatureRequest 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12615
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22251
   _ExtentY        =   13811
   SectionData     =   "RFeatureRequest.dsx":0000
End
Attribute VB_Name = "RFeatureRequest"
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
  adc.Connection = pCN
  If ConnType = Access Then
    adc.Source = "SELECT TCompany.Name AS Company, " & _
      "[FirstName] + ' ' + [LastName] AS Name, " & _
      "tblSupportActs.Type, tblSupportActs.Date, tblSupportActs.[User], " & _
      "tblSupportActs.Subject, tblSupportActs.Results " & _
      "FROM (TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID) LEFT JOIN tblSupportActs ON TContact.ID = tblSupportActs.CustRecID " & _
      "WHERE (((tblSupportActs.Type) = 'Feature Request')) " & _
      "AND (Date Between #" & pdStartDate & "# AND #" & pdEndDate & "#) " & _
      "ORDER BY tblSupportActs.Date"
  Else
    adc.Source = "SELECT TCompany.Name AS Company, " & _
      "[FirstName] + ' ' + [LastName] AS Name, " & _
      "TSupportActs.Type, TSupportActs.Date, TSupportActs.[User], " & _
      "TSupportActs.Subject, TSupportActs.Results " & _
      "FROM (TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID) LEFT JOIN TSupportActs ON TContact.ID = TSupportActs.CustRecID " & _
      "WHERE (((TSupportActs.Type) = 'Feature Request')) " & _
      "AND (Date Between '" & pdStartDate & "' AND '" & pdEndDate & "') " & _
      "ORDER BY TSupportActs.Date"
  End If
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description & " in Weekly Sales Report Setup"
End Sub

Private Sub GroupHeader1_Format()
  lblPeriod.Caption = "From " & dStartDate & " to " & dEndDate
End Sub
