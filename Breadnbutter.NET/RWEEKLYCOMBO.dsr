VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RWeeklyCombo 
   Caption         =   "Weekly Combo Report"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   18283
   _ExtentY        =   12356
   SectionData     =   "RWeeklyCombo.dsx":0000
End
Attribute VB_Name = "RWeeklyCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CN As ADODB.Connection
Private dStartDate As Date
Private dEndDate As Date

Public Sub Setup(pCN As ADODB.Connection, pdStartDate As Date, pdEndDate As Date)
  On Error GoTo ErrCall
  '
  Set CN = pCN
  dStartDate = pdStartDate
  dEndDate = pdEndDate
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description & " in Weekly Shipments, Evals, Sales Report Setup"
End Sub

Private Sub Detail_Format()
  On Error GoTo ErrCall
  '
  Dim Shipments As New RWeeklyFirstDiskShipments
  Shipments.Setup CN, dStartDate, dEndDate
  Set srptShipments.object = Shipments
  Set Shipments = Nothing
  '
  Dim Evals As New RWeeklyEvalAuthorizations
  Evals.Setup CN, dStartDate, dEndDate
  Set srptEvals.object = Evals
  Set Evals = Nothing
  '
  Dim Sales As New RWeeklySales
  Sales.Setup CN, dStartDate, dEndDate
  Set srptSales.object = Sales
  Set Sales = Nothing
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description & " in Weekly Combo Report Detail Format."
End Sub
