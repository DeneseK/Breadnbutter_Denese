VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FDailyReports 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Daily Report"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   Icon            =   "FDailyReports.frx":0000
   LinkTopic       =   "FDailyReports"
   ScaleHeight     =   5955
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton PrintButton 
      Caption         =   "Print Report"
      Height          =   315
      Left            =   4110
      TabIndex        =   3
      Top             =   2790
      Width           =   2025
   End
   Begin VB.CommandButton RefreshData 
      Caption         =   "Refresh Data"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2790
      Width           =   2055
   End
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   315
      Left            =   6120
      TabIndex        =   4
      Top             =   2790
      Width           =   2025
   End
   Begin VB.CommandButton Preview 
      Caption         =   "Preview Report"
      Height          =   315
      Left            =   2070
      TabIndex        =   2
      Top             =   2790
      Width           =   2025
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FDailyReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Report  As New CReport
'
Private Sub ExitButton_Click()
  On Error GoTo EH
  '
  Unload Me
  '
  Exit Sub
EH:
 MsgBox Err.Description & " in FDailyReports.Form_Load."
End Sub
'
Private Sub Form_Load()
  On Error GoTo EH
  '
  'Set Report = New CReport
  Report.Rtype = daily
  RefreshList
  '
  Exit Sub
EH:
 MsgBox Err.Description & " in FDailyReports.Form_Load."
End Sub
'
Private Sub RefreshList()
  On Error GoTo EH
  '
  ListView1.ListItems.Clear
  ListView1.ColumnHeaders.Clear
  Report.FillList Report.rsReport, ListView1
  '
  Exit Sub
EH:
 MsgBox Err.Description & " in FDailyReports.RefreshList."
End Sub
'
Private Sub SetupList(list As ListView, rs As Recordset)
  On Error GoTo EH
  '
  list.ListItems.Clear
  list.ColumnHeaders.Clear
  '
  Exit Sub
EH:
 MsgBox Err.Description & " in FDailyReports.SetupList."
End Sub
'
Private Sub Form_Resize()
  On Error GoTo EH
  '
  ListView1.Width = Me.Width - 100
  '
  If Me.Height > 1000 Then
  '
    ListView1.Height = Me.Height - 750
    Me.RefreshData.Move 0, Me.Height - 700
    Me.Preview.Move 2030, Me.Height - 700
    Me.ExitButton.Move 6090, Me.Height - 700
    Me.PrintButton.Move 4060, Me.Height - 700
    '
  End If
  Exit Sub
EH:
 MsgBox Err.Description & " in FDailyReports.Form_Resize."
End Sub
'
Private Sub Preview_Click()
  On Error GoTo EH
  '
  Report.PreviewReport "Daily"
  Exit Sub
EH:
 MsgBox Err.Description & " in FDailyReports.Preview_Click."
End Sub
'
Private Sub PrintButton_Click()
  On Error GoTo EH
  '
  Report.PrintReport "Daily"
  Exit Sub
EH:
 MsgBox Err.Description & " in FDailyReports.PrintButton_Click."
End Sub
'
Private Sub RefreshData_Click()
  On Error GoTo EH
  '
  Report.Rtype = daily
  RefreshList
  Exit Sub
EH:
 MsgBox Err.Description & " in FDailyReports.RefreshData_Click."
End Sub
