VERSION 5.00
Begin VB.Form FPrintLabels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Labels"
   ClientHeight    =   5865
   ClientLeft      =   7335
   ClientTop       =   4200
   ClientWidth     =   4155
   Icon            =   "FPrintLabels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   4155
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClearSelected 
      Caption         =   "Clear Selected"
      Height          =   315
      Left            =   1170
      TabIndex        =   27
      Top             =   60
      Width           =   1875
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   345
      Left            =   1500
      TabIndex        =   26
      Top             =   5430
      Width           =   1125
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   345
      Left            =   210
      TabIndex        =   22
      Top             =   5430
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   2820
      TabIndex        =   0
      Top             =   5430
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Labels to skip"
      Height          =   4905
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   3945
      Begin VB.CommandButton cmdPrintAll 
         Caption         =   "Fill All"
         Height          =   345
         Left            =   2460
         TabIndex        =   25
         Top             =   4440
         Width           =   1185
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear All"
         Height          =   345
         Left            =   270
         TabIndex        =   24
         Top             =   4440
         Width           =   1185
      End
      Begin VB.CheckBox chkPageNum 
         Caption         =   "One Page"
         Height          =   315
         Left            =   510
         TabIndex        =   23
         Top             =   270
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   0
         Left            =   450
         TabIndex        =   21
         Top             =   630
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   1
         Left            =   1980
         TabIndex        =   20
         Top             =   630
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   2
         Left            =   450
         TabIndex        =   19
         Top             =   990
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   3
         Left            =   1980
         TabIndex        =   18
         Top             =   990
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   4
         Left            =   450
         TabIndex        =   17
         Top             =   1350
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   5
         Left            =   1980
         TabIndex        =   16
         Top             =   1350
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   6
         Left            =   450
         TabIndex        =   15
         Top             =   1710
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   7
         Left            =   1980
         TabIndex        =   14
         Top             =   1710
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   8
         Left            =   450
         TabIndex        =   13
         Top             =   2070
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   9
         Left            =   1980
         TabIndex        =   12
         Top             =   2070
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   10
         Left            =   450
         TabIndex        =   11
         Top             =   2430
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   11
         Left            =   1980
         TabIndex        =   10
         Top             =   2430
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   12
         Left            =   450
         TabIndex        =   9
         Top             =   2790
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   13
         Left            =   1980
         TabIndex        =   8
         Top             =   2790
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   14
         Left            =   450
         TabIndex        =   7
         Top             =   3150
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   15
         Left            =   1980
         TabIndex        =   6
         Top             =   3150
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   16
         Left            =   450
         TabIndex        =   5
         Top             =   3510
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   17
         Left            =   1980
         TabIndex        =   4
         Top             =   3510
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   18
         Left            =   450
         TabIndex        =   3
         Top             =   3870
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   19
         Left            =   1980
         TabIndex        =   2
         Top             =   3870
         Width           =   1515
      End
   End
End
Attribute VB_Name = "FPrintLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sheet(20) As Boolean

Private Sub chkPageNum_KeyUp(KeyCode As Integer, Shift As Integer)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If chkPageNum.Value = True Then
110     EnableLabelButtons
120   Else
130     EnableLabelButtons
140   End If
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "chkPageNum_KeyUp", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub chkPageNum_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If chkPageNum.Value = 1 Then
110     EnableLabelButtons
120   Else
130     DisableLabelButtons
140   End If
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "chkPageNum_MouseUp", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub DisableLabelButtons()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim i As Integer
110   For i = 0 To 19
120     cmdLabel(i).Enabled = False
130   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "DisableLabelButtons", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub EnableLabelButtons()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim i As Integer
110   For i = 0 To 19
120     cmdLabel(i).Enabled = True
130   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "EnableLabelButtons", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub


Private Sub FlipEnableLabelButtons()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim i As Integer
110   For i = 0 To 19
120     cmdLabel(i).Enabled = Not cmdLabel(i).Enabled
130   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "FlipEnableLabelButtons", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Unload Me
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "cmdCancel_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdClearAll_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim i As Integer
110   For i = 0 To 19
120     cmdLabel(i).Caption = ""
130   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "cmdClearAll_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdClearSelected_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100  Dim rslabels As New ADODB.Recordset
110  rslabels.Open "SELECT BetaTester FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID WHERE BetaTester = 1", cnMain, adOpenDynamic, adLockBatchOptimistic
120  With rslabels
130   While Not .eof
140     !betatester = 0
150     .UpdateBatch
160     .MoveNext
170   Wend
180  End With
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "cmdClearSelected_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdLabel_Click(Index As Integer)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If cmdLabel(Index).Caption = "Print" Then
110     cmdLabel(Index).Caption = ""
120   Else
130     cmdLabel(Index).Caption = "Print"
140   End If
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "cmdLabel_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdLoadGroup_Click()
  RLabels.Clear
  CreateSheetArray
  RLabels.SetSheet (Sheet)
  RLabels.SetPages (Me.chkPageNum)
  RLabels.SetDBGroup FChooseGroup.GetGroup
  RLabels.Show vbModal, FMain
End Sub

Private Sub cmdPreview_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SetupLabels
110   RLabels.Show vbModal, FMain
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "cmdPreview_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdPrint_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SetupLabels
110   RLabels.PrintReport True
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FPrintLabels", "cmdPrint_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub SetupLabels()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
90    RLabels.Clear
100   CreateSheetArray
110   RLabels.SetSheet (Sheet)
120   RLabels.SetPages (Me.chkPageNum)
130   RLabels.SetDBGroup FChooseGroup.GetGroup 'SetDB
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "SetupLabels", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdPrintAll_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim i As Integer
110   For i = 0 To 19
120     cmdLabel(i).Caption = "Print"
130   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "cmdPrintAll_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub Form_Load()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim i As Integer
110   For i = 0 To 19
120     cmdLabel(i).Caption = "Print"
130   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "Form_Load", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub CreateSheetArray()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100 Dim i As Integer
110   For i = 0 To 19
120     If FPrintLabels.cmdLabel(i).Caption = "Print" Then
130       Sheet(i + 1) = True
140     Else
150       Sheet(i + 1) = False
160     End If
170   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FPrintLabels", "CreateSheetArray", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
