VERSION 5.00
Begin VB.Form FSingleLabel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Single Label"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   Icon            =   "FSingleLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   345
      Left            =   1230
      TabIndex        =   23
      Top             =   4170
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pick Label"
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   3345
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   19
         Left            =   1680
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   18
         Left            =   150
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   17
         Left            =   1680
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   16
         Left            =   150
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   15
         Left            =   1680
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   14
         Left            =   150
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   13
         Left            =   1680
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   12
         Left            =   150
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   11
         Left            =   1680
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   10
         Left            =   150
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   9
         Left            =   1680
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   8
         Left            =   150
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   7
         Left            =   1680
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   6
         Left            =   150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   5
         Left            =   1680
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   4
         Left            =   150
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   3
         Left            =   1680
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   2
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Unknown"
         Height          =   345
         Index           =   0
         Left            =   150
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   2430
      TabIndex        =   1
      Top             =   4170
      Width           =   1125
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   4170
      Width           =   1125
   End
End
Attribute VB_Name = "FSingleLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sheet(20) As Boolean
Private lContactID As Long

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
      ErrorMgr.Raise "FSingleLabel", "cmdCancel_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdLabel_Click(Index As Integer)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100  Dim i
110  For i = 0 To 19
120     cmdLabel(i).Caption = ""
130   Next i
140  cmdLabel(Index).Caption = "Print"
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FSingleLabel", "cmdLabel_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
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
      ErrorMgr.Raise "FSingleLabel", "cmdPreview_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Sub PrintSingle(plContactID As Long)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If plContactID > 0 Then
110     lContactID = plContactID
120     Me.Show vbModal
130   Else
140     MsgBox "Contact not ready"
150   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FSingleLabel", "PrintSingle", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub


Private Sub Form_Load()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
90      Set RLabels = New RLabels
100   Dim i As Integer
110   For i = 0 To 19
120     cmdLabel(i).Caption = ""
130   Next i
140     cmdLabel(0).Caption = "Print"
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FSingleLabel", "Form_Load", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdPrint_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SetupLabels
110   RLabels.PrintReport True
      'RLabels.Show vbModal, FMain
120   Unload Me
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FSingleLabel", "cmdPrint_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub CreateSheetArray()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100 Dim i As Integer
110   For i = 0 To 19
120     If cmdLabel(i).Caption = "Print" Then
130       Sheet(i + 1) = True
140     Else
150       Sheet(i + 1) = False
160     End If
170   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FSingleLabel", "CreateSheetArray", Err.Number, Err.Description, Erl
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
120   RLabels.SetPages 1
130   RLabels.SetDBCurrent lContactID
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "FSingleLabel", "SetupLabels", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
