VERSION 5.00
Begin VB.Form FPrinterSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Printer"
   ClientHeight    =   1215
   ClientLeft      =   5130
   ClientTop       =   3885
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtCopies 
      Height          =   285
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copies"
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Printer"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "FPrinterSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bPrintCancel As Boolean

Private Sub cmdCancel_Click()
  bPrintCancel = True
  Me.Hide
End Sub

Private Sub cmdPrint_Click()
  If cboPrinter.Text <> "" Then
    sPrinterName = cboPrinter.Text
    iNumofCopies = Val(txtCopies.Text)
  Else
    sPrinterName = Printer.DeviceName
    iNumofCopies = Val(txtCopies.Text)
  End If
  bPrintCancel = False
  Me.Hide
End Sub

Private Sub Form_Activate()
  txtCopies.Text = 1
End Sub

Private Sub Form_Load()
  Dim prt As Printer
  Dim X As Integer
  X = 0
  For Each prt In Printers
    cboPrinter.AddItem prt.DeviceName, X
    X = X + 1
  Next
  cboPrinter.ListIndex = 0
  txtCopies.Text = 1
End Sub

