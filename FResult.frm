VERSION 5.00
Begin VB.Form FResult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Result"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "FResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   2070
      TabIndex        =   1
      Top             =   2220
      Width           =   2025
   End
   Begin VB.TextBox TextResult 
      Height          =   2115
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FResult.frx":0442
      Top             =   30
      Width           =   6135
   End
End
Attribute VB_Name = "FResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  On Error GoTo ErrCall
    Unload Me
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FResult.cmdClose", vbCritical, "Error"
End Sub


