VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RDaysLeft 
   Caption         =   "ActiveReport3"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13470
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   23760
   _ExtentY        =   10530
   SectionData     =   "RDaysLeft.dsx":0000
End
Attribute VB_Name = "RDaysLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lCount As Long

Private Sub Detail_Format()
  lCount = lCount + 1
End Sub

Private Sub GroupFooter1_Format()
  txtCount.DataValue = lCount
End Sub
