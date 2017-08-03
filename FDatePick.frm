VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FDatePick 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Double Click to Set"
   ClientHeight    =   2760
   ClientLeft      =   6495
   ClientTop       =   4740
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   345
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2700
   End
   Begin MSComCtl2.MonthView Calendar1 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      ShowToday       =   0   'False
      StartOfWeek     =   59506689
      CurrentDate     =   38113
   End
End
Attribute VB_Name = "FDatePick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dtValue As Date

Public Property Get DateValue(ByVal CurrentValue As Variant) As Date
  On Error GoTo ErrCall
  '
  Me.Show vbModal
  '
  DateValue = dtValue
  '
  Exit Property
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.DateValue", vbCritical, "Error"

End Property

Public Property Get DateText(ByVal CurrentValue As Variant) As String
  On Error GoTo ErrCall
  '
  If IsDate(CurrentValue) Then
  '
    Calendar1.Value = CurrentValue
    dtValue = CurrentValue
  '
  Else
  '
    Calendar1.Value = Date
    dtValue = #12:00:00 AM#
  '
  End If
  '
  'dtValue = CurrentValue 'Calendar1.Value
  Me.Show vbModal
  '
  If (dtValue <> #12:00:00 AM#) And (dtValue <> CurrentValue) Then
    DateText = Format(dtValue, "mm/dd/yyyy")
  Else
    DateText = CurrentValue
  End If
  '
  Exit Property
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.DateText", vbCritical, "Error"

End Property

Private Sub Calendar1_DateClick(ByVal DateClicked As Date)
   Me.cmdSet.SetFocus
End Sub

Private Sub Calendar1_DateDblClick(ByVal DateDblClicked As Date)
  On Error GoTo ErrCall
    '
    dtValue = Calendar1.Value
    Unload Me
    '
    Exit Sub
ErrCall:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.Calendar1_DblClick", vbCritical, "Error"
End Sub

'Private Sub Calendar1_DblClick()
'  On Error GoTo ErrCall
'  '
'  dtValue = Calendar1.Value
'  Unload Me
'  '
'  Exit Sub
'ErrCall:
'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.Calendar1_DblClick", vbCritical, "Error"
'
'End Sub

Private Sub cmdSet_Click()
  On Error GoTo ErrCall
  '
  dtValue = Calendar1.Value
  '
  Unload Me
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FDatePick.cmdSet_Click", vbCritical, "Error"

End Sub

