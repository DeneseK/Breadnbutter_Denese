VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sngMinHeight As Single
Private sngMinWidth As Single

Private bDataForm As Boolean

Private sDescription As String

Public AlwaysVisible As Boolean

Public Event SwitchFrom(bCancel As Boolean)
Public Event SwitchTo(bCancel As Boolean)

Public Function SwitchFrom() As Boolean
  On Error GoTo ErrCall
  '
  Dim bCancel As Boolean
  '
  RaiseEvent SwitchFrom(bCancel)
  '
  SwitchFrom = Not bCancel
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormControl.SwitchFrom.", vbCritical, "Error"
End Function

Public Function SwitchTo() As Boolean
  On Error GoTo ErrCall
  '
  Dim bCancel As Boolean
  '
  RaiseEvent SwitchTo(bCancel)
  '
  SwitchTo = Not bCancel
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormControl.SwitchTo.", vbCritical, "Error"
End Function

Public Property Get MinHeight() As Single
  MinHeight = sngMinHeight
End Property

Public Property Let MinHeight(ByVal psngMinHeight As Single)
  sngMinHeight = psngMinHeight
End Property

Public Property Get MinWidth() As Single
  MinWidth = sngMinWidth
End Property

Public Property Let MinWidth(ByVal psngMinWidth As Single)
  sngMinWidth = psngMinWidth
End Property

Public Property Get DataForm() As Boolean
  DataForm = bDataForm
End Property

Public Property Let DataForm(ByVal pbDataForm As Boolean)
  bDataForm = pbDataForm
End Property

Public Sub Setup(pForm As Form, pbData As Boolean, Optional psngHeight As Single, Optional psngWidth As Single, Optional psDescription As String, Optional pfAlwaysVisible As Boolean)
  On Error GoTo ErrCall
  '
  sngMinHeight = IIf(psngHeight = 0, pForm.Height, psngHeight)
  sngMinWidth = IIf(psngWidth = 0, pForm.Width, psngWidth)
  bDataForm = pbData
  sDescription = psDescription
  AlwaysVisible = pfAlwaysVisible
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormControl.Setup.", vbCritical, "Error"
End Sub

Public Property Get Description() As String
  Description = sDescription
End Property
