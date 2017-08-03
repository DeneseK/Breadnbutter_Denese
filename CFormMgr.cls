VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private frmMainForm As MDIForm

Public Event SetStatus(Status As String) '0=Working 1=Ready
Public Event SetDescription(FormName As String, FormDescription As String)

Public Function ShowForm(pCurForm As Form, pShowForm As Form, Optional bHideCurForm As Boolean) As Boolean
  On Error GoTo ErrCall
  '
  Dim bCancel As Boolean
  '
  RaiseEvent SetStatus(0)
  '
  If Not pCurForm Is Nothing Then
    
      bCancel = Not pCurForm.FormControl.SwitchFrom
      
    '
    If bCancel Then
      Dim iRsp As Integer
      Dim sMsg As String
      '
      sMsg = "Data could not be saved or contains invalid entries. Would you like to continue editing?" & vbCrLf & vbCrLf & _
              "NOTE: If you choose no, changes to your data may not be saved."
      '
      iRsp = MsgBox(sMsg, vbQuestion + vbYesNo)
      If iRsp = vbNo Then bCancel = False
    End If
  End If
  '
  If Not bCancel Then
    frmMainForm.BackColor = &H8000000F
    '
    Dim frm As Form
    Dim bFormLoaded
    '
    For Each frm In Forms
      If frm.Name = pShowForm.Name Then
        bFormLoaded = True
        Exit For
      End If
    Next
    '
    If Not bFormLoaded Then Load pShowForm
    ResizeForm pShowForm
    pShowForm.Show
    pShowForm.FormControl.SwitchTo
    RaiseEvent SetDescription(pShowForm.Name, pShowForm.FormControl.Description)
    '
    If Not pCurForm Is Nothing Then
      If pCurForm.Name <> pShowForm.Name Then
        If bHideCurForm Then
          pCurForm.Hide
        Else
          Unload pCurForm
        End If
      End If
    End If
    '
    ShowForm = True
  Else
    ShowForm = False
  End If
  '
  RaiseEvent SetStatus(1)
  '
  Exit Function
ErrCall:
  ShowForm = False
  MsgBox Err.Description
  
End Function

Public Sub Setup(pfrmMDI As MDIForm)
  Set frmMainForm = pfrmMDI
End Sub

Public Sub ResizeForm(pForm As Form)
  On Error GoTo ErrCall
  '
  Dim sngHeight As Single, sngWidth As Single
  '
  If Not pForm Is Nothing Then
    sngHeight = frmMainForm.ScaleHeight
    If (sngHeight < pForm.FormControl.MinHeight) And (sngHeight > 0) Then
      sngHeight = pForm.FormControl.MinHeight
    End If
    '
    'If frmMainForm.Width - frmMainForm.ScaleWidth > 181 Then
    '  sngWidth = frmMainForm.ScaleWidth + 195
    'Else
      sngWidth = frmMainForm.ScaleWidth - (2 * Screen.TwipsPerPixelX)
    'End If
    '
    If sngWidth < pForm.FormControl.MinWidth Then sngWidth = pForm.FormControl.MinWidth
    '
    pForm.Move 0, 0, sngWidth, sngHeight
  End If
  '
  Exit Sub
ErrCall:
  If Not Err.Number = 384 Then
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormMgr.ResizeForm.", vbCritical, "Error"
  End If
End Sub

Public Function CloseForm(pForm As Form) As Boolean
  On Error GoTo ErrCall
  '
  Dim bCancel As Boolean
  '
  CloseForm = False
  '
  If Not pForm Is Nothing Then
    bCancel = Not pForm.FormControl.SwitchFrom
    '
    If bCancel Then
      Dim iRsp As Integer
      Dim sMsg As String
      '
      sMsg = "Data could not be saved or contains invalid entries. Would you like to continue editing?" & vbCrLf & vbCrLf & _
              "NOTE: If you choose no, changes to your data may not be saved."
      '
      iRsp = MsgBox(sMsg, vbQuestion + vbYesNo)
      If iRsp = vbNo Then bCancel = False
    End If
    '
    If Not bCancel Then
      Unload pForm
      RaiseEvent SetDescription("", "Main")
      '
      If frmMainForm.ActiveForm Is Nothing Then
        frmMainForm.BackColor = &H8000000C
      End If
      '
      CloseForm = True
    Else
      CloseForm = False
    End If
  Else
    CloseForm = True
  End If
  '
  ' CSErrorHandler begin - please do not modify or remove this line
  Exit Function
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormMgr.CloseForm.", vbCritical, "Error"
  CloseForm = False
End Function

Private Sub Class_Terminate()
  On Error GoTo ErrCall
  '
  Set frmMainForm = Nothing
  '
  Exit Sub
ErrCall:
  MsgBox "Error: " & Err.Description & " in Form Manager Terminate."
End Sub
