VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FEmployeeMgt 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Employee Management"
   ClientHeight    =   8535
   ClientLeft      =   2460
   ClientTop       =   2475
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtIcon 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9840
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3840
         Width           =   255
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add New"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Frame fmeAddEdit 
         BackColor       =   &H000080FF&
         Height          =   2535
         Left            =   360
         TabIndex        =   10
         Top             =   4320
         Width           =   9495
         Begin VB.TextBox txtMail 
            Height          =   285
            Left            =   6720
            TabIndex        =   9
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            Left            =   7200
            TabIndex        =   21
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtFirst 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtMid 
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtLast 
            Height          =   285
            Left            =   3720
            TabIndex        =   7
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   300
            Left            =   7200
            TabIndex        =   22
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Frame fmeGroups 
            BackColor       =   &H000080FF&
            Caption         =   "Groups"
            Height          =   1335
            Left            =   360
            TabIndex        =   26
            Top             =   960
            Width           =   1575
            Begin VB.CheckBox chkAuthorizations 
               BackColor       =   &H000080FF&
               Caption         =   "Authorizations"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chkSales 
               BackColor       =   &H000080FF&
               Caption         =   "Sales"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox chkSupport 
               BackColor       =   &H000080FF&
               Caption         =   "Support"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   975
            End
            Begin VB.CheckBox chkOperator 
               BackColor       =   &H000080FF&
               Caption         =   "Operator"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Width           =   1095
            End
         End
         Begin VB.Frame fmeWorkGroups 
            BackColor       =   &H000080FF&
            Caption         =   "WorkGroups"
            Height          =   1335
            Left            =   2760
            TabIndex        =   25
            Top             =   960
            Width           =   1575
            Begin VB.CheckBox chkDev 
               BackColor       =   &H000080FF&
               Caption         =   "Development"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   960
               Width           =   1335
            End
            Begin VB.CheckBox chkWorkSupport 
               BackColor       =   &H000080FF&
               Caption         =   "Support"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   720
               Width           =   975
            End
            Begin VB.CheckBox chkWorkSales 
               BackColor       =   &H000080FF&
               Caption         =   "Sales"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox chkManagement 
               BackColor       =   &H000080FF&
               Caption         =   "Management"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox txtExt 
            Height          =   285
            Left            =   5880
            TabIndex        =   8
            Top             =   480
            Width           =   735
         End
         Begin VB.Frame fmeSecurity 
            BackColor       =   &H000080FF&
            Caption         =   "Security Level"
            Height          =   1335
            Left            =   5040
            TabIndex        =   24
            Top             =   960
            Width           =   1335
            Begin VB.OptionButton optLow 
               BackColor       =   &H000080FF&
               Caption         =   "Low"
               Height          =   255
               Left            =   240
               TabIndex        =   20
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton optHigh 
               BackColor       =   &H000080FF&
               Caption         =   "High"
               Height          =   255
               Left            =   240
               TabIndex        =   19
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   300
            Left            =   7200
            TabIndex        =   23
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblMail 
            BackColor       =   &H000080FF&
            Caption         =   "E-Mail  Address"
            Height          =   255
            Left            =   6720
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblPass 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Password"
            Height          =   195
            Left            =   7200
            TabIndex        =   31
            Top             =   960
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblFirst 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "First Name"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lblLast 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Last Name"
            Height          =   195
            Left            =   3720
            TabIndex        =   29
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblMid 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Middle Name"
            Height          =   195
            Left            =   1920
            TabIndex        =   28
            Top             =   240
            Width           =   930
         End
         Begin VB.Label lblExt 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Ext."
            Height          =   195
            Left            =   5880
            TabIndex        =   27
            Top             =   240
            Width           =   270
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ext"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Security Level"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Groups"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "WorkGroups"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "E-Mail"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "FEmployeeMgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
Private iTempID As Integer
Dim bAddNew As Boolean

Private Sub LoadEmployeeList()
  Dim Employees As New CEmployees
  Dim Employee As New CEmployee
  Dim iCounter As Integer
  Dim sKey As String
  '
  Employee.LoadCollection Employees
  '
  ListView1.ListItems.Clear
  '
  For iCounter = 1 To Employees.Count
    sKey = "A" & CStr(Employees.Item(iCounter).EmployeeID)
    ListView1.ListItems.Add , sKey, Employees.Item(iCounter).EmployeeID
    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , Employees.Item(iCounter).EmployeeFirst & " " & Employees.Item(iCounter).EmployeeLast, , Employees.Item(iCounter).EmployeeFirst & " " & Employees.Item(iCounter).EmployeeLast).ForeColor = vbBlack
    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , Employees.Item(iCounter).EmployeeExt, , Employees.Item(iCounter).EmployeeExt).ForeColor = vbBlack
    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , Employees.Item(iCounter).SecurityLevel, , Employees.Item(iCounter).SecurityLevel).ForeColor = vbBlack
    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , Employees.Item(iCounter).Groups, , Employees.Item(iCounter).Groups).ForeColor = vbBlack
    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , Employees.Item(iCounter).WorkGroups, , Employees.Item(iCounter).WorkGroups).ForeColor = vbBlack
    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , Employees.Item(iCounter).EMailAddress, , Employees.Item(iCounter).EMailAddress).ForeColor = vbBlack
  Next
  Set Employees = Nothing
  Set Employee = Nothing
End Sub

Private Sub LoadGroups(piTemp As Integer)
  '
  If piTemp >= 8 Then
    chkAuthorizations.Value = 1
    piTemp = piTemp - 8
  Else
    chkAuthorizations.Value = 0
  End If
  '
  If piTemp >= 4 Then
    chkSales.Value = 1
    piTemp = piTemp - 4
  Else
    chkSales.Value = 0
  End If
  '
  If piTemp >= 2 Then
    chkSupport.Value = 1
    piTemp = piTemp - 2
  Else
    chkSupport.Value = 0
  End If
  '
  If piTemp >= 1 Then
    chkOperator.Value = 1
  Else
    chkOperator.Value = 0
  End If
  '
End Sub

Private Function SaveGroup() As Integer
  Dim iGroupNumber As Integer
  '
  iGroupNumber = 0
  '
  If chkOperator.Value = 1 Then
    iGroupNumber = iGroupNumber + 1
  End If
  '
  If chkSupport.Value = 1 Then
    iGroupNumber = iGroupNumber + 2
  End If
  '
  If chkSales.Value = 1 Then
    iGroupNumber = iGroupNumber + 4
  End If
  '
  If chkAuthorizations.Value = 1 Then
    iGroupNumber = iGroupNumber + 8
  End If
  SaveGroup = iGroupNumber
End Function

Private Sub LoadWorkGroups(piTemp As Integer)
  '
  If piTemp >= 8 Then
    chkManagement.Value = 1
    piTemp = piTemp - 8
  Else
    chkManagement.Value = 0
  End If
  '
  If piTemp >= 4 Then
    chkWorkSales.Value = 1
    piTemp = piTemp - 4
  Else
    chkWorkSales.Value = 0
  End If
  '
  If piTemp >= 2 Then
    chkWorkSupport.Value = 1
    piTemp = piTemp - 2
  Else
    chkWorkSupport.Value = 0
  End If
  '
  If piTemp >= 1 Then
    chkDev.Value = 1
  Else
    chkDev.Value = 0
  End If
  '
End Sub

Private Function SaveWorkGroup() As Integer
  Dim iGroupNumber As Integer
  '
  iGroupNumber = 0
  '
  If chkDev.Value = 1 Then
    iGroupNumber = iGroupNumber + 1
  End If
  '
  If chkWorkSupport.Value = 1 Then
    iGroupNumber = iGroupNumber + 2
  End If
  '
  If chkWorkSales.Value = 1 Then
    iGroupNumber = iGroupNumber + 4
  End If
  '
  If chkManagement.Value = 1 Then
    iGroupNumber = iGroupNumber + 8
  End If
  SaveWorkGroup = iGroupNumber
End Function



Private Sub cmdAdd_Click()
  bAddNew = True
  EnableEdit
  cmdDelete.Enabled = False
  txtFirst.SetFocus
End Sub

Private Sub cmdCancel_Click()
  DisableEdit
  bAddNew = False
  cmdDelete.Enabled = True
End Sub

Private Sub cmdDelete_Click()
  Dim Employee As New CEmployee
  bAddNew = False
  If MsgBox("Confirm Delete", vbYesNo, "Delete Employee") = vbYes Then
    Employee.Delete ListView1.SelectedItem.Text
  End If
  Set Employee = Nothing
  LoadEmployeeList
End Sub

Private Sub cmdEdit_Click()
  Dim Employee As New CEmployee
  Dim EmployeeData As New CEmployeeData
  '
  Employee.Load EmployeeData, ListView1.SelectedItem.Text
  '
  bAddNew = False
  EnableEdit
  '
  iTempID = EmployeeData.EmployeeID
  txtFirst.Text = EmployeeData.EmployeeFirst
  txtMid.Text = EmployeeData.EmployeeMiddle
  txtLast.Text = EmployeeData.EmployeeLast
  txtExt.Text = EmployeeData.EmployeeExt
  txtPassword.Text = DecryptStr(EmployeeData.Password)
  LoadGroups EmployeeData.Groups
  txtMail.Text = EmployeeData.EMailAddress
  LoadWorkGroups EmployeeData.WorkGroups
  If EmployeeData.SecurityLevel = 1 Then
    optLow.Value = True
  Else
    optHigh.Value = True
  End If
  '
  Set EmployeeData = Nothing
  Set Employee = Nothing
  cmdDelete.Enabled = False
End Sub

Private Sub cmdSave_Click()
  Dim Employee As New CEmployee
  Dim EmployeeData As New CEmployeeData
  '
  EmployeeData.EmployeeID = iTempID
  EmployeeData.EmployeeFirst = txtFirst.Text
  EmployeeData.EmployeeMiddle = txtMid.Text
  EmployeeData.EmployeeLast = txtLast.Text
  EmployeeData.EmployeeExt = Val(txtExt.Text)
  EmployeeData.Groups = SaveGroup
  EmployeeData.WorkGroups = SaveWorkGroup
  EmployeeData.Password = txtPassword.Text
  EmployeeData.EMailAddress = txtMail.Text
  
  
  If optLow.Value = True Then
    EmployeeData.SecurityLevel = 1
  Else
    EmployeeData.SecurityLevel = 2
  End If
  '
  If bAddNew = False Then
    Employee.Save EmployeeData, iTempID
  Else
    Employee.AddNew EmployeeData
  End If
  '
  Set EmployeeData = Nothing
  Set Employee = Nothing
  '
  LoadEmployeeList
  DisableEdit
  bAddNew = False
  cmdDelete.Enabled = True
End Sub

Private Sub Form_Load()
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.Setup Me, False, , , "Employee Management"
  '
  LoadEmployeeList
  bAddNew = False
  DisableEdit
End Sub

Private Sub Form_Resize()
  On Error GoTo EH
  '
  Frame1.Move (Width - Frame1.Width) / 2
  Exit Sub
EH:
 MsgBox Err.Description & " in FEmployeeMgt.Form_Resize."
End Sub

Private Sub DisableEdit()
  txtFirst.Text = ""
  txtMid.Text = ""
  txtLast.Text = ""
  txtExt.Text = ""
  txtPassword.Text = ""
  txtMail.Text = ""
  optLow.Value = True
  chkManagement.Value = 0
  chkWorkSales.Value = 0
  chkWorkSupport.Value = 0
  chkDev.Value = 0
  chkAuthorizations.Value = 0
  chkSales.Value = 0
  chkSupport.Value = 0
  chkOperator.Value = 0
  '
  txtFirst.Enabled = False
  txtMid.Enabled = False
  txtLast.Enabled = False
  txtExt.Enabled = False
  txtMail.Enabled = False
    '
  txtPassword.Visible = False
  lblPass.Visible = False
  '
  optLow.Enabled = False
  optHigh.Enabled = False
  chkManagement.Enabled = False
  chkWorkSales.Enabled = False
  chkWorkSupport.Enabled = False
  chkDev.Enabled = False
  chkAuthorizations.Enabled = False
  chkSales.Enabled = False
  chkSupport.Enabled = False
  chkOperator.Enabled = False
  cmdSave.Enabled = False
  cmdCancel.Enabled = False
End Sub

Private Sub EnableEdit()
  txtFirst.Text = ""
  txtMid.Text = ""
  txtLast.Text = ""
  txtExt.Text = ""
  txtPassword.Text = ""
  txtMail.Text = ""
  optLow.Value = True
  chkManagement.Value = 0
  chkWorkSales.Value = 0
  chkWorkSupport.Value = 0
  chkDev.Value = 0
  chkAuthorizations.Value = 0
  chkSales.Value = 0
  chkSupport.Value = 0
  chkOperator.Value = 0
  '
  txtFirst.Enabled = True
  txtMid.Enabled = True
  txtLast.Enabled = True
  txtExt.Enabled = True
  txtMail.Enabled = True
  optLow.Enabled = True
  optHigh.Enabled = True
  chkManagement.Enabled = True
  chkWorkSales.Enabled = True
  chkWorkSupport.Enabled = True
  chkDev.Enabled = True
  chkAuthorizations.Enabled = True
  chkSales.Enabled = True
  chkSupport.Enabled = True
  chkOperator.Enabled = True
  cmdSave.Enabled = True
  cmdCancel.Enabled = True
  '
  'If bAddNew = True Then
    txtPassword.Visible = True
    lblPass.Visible = True
  'End If
  '
End Sub

Private Sub ListView1_Click()
  cmdCancel_Click
End Sub

Private Sub txtIcon_KeyUp(KeyCode As Integer, Shift As Integer)
  If Shift = 7 And KeyCode = 73 Then
    If GetSetting(App.Title, "Settings", "Icon", 0) = 1 Then
      SaveSetting App.Title, "Settings", "Icon", 0
      MsgBox "Icon Enabled"
      FMain.GetUserGroups
    Else
      SaveSetting App.Title, "Settings", "Icon", 1
      MsgBox "Icon Disabled"
    End If
  End If
  Shift = 0
  KeyCode = 0
  End Sub
