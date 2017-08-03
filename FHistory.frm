VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FHistory 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "History Report"
   ClientHeight    =   7530
   ClientLeft      =   1920
   ClientTop       =   1755
   ClientWidth     =   11700
   Icon            =   "FHistory.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7335
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11565
      Begin VB.Frame Frame10 
         BackColor       =   &H00FF8080&
         Caption         =   "Record Limit"
         Height          =   735
         Left            =   9240
         TabIndex        =   42
         Top             =   6480
         Width           =   2055
         Begin VB.CheckBox chkLimit 
            BackColor       =   &H00FF8080&
            Caption         =   "Check1"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   300
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox txtLimit 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   480
            MaxLength       =   6
            TabIndex        =   43
            Text            =   "1000"
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FF8080&
         Caption         =   "Order By"
         Height          =   735
         Left            =   6960
         TabIndex        =   41
         Top             =   6480
         Width           =   2055
         Begin VB.ComboBox cboOrder 
            Height          =   315
            ItemData        =   "FHistory.frx":0442
            Left            =   120
            List            =   "FHistory.frx":0444
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   280
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Results to Clipboad"
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   6930
         Width           =   1995
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FF8080&
         Caption         =   "Product"
         Height          =   615
         Left            =   2400
         TabIndex        =   39
         Top             =   5130
         Width           =   2175
         Begin VB.ComboBox cboProduct 
            Height          =   315
            Left            =   105
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   225
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdShowResults 
         Caption         =   "Show Results"
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   6530
         Width           =   1995
      End
      Begin VB.CommandButton cmdPreviewReport 
         Caption         =   "Preview Report"
         Height          =   315
         Left            =   3570
         TabIndex        =   15
         Top             =   6530
         Width           =   1995
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FF8080&
         Caption         =   "Category"
         Height          =   675
         Left            =   120
         TabIndex        =   34
         Top             =   4200
         Width           =   2115
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   90
            TabIndex        =   0
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   "Date"
         Height          =   855
         Left            =   2430
         TabIndex        =   29
         Top             =   4200
         Width           =   2160
         Begin VB.CommandButton cmdDateSet1 
            BackColor       =   &H00FF8080&
            Caption         =   "Set"
            Height          =   225
            Left            =   1560
            MaskColor       =   &H00FF8080&
            TabIndex        =   2
            Top             =   210
            Width           =   495
         End
         Begin VB.CommandButton cmdDateSet2 
            BackColor       =   &H00FF8080&
            Caption         =   "Set"
            Height          =   225
            Left            =   1560
            MaskColor       =   &H00FF8080&
            TabIndex        =   3
            Top             =   540
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FF8080&
            Caption         =   "From:"
            Height          =   255
            Left            =   75
            TabIndex        =   33
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FF8080&
            Caption         =   "To:"
            Height          =   255
            Left            =   210
            TabIndex        =   32
            Top             =   525
            Width           =   315
         End
         Begin VB.Label lblDate1 
            BackColor       =   &H00FF8080&
            Caption         =   "Label7"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   555
            TabIndex        =   31
            Top             =   195
            Width           =   975
         End
         Begin VB.Label lblDate2 
            BackColor       =   &H00FF8080&
            Caption         =   "Label8"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   540
            TabIndex        =   30
            Top             =   525
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF8080&
         Caption         =   "Report Type"
         Height          =   645
         Left            =   4830
         TabIndex        =   28
         Top             =   4200
         Width           =   1905
         Begin VB.ComboBox lstReport 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1665
         End
      End
      Begin VB.Frame User 
         BackColor       =   &H00FF8080&
         Caption         =   "User"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   4980
         Width           =   2115
         Begin VB.ComboBox cboUser 
            Height          =   315
            Left            =   90
            TabIndex        =   1
            Top             =   210
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         Caption         =   "Contact Status"
         Height          =   615
         Left            =   4830
         TabIndex        =   26
         Top             =   4950
         Width           =   1905
         Begin VB.ComboBox lstStatus 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   210
            Width           =   1665
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         Caption         =   "Other Criteria"
         Height          =   2145
         Left            =   6960
         TabIndex        =   21
         Top             =   4200
         Width           =   4365
         Begin VB.TextBox txtBranch 
            Height          =   315
            Left            =   990
            TabIndex        =   11
            Top             =   1320
            Width           =   3255
         End
         Begin VB.ComboBox cboType 
            Height          =   315
            Left            =   2550
            TabIndex        =   13
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtFirstName 
            Height          =   315
            Left            =   990
            TabIndex        =   8
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtLastName 
            Height          =   315
            Left            =   990
            TabIndex        =   9
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtCompany 
            Height          =   315
            Left            =   990
            TabIndex        =   10
            Top             =   960
            Width           =   3255
         End
         Begin VB.ComboBox cboState 
            Height          =   315
            Left            =   990
            TabIndex        =   12
            Text            =   "All"
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Branch:"
            Height          =   195
            Left            =   360
            TabIndex        =   40
            Top             =   1350
            Width           =   555
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FF8080&
            Caption         =   "Type:"
            Height          =   255
            Left            =   2040
            TabIndex        =   38
            Top             =   1710
            Width           =   465
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FF8080&
            Caption         =   "First Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   825
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FF8080&
            Caption         =   "Last Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   660
            Width           =   825
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FF8080&
            Caption         =   "Company:"
            Height          =   255
            Left            =   210
            TabIndex        =   23
            Top             =   990
            Width           =   765
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FF8080&
            Caption         =   "State:"
            Height          =   255
            Left            =   480
            TabIndex        =   22
            Top             =   1710
            Width           =   465
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Caption         =   "Search History Notes"
         Height          =   645
         Left            =   120
         TabIndex        =   20
         Top             =   5700
         Width           =   6705
         Begin VB.TextBox txtHistory 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   6495
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Report"
         Height          =   315
         Left            =   3570
         TabIndex        =   17
         Top             =   6930
         Width           =   1995
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdHistory 
         Bindings        =   "FHistory.frx":0446
         Height          =   3885
         Left            =   0
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   270
         Width           =   11565
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         ColumnHeaders   =   0   'False
         Col.Count       =   10
         UseGroups       =   -1  'True
         AllowUpdate     =   0   'False
         AllowColumnShrinking=   0   'False
         ForeColorEven   =   0
         BackColorEven   =   -2147483633
         BackColorOdd    =   -2147483633
         Levels          =   3
         RowHeight       =   1270
         Groups(0).Width =   20399
         Groups(0).Caption=   $"FHistory.frx":0461
         Groups(0).CaptionAlignment=   0
         Groups(0).Columns.Count=   10
         Groups(0).Columns(0).Width=   3200
         Groups(0).Columns(0).Visible=   0   'False
         Groups(0).Columns(0).Caption=   "RecID"
         Groups(0).Columns(0).Name=   "RecID"
         Groups(0).Columns(0).Alignment=   1
         Groups(0).Columns(0).CaptionAlignment=   1
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).DataType=   3
         Groups(0).Columns(0).FieldLen=   256
         Groups(0).Columns(0).Locked=   -1  'True
         Groups(0).Columns(1).Width=   3200
         Groups(0).Columns(1).Visible=   0   'False
         Groups(0).Columns(1).Caption=   "CustRecID"
         Groups(0).Columns(1).Name=   "CustRecID"
         Groups(0).Columns(1).Alignment=   1
         Groups(0).Columns(1).CaptionAlignment=   1
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).DataType=   3
         Groups(0).Columns(1).FieldLen=   256
         Groups(0).Columns(1).Locked=   -1  'True
         Groups(0).Columns(2).Width=   2937
         Groups(0).Columns(2).Caption=   "Date"
         Groups(0).Columns(2).Name=   "Date"
         Groups(0).Columns(2).CaptionAlignment=   1
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   7
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(2).HasBackColor=   -1  'True
         Groups(0).Columns(2).BackColor=   12632256
         Groups(0).Columns(3).Width=   3149
         Groups(0).Columns(3).Caption=   "Time"
         Groups(0).Columns(3).Name=   "Time"
         Groups(0).Columns(3).CaptionAlignment=   1
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   7
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(3).HasBackColor=   -1  'True
         Groups(0).Columns(3).BackColor=   12632256
         Groups(0).Columns(4).Width=   3810
         Groups(0).Columns(4).Caption=   "Type"
         Groups(0).Columns(4).Name=   "Type"
         Groups(0).Columns(4).CaptionAlignment=   0
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(4).HasBackColor=   -1  'True
         Groups(0).Columns(4).BackColor=   12632256
         Groups(0).Columns(5).Width=   3863
         Groups(0).Columns(5).Caption=   "User"
         Groups(0).Columns(5).Name=   "User"
         Groups(0).Columns(5).CaptionAlignment=   0
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   8
         Groups(0).Columns(5).FieldLen=   256
         Groups(0).Columns(5).HasBackColor=   -1  'True
         Groups(0).Columns(5).BackColor=   12632256
         Groups(0).Columns(6).Width=   6641
         Groups(0).Columns(6).Caption=   "Subject"
         Groups(0).Columns(6).Name=   "Subject"
         Groups(0).Columns(6).CaptionAlignment=   0
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   8
         Groups(0).Columns(6).FieldLen=   256
         Groups(0).Columns(6).HasBackColor=   -1  'True
         Groups(0).Columns(6).BackColor=   12632256
         Groups(0).Columns(7).Width=   12356
         Groups(0).Columns(7).Caption=   "Company"
         Groups(0).Columns(7).Name=   "Company"
         Groups(0).Columns(7).DataField=   "Column 7"
         Groups(0).Columns(7).DataType=   8
         Groups(0).Columns(7).Level=   1
         Groups(0).Columns(7).FieldLen=   256
         Groups(0).Columns(7).HasBackColor=   -1  'True
         Groups(0).Columns(7).BackColor=   16777215
         Groups(0).Columns(8).Width=   8043
         Groups(0).Columns(8).Caption=   "Name"
         Groups(0).Columns(8).Name=   "Name"
         Groups(0).Columns(8).DataField=   "Column 8"
         Groups(0).Columns(8).DataType=   8
         Groups(0).Columns(8).Level=   1
         Groups(0).Columns(8).FieldLen=   256
         Groups(0).Columns(8).HasBackColor=   -1  'True
         Groups(0).Columns(8).BackColor=   16777215
         Groups(0).Columns(9).Width=   20399
         Groups(0).Columns(9).Caption=   "Results"
         Groups(0).Columns(9).Name=   "Results"
         Groups(0).Columns(9).CaptionAlignment=   0
         Groups(0).Columns(9).DataField=   "Column 9"
         Groups(0).Columns(9).DataType=   8
         Groups(0).Columns(9).Level=   2
         Groups(0).Columns(9).FieldLen=   2000
         Groups(0).Columns(9).VertScrollBar=   -1  'True
         Groups(0).Columns(9).HasForeColor=   -1  'True
         Groups(0).Columns(9).HasBackColor=   -1  'True
         Groups(0).Columns(9).ForeColor=   -2147483640
         Groups(0).Columns(9).BackColor=   -2147483643
         _ExtentX        =   20399
         _ExtentY        =   6853
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Notes Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LblCount 
         BackColor       =   &H00FF8080&
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   36
         Top             =   0
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Report As New CReport
'
Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1

Private Sub chkLimit_Click()
  If chkLimit.Value = vbChecked Then
    txtLimit.Enabled = True
  Else
    txtLimit.Enabled = False
  End If
End Sub

'
Private Sub cmdDateSet1_Click()
  On Error GoTo EH
  Me.lblDate1.Caption = FDatePick.DateText(lblDate1.Caption)
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.cmdDateSet1_Click."
End Sub
'
Private Sub cmdDateSet2_Click()
  On Error GoTo EH
  Me.lblDate2.Caption = FDatePick.DateText(lblDate2.Caption)
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.cmdDateSet2_Click."
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  '
  Dim Count As Long
  Dim sText As String
  '
  On Error GoTo EH
  SendData
  Me.grdHistory.RemoveAll
  With Report.rsReport
      .MoveLast
      Do While Not .BOF
         Count = Count + 1
         sText = sText & !FirstName & " " & !LastName & ", "
         sText = sText & !Company & vbCrLf
         sText = sText & !Date & ", "
         sText = sText & !Time & ", "
         sText = sText & !User & vbCrLf
         sText = sText & !Type & ", "
         '
         If !Subject & vbNullString <> "" Then
            sText = sText & !Subject & ", "
         End If
         '
         sText = sText & !Results & vbCrLf & vbCrLf
         'grdHistory.AddItem !RecID & vbTab & !CustRecID & vbTab & !Date & vbTab _
         '& !Time & vbTab & !Type & vbTab & !User & vbTab & !Subject & vbTab _
         '& "Company: " & !Company & vbTab & "Contact: " & !FirstName & " " & !LastName _
         '& vbTab & !Results
         .MovePrevious
      Loop
  End With
  '
  LblCount.Caption = Count
  '
  Clipboard.SetText sText, vbCFText
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.cmdShowResults_Click."
End Sub

Private Sub Command1_Click()
'   Dim SupportAct As New CSupportACT
'   Dim SupportActData As New CSupportActData
'   '
'   SupportActData.CustRecID = 21815
'   SupportActData.ActDate = Date
'   SupportActData.Subject = "TEST SUBJECT"
'   SupportActData.Results = "TEST RESULTS"
'   SupportActData.ActUser = "TEST USER"
'   SupportActData.ActTime = "3:00"
'   SupportActData.ActType = "TEST TYPE"
'   SupportActData.ProductID = 1
'   SupportActData.ClosedTime = Now
'   SupportActData.OpenCall = False
'   SupportAct.Save SupportActData, True
'
''  'Password = Awesome
''  'User = Hurray
''  '
''  'Create Security Login
''  cnMain.Execute "EXEC sp_addlogin 'Hurray', 'awesome', 'BNB_DATA'"
''  'Give access to DB. Current one I guess.
''  cnMain.Execute "EXEC sp_grantdbaccess N'Hurray', N'Hurray'"
''  'Assign "User" role.
''  cnMain.Execute "EXEC sp_addrolemember N'User', N'Hurray'"
''  'BONUS: Change Password
''  cnMain.Execute "EXEC sp_password NULL, 'gnarly', 'Hurray'"
  
End Sub

Private Sub Form_Activate()
  On Error GoTo EH
  '
  Frame1.Move (Width - Frame1.Width) / 2
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.Form_Activate."
End Sub

Private Sub Form_Resize()
  On Error GoTo EH
  '
  Frame1.Move (Width - Frame1.Width) / 2
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.Form_Resize."
End Sub

Private Sub Form_Load()
  On Error GoTo EH
  Dim rs As New ADODB.Recordset
  '
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.Setup Me, False, , , "History Reports"
  '
  lblDate1.Caption = Date
  lblDate2.Caption = Date
  '
  rs.Open "SELECT * FROM tblactivities ORDER BY Activity", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  cboCategory.AddItem "All Categories"
  '
  Do While Not rs.eof
    cboCategory.AddItem rs!Activity
    rs.MoveNext
  Loop
  '
  rs.Close
  '
  cboCategory.AddItem "Note"
  cboCategory.Text = "All Categories"
  cboCategory.Refresh
  '
  lstReport.AddItem "Detail"
  lstReport.AddItem "Simple"
  lstReport.Text = "Detail"
  '
  rs.Open "SELECT * FROM tblStatus ORDER BY Status", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  lstStatus.AddItem "Everyone"
  '
  Do While Not rs.eof
    lstStatus.AddItem rs!Status
    rs.MoveNext
  Loop
  '
  rs.Close
  '
  lstStatus.Text = "Everyone"
  '
  rs.Open "SELECT * FROM tblEmployees ORDER BY EmployeeLast", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  cboUser.AddItem "All Users"
  '
  Do While Not rs.eof
    cboUser.AddItem rs!EmployeeFirst & " " & rs!EmployeeLast
    rs.MoveNext
  Loop
  '
  rs.Close
  '
  cboUser.Text = "All Users"
  '
  rs.Open "SELECT * FROM TType ORDER BY TypeID", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  cboType.AddItem "All Types"
  '
  Do While Not rs.eof
    cboType.AddItem rs!Description
    rs.MoveNext
  Loop
  '
  rs.Close
  '
  cboType.ListIndex = 0
  '
  FillProductBox
  '
  cboState.AddItem "All"
  cboState.AddItem "AL"
  cboState.AddItem "AK"
  cboState.AddItem "AZ"
  cboState.AddItem "AR"
  cboState.AddItem "CA"
  cboState.AddItem "CO"
  cboState.AddItem "CT"
  cboState.AddItem "DE"
  cboState.AddItem "DC"
  cboState.AddItem "FL"
  cboState.AddItem "GA"
  cboState.AddItem "HI"
  cboState.AddItem "ID"
  cboState.AddItem "IL"
  cboState.AddItem "IN"
  cboState.AddItem "IA"
  cboState.AddItem "KS"
  cboState.AddItem "KY"
  cboState.AddItem "LA"
  cboState.AddItem "ME"
  cboState.AddItem "MD"
  cboState.AddItem "MA"
  cboState.AddItem "MI"
  cboState.AddItem "MN"
  cboState.AddItem "MS"
  cboState.AddItem "MO"
  cboState.AddItem "MT"
  cboState.AddItem "NE"
  cboState.AddItem "NV"
  cboState.AddItem "NH"
  cboState.AddItem "NJ"
  cboState.AddItem "NM"
  cboState.AddItem "NY"
  cboState.AddItem "NC"
  cboState.AddItem "ND"
  cboState.AddItem "OH"
  cboState.AddItem "OK"
  cboState.AddItem "OR"
  cboState.AddItem "PA"
  cboState.AddItem "PR"
  cboState.AddItem "RI"
  cboState.AddItem "SC"
  cboState.AddItem "SD"
  cboState.AddItem "TN"
  cboState.AddItem "TX"
  cboState.AddItem "UT"
  cboState.AddItem "VT"
  cboState.AddItem "WA"
  cboState.AddItem "WV"
  cboState.AddItem "WI"
  cboState.AddItem "WY"
  '
  cboOrder.AddItem "Date/Time"
  cboOrder.AddItem "Company/Branch"
  cboOrder.AddItem "User"
  cboOrder.AddItem "Category"
  cboOrder.ListIndex = 0
  '
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.Form_Load."
End Sub


Private Sub grdHistory_DblClick()
  On Error GoTo EH
  Load FResult
  FResult.TextResult.Text = grdHistory.Columns(9).Value
  FResult.Show vbModal
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.grdHistory_DblClick."
End Sub
Private Sub SendData()
  On Error GoTo EH
  Report.FirstName = Me.txtFirstName.Text
  Report.LastName = Me.txtLastName.Text
  Report.State = Me.cboState.Text
  Report.Company = Me.txtCompany.Text
  Report.Branch = Me.txtBranch.Text
  Report.ResultsDateMin = Me.lblDate1.Caption
  Report.ResultsDateMax = Me.lblDate2.Caption
  Report.Results = Me.txtHistory.Text
  Report.ResultsType = Me.cboCategory.Text
  Report.User = Me.cboUser.Text
  Report.Status = Me.lstStatus
  Report.ContactType = Me.cboType.ListIndex
  Report.ProductID = Product.GetProductID(cboProduct.Text)
  Report.SortOrder = Me.cboOrder.Text
  If Me.chkLimit.Value = vbChecked Then
    Report.RecLimit = CInt(Me.txtLimit)
  Else
    Report.RecLimit = 0
  End If
  '
  Report.Rtype = History
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.SendData."
End Sub

Private Sub cmdPreviewReport_Click()
  On Error GoTo EH
  SendData
  If Me.lstReport.Text = "Detail" Then
    Report.PreviewReport ("History")
  Else
    Report.PreviewReport ("Simple Contact")
  End If
  Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.cmdPreviewReport_Click."
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo EH
   SendData
  If Me.lstReport.Text = "Detail" Then
    Report.PrintReport ("History")
  Else
    Report.PrintReport ("Simple Contact")
  End If
  Exit Sub
Exit Sub
EH:
 MsgBox Err.Description & " in FHistory.cmdPrint_Click."
End Sub

Private Sub cmdShowResults_Click()
  '
  Dim sCompanyandBranch As String
  Dim Count As Long
  'On Error GoTo EH
  SendData
  Me.grdHistory.RemoveAll
  With Report.rsReport
      Do While Not .eof
      Count = Count + 1
      If Not IsNull(!Branch) Then
        sCompanyandBranch = !Company & "   Branch: " & !Branch
      Else
        sCompanyandBranch = !Company
      End If
      grdHistory.AddItem !RecID & vbTab & !CustRecID & vbTab & !Date & vbTab _
      & !Time & vbTab & !Type & vbTab & !User & vbTab & !Subject & vbTab _
      & "Company: " & sCompanyandBranch & vbTab & "Contact: " & !FirstName & " " & !LastName _
      & vbTab & !Results
      .MoveNext
      Loop
  End With
  '
  LblCount.Caption = Count
  Exit Sub
'EH:
' MsgBox Err.Description & " in FHistory.cmdShowResults_Click."
End Sub

Private Sub FillProductBox()
  Dim Products As New CProducts
  Dim i As Integer
  '
  cboProduct.Clear
  '
  Product.LoadCollection Products
  '
  cboProduct.AddItem "All"
  '
  For i = 1 To Products.Count
    cboProduct.AddItem Products(i).Product
  Next i
  '
  cboProduct.ListIndex = 0
  '
  Set Products = Nothing
End Sub

Private Sub txtHistory_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdShowResults_Click
End Sub
Private Sub txtLimit_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 And KeyAscii <> 127 Then
    If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
    End If
  End If
End Sub
