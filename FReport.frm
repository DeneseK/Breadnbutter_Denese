VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FReport 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Report Printer"
   ClientHeight    =   8430
   ClientLeft      =   2130
   ClientTop       =   2235
   ClientWidth     =   11955
   Icon            =   "FReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame frame1 
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11685
      Begin VB.CommandButton cmdCopyToClipBoard 
         Caption         =   "Copy To Clipboard"
         Height          =   345
         Left            =   7200
         TabIndex        =   49
         Top             =   7560
         Width           =   1965
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdHistory 
         Bindings        =   "FReport.frx":0442
         Height          =   1725
         Left            =   4560
         TabIndex        =   48
         Top             =   5640
         Width           =   7065
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         ColumnHeaders   =   0   'False
         Col.Count       =   8
         UseGroups       =   -1  'True
         AllowUpdate     =   0   'False
         AllowColumnShrinking=   0   'False
         ForeColorEven   =   0
         BackColorEven   =   -2147483633
         BackColorOdd    =   -2147483633
         Levels          =   3
         RowHeight       =   1270
         Groups(0).Width =   11827
         Groups(0).Caption=   "Date/User                          Time /Subject                         Type"
         Groups(0).CaptionAlignment=   0
         Groups(0).Columns.Count=   8
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
         Groups(0).Columns(2).Width=   3440
         Groups(0).Columns(2).Caption=   "Date"
         Groups(0).Columns(2).Name=   "Date"
         Groups(0).Columns(2).CaptionAlignment=   1
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).DataType=   7
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(3).Width=   3704
         Groups(0).Columns(3).Caption=   "Time"
         Groups(0).Columns(3).Name=   "Time"
         Groups(0).Columns(3).CaptionAlignment=   1
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   7
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(4).Width=   4683
         Groups(0).Columns(4).Caption=   "Type"
         Groups(0).Columns(4).Name=   "Type"
         Groups(0).Columns(4).CaptionAlignment=   0
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   8
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(5).Width=   3440
         Groups(0).Columns(5).Caption=   "User"
         Groups(0).Columns(5).Name=   "User"
         Groups(0).Columns(5).CaptionAlignment=   0
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   8
         Groups(0).Columns(5).Level=   1
         Groups(0).Columns(5).FieldLen=   256
         Groups(0).Columns(6).Width=   8387
         Groups(0).Columns(6).Caption=   "Subject"
         Groups(0).Columns(6).Name=   "Subject"
         Groups(0).Columns(6).CaptionAlignment=   0
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   8
         Groups(0).Columns(6).Level=   1
         Groups(0).Columns(6).FieldLen=   256
         Groups(0).Columns(7).Width=   11827
         Groups(0).Columns(7).Caption=   "Results"
         Groups(0).Columns(7).Name=   "Results"
         Groups(0).Columns(7).CaptionAlignment=   0
         Groups(0).Columns(7).DataField=   "Column 7"
         Groups(0).Columns(7).DataType=   8
         Groups(0).Columns(7).Level=   2
         Groups(0).Columns(7).FieldLen=   256
         Groups(0).Columns(7).HasForeColor=   -1  'True
         Groups(0).Columns(7).HasBackColor=   -1  'True
         Groups(0).Columns(7).ForeColor=   -2147483640
         Groups(0).Columns(7).BackColor=   -2147483643
         _ExtentX        =   12462
         _ExtentY        =   3043
         _StockProps     =   79
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
      Begin VB.CommandButton ShowResults 
         Caption         =   "Show Results"
         Height          =   345
         Left            =   2880
         TabIndex        =   22
         Top             =   7560
         Width           =   1965
      End
      Begin VB.CommandButton PreviewReport 
         Caption         =   "Old Report"
         Height          =   345
         Left            =   480
         TabIndex        =   23
         Top             =   7560
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Choose Search"
         ForeColor       =   &H00000000&
         Height          =   7305
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   4320
         Begin VB.TextBox txtDays 
            Height          =   285
            Left            =   360
            TabIndex        =   53
            Text            =   "90"
            Top             =   1520
            Width           =   495
         End
         Begin VB.OptionButton optChoice 
            BackColor       =   &H00FFCCCC&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   52
            Top             =   1520
            Width           =   285
         End
         Begin VB.CheckBox chkTtls 
            BackColor       =   &H00FFCCCC&
            Caption         =   "Show &Totals"
            Height          =   225
            Left            =   2190
            TabIndex        =   12
            Top             =   6645
            Width           =   1185
         End
         Begin VB.CheckBox chkFilter 
            BackColor       =   &H00FFCCCC&
            Caption         =   "AM &Best"
            Height          =   225
            Index           =   1
            Left            =   1200
            TabIndex        =   11
            Top             =   6645
            Width           =   915
         End
         Begin VB.CheckBox chkFilter 
            BackColor       =   &H00FFCCCC&
            Caption         =   "&Standard"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   6645
            Width           =   975
         End
         Begin VB.ListBox lstGroups 
            Height          =   2985
            Left            =   120
            TabIndex        =   9
            Top             =   3240
            Width           =   4095
         End
         Begin VB.OptionButton optChoice 
            BackColor       =   &H00FFCCCC&
            Caption         =   "Group"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   45
            TabIndex        =   8
            Top             =   2880
            Width           =   945
         End
         Begin VB.ComboBox cboProduct 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1110
            Width           =   1695
         End
         Begin VB.OptionButton optChoice 
            BackColor       =   &H00FFE1E1&
            Caption         =   "Notes"
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   765
         End
         Begin VB.OptionButton optChoice 
            BackColor       =   &H00FFCCCC&
            Caption         =   "Days Not Authorized"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton optChoice 
            BackColor       =   &H00FFCCCC&
            Caption         =   "Days Left"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   1125
         End
         Begin VB.OptionButton optChoice 
            BackColor       =   &H00FFE1E1&
            Caption         =   "Basic"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   90
            Width           =   945
         End
         Begin VB.Frame Days 
            BackColor       =   &H00FFCCCC&
            BorderStyle     =   0  'None
            Caption         =   "Days"
            Height          =   975
            Left            =   1920
            TabIndex        =   37
            Top             =   510
            Width           =   2235
            Begin VB.TextBox TextDaysMax 
               Height          =   315
               Left            =   645
               TabIndex        =   5
               Text            =   "30"
               Top             =   495
               Width           =   1305
            End
            Begin VB.TextBox TextDaysMin 
               Height          =   315
               Left            =   645
               TabIndex        =   4
               Text            =   "0"
               Top             =   165
               Width           =   1305
            End
            Begin VB.Label Label10 
               BackColor       =   &H00FFCCCC&
               Caption         =   " Days"
               Height          =   225
               Left            =   105
               TabIndex        =   38
               Top             =   -15
               Width           =   465
            End
            Begin VB.Label Label22 
               BackColor       =   &H00FFCCCC&
               Caption         =   "Min:"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   225
               TabIndex        =   40
               Top             =   225
               Width           =   375
            End
            Begin VB.Label Label23 
               BackColor       =   &H00FFCCCC&
               Caption         =   "Max:"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   195
               TabIndex        =   39
               Top             =   555
               Width           =   405
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00FFCCCC&
               BorderColor     =   &H00C0C0C0&
               Height          =   825
               Left            =   30
               Shape           =   4  'Rounded Rectangle
               Top             =   75
               Width           =   2115
            End
         End
         Begin VB.Frame Text 
            BackColor       =   &H00FFE1E1&
            BorderStyle     =   0  'None
            Caption         =   "Text"
            Height          =   765
            Left            =   960
            TabIndex        =   35
            Top             =   1950
            Width           =   3135
            Begin VB.ComboBox textNotes 
               Height          =   315
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   2955
            End
            Begin VB.Label Label12 
               BackColor       =   &H00FFE1E1&
               Caption         =   "Text"
               Height          =   225
               Left            =   120
               TabIndex        =   36
               Top             =   0
               Width           =   465
            End
            Begin VB.Shape Shape5 
               BorderColor     =   &H00C0C0C0&
               Height          =   555
               Left            =   0
               Shape           =   4  'Rounded Rectangle
               Top             =   120
               Width           =   3105
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFE1E1&
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   1
            Left            =   0
            ScaleHeight     =   855
            ScaleWidth      =   4335
            TabIndex        =   44
            Top             =   1920
            Width           =   4335
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFCCCC&
            BorderStyle     =   0  'None
            Height          =   4455
            Index           =   2
            Left            =   0
            ScaleHeight     =   4455
            ScaleWidth      =   4335
            TabIndex        =   45
            Top             =   2760
            Width           =   4335
            Begin VB.CheckBox chkAlpha 
               BackColor       =   &H00FFCCCC&
               Caption         =   "Sort Alphabetically"
               Height          =   255
               Left            =   135
               TabIndex        =   13
               Top             =   3600
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFCCCC&
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   4
            Left            =   0
            ScaleHeight     =   1455
            ScaleWidth      =   4335
            TabIndex        =   46
            Top             =   480
            Width           =   4335
            Begin VB.ComboBox cboAction 
               Height          =   315
               Left            =   1800
               TabIndex        =   50
               Top             =   1020
               Width           =   1695
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFCCCC&
               Caption         =   "Days Since"
               Height          =   195
               Left            =   960
               TabIndex        =   54
               Top             =   1080
               Width           =   810
            End
            Begin VB.Label Label8 
               BackColor       =   &H00FFCCCC&
               Caption         =   "Contact"
               Height          =   255
               Left            =   3600
               TabIndex        =   51
               Top             =   1080
               Width           =   615
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFE1E1&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   5
            Left            =   0
            ScaleHeight     =   615
            ScaleWidth      =   4335
            TabIndex        =   47
            Top             =   0
            Width           =   4335
         End
      End
      Begin VB.Frame frmCriteria 
         BorderStyle     =   0  'None
         Caption         =   "Other Criteria"
         ForeColor       =   &H00000000&
         Height          =   1725
         Left            =   4560
         TabIndex        =   26
         Top             =   5640
         Width           =   6945
         Begin VB.TextBox TextSource 
            Height          =   285
            Left            =   4380
            MaxLength       =   100
            TabIndex        =   21
            Top             =   1305
            Width           =   2070
         End
         Begin VB.ComboBox ComboStatus 
            Height          =   315
            Left            =   1575
            TabIndex        =   14
            Text            =   "Customer"
            Top             =   210
            Width           =   2055
         End
         Begin VB.TextBox TextZip 
            Height          =   315
            Left            =   4380
            TabIndex        =   20
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox TextCity 
            Height          =   315
            Left            =   4380
            TabIndex        =   18
            Top             =   225
            Width           =   2055
         End
         Begin VB.TextBox TextCompany 
            Height          =   315
            Left            =   1575
            TabIndex        =   17
            Top             =   1260
            Width           =   2055
         End
         Begin VB.TextBox TextLastName 
            Height          =   315
            Left            =   1575
            TabIndex        =   16
            Top             =   900
            Width           =   2055
         End
         Begin VB.TextBox TextFirstName 
            Height          =   315
            Left            =   1575
            TabIndex        =   15
            Top             =   540
            Width           =   2055
         End
         Begin VB.ComboBox StateCombo 
            Height          =   315
            Left            =   4380
            TabIndex        =   19
            Text            =   "All"
            Top             =   585
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Source:"
            Height          =   255
            Left            =   3690
            TabIndex        =   43
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Status:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1005
            TabIndex        =   33
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label6 
            Caption         =   "Zip:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3990
            TabIndex        =   32
            Top             =   975
            Width           =   345
         End
         Begin VB.Label Label5 
            Caption         =   "State:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3840
            TabIndex        =   31
            Top             =   615
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "City:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3990
            TabIndex        =   30
            Top             =   255
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Company:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   765
            TabIndex        =   29
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label Label2 
            Caption         =   "Last Name:"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   675
            TabIndex        =   28
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "First Name:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   735
            TabIndex        =   27
            Top             =   600
            Width           =   825
         End
      End
      Begin VB.CommandButton PrintButton 
         Caption         =   "Print Report"
         Height          =   345
         Left            =   5040
         TabIndex        =   24
         Top             =   7560
         Width           =   1965
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5295
         Left            =   4530
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   270
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label LabelResults 
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   4560
         TabIndex        =   42
         Top             =   60
         Width           =   1245
      End
   End
End
Attribute VB_Name = "FReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Report As New CReport
'
Private Enum eFilter
  FilterNone
  FilterStandard
  FilterAMBest
  FilterAll
End Enum
'
Private fShowTotals       As Boolean
Private fSettingPrefs     As Boolean

Private lProspectGroupID  As Long
'
Private rsGroupCategories   As ADODB.Recordset
Private rsProspectGroup     As ADODB.Recordset
'
Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
'
Private Filter            As eFilter
'
Private WithEvents objLvPrint As clsPrintLV
Attribute objLvPrint.VB_VarHelpID = -1
'
Private iLastKey As Integer
'
'Private SortColumn        As eSortColumn
'
Private Sub Command2_Click()
  On Error GoTo EH
  '
  FHistory.Show
  '
  Exit Sub
EH:
 MsgBox Err.Description & " in FReport.Command2_Click."
End Sub

Private Sub chkAlpha_Click()
  If chkAlpha.Value = vbChecked Then
    SaveSetting App.Title, "ProspectMgt", "SortColumn", 1
  Else
    SaveSetting App.Title, "ProspectMgt", "SortColumn", 0
  End If
  '
  SetupGroups
  SelectGroup
End Sub

Private Sub cmdCopyToClipBoard_Click()
  Dim sText As String
  Dim lItemCount As Long
  Dim lSubItemCount As Long
  Dim lColumnCount As Long
  '
  Clipboard.Clear
  '
  Screen.MousePointer = vbHourglass
  SetupData
  ListView1.ListItems.Clear
  ListView1.ColumnHeaders.Clear
  Report.FillList Report.rsReport, ListView1
  '
  For lColumnCount = 1 To ListView1.ColumnHeaders.Count
    sText = sText & ListView1.ColumnHeaders(lColumnCount).Text & vbTab
  Next
  '
  sText = sText & vbCrLf
  '
  For lItemCount = 1 To ListView1.ListItems.Count
    sText = sText & ListView1.ListItems(lItemCount).Text & vbTab
    '
    For lSubItemCount = 1 To ListView1.ListItems(lItemCount).ListSubItems.Count
      sText = sText & ListView1.ListItems(lItemCount).ListSubItems(lSubItemCount) & vbTab
    Next
    '
    sText = sText & vbCrLf
  Next
  '
  'MsgBox sText
  Clipboard.SetText sText, vbCFText
  '
  Screen.MousePointer = vbDefault
End Sub

'Private Sub DateSet1_Click()
'  On Error GoTo EH
'  '
'  Me.Date1.Caption = FDatePick.DateText(Date1.Caption)
'  '
'  Exit Sub
'EH:
' MsgBox Err.Description & " in FReport.DateSet1_Click."
'End Sub

'Private Sub DateSet2_Click()
'  On Error GoTo EH
'  '
'  Me.Date2.Caption = FDatePick.DateText(Date2.Caption)
'  '
'  Exit Sub
'EH:
' MsgBox Err.Description & " in FReport.DateSet2_Click."
'End Sub

Private Sub Form_Activate()
  On Error GoTo EH
  '
  frame1.Move (Width - frame1.Width) / 2
  frame1.Visible = True
  Exit Sub
EH:
 MsgBox Err.Description & " in FReport.Form_Activate."
End Sub

'
Private Sub Form_Load()
  On Error GoTo EH
  '
  iLastKey = 1
  '
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.Setup Me, False, , , "Contact Reports"
  '
  'Date1.Caption = Date
  'Date2.Caption = Date
  'optChoice(0) = True
  '
  ReadPreferences
  '
  SetupGroups
  SelectGroup
  '
'  textNotes.AddItem "(Close 1)"
'  textNotes.AddItem "(Close 2)"
'  textNotes.AddItem "(Close 3)"
'  textNotes.AddItem "(Close 4)"
'  textNotes.AddItem "(Close 5)"
'  textNotes.AddItem "(Tech)"
  '
  Dim rsStatus As Recordset
  Set rsStatus = New ADODB.Recordset
  rsStatus.Open "SELECT * FROM tblStatus", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  ComboStatus.AddItem "Everyone"
 '
  Do While Not rsStatus.eof
    ComboStatus.AddItem "" & rsStatus!Status
    rsStatus.MoveNext
  Loop
  Set rsStatus = Nothing
  '
  FillProductBox
'
  StateCombo.AddItem "All"
  StateCombo.AddItem "AL"
  StateCombo.AddItem "AK"
  StateCombo.AddItem "AZ"
  StateCombo.AddItem "AR"
  StateCombo.AddItem "CA"
  StateCombo.AddItem "CO"
  StateCombo.AddItem "CT"
  StateCombo.AddItem "DE"
  StateCombo.AddItem "DC"
  StateCombo.AddItem "FL"
  StateCombo.AddItem "GA"
  StateCombo.AddItem "HI"
  StateCombo.AddItem "ID"
  StateCombo.AddItem "IL"
  StateCombo.AddItem "IN"
  StateCombo.AddItem "IA"
  StateCombo.AddItem "KS"
  StateCombo.AddItem "KY"
  StateCombo.AddItem "LA"
  StateCombo.AddItem "ME"
  StateCombo.AddItem "MD"
  StateCombo.AddItem "MA"
  StateCombo.AddItem "MI"
  StateCombo.AddItem "MN"
  StateCombo.AddItem "MS"
  StateCombo.AddItem "MO"
  StateCombo.AddItem "MT"
  StateCombo.AddItem "NE"
  StateCombo.AddItem "NV"
  StateCombo.AddItem "NH"
  StateCombo.AddItem "NJ"
  StateCombo.AddItem "NM"
  StateCombo.AddItem "NY"
  StateCombo.AddItem "NC"
  StateCombo.AddItem "ND"
  StateCombo.AddItem "OH"
  StateCombo.AddItem "OK"
  StateCombo.AddItem "OR"
  StateCombo.AddItem "PA"
  StateCombo.AddItem "PR"
  StateCombo.AddItem "RI"
  StateCombo.AddItem "SC"
  StateCombo.AddItem "SD"
  StateCombo.AddItem "TN"
  StateCombo.AddItem "TX"
  StateCombo.AddItem "UT"
  StateCombo.AddItem "VT"
  StateCombo.AddItem "WA"
  StateCombo.AddItem "WV"
  StateCombo.AddItem "WI"
  StateCombo.AddItem "WY"
  '
  LoadcboAction
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in FReport.Form_Load."
End Sub

Public Sub ReadPreferences()
  '*
  '
  On Error GoTo ErrHndlr
  '
  'fSettingPrefs = True
  '
  '\\ Filter
  Filter = GetSetting(App.Title, "ProspectMgt", "FilterGroups", FilterAll)
  '
  Dim iChoice As Integer
  '
  iChoice = nnNum(GetSetting(App.Title, "Reports", "TypeSelect", "0"))
  '
  optChoice(iChoice).Value = True
  '
  If Filter = FilterAll Then
    chkFilter(0).Value = vbChecked
    chkFilter(1).Value = vbChecked
  Else
    chkFilter(0).Value = IIf(Filter = FilterStandard, vbChecked, vbUnchecked)
    chkFilter(1).Value = IIf(Filter = FilterAMBest, vbChecked, vbUnchecked)
  End If
  '
  '\\ Totals
  fShowTotals = GetSetting(App.Title, "ProspectMgt", "ShowTotals", "False")
  chkTtls.Value = IIf(fShowTotals, vbChecked, vbUnchecked)
  If GetSetting(App.Title, "ProspectMgt", "SortColumn", 0) = 0 Then
    chkAlpha.Value = vbUnchecked
  Else
    chkAlpha.Value = vbChecked
  End If
  '
'  '\\ Sort
'  SortColumn = GetSetting(App.Title, "ProspectMgt", "SortColumn", SortByGroup)
'  '
'  If SortColumn = SortByGroup Then
'    optSort(0).value = True
'  Else
'    optSort(1).value = True
'  End If
  '
  'fSettingPrefs = False
  '
  Exit Sub
  '
ErrHndlr:
  fSettingPrefs = False
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.General.ReadPreferences", vbCritical, "Error"
End Sub

Private Sub SetupData()
  On Error GoTo EH
  '
  Dim DaysMaxTemp As Integer
  Dim DaysMinTemp As Integer
  '
  Report.SortDirection = GetSortDirection
  Report.SortField = GetSortField
  '
  Report.ProductID = Product.GetProductID(cboProduct.Text)
  Report.FirstName = Me.TextFirstName.Text
  Report.LastName = Me.TextLastName.Text
  Report.Company = Me.TextCompany.Text
  Report.Status = Me.ComboStatus.Text
  Report.City = Me.TextCity.Text
  Report.Zip = Me.TextZip.Text
  Report.State = Me.StateCombo.Text
  Report.Source = Me.TextSource.Text
  '
  If Val(Me.TextDaysMin.Text) > 10000 Or Val(Me.TextDaysMin.Text) < -10000 Then
    DaysMinTemp = 0
  Else
    DaysMinTemp = Val(Me.TextDaysMin.Text)
  End If
  '
  If Val(Me.TextDaysMax.Text) > 10000 Or Val(Me.TextDaysMax.Text) < -10000 Then
    DaysMaxTemp = 0
  Else
    DaysMaxTemp = Val(Me.TextDaysMax.Text)
  End If
  '
  If optChoice(0).Value = True Then
    Report.Rtype = SimpleContact
  End If
  If optChoice(1).Value = True Then
    Report.DaysMax = DaysMaxTemp ',Val(Me.TextDaysMax.Text)
    Report.DaysMin = DaysMinTemp 'Val(Me.TextDaysMin.Text)
    Report.Rtype = DaysRemaining
  End If
  If optChoice(2).Value = True Then
    Report.DaysMin = DaysMinTemp 'Val(Me.TextDaysMin.Text)
    Report.DaysMax = DaysMaxTemp 'Val(Me.TextDaysMax.Text)
    Report.Rtype = DaysNotAuth
  End If
  If optChoice(3).Value = True Then
      Report.Notes = Me.textNotes.Text
      Report.Rtype = NotesSearch
  End If
  If optChoice(4).Value = True Then
    Report.DaysMin = Val(Me.txtDays.Text)
    Report.ActionType = cboAction.Text
    Report.Rtype = NoContactIn
  End If
  If optChoice(5).Value = True Then
    Report.ListQuery = GetGroupQuery
    Report.Rtype = FromList
  End If
  Exit Sub
EH:
 MsgBox Err.Description & " in FReport.SetupData."
End Sub

Private Sub Form_Resize()
  On Error GoTo EH
  '
  frame1.Move (Width - frame1.Width) / 2
  Exit Sub
EH:
 MsgBox Err.Description & " in FReport.Form_Resize."
End Sub

Private Sub grdHistory_DblClick()
  Load FResult
  FResult.TextResult.Text = grdHistory.Columns(7).Value
  FResult.Show vbModal
  Exit Sub
End Sub

Private Sub ListView1_Click()
  On Error GoTo EH
  '
  ListView1.ListItems(iLastKey).ForeColor = vbBlack
  If optChoice(5).Value = True Then
    LoadHistory nnNum(GetCurrentContactID)
  End If
  ListView1.SelectedItem.ForeColor = vbBlue
  iLastKey = ListView1.SelectedItem.Index
  Exit Sub
EH:
End Sub

Private Sub ListView1_DblClick()
  On Error GoTo EH
  '
  FContact.LoadContact nnNum(GetCurrentContactID), True
  'Company.Fetch nnNum(GetCurrentCompanyID)
  'Company.Contact.Fetch nnNum(GetCurrentContactID)
  FormMgr.ShowForm FMain.ActiveForm, FContact, True
  'FormMgr.ShowForm Me, FContact
  Exit Sub
EH:
 MsgBox Err.Description & " in FReport.ListView1_DblClick."
End Sub

Private Function GetGroupQuery() As String
  Dim sSQL As String
  '
  Screen.MousePointer = vbHourglass
  DoEvents
  '
  rsGroupCategories.MoveFirst
  rsGroupCategories.Find "RecID = " & lstGroups.ItemData(lstGroups.ListIndex), , adSearchForward
  '
  If Not rsGroupCategories.eof Then
  '
  sSQL = "SELECT  TCompany.ID AS CompanyID, TContact.ID AS ContactID, " & _
    "TCompany.Name AS Company " & _
    ",TContact.FirstName" & _
    ",TContact.LastName" & _
     ",TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & _
    ",TContact.State, TContact.Status, TContact.AuthDate, TContact.Phone1 " & _
    ",TContact.VersionShipped, TContact.PVVersionShipped " & _
    ", TContact.Source, DATEADD(day,TContact.AuthRemaining,TContact.AuthDate) AS ExpDate" & _
    " FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
    "WHERE " & ConvertFormula(rsGroupCategories!Formula) '& " ORDER BY TContact.State"
    '
    'MsgBox sSQL
    '", TContact.ShipStatus, TContact.AuthStatus, " & _
    " TContact.AuthDays, TContact.ShipDate "
   GetGroupQuery = sSQL
  End If
End Function

Private Sub lstGroups_Click()
  If lstGroups.Enabled = True Then
    ShowResults_Click
    grdHistory.RemoveAll
  End If
  'optChoice(5).value = True
End Sub

Private Sub optChoice_Click(Index As Integer)
    '
    SaveSetting App.Title, "Reports", "TypeSelect", Index
    '
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    '
    TextDaysMin.Enabled = False
    TextDaysMax.Enabled = False
    textNotes.Enabled = False
    'DateSet1.Enabled = False
    'DateSet2.Enabled = False
    lstGroups.Enabled = False
    chkAlpha.Enabled = False
    chkFilter(0).Enabled = False
    chkFilter(1).Enabled = False
    chkTtls.Enabled = False
    frmCriteria.Visible = True
    ListView1.Height = 5295 '4695
    grdHistory.Visible = False
    txtDays.Enabled = False
    cboAction.Enabled = False
    TextFirstName.Enabled = False
    TextLastName.Enabled = False
    TextCity.Enabled = False
    StateCombo.Enabled = False
    TextZip.Enabled = False
    TextSource.Enabled = False
    'txtDays.BackColor = &H80000011
    'cboAction.BackColor = &H80000011
    TextFirstName.BackColor = &H80000011
    TextLastName.BackColor = &H80000011
    TextCity.BackColor = &H80000011
    StateCombo.BackColor = &H80000011
    TextZip.BackColor = &H80000011
    TextSource.BackColor = &H80000011

  If optChoice(0).Value = True Then
    TextDaysMin.Enabled = True
    TextDaysMax.Enabled = True
    TextFirstName.Enabled = True
    TextLastName.Enabled = True
    TextCity.Enabled = True
    StateCombo.Enabled = True
    TextZip.Enabled = True
    TextSource.Enabled = True
    txtDays.BackColor = &H80000005
    cboAction.BackColor = &H80000005
    TextFirstName.BackColor = &H80000005
    TextLastName.BackColor = &H80000005
    TextCity.BackColor = &H80000005
    StateCombo.BackColor = &H80000005
    TextZip.BackColor = &H80000005
    TextSource.BackColor = &H80000005
  End If
  If optChoice(1).Value = True Then
    TextDaysMin.Enabled = True
    TextDaysMax.Enabled = True
    TextFirstName.Enabled = True
    TextLastName.Enabled = True
    TextCity.Enabled = True
    StateCombo.Enabled = True
    TextZip.Enabled = True
    TextSource.Enabled = True
    txtDays.BackColor = &H80000005
    cboAction.BackColor = &H80000005
    TextFirstName.BackColor = &H80000005
    TextLastName.BackColor = &H80000005
    TextCity.BackColor = &H80000005
    StateCombo.BackColor = &H80000005
    TextZip.BackColor = &H80000005
    TextSource.BackColor = &H80000005
  End If
  If optChoice(2).Value = True Then
    TextDaysMin.Enabled = True
    TextDaysMax.Enabled = True
    TextFirstName.Enabled = True
    TextLastName.Enabled = True
    TextCity.Enabled = True
    StateCombo.Enabled = True
    TextZip.Enabled = True
    TextSource.Enabled = True
    txtDays.BackColor = &H80000005
    cboAction.BackColor = &H80000005
    TextFirstName.BackColor = &H80000005
    TextLastName.BackColor = &H80000005
    TextCity.BackColor = &H80000005
    StateCombo.BackColor = &H80000005
    TextZip.BackColor = &H80000005
    TextSource.BackColor = &H80000005
  End If
  If optChoice(3).Value = True Then
'    TextDaysMin.Enabled = False
'    TextDaysMax.Enabled = False
    textNotes.Enabled = True
    TextFirstName.Enabled = True
    TextLastName.Enabled = True
    TextCity.Enabled = True
    StateCombo.Enabled = True
    TextZip.Enabled = True
    TextSource.Enabled = True
    txtDays.BackColor = &H80000005
    cboAction.BackColor = &H80000005
    TextFirstName.BackColor = &H80000005
    TextLastName.BackColor = &H80000005
    TextCity.BackColor = &H80000005
    StateCombo.BackColor = &H80000005
    TextZip.BackColor = &H80000005
    TextSource.BackColor = &H80000005
  End If
  If optChoice(4).Value = True Then
    txtDays.Enabled = True
    cboAction.Enabled = True
  End If
  '
  If optChoice(5).Value = True Then
    grdHistory.Visible = True
    grdHistory.RemoveAll
    'ListView1.Height = 6500
    frmCriteria.Visible = False
    lstGroups.Enabled = True
    chkAlpha.Enabled = True
    chkFilter(0).Enabled = True
    chkFilter(1).Enabled = True
    chkTtls.Enabled = True
  End If
  grdHistory.RemoveAll
End Sub

Private Sub PreviewReport_Click()
  On Error GoTo EH
  '
  Screen.MousePointer = vbHourglass
  SetupData
  'If optChoice(4).Value = True Then
  '  Report.PreviewReport ("Frontier")
  'End If
  If (optChoice(1).Value = True) Or (optChoice(2).Value = True) Then
    Report.PreviewReport ("DaysLeft")
  End If
  If (optChoice(0).Value = True) Or (optChoice(3).Value = True) Or (optChoice(5).Value = True) Then
    Report.PreviewReport ("Simple Contact")
  End If
  Screen.MousePointer = vbDefault
  Exit Sub
EH:
 MsgBox Err.Description & " in FReport.PreviewReport_Click."
End Sub

Private Sub PrintButton_Click()
  Dim X As Integer
  On Error GoTo EH
  If GetPrinter = True Then Exit Sub
  '
  Screen.MousePointer = vbHourglass
  SetupData
  ListView1.ListItems.Clear
  ListView1.ColumnHeaders.Clear
  Report.FillList Report.rsReport, ListView1
  '
  'new code'''''''''''''''''''''''''''
  'Instantiate the PrintListView class
  Set objLvPrint = New clsPrintLV
  '
  For X = 1 To iNumofCopies
    'Call the Print command from the class
    objLvPrint.PrintListView ListView1, 0.1, 8, "ListView Report", Landscape, True, False
  Next X
  '
  'Destroy the object
  Set objLvPrint = Nothing
  '
  'End Code'''''''''''''''''''''''''''
  '
'  'If (optChoice(4).Value = True) = True Then
'  '  Report.PrintReport ("Frontier")
'  'End If
'  If (optChoice(1).Value = True) Or (optChoice(2).Value = True) Then
'    Report.PrintReport ("DaysLeft")
'  End If
'  If (optChoice(0).Value = True) Or (optChoice(3).Value = True) Or (optChoice(5).Value = True) Then
'    Report.PrintReport ("Simple Contact")
'  End If
  Screen.MousePointer = vbDefault
  Exit Sub
EH:
 MsgBox Err.Description & " in FReport.PrintButton_Click."
End Sub

Private Sub ShowResults_Click()
  On Error GoTo EH
  '
  Screen.MousePointer = vbHourglass
  SetupData
  ListView1.ListItems.Clear
  ListView1.ColumnHeaders.Clear
  Report.FillList Report.rsReport, ListView1
  LabelResults.Caption = Report.rsReport.RecordCount & " Results"
  Screen.MousePointer = vbDefault
  Exit Sub
EH:
 MsgBox Err.Description & " in FReport.ShowResults_Click."
End Sub

Private Function GetCurrentCompanyID() As String
  On Error GoTo EH
  '
  Dim sTemp As String
  Dim sTempChar As String
  Dim iCount As Integer
  Dim iLength As Integer
  iLength = Len(ListView1.SelectedItem.Key)
  Do
    iCount = iCount + 1
    sTempChar = Mid(ListView1.SelectedItem.Key, iCount, 1)
    sTemp = sTemp + sTempChar
  Loop While (sTempChar <> "A") And (iCount <= Len(ListView1.SelectedItem.Key))
  iLength = Len(sTemp) - 1
  GetCurrentCompanyID = Mid(sTemp, 1, iLength)
  '
  Exit Function
EH:
 MsgBox Err.Description & " in FReport.GetCurrentCompanyID."
End Function

Private Function GetCurrentContactID() As String
  On Error GoTo EH
  '
  Dim sTemp As String
  Dim sTempChar As String
  Dim iCount As Integer
  Dim iLength As Integer
  iLength = Len(ListView1.SelectedItem.Key)
  Do
    iCount = iCount + 1
    sTempChar = Mid(ListView1.SelectedItem.Key, iCount, 1)
    sTemp = sTemp + sTempChar
  Loop While (sTempChar <> "A") And (iCount <= Len(ListView1.SelectedItem.Key))
  '
  sTemp = vbNullString
  Do
    iCount = iCount + 1
    sTempChar = Mid(ListView1.SelectedItem.Key, iCount, 1)
    sTemp = sTemp + sTempChar
  Loop While (iCount <= iLength)
  GetCurrentContactID = sTemp
  '
  Exit Function
EH:
 MsgBox Err.Description & " in FReport.GetCurrentContactID."
End Function

Private Sub FillProductBox()
  Dim Products As New CProducts
  Dim i As Integer
  '
  cboProduct.Clear
  '
  Product.LoadCollection Products
  '
  'cboProduct.AddItem "All Products"
  '
  For i = 1 To Products.Count
    cboProduct.AddItem Products(i).Product
  Next i
  '
  cboProduct.ListIndex = 0
  '
  Set Products = Nothing
End Sub

Public Sub SetupGroups()
  '*
  '
  On Error GoTo ErrHndlr
  '
  Dim rsGroups As ADODB.Recordset
  '
  Screen.MousePointer = vbHourglass
  lstGroups.Clear
  '
  Set rsGroupCategories = New ADODB.Recordset
  '
  'SortColumn = GetSetting(App.Title, "ProspectMgt", "SortColumn", SortByGroup)
  '
'  If Filter = FilterAll Then
'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] ORDER BY [Priority]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
'  ElseIf Filter = FilterStandard Then
'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] NOT LIKE 'AM Best%' ORDER BY [Priority]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
'  ElseIf Filter = FilterAMBest Then
'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] LIKE 'AM Best%' ORDER BY [Priority]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
'  ElseIf Filter = FilterNone Then
'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] = 'Dummy'", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
'  End If
  '
  Dim sOrderField As String
  '
  If chkAlpha.Value = vbChecked Then
    sOrderField = "Label"
  Else
    sOrderField = "Priority"
  End If
  If Filter = FilterAll Then
    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] ORDER BY [" & sOrderField & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  ElseIf Filter = FilterStandard Then
    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] NOT LIKE 'AM Best%' ORDER BY [" & sOrderField & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  ElseIf Filter = FilterAMBest Then
    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] LIKE 'AM Best%' ORDER BY [" & sOrderField & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  ElseIf Filter = FilterNone Then
    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] = 'Dummy'", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  End If
  '
  Set rsGroups = New ADODB.Recordset
  '
  With rsGroupCategories
    Do While .eof = False
      If rsGroups.State = adStateOpen Then rsGroups.Close
      If fShowTotals Then
        If ConnType = Access Then
          rsGroups.Open "SELECT Count(*) as GroupCount FROM QProspectMgt WHERE " & _
            .Fields("Formula").Value, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
          '
        Else
          Dim sSQL As String
          '
          sSQL = "SELECT Count(*) as GroupCount " & _
            "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
            "WHERE " & ConvertFormula(.Fields("Formula").Value)
          '
          rsGroups.Open sSQL, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
        End If
        '
        lstGroups.AddItem .Fields("Label").Value & " (" & rsGroups.Fields("GroupCount").Value & ")"
      Else
        lstGroups.AddItem .Fields("Label").Value
      End If
      lstGroups.ItemData(lstGroups.NewIndex) = .Fields("RecID").Value
      .MoveNext
    Loop
  End With
  '
  Screen.MousePointer = vbDefault
  '
  Exit Sub
  '
ErrHndlr:
  Screen.MousePointer = vbDefault
  DoEvents
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmReports.SetupGroups.", vbCritical, "Error"
End Sub

Private Function ConvertFormula(ByVal psFormula As String) As String
  On Error GoTo EH
  '
  Dim sFormula As String
  '
  sFormula = psFormula
  '
  sFormula = Replace(sFormula, "ShipDays", "DateDiff(Day,[ShipDate],GETDATE())")
  sFormula = Replace(sFormula, "AuthDaysRemaining", "([AuthDays] - DateDiff(Day, [AuthDate], GETDATE()))")
  sFormula = Replace(sFormula, "isnull(shipdate)", "(ShipDate = Null)")
  '
  ConvertFormula = sFormula
  '
  Exit Function
EH:
  MsgBox Err.Description & " in Convert Formula.)"
End Function

Private Sub chkFilter_Click(Index As Integer)
  '*
  '
  On Error GoTo ErrHndlr
  '
  If fSettingPrefs = True Then Exit Sub
  '
  If chkFilter(0).Value = vbChecked And chkFilter(1).Value = vbChecked Then
    Filter = FilterAll
  ElseIf chkFilter(0).Value = vbChecked Then
    Filter = FilterStandard
  ElseIf chkFilter(1).Value = vbChecked Then
    Filter = FilterAMBest
  Else
    Filter = FilterNone
  End If
  '
  SaveSetting App.Title, "ProspectMgt", "FilterGroups", Filter
  '
  SetupGroups
  SelectGroup
  '
  Exit Sub
  '
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.chkFilter.Click", vbCritical, "Error"
End Sub

Private Sub chkTtls_Click()
  '*
  '
  On Error GoTo ErrHndlr
  '
  If fSettingPrefs Then Exit Sub
  '
'  If chkTtls.value = vbChecked Then
'    MsgBox "WARNING: This option is fast as dirt. You've been warned.", _
'           vbInformation, "I Wouldn't Do That If I Were You"
'  End If
  '
  fShowTotals = IIf(chkTtls.Value = vbChecked, True, False)
  SaveSetting App.Title, "ProspectMgt", "ShowTotals", fShowTotals
  '
  SetupGroups
  SelectGroup
  '
  Exit Sub
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.chkTtls.Click", vbCritical, "Error"
End Sub

Public Sub SelectGroup()
  '*
  '
  On Error GoTo ErrHndlr
  '
  '\\ Local Declarations
  Dim iCur      As Integer
  Dim iCt       As Integer
  '
  lProspectGroupID = GetSetting(App.Title, "ProspectMgt", "CurrentGroupRecID", 0)
  '
  With lstGroups
    iCt = lstGroups.ListCount - 1
    If iCt < 0 Then Exit Sub
    '
    For iCur = 0 To iCt
      If .ItemData(iCur) = lProspectGroupID Then
        .ListIndex = iCur
        Exit Sub
      End If
    Next
    '
    .ListIndex = 0
  End With
  '
  Exit Sub
  '
ErrHndlr:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmProspectMgt.General.SelectGroup", vbCritical, "Error"
End Sub

Private Function GetSortField() As String
  If ListView1.ColumnHeaders.Count >= (ListView1.SortKey + 1) Then
    GetSortField = ListView1.ColumnHeaders(ListView1.SortKey + 1).Text
  Else
    GetSortField = vbNullString
  End If
End Function

Private Function GetSortDirection() As Integer
  With ListView1
    If .SortOrder = lvwAscending Then
      GetSortDirection = 0
    Else
      GetSortDirection = 1
    End If
  End With
End Function

Private Sub SortListView(ByVal lvwCur As MSComctlLib.ListView, ByVal colHdr As MSComctlLib.ColumnHeader, Optional ByVal sSortOrder As String)
  On Error GoTo ErrorHandler
  '
  With lvwCur
    '
    'If .SortKey > -1 Then .ColumnHeaders.Item(.SortKey + 1).Icon = 0
    '
    .SortKey = colHdr.Index - 1
    '
    If sSortOrder <> vbNullString Then
      .SortOrder = IIf(sSortOrder = "Ascending", lvwAscending, lvwDescending)
    Else
      .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    End If
    '
    .Sorted = True
    '
    '.ColumnHeaders.Item(colHdr.Index).Icon = IIf(.SortOrder = lvwAscending, "imgAscending", "imgDescending")
    '
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox Error((Err.Number)), vbCritical + vbOKOnly, "Error: FReports.SortListView"
End Sub
  
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  On Error GoTo ErrorHandler
  '
  SortListView ListView1, ColumnHeader
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.ListView1.ColumnClick"
End Sub

Private Sub LoadHistory(ByVal plCustomerID As Long)
  '*
  '
  On Error GoTo EH
  '
  grdHistory.Redraw = False
  '
  Me.grdHistory.RemoveAll
  '
  If plCustomerID <> -1 Then
    Dim rsHistory As ADODB.Recordset
    '
    Set rsHistory = New ADODB.Recordset
    '
    If ConnType = Access Then
      rsHistory.Open "SELECT * FROM tblSupportActs WHERE CustRecID = " & plCustomerID & " ORDER BY Date DESC, Time DESC", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
      Dim cmdHistory As New ADODB.Command
      '
      With cmdHistory
        Set .ActiveConnection = cnMain
        .CommandText = "dbo.UpParmSelSupportActs"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("CustomerID", adInteger, adParamInput, , plCustomerID)
        Set rsHistory = .Execute
      End With 'cmdSupportAct
      '
      Set cmdHistory = Nothing
    End If
    '
    With rsHistory
    Do While Not .eof
      grdHistory.AddItem !RecID & vbTab & !CustRecID & vbTab & !Date & vbTab & !Time & vbTab & !Type & vbTab & !User & vbTab & !Subject & vbTab & !Results
      .MoveNext
    Loop
    End With 'rsHistory
    '
    rsHistory.Close
    Set rsHistory = Nothing
  End If
  '
  grdHistory.Redraw = True
  '
  Exit Sub
EH:
  grdHistory.Redraw = True
  MsgBox Err.Description
End Sub

Private Function GetPrinter() As Boolean
  FPrinterSelect.Show vbModal
  GetPrinter = FPrinterSelect.bPrintCancel
End Function

Private Sub LoadcboAction()
  Dim rsType As New Recordset
  '
  rsType.Open "SELECT * FROM tblActivities WHERE ActivityType = 1 ORDER BY Activity", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rsType.eof
    cboAction.AddItem rsType!Activity & vbNullString
    rsType.MoveNext
  Loop
  '
  rsType.Close
  '
  'cboAction.AddItem vbNullString
  '
  rsType.Open "SELECT * FROM tblActivities WHERE ActivityType = 0 ORDER BY Activity", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rsType.eof
    cboAction.AddItem rsType!Activity & vbNullString
    rsType.MoveNext
  Loop
  '
  rsType.Close
  '
  Set rsType = Nothing
  '
  cboAction.ListIndex = 1
End Sub

