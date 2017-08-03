VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FHours 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Employee Hours"
   ClientHeight    =   6855
   ClientLeft      =   3465
   ClientTop       =   3330
   ClientWidth     =   11385
   Icon            =   "FHours.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.TextBox txtHours 
         Height          =   315
         Left            =   9660
         TabIndex        =   8
         Top             =   4620
         Width           =   1095
      End
      Begin VB.TextBox txtLunch 
         Height          =   315
         Left            =   9660
         TabIndex        =   7
         Top             =   4980
         Width           =   1095
      End
      Begin VB.TextBox txtTotalHours 
         Height          =   315
         Left            =   9660
         TabIndex        =   6
         Top             =   5340
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalcHours 
         Caption         =   "Calculate Hours"
         Height          =   315
         Left            =   9120
         TabIndex        =   5
         Top             =   560
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   5040
         Width           =   1395
      End
      Begin VB.CommandButton cmdActual 
         Caption         =   "Hide/Unhide Actual Times"
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Top             =   5100
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkShow 
         BackColor       =   &H00008000&
         Caption         =   "Show Actual Log Times"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   1
         Top             =   560
         Width           =   3015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGridHours 
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         Top             =   975
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   6165
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   49152
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSAdodcLib.Adodc dcHours 
         Height          =   330
         Left            =   3960
         Tag             =   "tblHours"
         Top             =   4620
         Visible         =   0   'False
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Hours"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin SSDataWidgets_B.SSDBCombo cmbEmployee 
         Height          =   315
         Left            =   3420
         TabIndex        =   9
         Top             =   540
         Width           =   2355
         DataFieldList   =   "Column 0"
         _Version        =   196617
         DataMode        =   2
         Cols            =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnHeaders   =   0   'False
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   4154
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   0
         BackColor       =   16777215
      End
      Begin TDBDate6Ctl.TDBDate mskEndDate 
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   540
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FHours.frx":030A
         Caption         =   "FHours.frx":0422
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FHours.frx":048E
         Keys            =   "FHours.frx":04AC
         Spin            =   "FHours.frx":050A
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "02/10/2003"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   37662
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate mskBeginDate 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   540
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FHours.frx":0532
         Caption         =   "FHours.frx":064A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FHours.frx":06B6
         Keys            =   "FHours.frx":06D4
         Spin            =   "FHours.frx":0732
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "02/10/2003"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   37662
         CenturyMode     =   0
      End
      Begin Threed.SSCommand cmdEndDate 
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   540
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FHours.frx":075A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdBeginDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   540
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FHours.frx":0CF4
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Begin Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "End Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Employee:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3420
         TabIndex        =   17
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "Hours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8580
         TabIndex        =   16
         Top             =   4680
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
         Caption         =   "Lunch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8580
         TabIndex        =   15
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00008000&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8580
         TabIndex        =   14
         Top             =   5400
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FHours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
'
Private rsLog As ADODB.Recordset
Private rsEmployee As ADODB.Recordset
Attribute rsEmployee.VB_VarHelpID = -1
Private rsHours As ADODB.Recordset

Private Sub cmbEmployee_InitColumnProps()
  On Error GoTo ErrCall
  '
  Set rsEmployee = New ADODB.Recordset
  '
  If ConnType = Access Then
    rsEmployee.Open "SELECT EmployeeFirst & ' ' & EmployeeLast AS EmployeeName FROM tblEmployees ORDER BY EmployeeFirst, EmployeeLast", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  Else
    rsEmployee.Open "SELECT EmployeeFirst + ' ' + EmployeeLast AS EmployeeName FROM tblEmployees ORDER BY EmployeeFirst, EmployeeLast", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  End If
  '
  With rsEmployee
    Do Until .eof
      cmbEmployee.AddItem !EmployeeName
      .MoveNext
    Loop
  End With
  '
  Exit Sub
  '
ErrCall:
  MsgBox Err.Description
End Sub

Private Sub cmdActual_Click()
'  Me.grdHours.Columns("ActualIn").Visible = Not (Me.grdHours.Columns("ActualIn").Visible)
'  Me.grdHours.Columns("ActualOut").Visible = Not (Me.grdHours.Columns("ActualOut").Visible)
End Sub

Private Sub cmdCalcHours_Click()
  Dim X As Integer
  '
  GetEmployeeLog
  '
  If chkShow.value = 1 Then
    For X = 0 To 5
      FlexGridHours.ColWidth(X) = ((FlexGridHours.Width - 1180) / 6) - 8
    Next
    FlexGridHours.ColWidth(6) = 1080
  Else
    For X = 0 To 3
      FlexGridHours.ColWidth(X) = ((FlexGridHours.Width - 1180) / 4) - 15
    Next
    FlexGridHours.ColWidth(4) = 1080
  End If
End Sub

Private Sub cmdPrint_Click()
  Dim printDlg As PrinterDlg
  Set printDlg = New PrinterDlg
  ' Set the starting information for the dialog box based on the current
  ' printer settings.
  printDlg.PrinterName = Printer.DeviceName
  printDlg.DriverName = Printer.DriverName
  printDlg.Port = Printer.Port
  
  ' Set the default PaperBin so that a valid value is returned even
  ' in the Cancel case.
  printDlg.PaperBin = Printer.PaperBin
  
  ' Set the flags for the PrinterDlg object using the same flags as in the
  ' common dialog control. The structure starts with VBPrinterConstants.
  printDlg.FLAGS = VBPrinterConstants.cdlPDNoSelection _
                   Or VBPrinterConstants.cdlPDNoPageNums _
                   Or VBPrinterConstants.cdlPDReturnDC
  Printer.TrackDefault = False
  
  ' When CancelError is set to True the ShowPrinterDlg will return error
  ' 32755. You can handle the error to know when the Cancel button was
  ' clicked. Enable this by uncommenting the lines prefixed with "'**".
  '**printDlg.CancelError = True
  
  ' Add error handling for Cancel.
  '**On Error GoTo Cancel
  If Not printDlg.ShowPrinter(Me.hWnd) Then
      Debug.Print "Cancel Selected"
      Exit Sub
  End If
  
  'Turn off Error Handling for Cancel.
  '**On Error GoTo 0
  Dim NewPrinterName As String
  Dim objPrinter As Printer
  Dim strsetting As String
  
  ' Locate the printer that the user selected in the Printers collection.
  NewPrinterName = UCase$(printDlg.PrinterName)
  If Printer.DeviceName <> NewPrinterName Then
      For Each objPrinter In Printers
         If UCase$(objPrinter.DeviceName) = NewPrinterName Then
              Set Printer = objPrinter
         End If
      Next
  End If
  
  ' Copy user input from the dialog box to the properties of the selected printer.
  Printer.Copies = printDlg.Copies
  Printer.Orientation = printDlg.Orientation
  Printer.ColorMode = printDlg.ColorMode
  Printer.Duplex = printDlg.Duplex
  Printer.PaperBin = printDlg.PaperBin
  Printer.PaperSize = printDlg.PaperSize
  Printer.PrintQuality = printDlg.PrintQuality
  
  ' Display the results in the immediate (Debug) window.
  ' NOTE: Supported values for PaperBin and Size are printer specific. Some
  ' common defaults are defined in the Win32 SDK in MSDN and in Visual Basic.
  ' Print quality is the number of dots per inch.
  With Printer
      Debug.Print .DeviceName
      If .Orientation = 1 Then
          strsetting = "Portrait. "
      Else
          strsetting = "Landscape. "
      End If
      Debug.Print "Copies = " & .Copies, "Orientation = " & _
         strsetting
      If .ColorMode = 1 Then
          strsetting = "Black and White. "
      Else
          strsetting = "Color. "
      End If
      Debug.Print "ColorMode = " & strsetting
      If .Duplex = 1 Then
          strsetting = "None. "
      ElseIf .Duplex = 2 Then
          strsetting = "Horizontal/Long Edge. "
      ElseIf .Duplex = 3 Then
          strsetting = "Vertical/Short Edge. "
      Else
          strsetting = "Unknown. "
      End If
      Debug.Print "Duplex = " & strsetting
      Debug.Print "PaperBin = " & .PaperBin
      Debug.Print "PaperSize = " & .PaperSize
      Debug.Print "PrintQuality = " & .PrintQuality
      If (printDlg.FLAGS And VBPrinterConstants.cdlPDPrintToFile) = _
         VBPrinterConstants.cdlPDPrintToFile Then
           Debug.Print "Print to File Selected"
      Else
           Debug.Print "Print to File Not Selected"
      End If
      Debug.Print "hDC = " & printDlg.hDC
  End With
  '
  Dim old_width As Integer
  '
  old_width = FlexGridHours.Width
  FlexGridHours.Width = Printer.Width
  Printer.PaintPicture FlexGridHours.Picture, 0, 0
  Printer.EndDoc
  FlexGridHours.Width = old_width
  '
  Exit Sub
  '**Cancel:
  '**If Err.Number = 32755 Then
  '**    Debug.Print "Cancel Selected"
  '**Else
  '**    Debug.Print "A nonCancel Error Occured - "; Err.Number
  '**End If
End Sub

Private Sub Form_Load()
  'On Error GoTo ErrCall
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.Setup Me, False, , , "Hours"
  '
  Dim rsEmployee2 As ADODB.Recordset
  '
  Set rsEmployee2 = New ADODB.Recordset
  '
  Set rsHours = New ADODB.Recordset
  '
  rsEmployee2.Open "SELECT Password FROM tblEmployees WHERE (EmployeeFirst + ' ' + EmployeeLast = '" & User.Name & "')", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  '
  Set rsLog = New ADODB.Recordset
  rsLog.Open "SELECT * FROM tblHours WHERE RecID = NULL", cnMain, adOpenKeyset, adLockOptimistic, adCmdText
  Set dcHours.Recordset = rsLog
  dcHours.Password = Rot39(DecryptStr(rsEmployee2!Password & ""))
  '
  mskBeginDate.value = Date
  mskEndDate.value = Date
  Exit Sub
  '
  Set rsEmployee2 = Nothing
ErrCall:
  MsgBox Err.Description
End Sub

Private Sub cmdBeginDate_Click()
  mskBeginDate.value = FDatePick.DateText(mskBeginDate.value)
End Sub

Private Sub cmdEndDate_Click()
  mskEndDate.value = FDatePick.DateText(mskEndDate.value)
End Sub

Private Sub GetEmployeeLog()
  'On Error GoTo ErrCall
  '
  With rsEmployee
  If Not .BOF Then .MoveFirst
  .Find "EmployeeName = '" & cmbEmployee.Text & "'", , adSearchForward
  '
  If Not .eof Then
    'grdHours.Redraw = False
    '
    If chkShow.value = 1 Then
      If ConnType = Access Then
        rsHours.Open "SELECT  Employee, LogDate, InTime, ActualIn, OutTime, ActualOut, Hours FROM tblHours WHERE (Employee = '" & cmbEmployee.Text & "') AND (LogDate >= #" & Trim(mskBeginDate.Text) & "# AND LogDate <= #" & Trim(mskEndDate.Text) & "#)", cnMain, adOpenKeyset, adLockOptimistic, adCmdText
      Else 'SQL Server
        rsHours.Open "SELECT  Employee, LogDate, InTime, ActualIn, OutTime, ActualOut, Hours FROM tblHours WHERE (Employee = '" & cmbEmployee.Text & "') AND (LogDate >= '" & Trim(mskBeginDate.Text) & "' AND LogDate <= '" & Trim(mskEndDate.Text) & "')", cnMain, adOpenKeyset, adLockOptimistic, adCmdText
      End If
    Else
      If ConnType = Access Then
        rsHours.Open "SELECT  Employee, LogDate, InTime, OutTime, Hours FROM tblHours WHERE (Employee = '" & cmbEmployee.Text & "') AND (LogDate >= #" & Trim(mskBeginDate.Text) & "# AND LogDate <= #" & Trim(mskEndDate.Text) & "#)", cnMain, adOpenKeyset, adLockOptimistic, adCmdText
      Else 'SQL Server
        rsHours.Open "SELECT  Employee, LogDate, InTime, OutTime, Hours FROM tblHours WHERE (Employee = '" & cmbEmployee.Text & "') AND (LogDate >= '" & Trim(mskBeginDate.Text) & "' AND LogDate <= '" & Trim(mskEndDate.Text) & "')", cnMain, adOpenKeyset, adLockOptimistic, adCmdText
      End If
    End If
    '
    'dcHours.Refresh
    '
    If Not (rsHours.eof And rsHours.BOF) Then rsHours.MoveFirst
    If Not rsHours.RecordCount = 0 Then
      Dim TotalHours As Single
      '
      rsHours.MoveFirst
      Do While Not rsHours.eof
        If IsNull(rsHours!outtime) Or IsNull(rsHours!intime) Then
          MsgBox "An entry does not contain both an in-time and out-time. Please correct this before calculating final hours."
        Else
          rsHours!hours = ConvertToHours(CDate(0 & rsHours!outtime) - CDate(0 & rsHours!intime))
          'dcHours.Recordset!hours = 6 'Format(dcHours.Recordset!hours, "h:mm")
          rsHours.Update
          TotalHours = TotalHours + rsHours!hours
        End If
        'dcHours.Recordset!LogDate = DatePart("yyyy", dcHours.Recordset!LogDate) & "/" & DatePart("m", dcHours.Recordset!LogDate) & "/" & DatePart("d", dcHours.Recordset!LogDate)
        '
        rsHours.MoveNext
      Loop
      '
      'dcHours.Refresh
      '
      txtHours = TotalHours
    End If
    '
    'grdHours.Redraw = True
  Else
    MsgBox "Employee not found."
  End If
  End With
  '
  rsHours.MoveFirst
  PopulateFlexGrid FlexGridHours, rsHours
  '
  rsHours.Close
  Exit Sub
  '
'ErrCall:
  'grdHours.Redraw = True
  'MsgBox Err.Description
End Sub

Private Sub Form_Resize()
  Frame1.Move (Width - Frame1.Width) / 2
End Sub

Private Sub txtHours_Change()
  On Error GoTo ErrCall
  '
  txtTotalHours = Val(txtHours) - Val(txtLunch)
  '
  Exit Sub
  '
ErrCall:
  MsgBox Err.Description
End Sub

Private Sub txtLunch_Change()
  On Error GoTo ErrCall
  '
  txtTotalHours = Val(txtHours) - Val(txtLunch)
  '
  Exit Sub
  '
ErrCall:
  MsgBox Err.Description
End Sub

Private Function ConvertToHours(HoursIn As Double) As Single
  On Error GoTo ErrCall
  '
  Dim intDays As Integer
  Dim sngHours As Single
  
  Dim intHours As Integer
  Dim intMinutes As Integer
  
  intDays = Int(HoursIn)
  sngHours = HoursIn - intDays
  
  intHours = Hour(sngHours)
  intMinutes = Minute(sngHours)
  
  ConvertToHours = Format((intDays * 24) + intHours + (intMinutes / 60), "fixed")
  '
  Exit Function
  '
ErrCall:
  MsgBox Err.Description
End Function

Public Function PopulateFlexGrid(FlexGrid As Object, _
   rs As Object) As Boolean
'*******************************************************
'PURPOSE: Populate MSFlexGrid with data from an
'         ADO Recordset
'PARAMETERS:  FlexGrid: MsFlexGrid to Populate
'             RS: Open ADO Recordset
'RETURNS:     True if successful, false otherwise
'REQUIRES:    -- Reference to Microsoft Active Data Objects
'             -- Recordset should be open with cursor set at
'                first row when passed and must
'                support recordcount property
'             -- FlexGrid should be empty when passed
'EXAMPLE:
'Dim conn As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'Dim sConnString As String
'
'sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDatabase.mdb"
'conn.Open sConnString
'rs.Open " SELECT * FROM MyTable", oConn, adOpenKeyset, adLockOptimistic
'PopulateFlexGrid MSFlexGrid1, rs
'
'rs.Close
'conn.Close
'***********************************************************
  On Error GoTo ErrorHandler
  '
  If Not TypeOf FlexGrid Is MSHFlexGrid Then Exit Function
  If Not TypeOf rs Is ADODB.Recordset Then Exit Function
  '
  Dim i As Integer
  Dim J As Integer
  '
  FlexGrid.FixedRows = 1
  FlexGrid.FixedCols = 0
  '
  If Not rs.eof Then
    '
    FlexGrid.Rows = rs.RecordCount + 1
    FlexGrid.Cols = rs.Fields.Count
    '
    For i = 0 To rs.Fields.Count - 1
      FlexGrid.TextMatrix(0, i) = rs.Fields(i).Name
    Next

    i = 1
    Do While Not rs.eof
      '
      For J = 0 To rs.Fields.Count - 1
        If Not IsNull(rs.Fields(J).value) Then
          FlexGrid.TextMatrix(i, J) = rs.Fields(J).value
        End If
      Next
    '
    i = i + 1
    rs.MoveNext
    Loop
  End If
  '
  PopulateFlexGrid = True
  '
  Exit Function
ErrorHandler:
  Exit Function
End Function

