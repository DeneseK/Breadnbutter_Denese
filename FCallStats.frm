VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FCallStats 
   BorderStyle     =   0  'None
   Caption         =   "Call / BNB Notes Statistics"
   ClientHeight    =   4935
   ClientLeft      =   3060
   ClientTop       =   2715
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ListView1 
      Height          =   975
      Left            =   480
      TabIndex        =   24
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1720
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cboMins 
      Height          =   315
      ItemData        =   "FCallStats.frx":0000
      Left            =   3240
      List            =   "FCallStats.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1130
      Width           =   735
   End
   Begin VB.CheckBox chkUpdate 
      Caption         =   "Auto Update every"
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   7680
      Top             =   2640
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCallData 
      Height          =   2175
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   20
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      HighLight       =   2
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtAvg 
      Height          =   285
      Left            =   2880
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optDuration 
      Caption         =   "Duration"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton optBNB 
      Caption         =   "BNB/Call"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame fraGroup 
      Caption         =   "Ext or Group"
      Height          =   975
      Left            =   5160
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
      Begin VB.ComboBox cboCallDir 
         Height          =   315
         ItemData        =   "FCallStats.frx":0031
         Left            =   2400
         List            =   "FCallStats.frx":003E
         TabIndex        =   14
         Text            =   "Both"
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cboGroup 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optExt 
         Caption         =   "by Ext"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCallType 
         AutoSize        =   -1  'True
         Caption         =   "Call Direction"
         Height          =   195
         Left            =   2400
         TabIndex        =   15
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdGenReport 
      Caption         =   "Run Report"
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Print Report"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   3240
      TabIndex        =   7
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69009409
      CurrentDate     =   37811
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69009409
      CurrentDate     =   37811
   End
   Begin VB.Label Label5 
      Caption         =   "Mins"
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Average Call Time"
      Height          =   195
      Left            =   2880
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Call Time"
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Call / BNB Notes Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   11
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      Caption         =   "Start Date"
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   480
      Width           =   720
   End
   Begin VB.Label lblEnd 
      AutoSize        =   -1  'True
      Caption         =   "End Date"
      Height          =   195
      Left            =   3240
      TabIndex        =   9
      Top             =   480
      Width           =   675
   End
End
Attribute VB_Name = "FCallStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
' global in other prog
Private strSQL As String, DTStart As Date, DTEnd As Date, intEXT As Integer, intWorkgroup As Integer, strDateType As String, strGroupType As String, strDirection As String
'Public chrtArray()
Private lGreatestValue As Long
Private iExt() As Integer
Private sName() As String
Private iNumofCalls() As Integer
Private iBNBNotes() As Integer
Private lCallTime() As Long
Private sCallTime() As String
Private iFollowups() As Integer
Private iWalkThroughs() As Integer
Private iSales() As Integer
Private lAvgCallTime() As Long
Private sAvgCallTime() As String
Private iNumofUsers As Integer
Private strSQL2 As String
Private iMins As Integer

Private Sub Form_Initialize()
  On Error GoTo ErrCall
  '
  Set FormControl = New CFormControl
  '
  FormControl.MinHeight = Me.Height
  FormControl.MinWidth = Me.Width
  FormControl.DataForm = False
  '
  cboMins.ListIndex = 0
  '
  optAll.Value = True
  DTPicker1.Value = Date - 7
  DTPicker2.Value = Date
  optBNB.Value = True
  '
  grdCallData.Visible = False
  ListView1.Visible = False
  '
  LoadExtList
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSelect.Form_Initialize.", vbCritical, "Error"
End Sub


Private Sub Form_Load()

    ListView1.View = lvwReport
    ListView1.Sorted = True

    With ListView1.ColumnHeaders
        .Add Text:="Name"
        .Add Text:="# Calls"
        .Add Text:="# BNB Notes"
        .Add Text:="Calls Vs. Notes"
        .Add Text:="Total Call Time"
        .Add Text:="# Follow-ups"
        .Add Text:="# Walk Throughs"
        .Add Text:="# Sales"
        .Add Text:="Avg Call Time"
        .Add                    '-- Dummy column. No text in case user pulls out to view.
        
        .Item(1).Width = 1500.09
        .Item(2).Width = 700.15
        .Item(3).Width = 1120.25
        .Item(4).Width = 1239.87
        .Item(5).Width = 1230.23
        .Item(6).Width = 1080
        .Item(7).Width = 1429.79
        .Item(8).Width = 739.84
        .Item(9).Width = 1149.73
        .Item(10).Width = 0      '-- Dummy column.
    End With
    
End Sub

Private Sub Form_Resize()
grdCallData.Width = Me.Width - 1000
ListView1.Width = Me.Width - 1000
If Me.Height > 0 Then
  grdCallData.Height = Me.Height - 2900
  ListView1.Height = Me.Height - 2900
End If
  Label3.Top = grdCallData.Top + grdCallData.Height + 300
  Label4.Top = Label3.Top
  txtTotal.Top = Label3.Top + 200
  txtAvg.Top = txtTotal.Top
End Sub


Private Sub cboGroup_Click()
  '
  'Disable Print button
  '
  cmdPrintReport.Enabled = False
End Sub

Private Sub cboCallDir_Click()
  '
  'Disable Print button
  '
  cmdPrintReport.Enabled = False
End Sub

Private Sub cmdPrintReport_Click()
  Dim sDate1 As String
  Dim sDate2 As String
  'Dim old_width As Integer
  If optBNB.Value = True Then
    '
    sDate1 = "( " & DTPicker1 & " )"
    sDate2 = "( " & DTPicker2 & " )"
    '
    RCallsAndNotes.GetData sName, iNumofCalls, iBNBNotes, sCallTime, iFollowups, iWalkThroughs, iSales, sAvgCallTime, iNumofUsers, sDate1, sDate2
    '
  Else
    '
    RCallLog.GetData DTStart, DTEnd, intWorkgroup, strSQL, strDirection
    '
  End If
  '

'Dim objPrintLV As clsPrintLV
'    Set objPrintLV = New clsPrintLV
'    objPrintLV.PrintListView ListView1, 0.1, 8, "Sample ListView Report", landscape, True
'    Set objPrintLV = Nothing

End Sub

Private Sub DTPicker1_Change()
  '
  'Disable Print button
  '
  cmdPrintReport.Enabled = False
End Sub

Private Sub DTPicker2_Change()
  '
  'Disable Print button
  '
  cmdPrintReport.Enabled = False
End Sub

Private Sub cmdGenReport_Click()
  If optBNB.Value = True Then
    BNBReport
  Else
    DurationReport
  End If
End Sub
  
Private Sub grdCallData_Click()

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SortListView(ListView1, ColumnHeader)
End Sub

Private Sub optBNB_Click()
  fraGroup.Visible = False
  fraGroup.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  txtTotal.Visible = False
  txtAvg.Visible = False
End Sub

Private Sub optDuration_Click()
  fraGroup.Visible = True
  Label3.Visible = True
  Label4.Visible = True
  txtTotal.Visible = True
  txtAvg.Visible = True
End Sub

Private Sub optExt_Click()
  '
End Sub

Private Sub optAll_Click()
  '
End Sub

Private Sub LoadExtList()
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  Dim X As Integer
  '
  cmdCategories.CommandTimeout = 300
  '
  Set cmdCategories.ActiveConnection = cnMain
  cmdCategories.CommandText = "SELECT     EmployeeFirst + N' ' + EmployeeLast AS Name, EmployeeExt From tblEmployees GROUP BY EmployeeFirst + N' ' + EmployeeLast, EmployeeExt HAVING      (NOT (EmployeeFirst + N' ' + EmployeeLast IS NULL)) AND (NOT (EmployeeExt IS NULL)) ORDER BY EmployeeExt"
  rstCategories.CursorLocation = adUseClient
  rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
  '
  iNumofUsers = rstCategories.RecordCount
  ReDim iExt(1 To iNumofUsers)
  ReDim sName(1 To iNumofUsers)
  ReDim iNumofCalls(1 To iNumofUsers)
  ReDim iBNBNotes(1 To iNumofUsers)
  ReDim lCallTime(1 To iNumofUsers)
  ReDim sCallTime(1 To iNumofUsers)
  ReDim iFollowups(1 To iNumofUsers)
  ReDim iWalkThroughs(1 To iNumofUsers)
  ReDim iSales(1 To iNumofUsers)
  ReDim lAvgCallTime(1 To iNumofUsers)
  ReDim sAvgCallTime(1 To iNumofUsers)
  X = 0
  '
  Do While Not rstCategories.eof
    If rstCategories!EmployeeExt <> "" Then
      X = X + 1
      iExt(X) = rstCategories!EmployeeExt
      cboGroup.AddItem rstCategories!Name & " (" & rstCategories!EmployeeExt & ")"
      sName(X) = rstCategories!Name
    End If
    rstCategories.MoveNext
  Loop
  Set cmdCategories = Nothing
  rstCategories.Close
  '
  cboGroup.ListIndex = 0
End Sub

Private Sub LoadNumOfCalls()
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  Dim X As Integer
  '
  cmdCategories.CommandTimeout = 300
  '
  For X = 1 To iNumofUsers
    Set cmdCategories.ActiveConnection = cnMain
    '
    cmdCategories.CommandText = "SELECT COUNT(DISTINCT SESSID) AS Calls,  SUM(DISTINCT CALLDUR) AS Duration, AVG(CALLDUR) AS AvgTime From ICC_CDR WHERE (P1NO LIKE N'" & iExt(X) & "') AND (DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1970-01-01 00:00:00', 102)) BETWEEN CONVERT(DATETIME,'" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) AND (NOT (LEFT(TKRMNO, 3) LIKE N'270'))"
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
    '
    If Not rstCategories.eof Then
      If Not IsNull(rstCategories!Calls) Then
        iNumofCalls(X) = rstCategories!Calls
      End If
      If Not IsNull(rstCategories!Duration) Then
        sCallTime(X) = ConvertTime(rstCategories!Duration)
        'lCallTime(x) = rstCategories!Duration
      Else
        sCallTime(X) = "No Data"
      End If
      If Not IsNull(rstCategories!AvgTime) Then
        sAvgCallTime(X) = ConvertTime(rstCategories!AvgTime)
        'lAvgCallTime(x) = rstCategories!AvgTime
      Else
        sAvgCallTime(X) = "No Data"
      End If
    End If
    'rstCategories.MoveNext
    Set cmdCategories = Nothing
    rstCategories.Close
  Next
  '

End Sub

Private Sub LoadBNBData()
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  Dim X As Integer
  '
  cmdCategories.CommandTimeout = 300
  '
  For X = 1 To iNumofUsers
    Set cmdCategories.ActiveConnection = cnMain
    '
    cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS BNBNotes FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) AND (tblEmployees.EmployeeExt = " & iExt(X) & ") GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast HAVING (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL))"
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
    '
    If Not rstCategories.eof Then
      iBNBNotes(X) = rstCategories!BNBNotes
    End If
    'rstCategories.MoveNext
    Set cmdCategories = Nothing
    rstCategories.Close
  Next
  '
  For X = 1 To iNumofUsers
    Set cmdCategories.ActiveConnection = cnMain
    '
    cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS Followups FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) AND (tblEmployees.EmployeeExt = " & iExt(X) & ") GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, TSupportActs.Type HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) AND (TSupportActs.Type LIKE N'Follow-up call')"
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
    '
    If Not rstCategories.eof Then
      iFollowups(X) = rstCategories!Followups
    End If
    'rstCategories.MoveNext
    Set cmdCategories = Nothing
    rstCategories.Close
  Next
  '
  For X = 1 To iNumofUsers
    Set cmdCategories.ActiveConnection = cnMain
    '
    cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS WalkThroughs FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) AND (tblEmployees.EmployeeExt = " & iExt(X) & ") GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, TSupportActs.Type HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) AND (TSupportActs.Type LIKE N'Walk Through')"
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
    '
    If Not rstCategories.eof Then
      iWalkThroughs(X) = rstCategories!WalkThroughs
    End If
    'rstCategories.MoveNext
    Set cmdCategories = Nothing
    rstCategories.Close
  Next
  '
  For X = 1 To iNumofUsers
    Set cmdCategories.ActiveConnection = cnMain
    '
    cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS Sales FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) AND (tblEmployees.EmployeeExt = " & iExt(X) & ") GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, TSupportActs.Type HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) AND (TSupportActs.Type LIKE N'Sale')"
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
    '
    If Not rstCategories.eof Then
      iSales(X) = rstCategories!Sales
    End If
    'rstCategories.MoveNext
    Set cmdCategories = Nothing
    rstCategories.Close
  Next
  '
End Sub

Private Function ConvertTime(Seconds As Long) As String
  Dim lHrs As Long
  Dim lMinutes As Long
  Dim lSeconds As Long
  
  lSeconds = Seconds
  '
  If lSeconds >= 3600 Then
     'get hours which is equal to seconds divided by 3600
        lHrs = lSeconds / 3600
      
      'set the seconds to the numbers after the decimal sign
      'thats what mod does
        lSeconds = lSeconds Mod 3600
  Else
  'if not greater than 3600, just set it to 0
      lHrs = 0
  End If

  If lSeconds >= 60 Then
  'greater than or equal to 60
  'set the minutes equal to the value of (seconds divided by 60).
  'and get the remaining numbers after the decimal
  'which will be the seconds
   'using the mod sign
  
      lMinutes = lSeconds \ 60
      lSeconds = lSeconds Mod 60
  Else
  'if not set to 0
      lMinutes = 0
  End If
  '
  If lHrs > 0 Then
    ConvertTime = Format(CStr(lHrs), "#####0") & ":" & _
      Format(CStr(lMinutes), "00") & "." & _
      Format(CStr(lSeconds), "00")
  Else
    ConvertTime = Format(CStr(lMinutes), "#0") & "." & _
      Format(CStr(lSeconds), "00")
  End If
  '
End Function

Private Sub LoadBNBData2()
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  Dim X As Integer
  '
  cmdCategories.CommandTimeout = 300
  '
  Set cmdCategories.ActiveConnection = cnMain
  '
  cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS BNBNotes, tblEmployees.EmployeeExt FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, tblEmployees.EmployeeExt HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) ORDER BY tblEmployees.EmployeeExt"
  rstCategories.CursorLocation = adUseClient
  rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
  '
  Do Until rstCategories.eof
    For X = 1 To iNumofUsers
      If iExt(X) = rstCategories!EmployeeExt Then
        iBNBNotes(X) = rstCategories!BNBNotes
      End If
    Next
    rstCategories.MoveNext
  Loop
  Set cmdCategories = Nothing
  rstCategories.Close
  '
  Set cmdCategories.ActiveConnection = cnMain
  '
  cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS Followups, tblEmployees.EmployeeExt FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) AND (TSupportActs.Type LIKE N'Follow-up call') GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, tblEmployees.EmployeeExt HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) ORDER BY tblEmployees.EmployeeExt"
  rstCategories.CursorLocation = adUseClient
  rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
  '
  Do Until rstCategories.eof
    For X = 1 To iNumofUsers
      If iExt(X) = rstCategories!EmployeeExt Then
        iFollowups(X) = rstCategories!Followups
      End If
    Next
    rstCategories.MoveNext
  Loop
  Set cmdCategories = Nothing
  rstCategories.Close
  '
  Set cmdCategories.ActiveConnection = cnMain
  '
  cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS WalkThroughs, tblEmployees.EmployeeExt FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) AND (TSupportActs.Type LIKE N'Walk Through') GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, tblEmployees.EmployeeExt HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) ORDER BY tblEmployees.EmployeeExt"
  rstCategories.CursorLocation = adUseClient
  rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
  '
  Do Until rstCategories.eof
    For X = 1 To iNumofUsers
      If iExt(X) = rstCategories!EmployeeExt Then
        iWalkThroughs(X) = rstCategories!WalkThroughs
      End If
    Next
    rstCategories.MoveNext
  Loop
  Set cmdCategories = Nothing
  rstCategories.Close
  '
  Set cmdCategories.ActiveConnection = cnMain
  '
  cmdCategories.CommandText = "SELECT COUNT(TSupportActs.Type) AS Sales, tblEmployees.EmployeeExt FROM tblEmployees RIGHT OUTER JOIN TSupportActs ON tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast = TSupportActs.[User] WHERE     (TSupportActs.[Date] BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 & "', 102)) AND (TSupportActs.Type LIKE N'Sale') GROUP BY tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast, tblEmployees.EmployeeExt HAVING      (NOT (tblEmployees.EmployeeFirst + N' ' + tblEmployees.EmployeeLast IS NULL)) ORDER BY tblEmployees.EmployeeExt"
  rstCategories.CursorLocation = adUseClient
  rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
  '
  Do Until rstCategories.eof
    For X = 1 To iNumofUsers
      If iExt(X) = rstCategories!EmployeeExt Then
        iSales(X) = rstCategories!Sales
      End If
    Next
    rstCategories.MoveNext
  Loop
  Set cmdCategories = Nothing
  rstCategories.Close
  '
End Sub

Private Sub LoadNumOfCalls2()
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  Dim X As Integer
  '
  cmdCategories.CommandTimeout = 300
  '
  Set cmdCategories.ActiveConnection = cnMain
  '
  cmdCategories.CommandText = "SELECT COUNT(DISTINCT ICC_CDR.SESSID) AS Calls, SUM(DISTINCT ICC_CDR.CALLDUR) AS Duration, AVG(ICC_CDR.CALLDUR) AS AvgTime, ICC_CDR.P1NO AS EmployeeExt FROM ICC_CDR RIGHT OUTER JOIN tblEmployees ON ICC_CDR.P1NO = tblEmployees.EmployeeExt WHERE (DATEADD(ss, ICC_CDR.STARTTIME, CONVERT(DATETIME, '1970-01-01 00:00:00', 102)) BETWEEN CONVERT(DATETIME, '" & DTPicker1 & "', 102) AND CONVERT(DATETIME, '" & DTPicker2 + 1 & "', 102)) GROUP BY ICC_CDR.P1NO ORDER BY ICC_CDR.P1NO"
  rstCategories.CursorLocation = adUseClient
  rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
  '
  Do Until rstCategories.eof
    For X = 1 To iNumofUsers
      If iExt(X) = rstCategories!EmployeeExt Then
        If Not IsNull(rstCategories!Calls) Then
          iNumofCalls(X) = rstCategories!Calls
        End If
        If Not IsNull(rstCategories!Duration) Then
          sCallTime(X) = ConvertTime(rstCategories!Duration)
          'lCallTime(x) = rstCategories!Duration
        Else
          sCallTime(X) = "No Data"
        End If
        If Not IsNull(rstCategories!AvgTime) Then
          sAvgCallTime(X) = ConvertTime(rstCategories!AvgTime)
          'lAvgCallTime(x) = rstCategories!AvgTime
        Else
          sAvgCallTime(X) = "No Data"
        End If
      End If
    Next
    rstCategories.MoveNext
  Loop
  Set cmdCategories = Nothing
  rstCategories.Close
  '
End Sub

Private Sub BNBReport()
  Dim X As Integer
  Dim sKey As String
  Dim iHolder As Integer
  Dim sHolder As String
  '
  If DTPicker1.Value > Date Or DTPicker2.Value < DTPicker1.Value Then
    MsgBox "Incorrect Date Values!"
    Exit Sub
  End If
  grdCallData.Visible = False
  ListView1.Visible = True
  '
  For X = 1 To iNumofUsers
    'iExt(x) = 0
    'sName(x) = 0
    iNumofCalls(X) = 0
    iBNBNotes(X) = 0
    lCallTime(X) = 0
    sCallTime(X) = "0"
    iFollowups(X) = 0
    iWalkThroughs(X) = 0
    iSales(X) = 0
    lAvgCallTime(X) = 0
    sAvgCallTime(X) = "0"
  Next
  '
  Screen.MousePointer = vbHourglass
  '
  LoadNumOfCalls2
  LoadBNBData2
  '
  SetupListView
''
''Enable Print button
''
  cmdPrintReport.Enabled = True


  '
  Screen.MousePointer = vbDefault

End Sub

Private Sub DurationReport()
  
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  Dim sCallDir As String
  Dim X As Integer
  Dim iAvgTime As Long
  '
  cmdCategories.CommandTimeout = 300
  '
  Screen.MousePointer = vbHourglass
  '
  intWorkgroup = iExt(cboGroup.ListIndex + 1)
  '
  grdCallData.Clear
  grdCallData.Cols = 4
  grdCallData.ColHeader(0) = flexColHeaderOn
  grdCallData.ColHeaderCaption(0, 0) = "Caller's #"
  grdCallData.ColHeaderCaption(0, 1) = "Date"
  grdCallData.ColHeaderCaption(0, 2) = "Direction"
  grdCallData.ColHeaderCaption(0, 3) = "Duration"
  '
  grdCallData.ColWidth(0, 0) = 1500
  grdCallData.ColWidth(1, 0) = 1700
  grdCallData.ColWidth(2, 0) = 1000
  grdCallData.ColWidth(3, 0) = 1000
  grdCallData.ColWidth(4, 0) = 0
  '
  '
  'Set the date values
  '
  If DTPicker1.Value > DTPicker2.Value Then
    MsgBox "Incorrect Date Values!"
    Screen.MousePointer = vbDefault
    Exit Sub
  Else
    If DTPicker1.Value > Date Or DTPicker2.Value > Date Then
      MsgBox "Incorrect Date Values!"
      Screen.MousePointer = vbDefault
      Exit Sub
    Else
      DTStart = DTPicker1.Value
      DTEnd = DTPicker2.Value
    End If
  End If
  '
  grdCallData.Visible = True
  ListView1.Visible = False
'
'Set the sCallDir value
'
  Select Case cboCallDir
    Case "Incoming"
      sCallDir = " (TKDIR = 2) AND "
      strDirection = "Incoming "
    Case "Outgoing"
      sCallDir = " (TKDIR = 4) AND "
      strDirection = "Outgoing "
    Case "Both"
      sCallDir = " "
      strDirection = "All "
  End Select
  '
  'Set the Ext or Workgroup type and number
  '
  'If optExt.value = True Then
    strGroupType = "P1NO"
  'Else
    'strGroupType = "P1WGNO"
  'End If
  '
  'Create the SQL statement
  '
  strSQL = "SELECT TKRMNO as Phone, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)) AS Calls, TKDIR as Direction, CALLDUR as Duration FROM         ICC_CDR WHERE " & sCallDir & " (" & strGroupType & " LIKE '" & intWorkgroup & "')AND (DATEADD(ss, STARTTIME, '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, STARTTIME, '1969-12-31 19:00:00') < ('" & DTEnd + 1 & "')) GROUP BY CALLDUR, TKRMNO, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)), TKDIR HAVING      (TKRMNO IS NOT NULL) ORDER BY DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))"
  strSQL2 = "SELECT TKRMNO AS Phone, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)) AS Calls, Direction = CASE TKDIR WHEN '2' THEN 'Incoming' WHEN '4' THEN 'Outgoing' end,  convert( char(8), dateadd( ss, CALLDUR, '00:00:00' ), 108 ) AS Duration, CALLDUR  From ICC_CDR WHERE " & sCallDir & " (" & strGroupType & " LIKE '" & intWorkgroup & "')AND (DATEADD(ss, STARTTIME, '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, STARTTIME, '1969-12-31 19:00:00') < ('" & DTEnd + 1 & "')) GROUP BY CALLDUR, TKRMNO, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)), TKDIR HAVING      (TKRMNO IS NOT NULL) ORDER BY DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))"
  '
  'Select revelant data
  '
  Set cmdCategories.ActiveConnection = cnMain
  cmdCategories.CommandText = strSQL2
  rstCategories.CursorLocation = adUseClient
  rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
  If rstCategories.RecordCount = 0 Then
    MsgBox "No revelant data!"
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  '
  iAvgTime = 0
  rstCategories.MoveFirst
  For X = 1 To rstCategories.RecordCount
    If Not rstCategories.eof Then
      iAvgTime = iAvgTime + Val(rstCategories!CALLDUR)
      rstCategories.MoveNext
    End If
  Next
  txtTotal.Text = ConvertTime(iAvgTime)
  txtAvg.Text = ConvertTime(iAvgTime / rstCategories.RecordCount)
  '
  Set grdCallData.Recordset = rstCategories
  grdCallData.Refresh
'
'Enable Print button
'
  cmdPrintReport.Enabled = True
  '
  Screen.MousePointer = vbDefault
End Sub


Private Sub Timer1_Timer()
  If chkUpdate.Value = vbChecked Then
    If iMins = Val(cboMins.Text) Then
      cmdGenReport_Click
      iMins = 0
    End If
    iMins = iMins + 1
  End If
End Sub

Public Sub SortListView(ByRef oListView As MSComctlLib.ListView, _
                        ByRef oColumnHeader As MSComctlLib.ColumnHeader)
'-- Sorts all list items correctly according to data type.
'-- Requirements:
'--     Any items without tag data will be sorted alphabetically.
'--     When creating the list, add a dummy column to the end, width = 0.
'--     Must be the last column in the list.
'--     Create the dummy column subitems as you fill the loop.
'--     Set .Sorted property = True.

    Dim oListItem           As MSComctlLib.ListItem
    Dim i                   As Integer
    Dim iTempColIndex       As Integer
    Dim bNoTagInColumn      As Boolean
    
    With oListView
    
        '-- If 0 or 1 items or -1(uninitialized), then don't try to sort.
        If .ListItems.Count < 2 Then GoTo Exit_Point
        
        iTempColIndex = .ColumnHeaders.Count - 1


        '-- Add the tag data from the clicked-on column to the dummy column.
        If oColumnHeader.Index = 1 Then
            '-- First column gets special treatment.
            For i = 1 To .ListItems.Count
                Set oListItem = .ListItems(i)
                oListItem.ListSubItems(iTempColIndex) = oListItem.Tag
            Next
            If Len(Trim(oListItem.Tag)) = 0 Then bNoTagInColumn = True
        Else
            '-- Subcolumns.
            For i = 1 To .ListItems.Count
                Set oListItem = .ListItems(i)
                oListItem.ListSubItems(iTempColIndex) = oListItem.ListSubItems(oColumnHeader.Index - 1).Tag
            Next
            If Len(Trim(oListItem.ListSubItems(iTempColIndex))) = 0 Then bNoTagInColumn = True
        End If
        
        
        If bNoTagInColumn Then
            '-- If the tag is blank, sort by default - alphabetically.
            .SortKey = oColumnHeader.Index - 1
        Else
            '-- Otherwise sort by the dummy column.
            .SortKey = iTempColIndex
        End If
        
        '-- Sort.
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        
        
        '-- Remove the data so no peeking.
        For i = 1 To .ListItems.Count
            Set oListItem = .ListItems(i)
            oListItem.ListSubItems(iTempColIndex) = ""
        Next

    End With
    
Exit_Point:
    Set oListItem = Nothing
End Sub

Private Sub SetupListView()

    Dim oListItem       As MSComctlLib.ListItem
    Dim dblDate         As Double
    Dim X               As Integer
    Dim sKey            As String
    Dim iHolder         As Integer
    Dim sHolder         As String
    

  ListView1.ListItems.Clear
  '
'  For X = 1 To iNumofUsers
'  '
'  If iNumofCalls(X) = 0 Or iBNBNotes(X) = 0 Then
'    If iBNBNotes(X) < 0 Then
'        iHolder = 100
'      Else
'        iHolder = 0
'      End If
'    Else
'      iHolder = (iBNBNotes(X) / iNumofCalls(X)) * 100
'    End If
'    sHolder = iHolder & "%"
'    '
'    sKey = "A" & CStr(X)
'    Set oListItem = ListView1.ListItems.Add(, sKey, sName(X))
'
'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iNumofCalls(X), , iNumofCalls(X)).ForeColor = vbBlack
'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iBNBNotes(X), , iBNBNotes(X)).ForeColor = vbBlack
'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , sHolder, , sHolder).ForeColor = vbBlack
'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , sCallTime(X), , sCallTime(X)).ForeColor = vbBlack
'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iFollowups(X), , iFollowups(X)).ForeColor = vbBlack
'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iWalkThroughs(X), , iWalkThroughs(X)).ForeColor = vbBlack
'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , iSales(X), , iSales(X)).ForeColor = vbBlack
'    ListView1.ListItems.Item(sKey).ListSubItems.Add(, , sAvgCallTime(X), , sAvgCallTime(X)).ForeColor = vbBlack
'  Next
  
    '-- Put some data in the listview.
    With ListView1.ListItems
        For X = 1 To iNumofUsers
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          If iNumofCalls(X) = 0 Or iBNBNotes(X) = 0 Then
            If iBNBNotes(X) < 0 Then
                iHolder = 100
              Else
                iHolder = 0
              End If
            Else
              iHolder = (iBNBNotes(X) / iNumofCalls(X)) * 100
            End If
            sHolder = iHolder & "%"
            '
            sKey = "A" & CStr(X)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Set oListItem = .Add(, sKey, sName(X))
            'oListItem.Tag = Format(X, "00000")

            oListItem.ListSubItems.Add , , iNumofCalls(X)
            oListItem.ListSubItems(1).Tag = Format(iNumofCalls(X), "0000000")

            oListItem.ListSubItems.Add , , iBNBNotes(X)
            oListItem.ListSubItems(2).Tag = Format(iBNBNotes(X), "0000000")

            oListItem.ListSubItems.Add , , sHolder
            oListItem.ListSubItems(3).Tag = Format$(sHolder, "0000000000%")

            oListItem.ListSubItems.Add , , sCallTime(X)
            oListItem.ListSubItems(4).Tag = Format$(Val(sCallTime(X)), "yyyymmddHHMMSS")

            oListItem.ListSubItems.Add , , iFollowups(X)
            oListItem.ListSubItems(5).Tag = Format(iFollowups(X), "0000000")

            oListItem.ListSubItems.Add , , iWalkThroughs(X)
            oListItem.ListSubItems(6).Tag = Format(iWalkThroughs(X), "0000000")

            oListItem.ListSubItems.Add , , iSales(X)
            oListItem.ListSubItems(7).Tag = Format(iSales(X), "0000000")

            oListItem.ListSubItems.Add , , sAvgCallTime(X)
            oListItem.ListSubItems(8).Tag = Format$(Val(sAvgCallTime(X)), "yyyymmddHHMMSS")
            
            '-- Dummy ListSubItems column.
            oListItem.ListSubItems.Add

        Next
        
    End With
End Sub
