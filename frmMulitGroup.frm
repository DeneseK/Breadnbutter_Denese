VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmMultiChart 
   BorderStyle     =   0  'None
   Caption         =   "Multi-Line Chart"
   ClientHeight    =   8805
   ClientLeft      =   2055
   ClientTop       =   1890
   ClientWidth     =   12570
   Icon            =   "frmMulitGroup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCallType 
      Caption         =   "Calls and V-Mail"
      Height          =   1815
      Left            =   1800
      TabIndex        =   79
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton optVoiceMail 
         Caption         =   "V-Mail Only"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optCalls 
         Caption         =   "Calls Only"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optBoth 
         Caption         =   "Both"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkLablePoints 
      Caption         =   "Label Points"
      Height          =   255
      Left            =   9000
      TabIndex        =   78
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdAvg 
      Caption         =   "CHART IT!"
      Height          =   315
      Left            =   9000
      TabIndex        =   73
      Top             =   1080
      Width           =   1335
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   6600
      TabIndex        =   66
      Top             =   4080
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowToday       =   0   'False
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   70123521
      CurrentDate     =   38190
   End
   Begin VB.ComboBox cboChartType 
      Height          =   315
      ItemData        =   "frmMulitGroup.frx":030A
      Left            =   9000
      List            =   "frmMulitGroup.frx":032C
      TabIndex        =   52
      Text            =   "2D Line"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame fraMultiLines 
      Caption         =   "Multiple Lines"
      Height          =   1815
      Left            =   6360
      TabIndex        =   20
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optCallVsVoice 
         Caption         =   "Calls vs. V-Mails"
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optNone 
         Caption         =   "None"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "All Call Directions"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optMultiDates 
         Caption         =   "Multiple Dates"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMultiGroups 
         Caption         =   "Multiple Ext#s or Group#s"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdChart 
      Caption         =   "CHART IT!"
      Height          =   315
      Left            =   9000
      TabIndex        =   19
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrintChart 
      Caption         =   "PRINT CHART!"
      Enabled         =   0   'False
      Height          =   315
      Left            =   9000
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame fraGroup 
      Caption         =   "Ext# or Group#"
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1575
      Begin VB.ComboBox cboGroup 
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optExt 
         Caption         =   "by Ext"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optWorkgroup 
         Caption         =   "by Workgroup"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame fraVariables 
      Caption         =   "Variables"
      Height          =   1815
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton optTotal 
         Caption         =   "Total Count"
         Height          =   255
         Left            =   1320
         TabIndex        =   75
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optAvg 
         Caption         =   "Average"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   1440
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   315
         Left            =   1320
         TabIndex        =   72
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70123521
         CurrentDate     =   38194
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   315
         Left            =   1320
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70123521
         CurrentDate     =   38194
      End
      Begin VB.ComboBox cboCallDir 
         Height          =   315
         ItemData        =   "frmMulitGroup.frx":038A
         Left            =   120
         List            =   "frmMulitGroup.frx":0397
         TabIndex        =   3
         Text            =   "Both"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cboDateType 
         Height          =   315
         ItemData        =   "frmMulitGroup.frx":03B5
         Left            =   120
         List            =   "frmMulitGroup.frx":03C5
         TabIndex        =   2
         Text            =   "Hour"
         Top             =   480
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   1320
         TabIndex        =   64
         Text            =   "2003"
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboSunday 
         Height          =   315
         Left            =   2880
         TabIndex        =   65
         Text            =   "01"
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   2160
         TabIndex        =   69
         Text            =   "01"
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblCallType 
         AutoSize        =   -1  'True
         Caption         =   "Call Direction"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   930
      End
      Begin VB.Label lblDateType 
         AutoSize        =   -1  'True
         Caption         =   "By"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   180
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
         Height          =   195
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblMon 
         Caption         =   "Month"
         Height          =   255
         Left            =   2160
         TabIndex        =   70
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblWeek 
         AutoSize        =   -1  'True
         Caption         =   "Day"
         Height          =   195
         Left            =   2880
         TabIndex        =   68
         Top             =   840
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   1320
         TabIndex        =   67
         Top             =   840
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraGroups 
      Caption         =   "Ext#s or Group#s"
      Height          =   8535
      Left            =   10440
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ComboBox cboNum 
         Height          =   315
         ItemData        =   "frmMulitGroup.frx":03E1
         Left            =   1320
         List            =   "frmMulitGroup.frx":0400
         TabIndex        =   30
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   720
         TabIndex        =   29
         Top             =   7320
         Width           =   855
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   720
         TabIndex        =   28
         Top             =   6600
         Width           =   855
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   720
         TabIndex        =   27
         Top             =   5880
         Width           =   855
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   720
         TabIndex        =   26
         Top             =   5160
         Width           =   855
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   720
         TabIndex        =   25
         Top             =   4440
         Width           =   855
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   720
         TabIndex        =   16
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   720
         TabIndex        =   15
         Top             =   3000
         Width           =   855
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   13
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   63
         Top             =   7080
         Width           =   990
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   62
         Top             =   6360
         Width           =   990
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   61
         Top             =   5640
         Width           =   990
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   60
         Top             =   4920
         Width           =   990
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   59
         Top             =   4200
         Width           =   990
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   58
         Top             =   3480
         Width           =   990
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   57
         Top             =   2760
         Width           =   990
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   56
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label lblExtGroup 
         AutoSize        =   -1  'True
         Caption         =   "Ext or Group#"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   55
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblGroupNum 
         AutoSize        =   -1  'True
         Caption         =   "How Many?"
         Height          =   195
         Left            =   360
         TabIndex        =   54
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame fraMultiDates 
      Caption         =   "Multiple Dates"
      Height          =   8535
      Left            =   10440
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ComboBox cboDateNum 
         Height          =   315
         ItemData        =   "frmMulitGroup.frx":0420
         Left            =   600
         List            =   "frmMulitGroup.frx":043F
         TabIndex        =   50
         Top             =   360
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   32
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   34
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   3
         Left            =   360
         TabIndex        =   36
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   4
         Left            =   360
         TabIndex        =   38
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   40
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   6
         Left            =   360
         TabIndex        =   42
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   7
         Left            =   360
         TabIndex        =   44
         Top             =   6120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   8
         Left            =   360
         TabIndex        =   46
         Top             =   6960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   9
         Left            =   360
         TabIndex        =   48
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70123521
         CurrentDate     =   37809
      End
      Begin VB.Label lblLines 
         AutoSize        =   -1  'True
         Caption         =   "Lines"
         Height          =   195
         Left            =   1320
         TabIndex        =   51
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date# 10"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   49
         Top             =   7560
         Width           =   1050
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date #9"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   47
         Top             =   6720
         Width           =   960
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date #8"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   45
         Top             =   5880
         Width           =   960
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date #7"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   43
         Top             =   5040
         Width           =   960
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date #6"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   41
         Top             =   4200
         Width           =   960
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date #5"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   39
         Top             =   3360
         Width           =   960
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date #4"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   37
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date #3"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   35
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "Start Date #2"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   33
         Top             =   840
         Width           =   960
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6735
      Left            =   0
      OleObjectBlob   =   "frmMulitGroup.frx":045F
      TabIndex        =   76
      Top             =   2040
      Width           =   10335
   End
   Begin VB.Label lblChartType 
      AutoSize        =   -1  'True
      Caption         =   "Chart Type"
      Height          =   195
      Left            =   9000
      TabIndex        =   53
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmMultiChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Private rslabels As ADODB.Recordset

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
Public bChartData As Boolean
'

Public strSQL As String, DTStart As Date, DTEnd As Date, intEXT As Integer, intWorkgroup As Integer, strDateType As String, strGroupType As String, strDirection As String
'Public chrtArray()
Public lGreatestValue As Long


Private Sub cboCallDir_Click()
'
'Disable Print button
'
  cmdPrintChart.Enabled = False
End Sub

Private Sub cboChartType_Click()
  Select Case cboChartType.Text
    Case "2D Bar"
      MSChart1.chartType = VtChChartType2dBar
    Case "3D Bar"
      MSChart1.chartType = VtChChartType3dBar
    Case "2D Line"
      MSChart1.chartType = VtChChartType2dLine
    Case "3D Line"
      MSChart1.chartType = VtChChartType3dLine
    Case "2D Area"
      MSChart1.chartType = VtChChartType2dArea
    Case "3D Area"
      MSChart1.chartType = VtChChartType3dArea
    Case "2D Step"
      MSChart1.chartType = VtChChartType2dStep
    Case "3D Step"
      MSChart1.chartType = VtChChartType3dStep
    Case "2D Combo"
      MSChart1.chartType = VtChChartType2dCombination
    Case "3D Combo"
      MSChart1.chartType = VtChChartType3dCombination
  End Select
       
End Sub

Private Sub cboDateNum_Click()
  Dim x As Integer
  For x = 1 To 9
    DTPicker1(x).Enabled = False
  Next x
  For x = 2 To Val(cboDateNum.Text)
    DTPicker1(x - 1).Enabled = True
  Next x
'
'Disable Print button
'
  cmdPrintChart.Enabled = False
End Sub

Private Sub cboDateType_Click()
Dim sTemp As String
Dim x As Integer
  '
  Select Case cboDateType
      Case "Hour"
'        cboYear.Enabled = False
'        cboMonth.Enabled = False
'        cboSunday.Enabled = False
'        DTPicker1(0).Enabled = True
'        optMultiDates.Enabled = True
        DTPicker2.value = DTPicker1(0).value + 1
      Case "Day"
        MonthView1.value = DTPicker1(0).value
        MonthView1.DayOfWeek = mvwSunday
        DTPicker1(0).value = MonthView1.value
'        cboYear.Enabled = True
'        cboMonth.Enabled = True
'        cboSunday.Enabled = True
'        DTPicker1(0).Enabled = False
'        If optMultiDates.value = True Then
'          optNone.value = True
'        End If
'        optMultiDates.Enabled = False
'        GetSundays
        DTPicker2.value = DTPicker1(0).value + 7
      Case Else
'        cboYear.Enabled = True
'        cboMonth.Enabled = False
'        cboSunday.Enabled = False
'        DTPicker1(0).Enabled = False
'        optMultiDates.Enabled = False
'        If optMultiDates.value = True Then
'          optNone.value = True
'        End If
'        optMultiDates.Enabled = False
        sTemp = DTPicker1(0).year
        DTPicker1(0).value = "1/1/" & sTemp
        If DTPicker1(0).year = 2004 Then
          DTPicker2.value = DTPicker1(0).value + 366
        Else
          DTPicker2.value = DTPicker1(0).value + 365
        End If
  End Select
  For x = 1 To 9
    DTPicker1_Change x
  Next
'
'Disable Print button
'
  cmdPrintChart.Enabled = False
End Sub

Private Sub cboGroup_Click(Index As Integer)
'
'Disable Print button
'
  cmdPrintChart.Enabled = False
End Sub

Private Sub cboMonth_Click()
  GetSundays
  MonthView1.month = cboMonth.Text
  MonthView1.year = cboYear.Text
  cboSunday.Text = MonthView1.day
  'MonthView1.DayOfWeek = 1
  DTPicker1(0).value = MonthView1.value
  GetSundays
End Sub

Private Sub cboNum_Click()
  Dim x As Integer
  For x = 1 To 9
    cboGroup(x).Enabled = False
  Next x
  For x = 2 To Val(cboNum.Text)
    cboGroup(x - 1).Enabled = True
  Next x
'
'Disable Print button
'
  cmdPrintChart.Enabled = False
End Sub

Private Sub cboSunday_Click()
  MonthView1.year = cboYear.Text
  MonthView1.day = cboSunday.Text
  'MonthView1.DayOfWeek = 1
  cboMonth.Text = MonthView1.month
  DTPicker1(0).value = MonthView1.value
End Sub

Private Sub cboYear_Click()
  GetSundays
  If cboDateType = "Day" Then
    MonthView1.year = cboYear.Text
    MonthView1.day = cboSunday.Text
    'MonthView1.DayOfWeek = 1
    cboMonth.Text = MonthView1.month
  Else
    MonthView1.value = "1/1/" & cboYear.Text
    cboMonth.Text = MonthView1.month
    cboSunday.Text = MonthView1.day
  End If
  DTPicker1(0).value = MonthView1.value
End Sub

'Private Sub chkVoiceMail_Click()
''
''Disable Print button
''
'  cmdPrintChart.Enabled = False
''
''If checked Disable Call Direction and set to Incoming
''
'  If chkVoiceMail.Value = 1 Then
'    cboCallDir.Text = "Incoming"
'    cboCallDir.Enabled = False
'  Else
'    If optExt.Value = True Then
'      cboCallDir.Enabled = True
'    End If
'  End If
'End Sub

Private Sub cmdAvg_Click()
  Dim sVoiceMail As String
  Dim sCallDir As String
  Dim x As Integer
  Dim iInterval(1 To 100) As Integer
  Dim IAvg As Integer
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  '
  cmdCategories.CommandTimeout = 300
  '
  Screen.MousePointer = vbHourglass
  '
  intWorkgroup = cboGroup(0).Text
'
' VoiceMail
'
    If optVoiceMail.value = True Then
        sVoiceMail = "VMSTARTTM"
        strDirection = "VoiceMail "
    Else
        sVoiceMail = "STARTTIME"
    
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
            strDirection = ""
        End Select
        '
        If optCalls.value = True Then
          sCallDir = sCallDir + "VMSTARTTM = 0 AND "
        Else
          strDirection = "Total " & strDirection
        End If
        '
    End If
  
'
'Set the date values
'
    If DTPicker3.value > DTPicker4.value Then
        MsgBox "Incorrect Date Values!"
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        If DTPicker3.value > Date Or DTPicker4.value > Date Then
            MsgBox "Incorrect Date Values!"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            DTStart = DTPicker3.value
            DTEnd = DTPicker4.value
        End If
    End If
'
'Set the DATEPART value
'
    Select Case cboDateType
        Case "Hour"
            strDateType = "hh"
            lGreatestValue = 24
            For x = 1 To 24
              iInterval(x) = (DTEnd - DTStart)
            Next
        Case "Day"
            strDateType = "w"
            lGreatestValue = 7
            '
            For x = 0 To (DTEnd - DTStart)
              IAvg = DatePart("w", (DTStart + x), vbSunday, vbUseSystem)
              iInterval(IAvg) = iInterval(IAvg) + 1
            Next
        Case "Week"
            strDateType = "ww"
            lGreatestValue = 52
        Case "Month"
            strDateType = "m"
            lGreatestValue = 12
        Case "Year"
            'strDateType = "yyyy"
            Exit Sub
    End Select
'
'Set the Ext or Workgroup type and number
'
    If optExt.value = True Then
        strGroupType = "P1NO"
    Else
        strGroupType = "P1WGNO"
    End If
    
    


'
'Create the SQL statement
'
    strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & intWorkgroup & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & DTEnd & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
    'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
'Select revelant data

    Set cmdCategories.ActiveConnection = cnMain
    cmdCategories.CommandText = strSQL
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic

If rstCategories.RecordCount = 0 Then
    MsgBox "No revelant data!"
    Screen.MousePointer = vbDefault
    'Exit Sub
End If
bChartData = True 'set flag to true
'                    C H A R T
'
' Dynamic 2-dimensional array to store series
' The first index (x) is the total number of series
' x-axis value in the 1st slot (i.e. chrtArray(x,1)
' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
'
'ReDim chrtArray(1 To lGreatestValue, 1 To 2)

MSChart1.ShowLegend = True
'MSChart1.chartType = VtChChartType2dLine
'
'Chart Title centered on top
'
MSChart1.Title.Text = strDirection & "Calls between " & DTPicker3.value & " and " & DTPicker4.value
'
'Chart X and Y axis titles
'
MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle.Text = cboDateType.Text
MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle.Text = ""
'
'Find the minimum and maximum date value in the record
'
Dim iMinVal As Integer, iMaxVal As Integer
rstCategories.MoveFirst
iMinVal = rstCategories!datetype
rstCategories.MoveLast
iMaxVal = rstCategories!datetype
rstCategories.MoveFirst
If iMinVal = 0 Then
  iMinVal = 1
  iMaxVal = 24
End If
'
ReDim chrtArray(iMinVal To iMaxVal, 1 To 2)
'
'Load the array with 0s with correct # of rows
'
For x = iMinVal To iMaxVal
  If x = 0 Then
    chrtArray(24, 1) = 0
    chrtArray(24, 2) = 24
  Else
    chrtArray(x, 1) = 0
    chrtArray(x, 2) = x
End If
Next x
'
'Load the array with data
'
For x = 1 To rstCategories.RecordCount
  If strDateType = "hh" And rstCategories!datetype = 0 Then
    rstCategories!datetype = 24
  End If
  If iInterval(x) = 0 Then iInterval(x) = 1
  chrtArray(rstCategories!datetype, 1) = rstCategories!Calls / iInterval(x)
  rstCategories.MoveNext
Next x
'
'Attach the array of data to MS-CHART
'
With MSChart1
    .ChartData = chrtArray
    .ColumnCount = 1
    .ColumnLabelCount = 1
    .Column = 1
    .ColumnLabel = "Calls for " & cboGroup(0).Text
    For x = 1 To iMaxVal - iMinVal + 1
      .Row = x
      Select Case strDateType
        Case "hh"
          Select Case chrtArray(iMinVal - 1 + x, 2)
            Case 1
              .RowLabel = "1 AM"
            Case 2
              .RowLabel = "2 AM"
            Case 3
              .RowLabel = "3 AM"
            Case 4
              .RowLabel = "4 AM"
            Case 5
              .RowLabel = "5 AM"
            Case 6
              .RowLabel = "6 AM"
            Case 7
              .RowLabel = "7 AM"
            Case 8
              .RowLabel = "8 AM"
            Case 9
              .RowLabel = "9 AM"
            Case 10
              .RowLabel = "10 AM"
            Case 11
              .RowLabel = "11 AM"
            Case 12
              .RowLabel = "12 AM"
            Case 13
              .RowLabel = "1 PM"
            Case 14
              .RowLabel = "2 PM"
            Case 15
              .RowLabel = "3 PM"
            Case 16
              .RowLabel = "4 PM"
            Case 17
              .RowLabel = "5 PM"
            Case 18
              .RowLabel = "6 PM"
            Case 19
              .RowLabel = "7 PM"
            Case 20
              .RowLabel = "8 PM"
            Case 21
              .RowLabel = "9 PM"
            Case 22
              .RowLabel = "10 PM"
            Case 23
              .RowLabel = "11 PM"
            Case 24
              .RowLabel = "12 PM"
            End Select
          Case "w"
            Select Case chrtArray(iMinVal - 1 + x, 2)
              Case 1
                .RowLabel = "Sun"
              Case 2
               .RowLabel = "Mon"
              Case 3
               .RowLabel = "Tue"
              Case 4
                .RowLabel = "Wed"
              Case 5
                .RowLabel = "Thur"
              Case 6
                .RowLabel = "Fri"
              Case 7
                .RowLabel = "Sat"
            End Select
          Case "m"
            Select Case chrtArray(iMinVal - 1 + x, 2)
              Case 1
                .RowLabel = "Jan"
              Case 2
                .RowLabel = "Feb"
              Case 3
                .RowLabel = "Mar"
              Case 4
                .RowLabel = "Apr"
              Case 5
                .RowLabel = "May"
              Case 6
                .RowLabel = "Jun"
              Case 7
                .RowLabel = "Jul"
              Case 8
                .RowLabel = "Aug"
              Case 9
                .RowLabel = "Sep"
              Case 10
                .RowLabel = "Oct"
              Case 11
                .RowLabel = "Nov"
              Case 12
                .RowLabel = "Dec"
            End Select
          Case Else
            .RowLabel = chrtArray(iMinVal - 1 + x, 2)
        End Select
    Next x
End With
  If chkLablePoints.value = vbChecked Then LablePoints
'
'Enable Print button
'
  cmdPrintChart.Enabled = True
  '
  Screen.MousePointer = vbDefault
  '
End Sub

Private Sub cmdChart_Click()

  Dim sVoiceMail As String
  Dim sCallDir As String
  Dim x As Integer
  Dim y As Integer
  Dim iLines As Integer
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  '
  cmdCategories.CommandTimeout = 300
  '
  Screen.MousePointer = vbHourglass
  '
  'intWorkgroup = cboGroup.Text
  If optMultiDates.value = True And cboDateNum.Text = "" Or optMultiGroups.value = True And cboNum.Text = "" Then
      optNone.value = True
  End If
  '
'  If DTPicker1(0).value > DTPicker2.value Then
'      MsgBox "Incorrect Date Values!"
'      Exit Sub
'    Else
'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
'        MsgBox "Incorrect Date Values!"
'        Exit Sub
'      Else
'        DTStart = DTPicker1(0).value
'        DTEnd = DTPicker2.value
'      End If
'    End If
  '
  Select Case cboDateType
      Case "Hour"
        strDateType = "hh"
        lGreatestValue = 24
      Case "Day"
        strDateType = "w"
        lGreatestValue = 7
      Case "Week"
        strDateType = "ww"
        lGreatestValue = 53
      Case "Month"
        strDateType = "m"
        lGreatestValue = 12
      Case "Year"
        'strDateType = "yyyy"
        Exit Sub
    End Select
  '
  If optMultiGroups.value = True Then
    '
    ' VoiceMail
    '
    If optVoiceMail.value = True Then
      sVoiceMail = "VMSTARTTM"
      strDirection = "VoiceMail "
    Else
      sVoiceMail = "STARTTIME"
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
          strDirection = ""
      End Select
      '
      If optCalls.value = True Then
        sCallDir = sCallDir + "VMSTARTTM = 0 AND "
      Else
        strDirection = "Total " & strDirection
      End If
      '
    End If
    '
    'Set the date values
    '
'    If DTPicker1(0).value > DTPicker2.value Then
'      MsgBox "Incorrect Date Values!"
'      Exit Sub
'    Else
'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
'        MsgBox "Incorrect Date Values!"
'        Exit Sub
'      Else
        DTStart = DTPicker1(0).value
        DTEnd = DTPicker2.value
'      End If
'    End If
    '
    'Set the DATEPART value
    '
'    Select Case cboDateType
'      Case "Hour"
'        strDateType = "hh"
'        lGreatestValue = 24
'      Case "Day"
'        strDateType = "w"
'        lGreatestValue = 7
'      Case "Week"
'        strDateType = "ww"
'        lGreatestValue = 53
'      Case "Month"
'        strDateType = "m"
'        lGreatestValue = 12
'      Case "Year"
'        'strDateType = "yyyy"
'        Exit Sub
'    End Select
    '
    'Set the Ext or Workgroup type and number
    '
    If optExt.value = True Then
      strGroupType = "P1NO"
    Else
      strGroupType = "P1WGNO"
    End If
    '                    C H A R T
    '
    ' Dynamic 2-dimensional array to store series
    ' The first index (x) is the total number of series
    ' x-axis value in the 1st slot (i.e. chrtArray(x,1)
    ' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
    '
    If cboNum.Text = "" Then cboNum = 1
    ReDim chrtArray(1 To lGreatestValue, 1 To Val(cboNum))
    MSChart1.ShowLegend = True
    'MSChart1.chartType = VtChChartType2dLine
    '
    'Chart Title centered on top
    '
    MSChart1.Title.Text = strDirection & "Calls between " & DTPicker1(0).value & " and " & DTPicker2.value
    '
    'Chart X and Y axis titles
    '
    MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle.Text = cboDateType.Text
    MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle.Text = ""
    '
    'Chart Foot note
    '
    'MSChart1.FootnoteText = "footnote"
    '
    'Load the array with 0s with correct # of rows
    '
    For x = 1 To lGreatestValue
      For y = 1 To Val(cboNum)
        chrtArray(x, y) = 0
      Next y
    Next x
    For x = 0 To Val(cboNum) - 1
      '
      'Create the SQL statement
      '
      strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & cboGroup(x).Text & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & DTEnd + 1 & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
      'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
      '
      'Select revelant data
      '
      Set cmdCategories.ActiveConnection = cnMain
      cmdCategories.CommandText = strSQL
      rstCategories.CursorLocation = adUseClient
      rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
      If rstCategories.RecordCount = 0 Then
        MsgBox "No revelant data for " & cboGroup(x).Text & "!"
        Screen.MousePointer = vbDefault
        'Exit Sub
      End If
      bChartData = True 'set flag to true
      '
      'Load the array with data
      '
      For y = 1 To rstCategories.RecordCount
        If strDateType = "hh" And rstCategories!datetype = 0 Then 'this is for the 0 hour problem
          rstCategories!datetype = 24
        End If
        chrtArray(rstCategories!datetype, x + 1) = rstCategories!Calls
        rstCategories.MoveNext
      Next y
      rstCategories.Close
    Next x
    '
    'Attach the array of data to MS-CHART
    '
    With MSChart1
      .ChartData = chrtArray
      .ColumnCount = Val(cboNum)
      .ColumnLabelCount = Val(cboNum) - 1
      For x = 0 To Val(cboNum) - 1
        .Column = x + 1
        .ColumnLabel = "Calls for " & cboGroup(x).Text
      Next x
      Select Case strDateType
        Case "hh"
          .Row = 1
          .RowLabel = "1 AM"
          .Row = 2
          .RowLabel = "2 AM"
          .Row = 3
          .RowLabel = "3 AM"
          .Row = 4
          .RowLabel = "4 AM"
          .Row = 5
          .RowLabel = "5 AM"
          .Row = 6
          .RowLabel = "6 AM"
          .Row = 7
          .RowLabel = "7 AM"
          .Row = 8
          .RowLabel = "8 AM"
          .Row = 9
          .RowLabel = "9 AM"
          .Row = 10
          .RowLabel = "10 AM"
          .Row = 11
          .RowLabel = "11 AM"
          .Row = 12
          .RowLabel = "12 AM"
          .Row = 13
          .RowLabel = "1 PM"
          .Row = 14
          .RowLabel = "2 PM"
          .Row = 15
          .RowLabel = "3 PM"
          .Row = 16
          .RowLabel = "4 PM"
          .Row = 17
          .RowLabel = "5 PM"
          .Row = 18
          .RowLabel = "6 PM"
          .Row = 19
          .RowLabel = "7 PM"
          .Row = 20
          .RowLabel = "8 PM"
          .Row = 21
          .RowLabel = "9 PM"
          .Row = 22
          .RowLabel = "10 PM"
          .Row = 23
          .RowLabel = "11 PM"
          .Row = 24
          .RowLabel = "12 PM"
        Case "w"
          .Row = 1
          .RowLabel = "Sun"
          .Row = 2
          .RowLabel = "Mon"
          .Row = 3
          .RowLabel = "Tue"
          .Row = 4
          .RowLabel = "Wed"
          .Row = 5
          .RowLabel = "Thur"
          .Row = 6
          .RowLabel = "Fri"
          .Row = 7
          .RowLabel = "Sat"
        Case "m"
          .Row = 1
          .RowLabel = "Jan"
          .Row = 2
          .RowLabel = "Feb"
          .Row = 3
          .RowLabel = "Mar"
          .Row = 4
          .RowLabel = "Apr"
          .Row = 5
          .RowLabel = "May"
          .Row = 6
          .RowLabel = "Jun"
          .Row = 7
          .RowLabel = "Jul"
          .Row = 8
          .RowLabel = "Aug"
          .Row = 9
          .RowLabel = "Sep"
          .Row = 10
          .RowLabel = "Oct"
          .Row = 11
          .RowLabel = "Nov"
          .Row = 12
          .RowLabel = "Dec"
        Case Else
          For y = 1 To lGreatestValue
            .Row = y
            .RowLabel = chrtArray(y, 2)
          Next y
          .Refresh
      End Select
    End With
  End If
  '
  '
  '
  '
  '
  '
  If optMultiDates.value = True Then
    '
    ' VoiceMail
    '
    If optVoiceMail.value = True Then
      sVoiceMail = "VMSTARTTM"
      strDirection = "VoiceMail "
    Else
      sVoiceMail = "STARTTIME"
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
          strDirection = ""
      End Select
      '
      If optCalls.value = True Then
        sCallDir = sCallDir + "VMSTARTTM = 0 AND "
      Else
        strDirection = "Total " & strDirection
      End If
      '
    End If
    '
    'Set the DATEPART value
    '
'    Select Case cboDateType
'      Case "Hour"
'        strDateType = "hh"
'        lGreatestValue = 24
'      Case "Day"
'        strDateType = "w"
'        lGreatestValue = 7
'      Case "Week"
'        strDateType = "ww"
'        lGreatestValue = 53
'      Case "Month"
'        strDateType = "m"
'        lGreatestValue = 12
'      Case "Year"
'        'strDateType = "yyyy"
'        Exit Sub
'    End Select
    '
    'Set the Ext or Workgroup type and number
    '
    If optExt.value = True Then
      strGroupType = "P1NO"
    Else
      strGroupType = "P1WGNO"
    End If
    '                    C H A R T
    '
    ' Dynamic 2-dimensional array to store series
    ' The first index (x) is the total number of series
    ' x-axis value in the 1st slot (i.e. chrtArray(x,1)
    ' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
    '
    If Val(cboDateNum) = 0 Then
      cboDateNum.Text = "1"
    End If
    ReDim chrtArray(1 To lGreatestValue, 1 To Val(cboDateNum))
    MSChart1.ShowLegend = True
    'MSChart1.chartType = VtChChartType2dLine
    '
    'Chart X and Y axis titles
    '
    MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle.Text = cboDateType.Text
    MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle.Text = ""
    '
    'Set the date values
    '
'    If DTPicker1(0).value > DTPicker2.value Then
'      MsgBox "Incorrect Date Values!"
'      Exit Sub
'    Else
'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
'        MsgBox "Incorrect Date Values!"
'        Exit Sub
'      Else
        DTStart = DTPicker1(0).value
        DTEnd = DTPicker2.value - DTPicker1(0).value
'      End If
'    End If
    '
    'Chart Title centered on top
    '
    MSChart1.Title.Text = strDirection & "Calls for " & cboGroup(0).Text '& ", " & DTEnd & " " & cboDateType.Text & "s of data per line"
    '
    'Load the array with 0s with correct # of rows
    '
    For x = 1 To lGreatestValue
      For y = 1 To Val(cboDateNum)
        chrtArray(x, y) = 0
      Next y
    Next x
    For x = 0 To Val(cboDateNum) - 1
      '
      'Create the SQL statement
      '
      strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & cboGroup(0).Text & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTPicker1(x).value & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & DTPicker1(x).value + DTEnd + 1 & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
      'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
      '
      'Select revelant data
      '
      Set cmdCategories.ActiveConnection = cnMain
      cmdCategories.CommandText = strSQL
      rstCategories.CursorLocation = adUseClient
      rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
      If rstCategories.RecordCount = 0 Then
        MsgBox "No revelant data for " & DTPicker1(x).value & "!"
        Screen.MousePointer = vbDefault
        'Exit Sub
      End If
      bChartData = True 'set flag to true
      '
      'Load the array with data
      '
      For y = 1 To rstCategories.RecordCount
        If strDateType = "hh" And rstCategories!datetype = 0 Then
          rstCategories!datetype = 24
        End If
        chrtArray(rstCategories!datetype, x + 1) = rstCategories!Calls
        rstCategories.MoveNext
      Next y
      rstCategories.Close
    Next x
    '
    'Attach the array of data to MS-CHART
    '
    With MSChart1
      .ChartData = chrtArray
      .ColumnCount = Val(cboDateNum)
      .ColumnLabelCount = Val(cboDateNum)
      For x = 0 To Val(cboDateNum) - 1
        .Column = x + 1
        .ColumnLabel = cboGroup(0).Text & " on " & DTPicker1(x).value
      Next x
      Select Case strDateType
        Case "hh"
          .Row = 1
          .RowLabel = "1 AM"
          .Row = 2
          .RowLabel = "2 AM"
          .Row = 3
          .RowLabel = "3 AM"
          .Row = 4
          .RowLabel = "4 AM"
          .Row = 5
          .RowLabel = "5 AM"
          .Row = 6
          .RowLabel = "6 AM"
          .Row = 7
          .RowLabel = "7 AM"
          .Row = 8
          .RowLabel = "8 AM"
          .Row = 9
          .RowLabel = "9 AM"
          .Row = 10
          .RowLabel = "10 AM"
          .Row = 11
          .RowLabel = "11 AM"
          .Row = 12
          .RowLabel = "12 AM"
          .Row = 13
          .RowLabel = "1 PM"
          .Row = 14
          .RowLabel = "2 PM"
          .Row = 15
          .RowLabel = "3 PM"
          .Row = 16
          .RowLabel = "4 PM"
          .Row = 17
          .RowLabel = "5 PM"
          .Row = 18
          .RowLabel = "6 PM"
          .Row = 19
          .RowLabel = "7 PM"
          .Row = 20
          .RowLabel = "8 PM"
          .Row = 21
          .RowLabel = "9 PM"
          .Row = 22
          .RowLabel = "10 PM"
          .Row = 23
          .RowLabel = "11 PM"
          .Row = 24
          .RowLabel = "12 PM"
        Case "w"
          .Row = 1
          .RowLabel = "Sun"
          .Row = 2
          .RowLabel = "Mon"
          .Row = 3
          .RowLabel = "Tue"
          .Row = 4
          .RowLabel = "Wed"
          .Row = 5
          .RowLabel = "Thur"
          .Row = 6
          .RowLabel = "Fri"
          .Row = 7
          .RowLabel = "Sat"
        Case "m"
          .Row = 1
          .RowLabel = "Jan"
          .Row = 2
          .RowLabel = "Feb"
          .Row = 3
          .RowLabel = "Mar"
          .Row = 4
          .RowLabel = "Apr"
          .Row = 5
          .RowLabel = "May"
          .Row = 6
          .RowLabel = "Jun"
          .Row = 7
          .RowLabel = "Jul"
          .Row = 8
          .RowLabel = "Aug"
          .Row = 9
          .RowLabel = "Sep"
          .Row = 10
          .RowLabel = "Oct"
          .Row = 11
          .RowLabel = "Nov"
          .Row = 12
          .RowLabel = "Dec"
        Case Else
          For y = 1 To lGreatestValue
            .Row = y
            .RowLabel = chrtArray(y, 2)
          Next y
          .Refresh
      End Select
    End With
  End If
  '
  '
  '
  '
  '
  '
  '
  If optDirection.value = True Then
    '
    ' VoiceMail
    '
    If optVoiceMail.value = True Then
      sVoiceMail = "VMSTARTTM"
      strDirection = "VoiceMail "
    Else
      sVoiceMail = "STARTTIME"
      '
      'Set the sCallDir value
      '
'      Select Case cboCallDir
'        Case "Incoming"
'          sCallDir = " (TKDIR = 2) AND "
'          strDirection = "Incoming "
'        Case "Outgoing"
'          sCallDir = " (TKDIR = 4) AND "
'          strDirection = "Outgoing "
'        Case "Both"
'          sCallDir = " "
'          strDirection = ""
'      End Select
      '
      If optCalls.value = True Then
        sCallDir = sCallDir + "VMSTARTTM = 0 AND "
      Else
        strDirection = "Total " & strDirection
      End If
      '
    End If
    '
    'Set the date values
    '
'    If DTPicker1(0).value > DTPicker2.value Then
'      MsgBox "Incorrect Date Values!"
'      Exit Sub
'    Else
'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
'        MsgBox "Incorrect Date Values!"
'        Exit Sub
'      Else
        DTStart = DTPicker1(0).value
        DTEnd = DTPicker2.value
'      End If
'    End If
    '
    'Set the DATEPART value
    '
'    Select Case cboDateType
'      Case "Hour"
'        strDateType = "hh"
'        lGreatestValue = 24
'      Case "Day"
'        strDateType = "w"
'        lGreatestValue = 7
'      Case "Week"
'        strDateType = "ww"
'        lGreatestValue = 53
'      Case "Month"
'        strDateType = "m"
'        lGreatestValue = 12
'      Case "Year"
'        'strDateType = "yyyy"
'        Exit Sub
'    End Select
    '
    'Set the Ext or Workgroup type and number
    '
    If optExt.value = True Then
      strGroupType = "P1NO"
    Else
      strGroupType = "P1WGNO"
    End If
    '                    C H A R T
    '
    ' Dynamic 2-dimensional array to store series
    ' The first index (x) is the total number of series
    ' x-axis value in the 1st slot (i.e. chrtArray(x,1)
    ' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
    '

    ReDim chrtArray(1 To lGreatestValue, 1 To 3)
    MSChart1.ShowLegend = True
    'MSChart1.chartType = VtChChartType2dLine
    '
    'Chart Title centered on top
    '
    MSChart1.Title.Text = "Calls between " & DTPicker1(0).value & " and " & DTPicker2.value
    '
    'Chart X and Y axis titles
    '
    MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle.Text = cboDateType.Text
    MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle.Text = ""
    '
    'Chart Foot note
    '
    'MSChart1.FootnoteText = "footnote"
    '
    'Load the array with 0s with correct # of rows
    '
    For x = 1 To lGreatestValue
      For y = 1 To 3
        chrtArray(x, y) = 0
      Next y
    Next x
    For x = 1 To 3
      If x = 1 Then sCallDir = " (TKDIR = 2) AND "
      If x = 2 Then sCallDir = " (TKDIR = 4) AND "
      If x = 3 Then sCallDir = " "
      '
      'Create the SQL statement
      '
      strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & cboGroup(0).Text & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & DTEnd + 1 & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
      'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
      '
      'Select revelant data
      '
      Set cmdCategories.ActiveConnection = cnMain
      cmdCategories.CommandText = strSQL
      rstCategories.CursorLocation = adUseClient
      rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
      If rstCategories.RecordCount = 0 Then
        MsgBox "No revelant data for " & cboGroup(0).Text & "!"
        Screen.MousePointer = vbDefault
        'Exit Sub
      End If
      bChartData = True 'set flag to true
      '
      'Load the array with data
      '
      For y = 1 To rstCategories.RecordCount
        If strDateType = "hh" And rstCategories!datetype = 0 Then
          rstCategories!datetype = 24
        End If
        chrtArray(rstCategories!datetype, x) = rstCategories!Calls
        rstCategories.MoveNext
      Next y
      rstCategories.Close
    Next x
    '
    'Attach the array of data to MS-CHART
    '
    With MSChart1
      .ChartData = chrtArray
      .ColumnCount = 3
      .ColumnLabelCount = 3
      .Column = 1
      .ColumnLabel = "Incoming"
      .Column = 2
      .ColumnLabel = "Outgoing"
      .Column = 3
      .ColumnLabel = "Both"
      Select Case strDateType
        Case "hh"
          .Row = 1
          .RowLabel = "1 AM"
          .Row = 2
          .RowLabel = "2 AM"
          .Row = 3
          .RowLabel = "3 AM"
          .Row = 4
          .RowLabel = "4 AM"
          .Row = 5
          .RowLabel = "5 AM"
          .Row = 6
          .RowLabel = "6 AM"
          .Row = 7
          .RowLabel = "7 AM"
          .Row = 8
          .RowLabel = "8 AM"
          .Row = 9
          .RowLabel = "9 AM"
          .Row = 10
          .RowLabel = "10 AM"
          .Row = 11
          .RowLabel = "11 AM"
          .Row = 12
          .RowLabel = "12 AM"
          .Row = 13
          .RowLabel = "1 PM"
          .Row = 14
          .RowLabel = "2 PM"
          .Row = 15
          .RowLabel = "3 PM"
          .Row = 16
          .RowLabel = "4 PM"
          .Row = 17
          .RowLabel = "5 PM"
          .Row = 18
          .RowLabel = "6 PM"
          .Row = 19
          .RowLabel = "7 PM"
          .Row = 20
          .RowLabel = "8 PM"
          .Row = 21
          .RowLabel = "9 PM"
          .Row = 22
          .RowLabel = "10 PM"
          .Row = 23
          .RowLabel = "11 PM"
          .Row = 24
          .RowLabel = "12 PM"
        Case "w"
          .Row = 1
          .RowLabel = "Sun"
          .Row = 2
          .RowLabel = "Mon"
          .Row = 3
          .RowLabel = "Tue"
          .Row = 4
          .RowLabel = "Wed"
          .Row = 5
          .RowLabel = "Thur"
          .Row = 6
          .RowLabel = "Fri"
          .Row = 7
          .RowLabel = "Sat"
        Case "m"
          .Row = 1
          .RowLabel = "Jan"
          .Row = 2
          .RowLabel = "Feb"
          .Row = 3
          .RowLabel = "Mar"
          .Row = 4
          .RowLabel = "Apr"
          .Row = 5
          .RowLabel = "May"
          .Row = 6
          .RowLabel = "Jun"
          .Row = 7
          .RowLabel = "Jul"
          .Row = 8
          .RowLabel = "Aug"
          .Row = 9
          .RowLabel = "Sep"
          .Row = 10
          .RowLabel = "Oct"
          .Row = 11
          .RowLabel = "Nov"
          .Row = 12
          .RowLabel = "Dec"
        Case Else
          For y = 1 To lGreatestValue
            .Row = y
            .RowLabel = chrtArray(y, 2)
          Next y
          .Refresh
      End Select
    End With
  End If
  '
  '
  '
  '
  
If optNone.value = True Then

  intWorkgroup = cboGroup(0).Text
'
' VoiceMail
'
    If optVoiceMail.value = True Then
      sVoiceMail = "VMSTARTTM"
      strDirection = "VoiceMail "
    Else
      sVoiceMail = "STARTTIME"
  
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
          strDirection = ""
      End Select
      '
      If optCalls.value = True Then
        sCallDir = sCallDir + "VMSTARTTM = 0 AND "
      Else
        strDirection = "Total " & strDirection
      End If
      '
  End If
  
'
'Set the date values
'
'    If DTPicker1(0).value > DTPicker2.value Then
'        MsgBox "Incorrect Date Values!"
'        Exit Sub
'    Else
'        If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
'            MsgBox "Incorrect Date Values!"
'            Exit Sub
'        Else
            DTStart = DTPicker1(0).value
            DTEnd = DTPicker2.value
'        End If
'    End If
'
'Set the DATEPART value
'
'    Select Case cboDateType
'        Case "Hour"
'            strDateType = "hh"
'            lGreatestValue = 24
'        Case "Day"
'            strDateType = "w"
'            lGreatestValue = 7
'        Case "Week"
'            strDateType = "ww"
'            lGreatestValue = 53
'        Case "Month"
'            strDateType = "m"
'            lGreatestValue = 12
'        Case "Year"
'            'strDateType = "yyyy"
'            Exit Sub
'    End Select
'
'Set the Ext or Workgroup type and number
'
    If optExt.value = True Then
        strGroupType = "P1NO"
    Else
        strGroupType = "P1WGNO"
    End If
    
    


'
'Create the SQL statement
'
    strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & intWorkgroup & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & DTEnd + 1 & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
    'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
'Select revelant data

    Set cmdCategories.ActiveConnection = cnMain
    cmdCategories.CommandText = strSQL
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic

If rstCategories.RecordCount = 0 Then
    MsgBox "No revelant data!"
    Screen.MousePointer = vbDefault
    'Exit Sub
End If
bChartData = True 'set flag to true
'                    C H A R T
'
' Dynamic 2-dimensional array to store series
' The first index (x) is the total number of series
' x-axis value in the 1st slot (i.e. chrtArray(x,1)
' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
'
ReDim chrtArray(1 To lGreatestValue, 1 To 2)
MSChart1.ShowLegend = True
'MSChart1.chartType = VtChChartType2dLine
'
'Chart Title centered on top
'
MSChart1.Title.Text = strDirection & "Calls between " & DTPicker1(0).value & " and " & DTPicker2.value
'
'Chart X and Y axis titles
'
MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle.Text = cboDateType.Text
MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle.Text = ""
'
'Chart Foot note
'
'MSChart1.FootnoteText = "footnote"
'
'Load the array with 0s with correct # of rows
'
For x = 1 To lGreatestValue
  chrtArray(x, 1) = 0
  'chrtArray(X, 2) = X
Next x
'
'Load the array with data
'
For x = 1 To rstCategories.RecordCount
  If strDateType = "hh" And rstCategories!datetype = 0 Then
    rstCategories!datetype = 24
  End If
  chrtArray(rstCategories!datetype, 1) = rstCategories!Calls
  rstCategories.MoveNext
Next x
'
'Attach the array of data to MS-CHART
'
With MSChart1
    .ChartData = chrtArray
    .ColumnCount = 1
    .ColumnLabelCount = 1
    .Column = 1
    .ColumnLabel = "Calls for " & cboGroup(0).Text
    Select Case strDateType
        Case "hh"
          .Row = 1
          .RowLabel = "1 AM"
          .Row = 2
          .RowLabel = "2 AM"
          .Row = 3
          .RowLabel = "3 AM"
          .Row = 4
          .RowLabel = "4 AM"
          .Row = 5
          .RowLabel = "5 AM"
          .Row = 6
          .RowLabel = "6 AM"
          .Row = 7
          .RowLabel = "7 AM"
          .Row = 8
          .RowLabel = "8 AM"
          .Row = 9
          .RowLabel = "9 AM"
          .Row = 10
          .RowLabel = "10 AM"
          .Row = 11
          .RowLabel = "11 AM"
          .Row = 12
          .RowLabel = "12 AM"
          .Row = 13
          .RowLabel = "1 PM"
          .Row = 14
          .RowLabel = "2 PM"
          .Row = 15
          .RowLabel = "3 PM"
          .Row = 16
          .RowLabel = "4 PM"
          .Row = 17
          .RowLabel = "5 PM"
          .Row = 18
          .RowLabel = "6 PM"
          .Row = 19
          .RowLabel = "7 PM"
          .Row = 20
          .RowLabel = "8 PM"
          .Row = 21
          .RowLabel = "9 PM"
          .Row = 22
          .RowLabel = "10 PM"
          .Row = 23
          .RowLabel = "11 PM"
          .Row = 24
          .RowLabel = "12 PM"
        Case "w"
          .Row = 1
          .RowLabel = "Sun"
          .Row = 2
          .RowLabel = "Mon"
          .Row = 3
          .RowLabel = "Tue"
          .Row = 4
          .RowLabel = "Wed"
          .Row = 5
          .RowLabel = "Thur"
          .Row = 6
          .RowLabel = "Fri"
          .Row = 7
          .RowLabel = "Sat"
        Case "m"
          .Row = 1
          .RowLabel = "Jan"
          .Row = 2
          .RowLabel = "Feb"
          .Row = 3
          .RowLabel = "Mar"
          .Row = 4
          .RowLabel = "Apr"
          .Row = 5
          .RowLabel = "May"
          .Row = 6
          .RowLabel = "Jun"
          .Row = 7
          .RowLabel = "Jul"
          .Row = 8
          .RowLabel = "Aug"
          .Row = 9
          .RowLabel = "Sep"
          .Row = 10
          .RowLabel = "Oct"
          .Row = 11
          .RowLabel = "Nov"
          .Row = 12
          .RowLabel = "Dec"
        Case Else
          For x = 1 To lGreatestValue
            .Row = x
            .RowLabel = x '"Week " & X
            '.RowLabel = chrtArray(X, 2)
          Next x
    .Refresh
    End Select
End With
End If

If optCallVsVoice.value = True Then
  CallVsVoice
End If
  '
  If chkLablePoints.value = vbChecked Then LablePoints
  '
  '
  'Enable Print button
  '
  cmdPrintChart.Enabled = True
  '
  Screen.MousePointer = vbDefault
  '
End Sub

Private Sub cmdPrintChart_Click()
  On Error GoTo PrintErrHandler
  dlgCommon.CancelError = True
  dlgCommon.ShowPrinter
  Printer.PaperSize = vbPRPSLetter
  Printer.Orientation = vbPRORLandscape
  Printer.Copies = dlgCommon.Copies
  MSChart1.EditCopy
  Printer.Print " "
  'Printer.PaintPicture Clipboard.GetData(), 0, 0
  Printer.PaintPicture Clipboard.GetData(), 150, 0, 15000, 12000
  Printer.EndDoc
  Exit Sub
  
PrintErrHandler:
   Select Case Err.Number
   Case 32755
       'MsgBox "Print cancelled."
   End Select
   Exit Sub
End Sub

Private Sub DTPicker1_Change(Index As Integer)
  '
  If cboDateType = "Day" Then
    MonthView1.value = DTPicker1(Index).value
    MonthView1.DayOfWeek = mvwSunday
    DTPicker1(Index).value = MonthView1.value
  End If
  '
  If Index = 0 Then
    cboDateType_Click
  Else
    SetDate Index
  End If
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
End Sub

Private Sub DTPicker2_Change()
'
'Disable Print button
'
  cmdPrintChart.Enabled = False
End Sub

Private Sub Form_Initialize()

  On Error GoTo ErrCall
  '
  Set FormControl = New CFormControl
  '
  FormControl.MinHeight = Me.Height
  FormControl.MinWidth = Me.Width
  FormControl.DataForm = False
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmSelect.Form_Initialize.", vbCritical, "Error"
End Sub

Private Sub Form_Load()
  Dim x As Integer
  Dim y As Integer
  'frmMultiChart.Top = (Screen.Height - frmMultiChart.Height) / 2
  'frmMultiChart.Left = (Screen.Width - frmMultiChart.Width) / 2
  optExt.value = True
  optNone.value = True
  optBoth.value = True
  For x = 0 To 9
  DTPicker1(x).value = Date
  DTPicker2.value = Date
  DTPicker3.value = Date
  DTPicker4.value = Date
  Next x
  cboDateType_Click
  optTotal.value = True
  '
'  Y = 0
'  For X = 2003 To 2010
'    cboYear.AddItem X, Y
'    Y = Y + 1
'  Next
'  '
''  Y = 0
''  For x = 1 To 53
''    cboSunday.AddItem x, Y
''    Y = Y + 1
''  Next
'  '
'  Y = 0
'  For X = 1 To 12
'    cboMonth.AddItem X, Y
'    Y = Y + 1
'  Next
'  cboYear.ListIndex = 0
'  cboMonth.ListIndex = 0
'  GetSundays
'  '
'  bChartData = False 'set flag to false
'  cboDateType_Click
End Sub


Private Sub MSChart1_DblClick()
'  If bChartData = True Then
'    frmBigChart.Show
'    frmBigChart.MSChart2.ChartData = MSChart1.ChartData
'    frmBigChart.MSChart2.chartType = MSChart1.chartType
'    frmBigChart.MSChart2.Title = MSChart1.Title
'  End If
End Sub


Private Sub optAvg_Click()
  DTPicker1(0).Visible = False
  DTPicker2.Visible = False
  DTPicker3.Visible = True
  DTPicker4.Visible = True
  fraMultiLines.Enabled = False
  optNone.Enabled = False
  optMultiGroups.Enabled = False
  optMultiDates.Enabled = False
  optDirection.Enabled = False
  optCallVsVoice.Enabled = False
  cmdAvg.Visible = True
  cmdChart.Visible = False
  fraGroups.Visible = False
  fraMultiDates.Visible = False
  cmdPrintChart.Enabled = False
  optNone.value = True
End Sub

Private Sub optBoth_Click()
  If optExt.value = True Then
    cboCallDir.Enabled = True
  End If
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
End Sub

Private Sub optCalls_Click()
  If optExt.value = True Then
    cboCallDir.Enabled = True
  End If
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
End Sub

Private Sub optCallVsVoice_Click()
  fraGroups.Visible = False
  fraMultiDates.Visible = False
  fraCallType.Enabled = False
  optBoth.Enabled = False
  optCalls.Enabled = False
  optVoiceMail.Enabled = False
  '
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
End Sub

Private Sub optExt_Click()
    Dim cmdCategories As New ADODB.Command
    Dim rstCategories As New ADODB.Recordset
    Dim x As Integer
  '
  cmdCategories.CommandTimeout = 300
'
'Clear the combo box
'
  For x = 0 To 9
    cboGroup(x).Clear
  Next x
'
'Select all Ext#s and put in combo box
'
    Set cmdCategories.ActiveConnection = cnMain
    cmdCategories.CommandText = "SELECT P1NO FROM ICC_CDR GROUP BY P1NO ORDER BY P1NO"
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
'
'Add Ext# to drop down cbo
'
    Do While Not rstCategories.eof
        If rstCategories!P1NO <> "" Then
            For x = 0 To 9
            cboGroup(x).AddItem rstCategories!P1NO
          Next x
        End If
        rstCategories.MoveNext
    Loop
    Set cmdCategories = Nothing
    rstCategories.Close
'
'Set the first Ext# as default
'
    For x = 0 To 9
    cboGroup(x).ListIndex = 0
  Next x
'
'Disable Print button
'
  cmdPrintChart.Enabled = False
'
'Enable call direction
'
  If optVoiceMail.value = False Then
    cboCallDir.Enabled = True
  End If
End Sub

Private Sub optDirection_Click()
  DTPicker1(0).Visible = True
  DTPicker2.Visible = True
  DTPicker3.Visible = False
  DTPicker4.Visible = False
  cmdAvg.Visible = False
  cmdChart.Visible = True
  fraGroups.Visible = False
  fraMultiDates.Visible = False
  'frmMultiChart.Width = 10440
  fraCallType.Enabled = True
  optBoth.Enabled = True
  optCalls.Enabled = True
  optVoiceMail.Enabled = True
  '
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
End Sub

Private Sub optMultiDates_Click()
  DTPicker1(0).Visible = True
  DTPicker2.Visible = True
  DTPicker3.Visible = False
  DTPicker4.Visible = False
  cmdAvg.Visible = False
  cmdChart.Visible = True
  Dim x As Integer
  For x = 1 To 9
    DTPicker1_Change x
  Next
  fraGroups.Visible = False
  fraMultiDates.Visible = True
  'frmMultiChart.Width = 12705
  fraCallType.Enabled = True
  optBoth.Enabled = True
  optCalls.Enabled = True
  optVoiceMail.Enabled = True
  '
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
End Sub

Private Sub optMultiGroups_Click()
  DTPicker1(0).Visible = True
  DTPicker2.Visible = True
  DTPicker3.Visible = False
  DTPicker4.Visible = False
  cmdAvg.Visible = False
  cmdChart.Visible = True
  fraGroups.Visible = True
  fraMultiDates.Visible = False
  'frmMultiChart.Width = 12705
  fraCallType.Enabled = True
  optBoth.Enabled = True
  optCalls.Enabled = True
  optVoiceMail.Enabled = True
  '
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
End Sub

Private Sub optNone_Click()
  fraGroups.Visible = False
  fraMultiDates.Visible = False
  'frmMultiChart.Width = 10440
  fraCallType.Enabled = True
  optBoth.Enabled = True
  optCalls.Enabled = True
  optVoiceMail.Enabled = True
  '
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
End Sub

Private Sub optTotal_Click()
  DTPicker1(0).Visible = True
  DTPicker2.Visible = True
  DTPicker3.Visible = False
  DTPicker4.Visible = False
  cmdAvg.Visible = False
  cmdChart.Visible = True
  fraMultiLines.Enabled = True
  optNone.Enabled = True
  optMultiGroups.Enabled = True
  optMultiDates.Enabled = True
  optDirection.Enabled = True
  optCallVsVoice.Enabled = True
End Sub

Private Sub optVoiceMail_Click()
  '
  'Disable Print button
  '
  cmdPrintChart.Enabled = False
  '
  cboCallDir.Text = "Incoming"
  cboCallDir.Enabled = False
  '
End Sub

'Private Sub optMean_Click()
'  MSChart1.Plot.SeriesCollection(1).StatLine.Flag = VtChStatsMean
'End Sub
'
'Private Sub optMinMax_Click()
'  MSChart1.Plot.SeriesCollection(1).StatLine.Flag = VtChStatsMinimum Or VtChStatsMaximum
'End Sub
'
'Private Sub optRegression_Click()
'  MSChart1.Plot.SeriesCollection(1).StatLine.Flag = VtChStatsRegression
'End Sub
'
'Private Sub optStdDev_Click()
'  MSChart1.Plot.SeriesCollection(1).StatLine.Flag = VtChStatsStddev
'End Sub

Private Sub optWorkgroup_Click()
    Dim cmdCategories As New ADODB.Command
    Dim rstCategories As New ADODB.Recordset
    Dim x As Integer
'
'Clear the combo box
'
  For x = 0 To 9
    cboGroup(x).Clear
  Next x
'
'Select all Workgroup#s and put in combo box
'
    Set cmdCategories.ActiveConnection = cnMain
    cmdCategories.CommandText = "SELECT P1WGNO FROM ICC_CDR GROUP BY P1WGNO ORDER BY P1WGNO"
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
'
'Add Workgroup# to drop down cbo
'
    Do While Not rstCategories.eof
        If rstCategories!P1WGNO <> "" Then
          For x = 0 To 9
            cboGroup(x).AddItem rstCategories!P1WGNO
          Next x
        End If
        rstCategories.MoveNext
    Loop
    Set cmdCategories = Nothing
    rstCategories.Close
'
'Set the first Workgroup# as default id
'
  For x = 0 To 9
    cboGroup(x).ListIndex = 0
  Next x
'
'Disable Print button
'
  cmdPrintChart.Enabled = False
'
'Disable call direction and set to incoming. There are no outgoing calls from a workgroup only Ext
'
  cboCallDir.Text = "Incoming"
  cboCallDir.Enabled = False
End Sub

Private Sub GetSundays()
  Dim x As Integer
  Dim y As Integer
  Dim Z As Integer
  '
  cboSunday.Clear
  MonthView1.year = cboYear.Text
  MonthView1.month = cboMonth.Text
  '
  Select Case cboMonth.Text
    Case 4, 6, 9, 11
      y = 30
    Case 1, 3, 5, 7, 8, 10, 12
      y = 31
    Case Else
      If cboYear.Text = 2004 Then
        y = 29
      Else
        y = 28
      End If
  End Select
  '
  Z = 0
  For x = 1 To y
    MonthView1.day = x
    If MonthView1.DayOfWeek = 1 Then
      cboSunday.AddItem x, Z
      Z = Z + 1
    End If
  Next
  '
  cboSunday.ListIndex = 0
  MonthView1.day = cboSunday.Text
  '
End Sub

Private Function SetDate(iControl As Integer)
Dim sTemp As String
  '
  Select Case cboDateType
    Case "Hour"
      'DTPicker1_Change iControl
      'DTPicker2.value = DTPicker1(iControl).value + 1
    Case "Day"
      'DTPicker1_Change iControl
      'DTPicker2.value = DTPicker1(iControl).value + 7
    Case Else
      sTemp = DTPicker1(iControl).year
      DTPicker1(iControl).value = "1/1/" & sTemp
  End Select
  '
End Function

Private Sub CallVsVoice()
  '
  Dim sVoiceMail As String
  Dim sCallDir As String
  Dim x As Integer
  Dim y As Integer
  Dim iInterval(1 To 100) As Integer
  Dim IAvg As Integer
  Dim cmdCategories As New ADODB.Command
  Dim rstCategories As New ADODB.Recordset
  '
  ' VoiceMail
  '
'  If optVoiceMail.Value = True Then
'    sVoiceMail = "VMSTARTTM"
'    strDirection = "VoiceMail "
'  Else
'    sVoiceMail = "STARTTIME"
'    '
'    'Set the sCallDir value
'    '
'      Select Case cboCallDir
'        Case "Incoming"
'          sCallDir = " (TKDIR = 2) AND "
'          strDirection = "Incoming "
'        Case "Outgoing"
'          sCallDir = " (TKDIR = 4) AND "
'          strDirection = "Outgoing "
'        Case "Both"
'          sCallDir = " "
'          strDirection = ""
'      End Select
'  End If
  '
  'Set the date values
  '
'    If DTPicker1(0).value > DTPicker2.value Then
'      MsgBox "Incorrect Date Values!"
'      Exit Sub
'    Else
'      If DTPicker1(0).value > Date Or DTPicker2.value > Date Then
'        MsgBox "Incorrect Date Values!"
'        Exit Sub
'      Else
      DTStart = DTPicker1(0).value
      DTEnd = DTPicker2.value
'      End If
'    End If
  '
  'Set the DATEPART value
  '
    Select Case cboDateType
      Case "Hour"
        strDateType = "hh"
        lGreatestValue = 24
      Case "Day"
        strDateType = "w"
        lGreatestValue = 7
      Case "Week"
        strDateType = "ww"
        lGreatestValue = 53
      Case "Month"
        strDateType = "m"
        lGreatestValue = 12
      Case "Year"
        'strDateType = "yyyy"
        Exit Sub
    End Select
  '
  'Set the Ext or Workgroup type and number
  '
  If optExt.value = True Then
    strGroupType = "P1NO"
  Else
    strGroupType = "P1WGNO"
  End If
  '                    C H A R T
  '
  ' Dynamic 2-dimensional array to store series
  ' The first index (x) is the total number of series
  ' x-axis value in the 1st slot (i.e. chrtArray(x,1)
  ' y-axis value in the 2nd slot (i.e. chrtArray(x,2)
  '

  ReDim chrtArray(1 To lGreatestValue, 1 To 2)
  MSChart1.ShowLegend = True
  'MSChart1.chartType = VtChChartType2dLine
  '
  'Chart Title centered on top
  '
  MSChart1.Title.Text = "Calls between " & DTPicker1(0).value & " and " & DTPicker2.value
  '
  'Chart X and Y axis titles
  '
  MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle.Text = cboDateType.Text
  MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle.Text = ""
  '
  'Chart Foot note
  '
  'MSChart1.FootnoteText = "footnote"
  '
  'Load the array with 0s with correct # of rows
  '
  For x = 1 To lGreatestValue
    For y = 1 To 2
      chrtArray(x, y) = 0
    Next y
  Next x
  For x = 1 To 2
    If x = 1 Then
      sVoiceMail = "STARTTIME"
      strDirection = ""
      sCallDir = sCallDir + "VMSTARTTM = 0 AND "
    End If
    If x = 2 Then
    sVoiceMail = "VMSTARTTM"
    strDirection = "VoiceMail "
    End If
    'sCallDir = sCallDir + " "
    '
    'Create the SQL statement
    '
    strSQL = "SELECT DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DISTINCT SESSID) AS Calls FROM ICC_CDR WHERE " & sCallDir & "  (" & strGroupType & " LIKE '" & cboGroup(0).Text & "')AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') > ('" & DTStart & "')) AND (DATEADD(ss, " & sVoiceMail & ", '1969-12-31 19:00:00') < ('" & DTEnd + 1 & "')) GROUP BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(" & strDateType & ", DATEADD(ss, " & sVoiceMail & ", CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
    sCallDir = " "
    'MsgBox strSQL & vbCrLf & "SELECT     DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS DateType, COUNT(DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) AS calls FROM         ICC_CDR WHERE     (P1WGNO LIKE N'765') GROUP BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102))) ORDER BY DATEPART(ww, DATEADD(ss, STARTTIME, CONVERT(DATETIME, '1969-12-31 19:00:00', 102)))"
    '
    'Select revelant data
    '
    Set cmdCategories.ActiveConnection = cnMain
    cmdCategories.CommandText = strSQL
    rstCategories.CursorLocation = adUseClient
    rstCategories.Open cmdCategories, , adOpenStatic, adLockBatchOptimistic
    If rstCategories.RecordCount = 0 Then
      MsgBox "No revelant data for " & cboGroup(0).Text & "!"
      Screen.MousePointer = vbDefault
      'Exit Sub
    End If
    bChartData = True 'set flag to true
    '
    'Load the array with data
    '
    For y = 1 To rstCategories.RecordCount
      If strDateType = "hh" And rstCategories!datetype = 0 Then
        rstCategories!datetype = 24
      End If
      chrtArray(rstCategories!datetype, x) = rstCategories!Calls
      rstCategories.MoveNext
    Next y
    rstCategories.Close
  Next x
  '
  'Attach the array of data to MS-CHART
  '
  With MSChart1
    .ChartData = chrtArray
    .ColumnCount = 2
    .ColumnLabelCount = 2
    .Column = 1
    .ColumnLabel = "Calls"
    .Column = 2
    .ColumnLabel = "VoiceMails"
    Select Case strDateType
      Case "hh"
        .Row = 1
        .RowLabel = "1 AM"
        .Row = 2
        .RowLabel = "2 AM"
        .Row = 3
        .RowLabel = "3 AM"
        .Row = 4
        .RowLabel = "4 AM"
        .Row = 5
        .RowLabel = "5 AM"
        .Row = 6
        .RowLabel = "6 AM"
        .Row = 7
        .RowLabel = "7 AM"
        .Row = 8
        .RowLabel = "8 AM"
        .Row = 9
        .RowLabel = "9 AM"
        .Row = 10
        .RowLabel = "10 AM"
        .Row = 11
        .RowLabel = "11 AM"
        .Row = 12
        .RowLabel = "12 AM"
        .Row = 13
        .RowLabel = "1 PM"
        .Row = 14
        .RowLabel = "2 PM"
        .Row = 15
        .RowLabel = "3 PM"
        .Row = 16
        .RowLabel = "4 PM"
        .Row = 17
        .RowLabel = "5 PM"
        .Row = 18
        .RowLabel = "6 PM"
        .Row = 19
        .RowLabel = "7 PM"
        .Row = 20
        .RowLabel = "8 PM"
        .Row = 21
        .RowLabel = "9 PM"
        .Row = 22
        .RowLabel = "10 PM"
        .Row = 23
        .RowLabel = "11 PM"
        .Row = 24
        .RowLabel = "12 PM"
      Case "w"
        .Row = 1
        .RowLabel = "Sun"
        .Row = 2
        .RowLabel = "Mon"
        .Row = 3
        .RowLabel = "Tue"
        .Row = 4
        .RowLabel = "Wed"
        .Row = 5
        .RowLabel = "Thur"
        .Row = 6
        .RowLabel = "Fri"
        .Row = 7
        .RowLabel = "Sat"
      Case "m"
        .Row = 1
        .RowLabel = "Jan"
        .Row = 2
        .RowLabel = "Feb"
        .Row = 3
        .RowLabel = "Mar"
        .Row = 4
        .RowLabel = "Apr"
        .Row = 5
        .RowLabel = "May"
        .Row = 6
        .RowLabel = "Jun"
        .Row = 7
        .RowLabel = "Jul"
        .Row = 8
        .RowLabel = "Aug"
        .Row = 9
        .RowLabel = "Sep"
        .Row = 10
        .RowLabel = "Oct"
        .Row = 11
        .RowLabel = "Nov"
        .Row = 12
        .RowLabel = "Dec"
      Case Else
        For y = 1 To lGreatestValue
          .Row = y
          .RowLabel = chrtArray(y, 2)
        Next y
        .Refresh
    End Select
  End With
  'If chkLablePoints.Value = vbChecked Then LablePoints
End Sub

Private Sub LablePoints()
  Dim y As Integer
  With MSChart1
    For y = 1 To .Plot.SeriesCollection.Count
      With .Plot.SeriesCollection(y).DataPoints(-1).DataPointLabel
          .LocationType = VtChLabelLocationTypeAbovePoint
          .Component = VtChLabelComponentValue
          '.PercentFormat = "0%"
          .VtFont.Style = 1
          .VtFont.Size = 10
      End With
    Next y
  End With
End Sub
