VERSION 5.00
Object = "{F83FB95C-D981-11D2-A80A-00104BF191A4}#1.0#0"; "SKCL.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm FMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Bread 'n' Butter"
   ClientHeight    =   6660
   ClientLeft      =   3255
   ClientTop       =   2670
   ClientWidth     =   9135
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9135
   End
   Begin MSComctlLib.ImageList TrayIconList3 
      Left            =   1560
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":1EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":220E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":2528
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":2842
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":2B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":2E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":3190
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":34AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":37C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":3ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":3DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":4112
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":442C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":4746
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":4A60
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList TrayIconList2 
      Left            =   480
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   48
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":4D7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":5094
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":53AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":56C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":59E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":5CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":6016
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":6330
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":664A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":6964
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":6C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":6F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":72B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":75CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":78E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":7C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":7F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":8234
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":854E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":8868
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":8B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":8E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":91B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":94D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":97EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":9B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":9E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":A138
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":A452
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":A76C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":AA86
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":ADA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":B0BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":B3D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":B6EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":BA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":BD22
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":C03C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":C356
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":C670
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":C98A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":CCA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":CFBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":D2D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":D5F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":D90C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":DC26
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":DF40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList TrayIconList 
      Left            =   480
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":E25A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":E3B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":E50E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":E668
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":E7C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":E91C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":EA76
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":EBD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":ED2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":EE84
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":EFDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":F138
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":F292
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":F3EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":F546
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":F6A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":F7FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":F954
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":FAAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":FC08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrMessages 
      Interval        =   30000
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer tmrSecChk 
      Interval        =   60000
      Left            =   330
      Top             =   3480
   End
   Begin ActiveToolBars.SSActiveToolBars tbMain 
      Left            =   300
      Top             =   2490
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      PictureBackgroundStyle=   2
      ToolBarsCount   =   3
      ToolsCount      =   52
      PersonalizedMenus=   0
      DragFullToolBars=   2
      PictureBackgroundUseMask=   -1  'True
      Tools           =   "FMain.frx":FD62
      ToolBars        =   "FMain.frx":29937
   End
   Begin VB.Timer tmrTray 
      Interval        =   150
      Left            =   300
      Top             =   1380
   End
   Begin SKCLLibCtl.LFile License 
      Left            =   330
      OleObjectBlob   =   "FMain.frx":29F3A
      Top             =   3000
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private TrayHook As New CTrayHook
Dim iMessageCount1 As Integer
Dim iMessageCount2 As Integer
Dim iMessageCount3 As Integer
Dim iMessageCount4 As Integer
Dim iFlasher As Integer
Dim iVMailTotal As Integer


Private Sub MDIForm_Initialize()
  On Error GoTo ErrCall
  '
  'RemoveCancelMenuItem Me
  'TrayHook.Setup Me, tmrTray, scTray, mnuTray
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.MDIForm_Initialize.", vbCritical, "Error"
End Sub

Private Sub MDIForm_Load()
  On Error GoTo ErrCall
  '
  '
  Me.Caption = "Bread 'n' Butter" '"Track It!   Datafile: " & DBOps.DBName
  '
'  With sbMain
'  .Panels.Add
'  .Panels.Add
'  .Panels.Add
'  '
'  .Panels(1).Width = (Me.Width / Screen.TwipsPerPixelX) - 200
'  .Panels(1).Text = dbmMain.DbPath & dbmMain.DBName
'  .Panels(2).Style = sbrDate
'  .Panels(3).Style = sbrTime
'  End With
  '
  RemoveCancelMenuItem Me 'disables the X
  '
  If ConnType = Access Then
    tbMain.ToolBars(2).Tools("ID_PathFile").Name = "Database: " & sAccessDB
  Else
    tbMain.ToolBars(2).Tools("ID_PathFile").Name = "Server: " & sSQLServerName & "  Database: " & sSQLServerDB
  End If
  tbMain.ToolBars(2).Tools("ID_UserName").Name = "User: " & StrUser
  
  '
  InitLicense
  '
  tmrMessages_Timer
  '
  'the form must be fully visible before calling Shell_NotifyIcon
  Me.Show
  'Company.Contact.SearchID = 0
  FormMgr.ShowForm Me.ActiveForm, FContact, True
  'Me.Refresh
  With nid
   .cbSize = Len(nid)
   .hWnd = Me.hWnd
   .uID = vbNull
   .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   .uCallbackMessage = WM_MOUSEMOVE
   .hIcon = Me.Icon
   .szTip = "Bread'n'Butter" & vbNullChar
  End With
  Shell_NotifyIcon NIM_ADD, nid
  '
  'Me.WindowState = 1
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.MDIForm_Load.", vbCritical, "Error"
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  'this procedure receives the callbacks from the System Tray icon.
  Dim Result As Long
  Dim msg As Long
  'Debug.Print X, X / Screen.TwipsPerPixelX

  msg = X / Screen.TwipsPerPixelX
  '
  Select Case msg
   Case WM_LBUTTONUP        '514 restore form window
    'Me.WindowState = vbMaximized
    'Result = SetForegroundWindow(Me.hwnd)
    'Me.Show
   Case WM_LBUTTONDBLCLK    '515 restore form window
    Me.WindowState = vbMaximized
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
   Case WM_RBUTTONUP        '517 display popup menu
    'Result = SetForegroundWindow(Me.hwnd)
    'Me.PopupMenu Me.mnuTray
  End Select
  If Button = 2 And FMain.WindowState = 1 Then
    Result = SetForegroundWindow(Me.hWnd)
    Me.PopupMenu Me.mnuTray
  End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ErrCall
  '
  If UnloadMode <> vbFormControlMenu Then
    Unload FLogon
    FLogon.Mode = enLogout
    FLogon.Show vbModal
    Cancel = Not User.LogResults
    '
    If Cancel = 0 Then
      Shell_NotifyIcon NIM_DELETE, nid
      'End ' Will this work?
    End If
    '
  End If
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.MDIForm_QueryUnload.", vbCritical, "Error"
End Sub

Private Sub MDIForm_Resize()
  On Error GoTo ErrCall
  '
  If Not Me.ActiveForm Is Nothing Then
    FormMgr.ResizeForm Me.ActiveForm
  End If
  '
  'sbMain.Panels(1).Width = (Me.Width / Screen.TwipsPerPixelX) - 200
  '
  'this is necessary to assure that the minimized window is hidden
  If Me.WindowState = vbMinimized Then Me.Hide
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.MDIForm_Resize.", vbCritical, "Error"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 ' End
  'this removes the icon from the system tray
  'Shell_NotifyIcon NIM_DELETE, nid
  'Cancel = 1
End Sub

Private Sub mnuTrayExit_Click()
  Me.WindowState = vbMaximized
  Me.Show
  Unload Me
End Sub

Private Sub mnuTrayOpen_Click()
  'called when the user clicks the popup menu Restore command
  Dim Result As Long
  Me.WindowState = vbMaximized
  Result = SetForegroundWindow(Me.hWnd)
  Me.Show
End Sub
Public Sub tbMain_Go(sResult As String)
  On Error GoTo ErrCall
  '
  Dim bActiveForm As Boolean
  '
  bActiveForm = Not Me.ActiveForm Is Nothing
  If bActiveForm = True Then
    SaveSetting App.Title, "Miscellaneous", "ActiveForm", Me.ActiveForm.Name
  Else
    SaveSetting App.Title, "Miscellaneous", "ActiveForm", vbNullString
  End If
  '
'  If sResult = "ID_Lookup" Then
'        Company.Contact.SearchID = 0
'        FormMgr.ShowForm Me.ActiveForm, FContact
'  End If
Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.tbMain_ToolClick.", vbCritical, "Error"

End Sub

Public Sub tbMain_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
  On Error GoTo ErrCall
  '
  Dim Employee As New CEmployee
  Dim bActiveForm As Boolean
  '
  bActiveForm = Not Me.ActiveForm Is Nothing
  If bActiveForm = True Then
    SaveSetting App.Title, "Miscellaneous", "ActiveForm", Me.ActiveForm.Name
  Else
    SaveSetting App.Title, "Miscellaneous", "ActiveForm", vbNullString
  End If
  '
  Select Case Tool.ID
  '\\ File Menu
'  Case "ID_Cases"
'    bCases = True
'    FormMgr.ShowForm Me.ActiveForm, FCase, True
  Case "ID_VMail"
   ' InitializeVmail
    FormMgr.ShowForm Me.ActiveForm, FVMail, True
'  Case "ID_Reports"
'    FormMgr.ShowForm Me.ActiveForm, FReports, True
  Case "ID_HistoryReporter"
    FormMgr.ShowForm Me.ActiveForm, FHistory, True
  Case "ID_OpenCalls"
    If Employee.InGroup(User.Name, "Management") = True Or Employee.InGroup(User.Name, "Development") = True Then
      FormMgr.ShowForm Me.ActiveForm, FSupportOpen, True
    Else
      MsgBox "Access denied.", vbCritical, ""
    End If
  Case "ID_ContactReporter"
    FormMgr.ShowForm Me.ActiveForm, FReport, True
  Case "ID_LicenseFacility"
     FormMgr.ShowForm Me.ActiveForm, FLicense, True
  Case "ID_CallLog"
    FormMgr.ShowForm Me.ActiveForm, FSupportLog, True
  Case "ID_AuthorizationLog"
    FormMgr.ShowForm Me.ActiveForm, FAuthLog, True
  Case "ID_Customer"
    FormMgr.ShowForm Me.ActiveForm, FContact, True
  Case "ID_Close"
    If bActiveForm Then
      Me.ActiveForm.FormControl.SwitchFrom
      Unload Me.ActiveForm
      Me.BackColor = &H8000000C
    End If
'  Case "ID_ProspectView" 'Prospecting View
'    FormMgr.ShowForm Me.ActiveForm, FProspecting, True
  Case "ID_ProspectMgt" 'Prospect Management
    FormMgr.ShowForm Me.ActiveForm, FProspectMgt, True
  Case "ID_Pricing" 'Pricing
    FPricing.Show vbModal
  Case "ID_Lookup" 'Customer Lookup
    'Company.Contact.SearchID = 0
    FormMgr.ShowForm Me.ActiveForm, FContact, True
 ' Case "ID_KB"
  '  FKB.Height = Me.Height - 1500
   ' FKB.Width = Me.Width - 1500
    'FKB.Show vbModal, Me
  Case "ID_TechSupport"
    'frmTechSupport.Show vbModal, Me
  'Case "ID_List"
    'FormMgr.ShowForm Me.ActiveForm, FCustomerList, True
  Case "ID_LogIn"
    FLogon.Mode = enLogin
    FLogon.Show vbModal
  Case "ID_Logout"
    FLogon.Mode = enLogout
    FLogon.Show vbModal
  Case "ID_Prefs"
    FPrefs.Show vbModal, Me
  Case "ID_Hours"
    If Employee.InGroup(User.Name, "Management") = True Or Employee.InGroup(User.Name, "Development") = True Then
      'FHours.Show vbModal
      FormMgr.ShowForm Me.ActiveForm, FHours, True
    Else
      MsgBox "Access denied.", vbCritical, ""
    End If
  Case "ID_Password"
    SetPassword
  'Case "ID_BatchHist"
    'FBatchHistory.Show vbModal, Me
  Case "ID_MailingLbl"
    FPrintLabels.Show vbModal, Me
  Case "ID_SelectCust"
    FormMgr.ShowForm Me.ActiveForm, FSelect, True
  Case "ID_ShipRpt"
    RShipping.Show
  Case "ID_CustStatus"
    RCustStatus.DBName DBOps.DBName
    RCustStatus.Show
  Case "ID_OpenDB"
    'OpenDatabase
  Case "ID_Exit"
    Unload Me
    Unload FVMail
  Case "ID_Utility"
    FUtility.Show , Me
'  Case "ID_EMail"
'    FEMail.Show , Me
  Case "ID_CallTimes"
    FormMgr.ShowForm Me.ActiveForm, FCallStats, True
    'FCallStats.Show , Me
  Case "ID_PhoneChart"
    FormMgr.ShowForm Me.ActiveForm, frmMultiChart, True
    'FCallStats.Show , Me
  Case "ID_EmployeeMgt"
    If Employee.InGroup(User.Name, "Development") = True Then ' Or Employee.InGroup(User.Name, "Management") = True Then
      FormMgr.ShowForm Me.ActiveForm, FEmployeeMgt, True
    Else
      MsgBox "Access denied.", vbCritical, ""
    End If
    '
  End Select
  '
  Set Employee = Nothing
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.tbMain_ToolClick.", vbCritical, "Error"
End Sub

Public Property Let ControlsEnabled(ByVal pbStatus As Boolean)
  On Error GoTo ErrCall
  '
  Dim i As Integer, j As Integer
  '
  i = tbMain.Tools.Count
  '
  tbMain.Redraw = False
  For j = 1 To i
    tbMain.Tools(j).Enabled = pbStatus
  Next j
  tbMain.Redraw = True
  '
  'sbMain.Enabled = pbStatus
  '
  Exit Property
ErrCall:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FMain.ControlsEnabled.", vbCritical, "Error"
End Property

Private Sub SetPassword()
  'On Error Resume Next
  On Error GoTo EH
  '
  Dim rsEmployee As ADODB.Recordset
  Set rsEmployee = New ADODB.Recordset
  '
  If ConnType = Access Then
    rsEmployee.Open "Select EmployeeFirst & ' ' & EmployeeLast AS Employee, Password FROM tblEmployees", cnMain, adOpenDynamic, adLockOptimistic
  Else
    rsEmployee.Open "SELECT *, EmployeeFirst + ' ' + EmployeeLast AS Employee FROM tblEmployees", cnMain, adOpenDynamic, adLockOptimistic, adCmdText
    'rsEmployee.Open "UpSelectEmployeeList", cnMain, adOpenDynamic, adLockOptimistic, adCmdStoredProc
  End If
  '
  rsEmployee.Find "Employee = '" & User.Name & "'" ', , adSearchForward
  '
  If Not rsEmployee.eof Then
    FSetPassword.Setup DecryptStr(rsEmployee!Password & ""), 0, 50, False
    FSetPassword.Show vbModal, FMain
    '
    If FSetPassword.PwdOK And Not FSetPassword.Cancelled Then
      rsEmployee!Password = EncryptStr(FSetPassword.NewPwd)
      rsEmployee.Update
      cnMain.Execute "EXEC sp_password NULL, '" & Rot39(FSetPassword.NewPwd) & "', '" & Replace(User.Name, " ", "") & "'"
    End If
    '
    Unload FSetPassword
    Set FSetPassword = Nothing
  End If
  '
  rsEmployee.Close
  Set rsEmployee = Nothing
  Exit Sub
EH:
  MsgBox Err.Description & " in Reset Password."
End Sub

Private Sub OpenDatabase()
  On Error GoTo EH
  '
  Dim frm As Form
  '
  For Each frm In Forms
    If frm.Name <> "FMain" Then
      Unload frm
      Me.BackColor = &H8000000C
    End If
  Next
  '
  Dim sPath As String, sFile As String
  '
  sPath = FileOps.IsolatePath(DBOps.DBName)
  sFile = FileOps.IsolateFile(DBOps.DBName)
  '
  If DBOps.GetPathFile(sPath, sFile, "PowerClaim Customers") Then
    If Not DBOps.OpenConnection(cnMain, sPath, sFile, "Bread 'n' Butter Data") Then
      MsgBox "Invalid database"
      End
    Else
      SaveSetting App.Title, "File", "PCCustomersName", sFile
      SaveSetting App.Title, "File", "PCCustomersPath", sPath
      '
      tbMain.ToolBars(2).Tools("ID_PathFile").Name = FileOps.IsolatePath(DBOps.DBName) & FileOps.IsolateFile(DBOps.DBName)
    End If
  End If
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in Open Database."
End Sub

Public Sub tmrMessages_Timer()
Dim rsUser As New ADODB.Recordset
Dim rsMessages As New ADODB.Recordset

  iMessageCount1 = 0
  iMessageCount2 = 0
  iMessageCount3 = 0
  iMessageCount4 = 0
  '
  'StrUser = cmbUser
  rsUser.Open "select [EmployeeFirst], [EmployeeLast], [Groups] from tblEmployees", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  With rsUser
    Do While Not .eof
      If LCase(StrUser) = LCase(!EmployeeFirst & " " & !EmployeeLast) Then
        iGroupNumber = !Groups '& vbNullString
      End If
      .MoveNext
    Loop
    .Close
  End With
  
  rsMessages.Open "SELECT [Group], [Completed] from TVMailMessages WHERE Completed = 'False'", cnMain, adOpenKeyset, adLockBatchOptimistic
  '
  
  With rsMessages
    If Not .eof Then
      .MoveFirst
      While Not .eof
        If !Group = "Authorizations" Then
            iMessageCount1 = iMessageCount1 + 1
        End If
        If !Group = "Sales" Then
            iMessageCount2 = iMessageCount2 + 1
        End If
        If !Group = "Support" Then
            iMessageCount3 = iMessageCount3 + 1
        End If
        If !Group = "Operator" Then
            iMessageCount4 = iMessageCount4 + 1
        End If
        .MoveNext
      Wend
    End If
  End With
  '
    If iFlasher > 2 Then
      iFlasher = 1
    End If
    '
    
    If DateDiff("n", (Mid(GetLastUpdate, 1, 11)), Time) > 30 Or DateDiff("d", (Mid(GetLastUpdate, 12, 11)), Date) > 0 Then
      tmrMessages.Interval = 100
      If iFlasher = 1 Then
        tbMain.ToolBars(2).Tools("ID_Messages").Name = " ***Server Is Down, Tell Supervisor***"
      Else
        tbMain.ToolBars(2).Tools("ID_Messages").Name = "       Server Is Down, Tell Supervisor"
      End If
      '
      iFlasher = iFlasher + 1
    Else
      tmrMessages.Interval = 30000
      '
      tbMain.ToolBars(2).Tools("ID_Messages").Name = GetUserGroups '"Msgs: Auth-" & iMessageCount1 & "  Sales-" & iMessageCount2 & "  Support-" & iMessageCount3 & "  Operator-" & iMessageCount4
      '
      'If iMessageCount1 > 0 Then
       ' tbMain.ToolBars(2).Tools("ID_Messages").ForeColor = vbRed
      'Else
        'tbMain.ToolBars(2).Tools("ID_Messages").ForeColor = vbBlack
      'End If
    'iFlasher = iFlasher + 1
  End If
    
    
    
    
    
    
    'tbMain.ToolBars(2).Tools("ID_Messages").Name = "New Messages: " & iMessageCount1
    '
End Sub

Private Sub tmrSecChk_Timer()
  On Error GoTo ErrorHandler
  '
  bLicTimer = True
  FMain.License.ForceStatusChanged
  bLicTimer = False
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.tmrSecChk.Timer"
End Sub

Private Sub License_Trigger(ByVal event_num As Long, ByVal event_data As Long)
  On Error GoTo ErrorHandler:
  '
  FLicense.NotifyResult event_num, event_data
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.License.Trigger"
End Sub

Private Sub License_StatusChanged(ByVal startup As Boolean)
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iSecValPair As Integer
  Dim iDlgRsp     As Integer
  Dim sDlgMsg     As String
  '
  With FMain.License
    '
    iSecValPair = Int((3 - 1 + 1) * Rnd + 1)
    '
    If .LibTest(lSecValCode(iSecValPair)) <> lSecValRslt(iSecValPair) Then
      If bSecDisp = False Then
        If bLicTimer = False Then FLicense.Show vbModal, FMain
      Else
        FLicense.NotifyStatus "Failed"
      End If
    End If
    '
    If CPCheck = (Date - dSecVar) Then
      If .IsExpired Then
        If bSecDisp = False Then
          If bLicTimer = False Then FormMgr.ShowForm Me.ActiveForm, FLicense
          'If bLicTimer = False Then FLicense.Show vbModal, FMain
        Else
          FLicense.NotifyStatus "Expired"
        End If
      Else
        If .IsClockTurnedBack Then
          If bSecDisp = False Then
            If bLicTimer = False Then FormMgr.ShowForm Me.ActiveForm, FLicense
            'If bLicTimer = False Then FLicense.Show vbModal, FMain
          Else
            FLicense.NotifyStatus "ClockTurnedBack"
          End If
        Else
          If bSecDisp = True Then
            FLicense.NotifyStatus "Licensed"
          ElseIf bLicChecked = False Then
            Select Case .DaysLeft
              Case 1
                sDlgMsg = "Your license will expire after today!"
                sDlgMsg = sDlgMsg & vbCrLf & vbCrLf
                sDlgMsg = sDlgMsg & "Contact Hawkins Research, Inc. at 1-800-736-1246 to extend your license."
                iDlgRsp = MsgBox(sDlgMsg, vbExclamation + vbOKOnly, "ATTENTION: License Expires Today")
              Case 30, 15, Is <= 10
                If GetSetting(App.Title, "Lic", "NotifyExp" & Trim(CStr(.DaysLeft)), True) Then
                  sDlgMsg = "You only have " & Trim(CStr(.DaysLeft)) & " remaining before your license expires."
                  sDlgMsg = sDlgMsg & "Your license will expire on " & Trim(CStr(.ExpireDateSoft)) & "."
                  sDlgMsg = sDlgMsg & vbCrLf & vbCrLf
                  sDlgMsg = sDlgMsg & "Contact Hawkins Research, Inc. at 1-800-736-1246 to extend your license."
                  sDlgMsg = sDlgMsg & vbCrLf
                  sDlgMsg = sDlgMsg & vbCrLf & vbCrLf & "Do you want to see this message the next time you start " & App.Title & " today?"
                  iDlgRsp = MsgBox(sDlgMsg, vbExclamation + vbYesNo, "ATTENTION: License Expires Soon")
                  SaveSetting App.Title, "Lic", "NotifyExp" & Trim(CStr(.DaysLeft)), IIf(iDlgRsp = vbYes, True, False)
                End If
            End Select
          End If
        End If
      End If
    Else
      'MsgBox .ExpireMode
      If .ExpireMode = "D" Then
        If .IsClockTurnedBack Then
          If bSecDisp = False Then
            If bLicTimer = False Then FormMgr.ShowForm Me.ActiveForm, FLicense
            'If bLicTimer = False Then FLicense.Show vbModal, FMain
          Else
            FLicense.NotifyStatus "ClockTurnedBack"
          End If
        Else
          If bSecDisp = False Then
            'If bLicTimer = False Then FLicense.Show vbModal, FMain
             If bLicTimer = False Then FormMgr.ShowForm Me.ActiveForm, FLicense
          Else
            FLicense.NotifyStatus "NeverAuthorized"
          End If
        End If
      Else
        If bSecDisp = False Then
          FormMgr.ShowForm Me.ActiveForm, FLicense
          'FLicense.Show vbModal, FMain
        Else
          FLicense.NotifyStatus "SystemFailure"
        End If
      End If
    End If
    '
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.License.StatusChanged"
End Sub

Private Sub License_Error()
  On Error GoTo ErrorHandler
  '
  If bLicError = True Then
    bLicError = False
    FLicense.NotifyStatus "Error"
    Exit Sub
  End If
  '
  bLicError = True
  '
  If bSecDisp = False Then FormMgr.ShowForm Me.ActiveForm, FLicense
 ' If bSecDisp = False Then FLicense.Show vbModal, FMain
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.License.Error"
End Sub

Private Sub tmrTray_Timer()
Dim iAddon As Integer
  '
  If GetSetting(App.Title, "Settings", "Icon", 0) = 1 Then
    iVMailTotal = 0
  End If
  '
  If iVMailTotal > 0 Then
    Select Case iVMailTotal
      Case Is < 3
        iAddon = 0
      Case 3, 4
        iAddon = 8 '4
      Case 5, 6
        iAddon = 16 '8
      Case 7, 8
        iAddon = 24 '12
      Case 9, 10
        iAddon = 32 '16
      Case Else
        iAddon = 40 '20
    End Select
      
    iFlasher = iFlasher + 1
    '
    nid.hIcon = TrayIconList2.ListImages(iFlasher + iAddon).Picture
    Shell_NotifyIcon NIM_MODIFY, nid
    If iFlasher = 8 Then iFlasher = 0
    '
  Else
    nid.hIcon = Me.Icon
    Shell_NotifyIcon NIM_MODIFY, nid
  End If
End Sub

Public Function GetUserGroups() As String
  Dim iTemp As Integer
  Dim sTempMsg As String
  '
  iVMailTotal = 0
  '
  iTemp = iGroupNumber
  If iTemp >= 8 Then
    iVMailTotal = iMessageCount1
    iTemp = iTemp - 8
  End If
  '
  If iTemp >= 4 Then
    iVMailTotal = iVMailTotal + iMessageCount2
    iTemp = iTemp - 4
  End If
  '
  If iTemp >= 2 Then
    iVMailTotal = iVMailTotal + iMessageCount3
    iTemp = iTemp - 2
  End If
  '
  If iTemp >= 1 Then
    iVMailTotal = iVMailTotal + iMessageCount4
  End If
  '
  sTempMsg = "Msgs: Auth-" & iMessageCount1 & "  Sales-" & iMessageCount2 & "  Support-" & iMessageCount3 & "  Operator-" & iMessageCount4
    
  GetUserGroups = sTempMsg
  '
End Function

