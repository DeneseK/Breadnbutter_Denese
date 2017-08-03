VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FDetails 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pac Man"
   ClientHeight    =   2190
   ClientLeft      =   3930
   ClientTop       =   3405
   ClientWidth     =   6075
   ControlBox      =   0   'False
   Icon            =   "FDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6075
   Begin VB.PictureBox fmePVAuthStatus 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   5685
      TabIndex        =   2
      Top             =   0
      Width           =   5685
      Begin VB.TextBox txtPVAuthDays 
         DataField       =   "AuthDays"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   8
         Tag             =   "1"
         Top             =   1560
         Width           =   675
      End
      Begin VB.TextBox txtPVVersionShipped 
         DataField       =   "VersionShipped"
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Tag             =   "1"
         Top             =   1245
         Width           =   675
      End
      Begin VB.TextBox txtPVAuths 
         Height          =   315
         Left            =   3360
         MaxLength       =   1
         TabIndex        =   6
         Top             =   75
         Width           =   375
      End
      Begin VB.TextBox txtPVPendingDays 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5070
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   5
         Top             =   510
         Width           =   555
      End
      Begin VB.TextBox txtPVGraceDays 
         DataField       =   "VersionShipped"
         Height          =   315
         Left            =   5085
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "1"
         Top             =   90
         Width           =   555
      End
      Begin VB.TextBox txtPVSaleDays 
         Height          =   315
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin SSDataWidgets_B.SSDBCombo cboPVAuthStatus 
         DataField       =   "AuthStatus"
         Height          =   315
         Left            =   810
         TabIndex        =   9
         Tag             =   "1"
         Top             =   1575
         Width           =   1860
         DataFieldList   =   "Column 0"
         ListAutoValidate=   0   'False
         _Version        =   196617
         DataMode        =   2
         Cols            =   1
         ColumnHeaders   =   0   'False
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   3281
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboPVShipStatus 
         DataField       =   "ShipStatus"
         Height          =   315
         Left            =   810
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1245
         Width           =   1860
         DataFieldList   =   "Column 0"
         ListAutoValidate=   0   'False
         _Version        =   196617
         DataMode        =   2
         Cols            =   1
         ColumnHeaders   =   0   'False
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   3281
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin TDBDate6Ctl.TDBDate mskPVShipDate 
         Height          =   315
         Left            =   2700
         TabIndex        =   11
         Top             =   1245
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FDetails.frx":014A
         Caption         =   "FDetails.frx":0262
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FDetails.frx":02CE
         Keys            =   "FDetails.frx":02EC
         Spin            =   "FDetails.frx":034A
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
      Begin TDBDate6Ctl.TDBDate mskPVAuthDate 
         Height          =   315
         Left            =   2700
         TabIndex        =   12
         Top             =   1560
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FDetails.frx":0372
         Caption         =   "FDetails.frx":048A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FDetails.frx":04F6
         Keys            =   "FDetails.frx":0514
         Spin            =   "FDetails.frx":0572
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
         Enabled         =   0
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
      Begin Threed.SSCommand cmdPVShipDate 
         Height          =   315
         Left            =   3780
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1245
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FDetails.frx":059A
      End
      Begin Threed.SSCommand cmdPVAuthDate 
         Height          =   315
         Left            =   3780
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1575
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FDetails.frx":0B34
      End
      Begin SSDataWidgets_B.SSDBCombo cboPVDownloadStatus 
         DataField       =   "ShipStatus"
         Height          =   315
         Left            =   810
         TabIndex        =   15
         Tag             =   "1"
         Top             =   900
         Width           =   1860
         DataFieldList   =   "Column 0"
         ListAutoValidate=   0   'False
         _Version        =   196617
         DataMode        =   2
         Cols            =   1
         ColumnHeaders   =   0   'False
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   3281
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin TDBDate6Ctl.TDBDate mskPVDownloadDate 
         Height          =   315
         Left            =   2700
         TabIndex        =   16
         Top             =   900
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FDetails.frx":10CE
         Caption         =   "FDetails.frx":11E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FDetails.frx":1252
         Keys            =   "FDetails.frx":1270
         Spin            =   "FDetails.frx":12CE
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
      Begin Threed.SSCommand cmdPVDownloadDate 
         Height          =   315
         Left            =   3780
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   900
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FDetails.frx":12F6
      End
      Begin TDBDate6Ctl.TDBDate mskPVSaleDate 
         Height          =   315
         Left            =   885
         TabIndex        =   18
         Top             =   465
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FDetails.frx":1890
         Caption         =   "FDetails.frx":19A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FDetails.frx":1A14
         Keys            =   "FDetails.frx":1A32
         Spin            =   "FDetails.frx":1A90
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
      Begin Threed.SSCommand cmdPVSalesDate 
         Height          =   315
         Left            =   1965
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   465
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FDetails.frx":1AB8
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Download:"
         Height          =   285
         Index           =   5
         Left            =   45
         TabIndex        =   31
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping:"
         Height          =   285
         Index           =   1
         Left            =   45
         TabIndex        =   30
         Top             =   1275
         Width           =   765
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ver.:"
         Height          =   285
         Index           =   6
         Left            =   4230
         TabIndex        =   29
         Top             =   1275
         Width           =   645
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Auth:"
         Height          =   285
         Index           =   2
         Left            =   45
         TabIndex        =   28
         Top             =   1605
         Width           =   1005
      End
      Begin VB.Label lblPVAuthRemaining 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "AuthRemaining"
         Height          =   315
         Left            =   4920
         TabIndex        =   27
         Tag             =   "1"
         Top             =   1575
         Width           =   675
      End
      Begin VB.Label lblPVExpires 
         BackStyle       =   0  'Transparent
         Caption         =   "Expires:"
         Height          =   285
         Left            =   4215
         TabIndex        =   26
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "PowerClaim XML"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   25
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Online Auths:"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Date:"
         Height          =   285
         Left            =   45
         TabIndex        =   23
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Pending Days:"
         Height          =   285
         Index           =   8
         Left            =   3960
         TabIndex        =   22
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Grace Period:"
         Height          =   285
         Index           =   9
         Left            =   4080
         TabIndex        =   21
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Days:"
         Height          =   285
         Index           =   12
         Left            =   2400
         TabIndex        =   20
         Top             =   525
         Width           =   1035
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   650
      Left            =   1080
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6360
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   " Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      Caption         =   "Game "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   1200
      Picture         =   "FDetails.frx":2052
      Top             =   1560
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   2040
      Picture         =   "FDetails.frx":2B68
      Top             =   1920
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   120
      Picture         =   "FDetails.frx":367E
      Top             =   1200
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   720
      Picture         =   "FDetails.frx":4194
      Top             =   960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   120
      Picture         =   "FDetails.frx":4CAA
      Top             =   1920
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   5640
      Picture         =   "FDetails.frx":57C0
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   6360
      Picture         =   "FDetails.frx":6672
      Top             =   720
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1920
      Picture         =   "FDetails.frx":7524
      Top             =   3000
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Picture         =   "FDetails.frx":83D6
      Top             =   3000
      Width           =   420
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Shape Shape6 
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   5760
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "FDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim x As Integer
Dim n As Integer
Dim y As Integer
Dim l As Integer
Dim g As Integer
Dim q As Integer


Private Sub Timer1_Timer()
   
  If Not i > (Me.Width - 500) Then
    Select Case x
      Case 1
        Image1.Move i, 3000
        Image2.Visible = False
        Image1.Visible = True
        x = 2
        i = i + 200
      Case 2
        Image1.Move i, 3000
        Image2.Visible = False
        Image1.Visible = True
        x = 3
      Case 3
        Image2.Move i, 3000
        Image2.Visible = True
        Image1.Visible = False
        x = 1
      Case Else
        x = 1
    End Select
  End If
  '
  If Not n > (Me.Height - 1000) Then
    Image3.Move l, n
    n = n + 100
  Else
    Image3.Visible = False
    Image4.Move l, n
    Image4.Visible = True
    l = l - 100
  End If
 
 If l = 3420 Then
  l = l - 20
  Image4.Move l, n
  Timer1.Enabled = False
  Timer2.Enabled = True
  Exit Sub
 End If
 
    
End Sub

Private Sub form_Load()
  x = 1
  y = 1
  n = 720
  l = 6320
  g = 1
  i = 0
  Timer1_Timer
  GetWavFiles
  Timer4_Timer
End Sub

Private Sub GameOver()
  Image4.Visible = False
End Sub

Private Sub Timer2_Timer()
  Image4.Visible = False
  Timer2.Enabled = False
  Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
  Timer4.Enabled = False
  Select Case g
    Case 1
      Call sndPlaySound(App.Path & "\killed.wav", &H1)
      Image1.Move i, 3000
      g = 2
    Case 2
      Image5.Move i, 3000
      Image1.Visible = False
      Image5.Visible = True
      g = 3
    Case 3
      Image6.Move i, 3000
      Image5.Visible = False
      Image6.Visible = True
      g = 4
    Case 4
      Image7.Move i, 3000
      Image6.Visible = False
      Image7.Visible = True
      g = 5
    Case 5
      Image8.Move i, 3000
      Image7.Visible = False
      Image8.Visible = True
      g = 6
    Case 6
      Image9.Move i, 3000
      Image8.Visible = False
      Image9.Visible = True
      g = 7
      Timer3.Interval = 500
    Case 7
      Image9.Visible = False
      g = 8
    Case 8
      Label1.Visible = True
      g = 9
    Case 9
      Label2.Visible = True
      g = 10
      Timer3.Interval = 1000
    Case 10
      Unload Me
  End Select
End Sub

Private Sub Timer4_Timer()
  Call sndPlaySound(App.Path & "\pacchomp.wav", &H1)
End Sub

Private Sub GetWavFiles()
  Dim rsWav As New ADODB.Recordset
  Dim strStream As ADODB.Stream
  rsWav.Open "Select * from TVMailMessages where [MessageName] = 'pacchomp.wav'", cnMain, adOpenDynamic, adLockBatchOptimistic
              If Not rsWav.eof Then
                If Not rsWav.BOF Then
              Set strStream = New ADODB.Stream
             strStream.Type = adTypeBinary
             strStream.Open
             strStream.Write rsWav!Attachment
             strStream.SaveToFile App.Path & "\" & rsWav!MessageName, adSaveCreateOverWrite
             strStream.Close
             Set strStream = Nothing
             End If
            End If
  rsWav.Close
  rsWav.Open "Select * from TVMailMessages where [MessageName] = 'killed.wav'", cnMain, adOpenDynamic, adLockBatchOptimistic
              If Not rsWav.eof Then
                If Not rsWav.BOF Then
              Set strStream = New ADODB.Stream
             strStream.Type = adTypeBinary
             strStream.Open
             strStream.Write rsWav!Attachment
             strStream.SaveToFile App.Path & "\" & rsWav!MessageName, adSaveCreateOverWrite
             strStream.Close
             Set strStream = Nothing
             End If
            End If
  rsWav.Close
  
  
End Sub
