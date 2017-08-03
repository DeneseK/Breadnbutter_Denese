VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{3D6D5D32-B9F2-101C-AED5-00608CF525A5}#1.4#0"; "Tx4ole.ocx"
Begin VB.Form FLicense 
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   ControlBox      =   0   'False
   Icon            =   "FLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   2265
      Left            =   150
      ScaleHeight     =   2205
      ScaleWidth      =   7695
      TabIndex        =   16
      Top             =   150
      Width           =   7755
      Begin Tx4oleLib.TXTextControl txtDescript 
         Height          =   2205
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   7695
         _Version        =   65540
         _ExtentX        =   13573
         _ExtentY        =   3889
         _StockProps     =   73
         Language        =   1
         BorderStyle     =   0
         BackStyle       =   1
         ControlChars    =   0   'False
         EditMode        =   2
         HideSelection   =   -1  'True
         InsertionMode   =   -1  'True
         MousePointer    =   0
         ZoomFactor      =   100
         ViewMode        =   0
         ClipChildren    =   0   'False
         ClipSiblings    =   -1  'True
         SizeMode        =   0
         TabKey          =   -1  'True
         FormatSelection =   0   'False
         VTSpellDictionary=   ""
         ScrollBars      =   2
         PageWidth       =   0
         PageHeight      =   6000
         PageMarginL     =   1440
         PageMarginT     =   1440
         PageMarginR     =   1440
         PageMarginB     =   1440
         PrintZoom       =   100
         PrintOffset     =   0   'False
         PrintColors     =   -1  'True
         FontName        =   "Arial"
         FontSize        =   12
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Baseline        =   0
         TextBkColor     =   -2147483633
         Alignment       =   0
         LineSpacing     =   100
         LineSpacingT    =   0
         FrameStyle      =   32
         FrameDistance   =   0
         FrameLineWidth  =   20
         IndentL         =   120
         IndentR         =   0
         IndentFL        =   0
         IndentT         =   0
         IndentB         =   0
         Text            =   ""
         WordWrapMode    =   1
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   555
      Begin ActiveToolBars.SSActiveToolBars tbLic 
         Left            =   30
         Top             =   90
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   131083
         MenuAnimations  =   3
         ToolBarsCount   =   12
         ToolsCount      =   4
         ShowShortcutsInToolTips=   -1  'True
         Visible         =   0   'False
         Tools           =   "FLicense.frx":000C
         ToolBars        =   "FLicense.frx":01DF
      End
   End
   Begin VB.Frame fmeAdv 
      Caption         =   "Advanced Operations"
      Height          =   1485
      Left            =   180
      TabIndex        =   3
      Top             =   4200
      Width           =   7695
      Begin Threed.SSCommand cmdDeauthorize 
         Height          =   285
         Left            =   270
         TabIndex        =   10
         Top             =   1020
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FLicense.frx":04E3
         Caption         =   "&Deauthorize"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdImprint 
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   1740
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FLicense.frx":063D
         Caption         =   "Im&print"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdExport 
         Height          =   315
         Left            =   3000
         TabIndex        =   12
         Top             =   1740
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FLicense.frx":0797
         Caption         =   "&Export"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdImport 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   1740
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FLicense.frx":08F1
         Caption         =   "&Import"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin TDBMask6Ctl.TDBMask mskConfirm 
         Height          =   285
         Left            =   3450
         TabIndex        =   15
         Top             =   1020
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   503
         Caption         =   "FLicense.frx":0A4B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "FLicense.frx":0AB7
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         AllowSpace      =   -1
         AutoConvert     =   1
         BackColor       =   255
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "?????????????????????????"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   1
         LookupTable     =   "0-9"
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                         "
         Value           =   ""
      End
      Begin Threed.SSCommand cmdRefresh 
         Height          =   285
         Left            =   270
         TabIndex        =   18
         Top             =   390
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FLicense.frx":0AF9
         Caption         =   "&Refresh Site Code"
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin VB.Label Label2 
         Caption         =   "Transfer Functions:"
         Height          =   225
         Left            =   510
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmation Code:"
         Height          =   195
         Index           =   88
         Left            =   2010
         TabIndex        =   4
         Top             =   1050
         WhatsThisHelpID =   10313
         Width           =   1365
      End
   End
   Begin VB.TextBox txtSiteCode 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2730
      Width           =   4725
   End
   Begin Threed.SSCommand cmdAuthorize 
      Default         =   -1  'True
      Height          =   375
      Left            =   6540
      TabIndex        =   1
      Top             =   2730
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   196610
      PictureFrames   =   1
      Picture         =   "FLicense.frx":0C53
      Caption         =   "  &Authorize"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin TDBMask6Ctl.TDBMask mskSiteKey 
      Height          =   345
      Left            =   1680
      TabIndex        =   2
      Top             =   3150
      Width           =   4725
      _Version        =   65536
      _ExtentX        =   8334
      _ExtentY        =   609
      Caption         =   "FLicense.frx":0E3D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FLicense.frx":0EA9
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      AllowSpace      =   -1
      AutoConvert     =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "?????????????????????????"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   1
      LookupTable     =   " ,0-9"
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "                         "
      Value           =   ""
   End
   Begin Threed.SSCommand cmdDone 
      Height          =   375
      Left            =   6540
      TabIndex        =   5
      Top             =   3150
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   196610
      PictureFrames   =   1
      Picture         =   "FLicense.frx":0EEB
      Caption         =   "   &Close"
      Alignment       =   1
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin Threed.SSCommand cmdAdv 
      Height          =   375
      Left            =   6540
      TabIndex        =   6
      Top             =   3690
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   196610
      PictureFrames   =   1
      Picture         =   "FLicense.frx":1045
      Caption         =   "Advanced >>"
      ButtonStyle     =   3
      PictureAlignment=   9
   End
   Begin VB.Image imgSec 
      Height          =   480
      Left            =   240
      Picture         =   "FLicense.frx":119F
      Top             =   2850
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Site Key:"
      Height          =   195
      Index           =   6
      Left            =   840
      TabIndex        =   8
      Top             =   3210
      WhatsThisHelpID =   10304
      Width           =   675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code:"
      Height          =   195
      Index           =   46
      Left            =   840
      TabIndex        =   7
      Top             =   2790
      WhatsThisHelpID =   10303
      Width           =   735
   End
End
Attribute VB_Name = "FLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1
'
'\\ General
Private lSessionID            As Long
Private sngFrmHtStd           As Single
Private sngFrmHtAdv           As Single
Private sLicExpireDays        As String
Private sLicExpireDate        As String
Private ctrlAct               As Control
'
'\\ License Messages
Private Const csLicMsgHeader                  As String = _
  "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman Times New Roman;}{\f3\froman Times New Roman;}}{\colortbl\red0\green0\blue0;}\deflang1033\pard\plain\f2\fs20"
'
Private Const csLicMsgApp                     As String = _
  "\plain\f2\fs20\b\i PowerKey\plain\f2\fs20"
'
Private Const csLicMsgContact                 As String = _
  "ontact the IT Manager"
'
Private Const csLicMsgError                   As String = _
  csLicMsgContact & " to report this error."
'
Private Const csLicMsgLimited                 As String = _
  "You cannot authorize or extend licenses."
'
Private Const csLicMsgUnauthorized            As String = csLicMsgHeader & _
  "\b Your license is not authorized.\par\par\plain\f2\fs20 " & _
  "You cannot authorize or extend licenses.\par\par\plain\f2\fs20\b " & _
  "C" & csLicMsgContact & " to authorize your license and unlock all the power of " & csLicMsgApp & ".\par }"
'
Private Const csLicMsgNeverAuthorized         As String = csLicMsgHeader & _
  "\b Welcome to " & csLicMsgApp & "!\par\par\plain\f2\fs20 " & _
  "To authorize " & csLicMsgApp & " , c" & csLicMsgContact & "\plain\f2\fs20.\par\par " & _
  "Until then, you will not be able to authorize licenses.\par }"
'
Private Const csLicMsg30DayEval               As String = csLicMsgHeader & _
  "\b Your license could not be authorized because it has previously been authorized for a 30-day trial.\par\par\par }"
'
Private Const csLicMsg15DayEval               As String = csLicMsgHeader & _
  "\b Your license could not be authorized because it has previously been authorized for 15-day trial extension.\par\par\par }"
'
Private Const csLicMsgValidDeauthorization    As String = csLicMsgHeader & _
  " Your license has been deauthorized.\par\par\par }"
'
Private Const csLicMsgInvalidDeauthorization        As String = csLicMsgHeader & _
  "\b Your license cannot be deauthorized because it is either currently unauthorized or in trial mode.\par\par\par }"
'
Private Const csLicMsgKeyValidAuthorization   As String = csLicMsgHeader & _
  "\b Your license has been authorized.\par\par\par }"
'
Private Const csLicMsgKeyValidExtension       As String = csLicMsgHeader & _
  "\b Your license has been extended.\par\par\par }"
'
Private Const csLicMsgInvalidExtension        As String = csLicMsgHeader & _
  "\b Your license cannot be extended because it is not authorized.\par\par\par }"
'
Private Const csLicMsgSiteKeyNotSpecified     As String = csLicMsgHeader & _
  "\b You must enter a site key obtained from Jason or Eric in order to authorize a license.\par\par\par }"
'
Private Const csLicMsgSiteCodeCompacted       As String = csLicMsgHeader & _
  "\b Your license could not be authorized because the site code you entered does not contain a space. Please check you code and attempt to authorize your license again.\par\par\par }"
'
Private Const csLicMsgKeyInvalid             As String = csLicMsgHeader & _
  "\b Your license could not be authorized because the site key you entered is invalid. Please verify your site key and attempt to authorize your license again.\par\par\par }"
'
Private Const csLicMsgKey1Invalid             As String = csLicMsgHeader & _
  "\b Your license could not be authorized because part one of the site key you entered is invalid. Please verify your site key and attempt to authorize your license again.\par\par\par }"
'
Private Const csLicMsgKey2Invalid             As String = csLicMsgHeader & _
  "\b Your license could not be authorized because part two of the site key you entered is invalid. Please verify your site key and attempt to authorize your license again.\par\par\par }"
'
Private Const csLicMsgSystemFailure           As String = csLicMsgHeader & _
  "\b Your license is damaged or security has been compromised. " & _
  "C" & csLicMsgError & "\par\par " & _
  csLicMsgUnauthorized & "\par }"
'
Private Const csLicMsgClockTurnedBack         As String = csLicMsgHeader & _
  "\b Your system calendar and/or clock has been turned back. " & _
  "Please correct your system's date and time and restart " & csLicMsgApp & ".\par\par " & _
  "If this error persists, c" & csLicMsgContact & "\par\par " & _
  csLicMsgUnauthorized & "\par }"
'
Private Const csLicMsgSiteCodeReset           As String = csLicMsgHeader & _
  "\b Your site code has been refreshed.\par\par\par }"
'
Private Const csLicMsgClockReset              As String = csLicMsgHeader & _
  "\b Your license has been synchronized with your system's date and time.\par\par\par }"
'
Private Const csLicMsgAuthorized              As String = csLicMsgHeader & _
  " To extend your license, c" & csLicMsgContact & ".\par }"
'
Private Const csLicMsgExpired                 As String = csLicMsgHeader & _
  "\plain\f2\fs20\  " & csLicMsgLimited & "\par\par " & _
  "To renew your license, c" & csLicMsgContact & "\par }"
Public Sub NotifyStatus(psLicSts As String)
  On Error GoTo ErrorHandler
  '
  With FMain.License
    '
    If psLicSts = "Error" Then
      DisplayStatus csLicMsgHeader & "\b The following error has occurred: Error Number " & CStr(.LastErrorNumber) & " -- " & .LastErrorString & ".\par\par\plain\f2\fs20\  " & _
        "C" & csLicMsgError & "\par\par " & _
        "Your license may not be authorized. \par }"
      Exit Sub
    End If
    '
    '\\ Calculate and Display Site Code
    lSessionID = .UserNumber(5)
    If lSessionID = 0 Then lSessionID = .TCSessionCode
    'txtSiteCode.Text = Trim(CStr(FMain.lSecCompID)) & " " & Trim(CStr(lSessionID))
    txtSiteCode.Text = Trim(CStr(lSecCompID)) & " " & Trim(CStr(lSessionID))
    '
    sLicExpireDays = .DaysLeft
    sLicExpireDate = .ExpireDateSoft
    '
    Select Case psLicSts
      Case "Licensed"
        Select Case .DaysLeft
          Case Is < 0
            DisplayStatus csLicMsgExpired
          Case Is > 1
            DisplayStatus csLicMsgHeader & "\b You have " & .DaysLeft & " days remaining before your license expires. The expiration date is " & sLicExpireDate & ".\par\par" & csLicMsgAuthorized
          Case 1
            DisplayStatus csLicMsgHeader & "\b Your license will expire today. The expiration date is " & sLicExpireDate & ".\par\par" & csLicMsgAuthorized
        End Select
      Case "Expired"
        DisplayStatus csLicMsgHeader & "\b Your license expired on " & FMain.License.ExpireDateSoft & ".\par\par\ }" & csLicMsgExpired
      Case "30DayEval"
        DisplayStatus csLicMsg30DayEval
      Case "15DayEval"
        DisplayStatus csLicMsg15DayEval
      Case "NeverAuthorized"
        DisplayStatus csLicMsgNeverAuthorized
      Case "Unauthorized"
        DisplayStatus csLicMsgUnauthorized
      Case "ClockTurnedBack"
        DisplayStatus csLicMsgClockTurnedBack
      Case "SystemFailure"
        DisplayStatus csLicMsgSystemFailure
      Case "Error"
        DisplayStatus csLicMsgHeader & "\b The following error has occurred: Error Number " & CStr(.LastErrorNumber) & " -- " & .LastErrorString & ".\par\par\plain\f2\fs20\  " & _
          "C" & csLicMsgError & "\par\par " & _
          "Your license may not be authorized. \par }"
    End Select
    '
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.General.NotifyStatus"
End Sub
Public Sub DisplayStatus(psStatus As String)
  On Error GoTo ErrorHandler:
  '
  txtDescript.SelStart = Len(txtDescript.Text)
  txtDescript.SelLength = Len(txtDescript.Text)
  txtDescript.RTFSelText = psStatus
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.General.DisplayStatus"
End Sub
Private Sub cmdAdv_Click()
  On Error GoTo ErrorHandler
  '
  Select Case Trim(cmdAdv.Caption)
    Case "Advanced >>"
      cmdAdv.Caption = "<< Standard"
      FLicense.Height = sngFrmHtAdv
      CenterForm FLicense
    Case "<< Standard"
      cmdAdv.Caption = "Advanced >>"
      FLicense.Height = sngFrmHtStd
      CenterForm FLicense
  End Select
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.cmdAdv.Click"
End Sub
Private Sub cmdAuthorize_Click()
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iDelPos   As Integer
  Dim sSiteCode As String
  Dim sRegKey1  As String
  Dim sRegKey2  As String
  '
  sSiteCode = Trim(mskSiteKey.Text)
  '
  If sSiteCode = vbNullString Then
    DisplayResult csLicMsgSiteKeyNotSpecified
    FMain.License.ForceStatusChanged
    Exit Sub
  End If
  '
  With FMain.License
    iDelPos = InStr(1, sSiteCode, " ", vbBinaryCompare)
    If iDelPos > 0 Then
      sRegKey1 = Left$(sSiteCode, iDelPos - 1)
      sRegKey2 = Trim(Right$(sSiteCode, Len(sSiteCode) - iDelPos))
    Else
      sRegKey1 = sSiteCode
    End If
    .TCode Val(sRegKey1), Val(sRegKey2), lSessionID, 0, 0
  End With
  '
  Exit Sub
  '
ErrorHandler:
  If Err.number = 6 Then '\\ No Space in Site Code
    If InStr(1, sSiteCode, " ", vbBinaryCompare) > 0 Then
      DisplayResult csLicMsgKeyInvalid
    Else
      DisplayResult csLicMsgSiteCodeCompacted
    End If
    FMain.License.ForceStatusChanged
    Exit Sub
  End If
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.cmdAuthorize.Click"
End Sub
Private Sub cmdDeauthorize_Click()
  On Local Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim lRes      As Long
  Dim sDaysLeft As String
  '
  If MsgBox("Are you certain you want to deauthorize the license? If you do, you will have to contact Hawkins Research before you will be able to create new claims using PowerClaim.", vbYesNo, "CONFIRM: Deauthorize License") = vbNo Then Exit Sub
  '
  With FMain.License
    '
    lRes = FMain.License.CPDelete(-1)
    '
    Select Case lRes
      Case 1
        If .ExpireMode = "P" Then
          If .ExpireDateSoft <> "00/00/00" Then
            If .IsExpired = False Then
              DisplayResult csLicMsgValidDeauthorization
              sDaysLeft = CStr(.DaysLeft / 1.27)
              If InStr(1, sDaysLeft, ".", vbBinaryCompare) = 0 Then sDaysLeft = sDaysLeft & ".0"
              mskConfirm.Text = Replace(CStr(CDbl(Date + Time)), ".", " ") & " " & Replace(sDaysLeft, ".", " ")
              .ExpireMode = "D"
              .ExpireDateSoft = "00/00/00"
              ResetSiteCode
              .ForceStatusChanged
            Else
              mskConfirm.Text = vbNullString
              DisplayResult csLicMsgInvalidDeauthorization
              .ForceStatusChanged
            End If
          Else
            mskConfirm.Text = vbNullString
            DisplayResult csLicMsgInvalidDeauthorization
            .ForceStatusChanged
          End If
        Else
          mskConfirm.Text = vbNullString
          DisplayResult csLicMsgInvalidDeauthorization
          .ForceStatusChanged
        End If
      Case Else
        mskConfirm.Text = vbNullString
        DisplayResult csLicMsgError
    End Select
    '
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.cmdDeauthorize.Click"
End Sub
Private Sub cmdDone_Click()
  On Error GoTo ErrorHandler
  '
  FMain.License.UserNumber(5) = lSessionID
  FMain.tmrSecChk.Enabled = True
  'FPrimary.bSecDisp = False
  bSecDisp = False
  Unload FLicense
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.cmdDone.Click"
End Sub
Private Sub cmdRefresh_Click()
  On Error GoTo ErrorHandler
  '
  ResetSiteCode
  DisplayResult csLicMsgSiteCodeReset
  FMain.License.ForceStatusChanged
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.cmdRefresh.Click"
End Sub
Private Sub Form_Activate()
  On Error GoTo ErrorHandler
  '
  FMain.License.ForceStatusChanged
  'If FPrimary.bLicError = True Then NotifyStatus "Error"
  If bLicError = True Then NotifyStatus "Error"
  '
  'mskSiteKey.SetFocus
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.Form.Activate"
End Sub

Private Sub Form_Initialize()
  On Error GoTo ErrorHandler
  Set FormData = New CFormData
  Set FormControl = New CFormControl
  '
  FormControl.MinHeight = 1965
  FormControl.MinWidth = Me.Width
  FormControl.DataForm = True
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.Form_Initialize"

End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
  '
  bSecDisp = True
  'FPrimary.bSecDisp = True
  FMain.tmrSecChk.Enabled = False
  '
  sngFrmHtStd = cmdAdv.Top + cmdAdv.Height + 255
  sngFrmHtAdv = fmeAdv.Top + fmeAdv.Height + 255
  '
  FLicense.Height = sngFrmHtStd
  CenterForm FLicense
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.Form.Load"
End Sub
Public Sub CenterForm(pFCur As Form, Optional pFRef As Form, Optional pbCmnDlg As Boolean)
  On Error GoTo ErrorHandler
  '
  With pFCur
    If Not pbCmnDlg Then
      If pFRef Is Nothing Then
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
      Else
        .Move pFRef.Left + (pFRef.Width - .Width) / 2, pFRef.Top + (pFRef.Height - .Height) / 2
      End If
    Else
      .Move (Screen.Width - .Width - 1000) / 2, (Screen.Height - .Height - 175) / 2
    End If
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.General.CenterForm"
End Sub
Private Sub mskConfirm_GotFocus()
  On Error GoTo ErrorHandler
  '
  Set ctrlAct = mskConfirm
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.mskConfirm.GotFocus"
End Sub
Private Sub mskSiteKey_GotFocus()
  On Error GoTo ErrorHandler
  '
  Set ctrlAct = mskSiteKey
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.mskSiteKey.GotFocus"
End Sub
Private Sub mskSiteKey_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrorHandler
  '
  If KeyAscii = 22 Then
    Clipboard.GetText
    KeyAscii = 0
  End If
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.mskSiteKey.KeyPress"
End Sub
Private Sub tbLic_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
  On Error GoTo ErrorHandler
  '
  Select Case Tool.ID
    Case "MnuCtxTxtCut"
      Clipboard.SetText ctrlAct.SelText
      ctrlAct.SelText = vbNullString
    Case "MnuCtxTxtCopy"
      Clipboard.SetText ctrlAct.SelText
    Case "MnuCtxTxtPaste"
      If ctrlAct.Name = "mskSiteCode" Or ctrlAct.Name = "txtConfirmCode" Then Exit Sub
      ctrlAct.SelText = Clipboard.GetText
  End Select
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.tbLic.ToolClick"
End Sub
Private Sub txtSiteCode_GotFocus()
  On Error GoTo ErrorHandler
  '
  Set ctrlAct = txtSiteCode
  SelectText txtSiteCode
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.txtSiteCode.GotFocus"
End Sub
Public Sub TransferLicense(sTransferMode As String)
  'On Error GoTo ErrorHandler
  ''
  ''\\ Local Declarations
  'Dim lRes      As String
  'Dim sFolder   As String
  'Dim sTxfrMode As String
  'Dim cdLic     As CCommonDialog
  ''
  ''\\ Obtain Directory From User
  'Select Case sTransferMode
  '  Case "Imprint"
  '    sTxfrMode = "Imprint License"
  '  Case "Export"
  '    sTxfrMode = "Export License"
  '  Case "Import"
  '    sTxfrMode = "Import License"
  'End Select
  ''
  'Set cdLic = New CCommonDialog
  'cdLic.hWnd = FLicense.hWnd
  'cdLic.DialogTitle = sTxfrMode & ": Select Folder"
  'cdLic.ShowOpen
  'sFolder = cdLic.FileName
  'If sFolder = vbNullString Then Exit Sub
  ''
  'sFolder = sFolder & "\sample.ini"
  'MsgBox sFolder
  ''
  ''\\ Perform Operation
  ''lstOps.AddItem "Authorize License" & vbTab & "Failed" & vbTab & License.ErrorMessage
  'Select Case sTransferMode
  '  Case "Imprint"
  '    lRes = FPrimary.License.Transfer(1, sFolder)
  '    'If lRes = 1 Then MsgBox("Proceed to Step 2, "safsdf")
  '  Case "Export"
  '    lRes = FPrimary.License.Transfer(2, sFolder)
  '  Case "Import"
  '    lRes = 1
  '    FileCopy sFolder, App.Path & "\sample.ini"
  'End Select
  ''
  'Select Case lRes
  '  Case 1
  '
  '  Case Else
  '
  'End Select
  '
  'Exit Sub
  ''
'ErrorHandler:
'  ErrorMgr.Raise "FLicense", "cmdImprint.Click", Err.Number, Err.Description
End Sub
Public Sub DisplayResult(psResult As String)
  On Error GoTo ErrorHandler
  '
  txtDescript.SelStart = 0
  txtDescript.SelLength = Len(txtDescript.Text)
  txtDescript.RTFSelText = psResult
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.General.DisplayResult"
End Sub
Public Sub NotifyResult(plAction As Long, plData As Long)
  On Error GoTo ErrorHandler
  '
  With FMain.License
    '
    .Enabled = False
    '
    If plAction <> 7 Then
      If plData = -1 Then
        DisplayResult csLicMsgKey2Invalid
        .Enabled = True
        .ForceStatusChanged
        Exit Sub
      End If
    End If
    '
    Select Case plAction
      Case 0
        DisplayResult csLicMsgKey1Invalid
      Case 1 '\\ Authorize License
        If plData = 30 Then
          If CBool(GetSetting(App.Title, "License", "30DayEval", False)) = True Then
            mskSiteKey.Text = vbNullString
            DisplayResult csLicMsg30DayEval
            .Enabled = True
            Exit Sub
          End If
        ElseIf plData = 15 Then
          If CBool(GetSetting(App.Title, "License", "15DayEval", False)) = True Then
            mskSiteKey.Text = vbNullString
            DisplayResult csLicMsg15DayEval
            .Enabled = True
            Exit Sub
          End If
        End If
        .CPAdd 0, 0
        .ExpireMode = "P"
        .ExpireDateSoft = Date + plData
        If plData = 30 Then
          SaveSetting App.Title, "License", "30DayEval", True
        ElseIf plData = 15 Then
          SaveSetting App.Title, "License", "15DayEval", True
        End If
        ResetExpirationNotifications
        ResetSiteCode
        DisplayResult csLicMsgKeyValidAuthorization
      Case 2 '\\ Extend License
        If .ExpireDateSoft <> "0/0/0" Then
          If .IsExpired = False Then
            .CPAdd 0, 0
            .ExpireMode = "P"
            .ExpireDateSoft = CDate(.ExpireDateSoft) + plData
            ResetExpirationNotifications
            ResetSiteCode
            DisplayResult csLicMsgKeyValidExtension
          Else
            mskSiteKey.Text = vbNullString
            DisplayResult csLicMsgInvalidExtension
          End If
        Else
          mskSiteKey.Text = vbNullString
          DisplayResult csLicMsgInvalidExtension
        End If
      Case 7
        .ResetLastUsedInfo
        DisplayResult csLicMsgClockReset
      Case Else
        DisplayResult csLicMsgError
    End Select
    '
    .Enabled = True
    '
  End With
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.General.NotifyResult"
End Sub
Public Sub ResetExpirationNotifications()
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iCt As Integer
  '
  SaveSetting App.Title, "License", "NotifyExp30", True
  SaveSetting App.Title, "License", "NotifyExp15", True
  '
  For iCt = 2 To 9
    SaveSetting App.Title, "Lic", "NotifyExp" & CStr(iCt), True
  Next
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.General.ResetExpirationNotifications"
End Sub
Public Sub ResetSiteCode()
  On Error GoTo ErrorHandler
  '
  FMain.License.UserNumber(5) = 0
  lSessionID = FMain.License.TCSessionCode
  'txtSiteCode.Text = Trim(CStr(FPrimary.lSecCompID)) & " " & Trim(CStr(lSessionID))
  txtSiteCode.Text = Trim(CStr(lSecCompID)) & " " & Trim(CStr(lSessionID))
  mskSiteKey.Text = vbNullString
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FLicense.General.ResetSiteCode"
End Sub
