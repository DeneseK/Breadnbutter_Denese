VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9CD56991-2E37-11D2-8C87-00104B9E072A}#3.0#0"; "ssscrl30.ocx"
Begin VB.Form FContact2 
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmCompanyActions 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   6000
      TabIndex        =   140
      Top             =   720
      Width           =   4095
      Begin Threed.SSCommand cmdNewCompany 
         Height          =   315
         Left            =   0
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":0000
         Caption         =   "New Company"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdEditCompany 
         Height          =   315
         Left            =   1440
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   0
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":015A
         Caption         =   "View/Edit Company"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdWarm 
         Height          =   315
         Left            =   3510
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   0
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "W"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdCool 
         Height          =   315
         Left            =   3270
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   0
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "C"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdRand 
         Height          =   315
         Left            =   3750
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   0
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "R"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
   End
   Begin VB.Frame frmNavigation 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   960
      TabIndex        =   133
      Top             =   120
      Width           =   12375
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   5640
         TabIndex        =   139
         Text            =   "Combo1"
         Top             =   45
         Width           =   1695
      End
      Begin VB.CommandButton cmdBack 
         Height          =   375
         Left            =   0
         Picture         =   "FContact2.frx":02B4
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdForward 
         Height          =   375
         Left            =   555
         Picture         =   "FContact2.frx":03FE
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   1965
         TabIndex        =   134
         Top             =   45
         Width           =   3615
      End
      Begin Threed.SSCommand cmdRefresh 
         Height          =   315
         Left            =   7350
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":0548
         Caption         =   "Refresh"
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin VB.Label Label13 
         Caption         =   "Search:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         TabIndex        =   138
         Top             =   60
         Width           =   825
      End
   End
   Begin VB.Frame frmContactActions 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   120
      TabIndex        =   125
      Top             =   960
      Width           =   2655
      Begin Threed.SSCommand cmdNewContact 
         Height          =   315
         Left            =   0
         TabIndex        =   126
         TabStop         =   0   'False
         ToolTipText     =   "New Contact"
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":0AE2
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdEditContact 
         Height          =   315
         Left            =   360
         TabIndex        =   127
         TabStop         =   0   'False
         ToolTipText     =   "Edit Contact"
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":0C3C
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdSaveContact 
         Height          =   315
         Left            =   720
         TabIndex        =   128
         TabStop         =   0   'False
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":0D96
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdCancelContact 
         Height          =   315
         Left            =   1080
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":0EF0
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdDeleteContact 
         Height          =   315
         Left            =   1440
         TabIndex        =   130
         TabStop         =   0   'False
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":104A
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdPrintLabel 
         Height          =   315
         Left            =   1800
         TabIndex        =   131
         TabStop         =   0   'False
         ToolTipText     =   "Print Label"
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":11A4
         ButtonStyle     =   3
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdSetAppt 
         Height          =   315
         Left            =   2160
         TabIndex        =   132
         ToolTipText     =   "Set App. "
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":173E
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   1
      End
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Action"
      Height          =   495
      Left            =   9360
      TabIndex        =   124
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox fmeAuthStatus 
      Height          =   1230
      Left            =   3360
      ScaleHeight     =   1170
      ScaleWidth      =   5895
      TabIndex        =   106
      Top             =   5640
      Width           =   5955
      Begin VB.TextBox txtVersionShipped 
         DataField       =   "VersionShipped"
         Height          =   315
         Left            =   5055
         TabIndex        =   108
         Tag             =   "1"
         Top             =   450
         Width           =   765
      End
      Begin VB.TextBox txtAuthDays 
         DataField       =   "AuthDays"
         Height          =   315
         Left            =   4200
         TabIndex        =   107
         Tag             =   "1"
         Top             =   780
         Width           =   765
      End
      Begin SSDataWidgets_B.SSDBCombo cboAuthStatus 
         DataField       =   "AuthStatus"
         Height          =   315
         Left            =   810
         TabIndex        =   109
         Tag             =   "1"
         Top             =   780
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
      Begin SSDataWidgets_B.SSDBCombo cboShipStatus 
         DataField       =   "ShipStatus"
         Height          =   315
         Left            =   810
         TabIndex        =   110
         Tag             =   "1"
         Top             =   450
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
      Begin TDBDate6Ctl.TDBDate mskShipDate 
         Height          =   315
         Left            =   2700
         TabIndex        =   111
         Top             =   450
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FContact2.frx":1898
         Caption         =   "FContact2.frx":19B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FContact2.frx":1A1C
         Keys            =   "FContact2.frx":1A3A
         Spin            =   "FContact2.frx":1A98
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
      Begin TDBDate6Ctl.TDBDate mskAuthDate 
         Height          =   315
         Left            =   2700
         TabIndex        =   112
         Top             =   780
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FContact2.frx":1AC0
         Caption         =   "FContact2.frx":1BD8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FContact2.frx":1C44
         Keys            =   "FContact2.frx":1C62
         Spin            =   "FContact2.frx":1CC0
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
      Begin Threed.SSCommand cmdShipDate 
         Height          =   315
         Left            =   3780
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   450
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":1CE8
      End
      Begin Threed.SSCommand cmdAuthDate 
         Height          =   315
         Left            =   3780
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   780
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":2282
      End
      Begin SSDataWidgets_B.SSDBCombo cboDownloadStatus 
         DataField       =   "ShipStatus"
         Height          =   315
         Left            =   810
         TabIndex        =   115
         Tag             =   "1"
         Top             =   105
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
      Begin TDBDate6Ctl.TDBDate mskDownloadDate 
         Height          =   315
         Left            =   2700
         TabIndex        =   116
         Top             =   105
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FContact2.frx":281C
         Caption         =   "FContact2.frx":2934
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FContact2.frx":29A0
         Keys            =   "FContact2.frx":29BE
         Spin            =   "FContact2.frx":2A1C
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
      Begin Threed.SSCommand cmdDownloadDate 
         Height          =   315
         Left            =   3780
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   105
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact2.frx":2A44
      End
      Begin VB.Label lblExpires 
         BackStyle       =   0  'Transparent
         Caption         =   "Expires:"
         Height          =   285
         Left            =   4215
         TabIndex        =   123
         Top             =   150
         Width           =   1590
      End
      Begin VB.Label lblAuthRemaining 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "AuthRemaining"
         Height          =   315
         Left            =   5055
         TabIndex        =   122
         Tag             =   "1"
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Auth:"
         Height          =   285
         Index           =   2
         Left            =   45
         TabIndex        =   121
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   285
         Index           =   6
         Left            =   4230
         TabIndex        =   120
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping:"
         Height          =   285
         Index           =   1
         Left            =   45
         TabIndex        =   119
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Download:"
         Height          =   285
         Index           =   5
         Left            =   45
         TabIndex        =   118
         Top             =   135
         Width           =   765
      End
   End
   Begin VB.Frame FControls 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   4485
      Left            =   3360
      TabIndex        =   33
      Top             =   1080
      Width           =   9225
      Begin ActiveScroll.SSScroll scrollFrame 
         Height          =   3375
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5953
         _Version        =   196610
         BorderStyle     =   0
         HScrollType     =   0
         VScrollType     =   1
         Begin VB.PictureBox picCanvas 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0F2F8&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5655
            Left            =   120
            ScaleHeight     =   5655
            ScaleWidth      =   8295
            TabIndex        =   35
            Top             =   930
            Width           =   8295
            Begin VB.PictureBox picBranch 
               BackColor       =   &H0080FF80&
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   45
               ScaleHeight     =   855
               ScaleWidth      =   4095
               TabIndex        =   99
               Top             =   4680
               Width           =   4095
               Begin VB.CommandButton cmdEditBranch 
                  Caption         =   "Edit"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   103
                  Top             =   480
                  Width           =   495
               End
               Begin VB.CommandButton cmdDeleteBranch 
                  Caption         =   "Delete"
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   102
                  Top             =   480
                  Width           =   615
               End
               Begin VB.CommandButton cmdAddBranch 
                  Caption         =   "Add"
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   101
                  Top             =   480
                  Width           =   495
               End
               Begin VB.ComboBox cboBranch 
                  Height          =   315
                  Left            =   840
                  Style           =   2  'Dropdown List
                  TabIndex        =   100
                  Top             =   120
                  Width           =   3015
               End
               Begin VB.Label Label32 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Branch"
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
                  Left            =   30
                  TabIndex        =   104
                  Top             =   60
                  Width           =   615
               End
            End
            Begin VB.PictureBox picPhoneEmail 
               BackColor       =   &H00FFCCCC&
               BorderStyle     =   0  'None
               Height          =   1380
               Left            =   4155
               ScaleHeight     =   1380
               ScaleWidth      =   3300
               TabIndex        =   90
               Top             =   45
               Width           =   3300
               Begin VB.TextBox txtEmail 
                  DataField       =   "email"
                  Height          =   285
                  Left            =   840
                  MaxLength       =   100
                  TabIndex        =   91
                  Tag             =   "1"
                  Top             =   1035
                  Width           =   2385
               End
               Begin TDBMask6Ctl.TDBMask mskPhone2 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   92
                  Top             =   375
                  Width           =   2385
                  _Version        =   65536
                  _ExtentX        =   4207
                  _ExtentY        =   556
                  Caption         =   "FContact2.frx":2FDE
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "FContact2.frx":304A
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   1
                  AllowSpace      =   -1
                  AutoConvert     =   -1
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  ClipMode        =   1
                  CursorPosition  =   -1
                  DataProperty    =   0
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   -2147483640
                  Format          =   "999-999-9999 x 99999"
                  HighlightText   =   0
                  IMEMode         =   0
                  IMEStatus       =   0
                  LookupMode      =   0
                  LookupTable     =   ""
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  PromptChar      =   "_"
                  ReadOnly        =   0
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "___-___-____ x _____"
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask mskFAX 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   93
                  Top             =   705
                  Width           =   2385
                  _Version        =   65536
                  _ExtentX        =   4207
                  _ExtentY        =   556
                  Caption         =   "FContact2.frx":308C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "FContact2.frx":30F8
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   1
                  AllowSpace      =   -1
                  AutoConvert     =   -1
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  ClipMode        =   1
                  CursorPosition  =   -1
                  DataProperty    =   0
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   -2147483640
                  Format          =   "999-999-9999 x 99999"
                  HighlightText   =   0
                  IMEMode         =   0
                  IMEStatus       =   0
                  LookupMode      =   0
                  LookupTable     =   ""
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  PromptChar      =   "_"
                  ReadOnly        =   0
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "___-___-____ x _____"
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask mskPhone1 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   94
                  Top             =   45
                  Width           =   2385
                  _Version        =   65536
                  _ExtentX        =   4207
                  _ExtentY        =   556
                  Caption         =   "FContact2.frx":313A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "FContact2.frx":31A6
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   1
                  AllowSpace      =   -1
                  AutoConvert     =   -1
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  ClipMode        =   1
                  CursorPosition  =   -1
                  DataProperty    =   0
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   -2147483640
                  Format          =   "999-999-9999 x 99999"
                  HighlightText   =   0
                  IMEMode         =   0
                  IMEStatus       =   0
                  LookupMode      =   0
                  LookupTable     =   ""
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  PromptChar      =   "_"
                  ReadOnly        =   0
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "___-___-____ x _____"
                  Value           =   ""
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Phone1:"
                  Height          =   255
                  Index           =   11
                  Left            =   90
                  TabIndex        =   98
                  Top             =   105
                  Width           =   675
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Phone2:"
                  Height          =   255
                  Index           =   12
                  Left            =   90
                  TabIndex        =   97
                  Top             =   435
                  Width           =   675
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fax:"
                  Height          =   255
                  Index           =   13
                  Left            =   90
                  TabIndex        =   96
                  Top             =   765
                  Width           =   705
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-mail:"
                  Height          =   255
                  Index           =   0
                  Left            =   90
                  TabIndex        =   95
                  Top             =   1065
                  Width           =   615
               End
            End
            Begin VB.PictureBox picGeneral 
               BackColor       =   &H00FFCCCC&
               BorderStyle     =   0  'None
               Height          =   3540
               Left            =   45
               ScaleHeight     =   3540
               ScaleWidth      =   4065
               TabIndex        =   68
               Top             =   45
               Width           =   4065
               Begin VB.ComboBox cboContactType 
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   75
                  Top             =   690
                  Width           =   2775
               End
               Begin VB.ComboBox cboSalutation 
                  Height          =   315
                  ItemData        =   "FContact2.frx":31E8
                  Left            =   1230
                  List            =   "FContact2.frx":31F2
                  Style           =   2  'Dropdown List
                  TabIndex        =   74
                  Top             =   360
                  Width           =   765
               End
               Begin VB.TextBox txtTitle 
                  DataField       =   "Title"
                  Height          =   285
                  Left            =   2430
                  TabIndex        =   73
                  Tag             =   "1"
                  Top             =   390
                  Width           =   1545
               End
               Begin VB.TextBox txtFirstName 
                  DataField       =   "FirstName"
                  Height          =   285
                  Left            =   1230
                  TabIndex        =   72
                  Tag             =   "1"
                  Top             =   60
                  Width           =   1185
               End
               Begin VB.TextBox txtLastName 
                  DataField       =   "LastName"
                  Height          =   285
                  Left            =   2430
                  TabIndex        =   71
                  Tag             =   "1"
                  Top             =   60
                  Width           =   1545
               End
               Begin VB.CheckBox chkSelected 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFCCCC&
                  Caption         =   "Select:"
                  DataField       =   "BetaTester"
                  Height          =   285
                  Left            =   0
                  TabIndex        =   70
                  Tag             =   "1"
                  Top             =   1380
                  Width           =   945
               End
               Begin VB.TextBox txtNotes 
                  DataField       =   "Notes"
                  Height          =   1200
                  Left            =   675
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   69
                  Tag             =   "1"
                  Top             =   1755
                  Width           =   3345
               End
               Begin TDBDate6Ctl.TDBDate mskCreated 
                  Height          =   315
                  Left            =   3015
                  TabIndex        =   76
                  Top             =   1035
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
                  _ExtentY        =   556
                  Calendar        =   "FContact2.frx":3200
                  Caption         =   "FContact2.frx":3318
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "FContact2.frx":3384
                  Keys            =   "FContact2.frx":33A2
                  Spin            =   "FContact2.frx":3400
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   1
                  BackColor       =   14737632
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
                  ReadOnly        =   -1
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "02/10/2003"
                  ValidateMode    =   0
                  ValueVT         =   7
                  Value           =   37662
                  CenturyMode     =   0
               End
               Begin SSDataWidgets_B.SSDBCombo cboStatus 
                  DataField       =   "Status"
                  Height          =   315
                  Left            =   1215
                  TabIndex        =   77
                  Tag             =   "1"
                  Top             =   1035
                  Width           =   1785
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
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _StockProps     =   93
                  ForeColor       =   -2147483640
                  BackColor       =   65280
               End
               Begin SSDataWidgets_B.SSDBCombo cboSource 
                  DataField       =   "Source"
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   78
                  Tag             =   "1"
                  Top             =   1395
                  Width           =   2235
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
                  _ExtentX        =   3942
                  _ExtentY        =   556
                  _StockProps     =   93
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
               End
               Begin TDBNumber6Ctl.TDBNumber tnmRate 
                  Height          =   315
                  Left            =   525
                  TabIndex        =   79
                  Top             =   3045
                  Width           =   825
                  _Version        =   65536
                  _ExtentX        =   1455
                  _ExtentY        =   556
                  Calculator      =   "FContact2.frx":3428
                  Caption         =   "FContact2.frx":3448
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "FContact2.frx":34B4
                  Keys            =   "FContact2.frx":34D2
                  Spin            =   "FContact2.frx":351C
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   1
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "$ ###,###.00;($ ###,###.00);$0.00;$0.00"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   -2147483640
                  Format          =   "#####0"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   999999
                  MinValue        =   0
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  NegativeColor   =   255
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  ReadOnly        =   0
                  Separator       =   ","
                  ShowContextMenu =   -1
                  ValueVT         =   2012807169
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin GTMaskDate.GTMaskDate mskRateExpDate 
                  Height          =   315
                  Left            =   2730
                  TabIndex        =   80
                  Tag             =   "1"
                  Top             =   3045
                  Width           =   1275
                  _Version        =   65537
                  _ExtentX        =   2249
                  _ExtentY        =   556
                  _StockProps     =   77
                  BackColor       =   -2147483643
                  KeepFocusOnError=   -1  'True
                  BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  MaskCentury     =   2
                  DataField       =   "ShipDate"
                  BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.Label Label10 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Status:"
                  Height          =   285
                  Index           =   0
                  Left            =   15
                  TabIndex        =   89
                  Top             =   1035
                  Width           =   495
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "First/Last Name:"
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   88
                  Top             =   90
                  Width           =   1215
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mr./Ms.:"
                  Height          =   255
                  Index           =   4
                  Left            =   0
                  TabIndex        =   87
                  Top             =   420
                  Width           =   885
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Title:"
                  Height          =   255
                  Index           =   5
                  Left            =   2040
                  TabIndex        =   86
                  Top             =   420
                  Width           =   375
               End
               Begin VB.Label Label21 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contact Type:"
                  Height          =   285
                  Left            =   0
                  TabIndex        =   85
                  Top             =   750
                  Width           =   1065
               End
               Begin VB.Label Label2 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Source:"
                  Height          =   285
                  Left            =   1095
                  TabIndex        =   84
                  Top             =   1425
                  Width           =   615
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Notes:"
                  Height          =   255
                  Index           =   1
                  Left            =   75
                  TabIndex        =   83
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.Label Label10 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rate:"
                  Height          =   285
                  Index           =   3
                  Left            =   75
                  TabIndex        =   82
                  Top             =   3075
                  Width           =   1005
               End
               Begin VB.Label Label23 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exp Date:"
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   81
                  Top             =   3075
                  Width           =   825
               End
            End
            Begin VB.PictureBox picShipping 
               BackColor       =   &H00ABDDF8&
               BorderStyle     =   0  'None
               Height          =   1575
               Left            =   4155
               ScaleHeight     =   1575
               ScaleWidth      =   3975
               TabIndex        =   55
               Top             =   1455
               Width           =   3975
               Begin VB.CheckBox chkPreferredAddress 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFCCCC&
                  Caption         =   "Preferred Address:"
                  DataField       =   "BetaTester"
                  Height          =   285
                  Index           =   0
                  Left            =   2205
                  TabIndex        =   61
                  Tag             =   "1"
                  Top             =   1230
                  Width           =   1695
               End
               Begin VB.TextBox txtZIP 
                  Height          =   315
                  Left            =   855
                  TabIndex        =   60
                  Top             =   1170
                  Width           =   1215
               End
               Begin VB.TextBox txtCity 
                  DataField       =   "City"
                  Height          =   285
                  Left            =   855
                  TabIndex        =   59
                  Tag             =   "1"
                  Top             =   870
                  Width           =   1995
               End
               Begin VB.TextBox txtAddress2 
                  DataField       =   "Address2"
                  Height          =   285
                  Left            =   855
                  TabIndex        =   58
                  Tag             =   "1"
                  Top             =   570
                  Width           =   3075
               End
               Begin VB.TextBox txtAddress1 
                  DataField       =   "Address1"
                  Height          =   285
                  Left            =   855
                  TabIndex        =   57
                  Tag             =   "1"
                  Top             =   270
                  Width           =   3075
               End
               Begin VB.TextBox txtState 
                  DataField       =   "State"
                  Height          =   315
                  Left            =   3405
                  TabIndex        =   56
                  Tag             =   "1"
                  Top             =   870
                  Width           =   525
               End
               Begin VB.Label Label28 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Shipping Address"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   30
                  TabIndex        =   67
                  Top             =   15
                  Width           =   1935
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address 1:"
                  Height          =   255
                  Index           =   6
                  Left            =   45
                  TabIndex        =   66
                  Top             =   300
                  Width           =   885
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address 2:"
                  Height          =   255
                  Index           =   7
                  Left            =   45
                  TabIndex        =   65
                  Top             =   630
                  Width           =   885
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "City:"
                  Height          =   255
                  Index           =   8
                  Left            =   45
                  TabIndex        =   64
                  Top             =   900
                  Width           =   885
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "State:"
                  Height          =   255
                  Index           =   0
                  Left            =   2925
                  TabIndex        =   63
                  Top             =   900
                  Width           =   885
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ZIP:"
                  Height          =   255
                  Index           =   1
                  Left            =   45
                  TabIndex        =   62
                  Top             =   1230
                  Width           =   375
               End
            End
            Begin VB.PictureBox picMailing 
               BackColor       =   &H00FFCCCC&
               BorderStyle     =   0  'None
               Height          =   1530
               Left            =   4155
               ScaleHeight     =   1530
               ScaleWidth      =   3990
               TabIndex        =   42
               Top             =   3105
               Width           =   3990
               Begin VB.TextBox txtMailState 
                  DataField       =   "PermMailState"
                  Height          =   315
                  Left            =   3390
                  TabIndex        =   48
                  Tag             =   "1"
                  Top             =   840
                  Width           =   525
               End
               Begin VB.TextBox txtMailAddress1 
                  DataField       =   "PermMailAddress1"
                  Height          =   285
                  Left            =   840
                  TabIndex        =   47
                  Tag             =   "1"
                  Top             =   240
                  Width           =   3075
               End
               Begin VB.TextBox txtMailAddress2 
                  DataField       =   "PermMailAddress2"
                  Height          =   285
                  Left            =   840
                  TabIndex        =   46
                  Tag             =   "1"
                  Top             =   540
                  Width           =   3075
               End
               Begin VB.TextBox txtMailCity 
                  DataField       =   "PermMailCity"
                  Height          =   285
                  Left            =   840
                  TabIndex        =   45
                  Tag             =   "1"
                  Top             =   840
                  Width           =   1995
               End
               Begin VB.TextBox txtMailZIP 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   44
                  Top             =   1140
                  Width           =   1215
               End
               Begin VB.CheckBox chkPreferredAddress 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFCCCC&
                  Caption         =   "Preferred Address:"
                  DataField       =   "BetaTester"
                  Height          =   285
                  Index           =   1
                  Left            =   2190
                  TabIndex        =   43
                  Tag             =   "1"
                  Top             =   1200
                  Width           =   1695
               End
               Begin VB.Label Label29 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mailing Address"
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
                  Left            =   60
                  TabIndex        =   54
                  Top             =   15
                  Width           =   1695
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ZIP:"
                  Height          =   255
                  Index           =   10
                  Left            =   30
                  TabIndex        =   53
                  Top             =   1200
                  Width           =   375
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "State:"
                  Height          =   255
                  Index           =   14
                  Left            =   2910
                  TabIndex        =   52
                  Top             =   870
                  Width           =   435
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "City:"
                  Height          =   255
                  Index           =   15
                  Left            =   30
                  TabIndex        =   51
                  Top             =   870
                  Width           =   885
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address 2:"
                  Height          =   255
                  Index           =   16
                  Left            =   30
                  TabIndex        =   50
                  Top             =   600
                  Width           =   885
               End
               Begin VB.Label lblLabels 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address 1:"
                  Height          =   255
                  Index           =   17
                  Left            =   30
                  TabIndex        =   49
                  Top             =   300
                  Width           =   885
               End
            End
            Begin VB.PictureBox picPCEmail 
               BackColor       =   &H00FFCCCC&
               BorderStyle     =   0  'None
               Height          =   960
               Left            =   45
               ScaleHeight     =   960
               ScaleWidth      =   4065
               TabIndex        =   36
               Top             =   3645
               Width           =   4065
               Begin VB.TextBox txtPCEmail 
                  Height          =   285
                  Left            =   1125
                  TabIndex        =   38
                  Top             =   285
                  Width           =   2775
               End
               Begin VB.TextBox txtPCEmailPassword 
                  Height          =   285
                  Left            =   1125
                  TabIndex        =   37
                  Top             =   585
                  Width           =   2775
               End
               Begin VB.Label Label6 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address:"
                  Height          =   285
                  Left            =   45
                  TabIndex        =   41
                  Top             =   315
                  Width           =   795
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFCCCC&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Password:"
                  Height          =   255
                  Left            =   45
                  TabIndex        =   40
                  Top             =   615
                  Width           =   765
               End
               Begin VB.Label Label30 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PowerClaim Email Address"
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
                  Left            =   30
                  TabIndex        =   39
                  Top             =   30
                  Width           =   2775
               End
            End
            Begin VB.Label lblDoNotContact 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H000000FF&
               Height          =   855
               Left            =   7500
               TabIndex        =   105
               Top             =   120
               Width           =   675
            End
         End
      End
   End
   Begin VB.TextBox txtCompanyNote 
      BackColor       =   &H8000000B&
      Height          =   615
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintList 
      Caption         =   "Print List"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdClipBoard 
      Caption         =   "Copy to Clipboard"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   1695
   End
   Begin TabDlg.SSTab tbContacts 
      Height          =   4215
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   7435
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Company"
      TabPicture(0)   =   "FContact2.frx":3544
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstBranch"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwContact"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Groups"
      TabPicture(1)   =   "FContact2.frx":3560
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwPMContacts"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frmGroupList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Search"
      TabPicture(2)   =   "FContact2.frx":357C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lvwSearchContacts"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frmSearch"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame frmGroupList 
         BorderStyle     =   0  'None
         Caption         =   "frmGroupList"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   25
         Top             =   500
         Width           =   3015
         Begin VB.ListBox lstGroups 
            Height          =   1035
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   3015
         End
         Begin VB.CheckBox chkAlpha 
            Caption         =   "Sort Alphabetically"
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   1290
            Width           =   1695
         End
      End
      Begin VB.Frame frmSearch 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2000
         Left            =   120
         TabIndex        =   3
         Top             =   500
         Width           =   3015
         Begin VB.CommandButton cmdSearch 
            Appearance      =   0  'Flat
            Caption         =   "Search"
            Height          =   300
            Left            =   2235
            TabIndex        =   4
            Top             =   1650
            Width           =   735
         End
         Begin ActiveScroll.SSScroll scrollSearchFields 
            Height          =   1215
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   2143
            _Version        =   196610
            BorderStyle     =   2
            HScrollType     =   0
            ScrollingHeight =   3425
            Begin VB.TextBox txtSearchNotes 
               Height          =   285
               Left            =   600
               MaxLength       =   100
               TabIndex        =   14
               Top             =   2970
               Width           =   2055
            End
            Begin VB.TextBox txtSearchCity 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   13
               Top             =   1560
               Width           =   2055
            End
            Begin VB.TextBox txtSearchZip 
               Height          =   315
               Left            =   600
               MaxLength       =   25
               TabIndex        =   12
               Top             =   2280
               Width           =   2055
            End
            Begin VB.TextBox txtSearchSource 
               Height          =   285
               Left            =   600
               MaxLength       =   100
               TabIndex        =   11
               Top             =   2640
               Width           =   2070
            End
            Begin VB.TextBox txtSearchFirstName 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   10
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtSearchLastName 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   9
               Top             =   840
               Width           =   2055
            End
            Begin VB.TextBox txtSearchCompany 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   8
               Top             =   1200
               Width           =   2055
            End
            Begin VB.ComboBox cboSearchStatus 
               Height          =   315
               Left            =   600
               TabIndex        =   7
               Text            =   "Customer"
               Top             =   120
               Width           =   2055
            End
            Begin VB.TextBox txtSearchState 
               Height          =   315
               Left            =   600
               MaxLength       =   5
               TabIndex        =   6
               Top             =   1920
               Width           =   2055
            End
            Begin VB.Label Label31 
               Caption         =   "Notes:"
               Height          =   255
               Left            =   15
               TabIndex        =   23
               Top             =   3000
               Width           =   495
            End
            Begin VB.Label Label4 
               Caption         =   "City:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   22
               Top             =   1590
               Width           =   375
            End
            Begin VB.Label Label5 
               Caption         =   "State:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   21
               Top             =   1950
               Width           =   495
            End
            Begin VB.Label Label24 
               Caption         =   "Zip:"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   15
               TabIndex        =   20
               Top             =   2310
               Width           =   345
            End
            Begin VB.Label Label12 
               Caption         =   "Source:"
               Height          =   255
               Left            =   15
               TabIndex        =   19
               Top             =   2655
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "First:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   15
               TabIndex        =   18
               Top             =   510
               Width           =   825
            End
            Begin VB.Label Label8 
               Caption         =   "Last:"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   15
               TabIndex        =   17
               Top             =   870
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Comp.:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   15
               TabIndex        =   16
               Top             =   1230
               Width           =   525
            End
            Begin VB.Label Label7 
               Caption         =   "Status:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   15
               TabIndex        =   15
               Top             =   150
               Width           =   525
            End
         End
         Begin VB.Label lblResults 
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1710
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lvwContact 
         Height          =   2325
         Left            =   -74925
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1740
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilContact"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwPMContacts 
         Height          =   1245
         Left            =   -74880
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2180
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2196
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilContact"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwSearchContacts 
         Height          =   1365
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2660
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2408
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilContact"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstBranch 
         Height          =   1125
         Left            =   -74880
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1984
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilContact"
         SmallIcons      =   "ilContact"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
         EndProperty
      End
   End
End
Attribute VB_Name = "FContact2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
