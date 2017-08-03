VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{9CD56991-2E37-11D2-8C87-00104B9E072A}#3.0#0"; "ssscrl30.ocx"
Begin VB.Form FContact 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   2085
   ClientTop       =   3915
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox StarPicture 
      Height          =   255
      Index           =   0
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   204
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   13320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   203
      Top             =   520
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   13800
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   202
      Top             =   520
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   13560
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   201
      Top             =   520
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   14040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   200
      Top             =   520
      Width           =   255
   End
   Begin VB.PictureBox StarPicture 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   14280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   199
      Top             =   520
      Width           =   255
   End
   Begin VB.PictureBox fmePVAuthStatus 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   9195
      ScaleHeight     =   1995
      ScaleWidth      =   5685
      TabIndex        =   163
      Top             =   4470
      Width           =   5685
      Begin VB.TextBox txtPVSaleDays 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   169
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtPVGraceDays 
         DataField       =   "VersionShipped"
         Height          =   315
         Left            =   5085
         MaxLength       =   5
         TabIndex        =   168
         Tag             =   "1"
         Top             =   90
         Width           =   555
      End
      Begin VB.TextBox txtPVPendingDays 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5070
         MaxLength       =   4
         TabIndex        =   167
         Top             =   510
         Width           =   555
      End
      Begin VB.TextBox txtPVAuths 
         Height          =   315
         Left            =   3120
         MaxLength       =   1
         TabIndex        =   166
         Top             =   75
         Width           =   375
      End
      Begin VB.TextBox txtPVVersionShipped 
         DataField       =   "VersionShipped"
         Height          =   315
         Left            =   4920
         TabIndex        =   165
         Tag             =   "1"
         Top             =   1245
         Width           =   675
      End
      Begin VB.TextBox txtPVAuthDays 
         DataField       =   "AuthDays"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   164
         Tag             =   "1"
         Top             =   1560
         Width           =   675
      End
      Begin SSDataWidgets_B.SSDBCombo cboPVAuthStatus 
         DataField       =   "AuthStatus"
         Height          =   315
         Left            =   810
         TabIndex        =   170
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
         TabIndex        =   171
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
         TabIndex        =   172
         Top             =   1245
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FContact.frx":0000
         Caption         =   "FContact.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FContact.frx":0184
         Keys            =   "FContact.frx":01A2
         Spin            =   "FContact.frx":0200
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
         TabIndex        =   173
         Top             =   1560
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FContact.frx":0228
         Caption         =   "FContact.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FContact.frx":03AC
         Keys            =   "FContact.frx":03CA
         Spin            =   "FContact.frx":0428
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
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   1245
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact.frx":0450
      End
      Begin Threed.SSCommand cmdPVAuthDate 
         Height          =   315
         Left            =   3780
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   1575
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact.frx":09EA
      End
      Begin SSDataWidgets_B.SSDBCombo cboPVDownloadStatus 
         DataField       =   "ShipStatus"
         Height          =   315
         Left            =   810
         TabIndex        =   176
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
         TabIndex        =   177
         Top             =   900
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FContact.frx":0F84
         Caption         =   "FContact.frx":109C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FContact.frx":1108
         Keys            =   "FContact.frx":1126
         Spin            =   "FContact.frx":1184
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
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   900
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact.frx":11AC
      End
      Begin TDBDate6Ctl.TDBDate mskPVSaleDate 
         Height          =   315
         Left            =   885
         TabIndex        =   179
         Top             =   465
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   556
         Calendar        =   "FContact.frx":1746
         Caption         =   "FContact.frx":185E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FContact.frx":18CA
         Keys            =   "FContact.frx":18E8
         Spin            =   "FContact.frx":1946
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
         Left            =   1980
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   465
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FContact.frx":196E
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Days:"
         Height          =   285
         Index           =   13
         Left            =   2400
         TabIndex        =   192
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Grace Period:"
         Height          =   285
         Index           =   11
         Left            =   3600
         TabIndex        =   191
         Top             =   120
         Width           =   1485
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Pending Days:"
         Height          =   285
         Index           =   10
         Left            =   3960
         TabIndex        =   190
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Date:"
         Height          =   285
         Left            =   45
         TabIndex        =   189
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Online Auths:"
         Height          =   255
         Left            =   1440
         TabIndex        =   188
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "PowerClaim PV"
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
         TabIndex        =   187
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label lblPVExpires 
         BackStyle       =   0  'Transparent
         Caption         =   "Expires:"
         Height          =   285
         Left            =   4215
         TabIndex        =   186
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label lblPVAuthRemaining 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "AuthRemaining"
         Height          =   315
         Left            =   4920
         TabIndex        =   185
         Tag             =   "1"
         Top             =   1575
         Width           =   675
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Auth:"
         Height          =   285
         Index           =   7
         Left            =   45
         TabIndex        =   184
         Top             =   1605
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ver.:"
         Height          =   285
         Index           =   1
         Left            =   4230
         TabIndex        =   183
         Top             =   1275
         Width           =   645
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping:"
         Height          =   285
         Index           =   6
         Left            =   45
         TabIndex        =   182
         Top             =   1275
         Width           =   765
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Download:"
         Height          =   285
         Index           =   4
         Left            =   45
         TabIndex        =   181
         Top             =   930
         Width           =   765
      End
   End
   Begin Threed.SSCommand cmdDeleteCompany 
      Height          =   315
      Left            =   11700
      TabIndex        =   15
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   196610
      Caption         =   "Delete Company"
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0F2F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   3360
      ScaleHeight     =   5715
      ScaleWidth      =   11700
      TabIndex        =   101
      Top             =   960
      Visible         =   0   'False
      Width           =   11700
      Begin VB.PictureBox picCustGroups 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   45
         ScaleHeight     =   1335
         ScaleWidth      =   4065
         TabIndex        =   193
         Top             =   2220
         Width           =   4065
         Begin MSComctlLib.ListView lstCustGroups 
            Height          =   1230
            Left            =   1200
            TabIndex        =   194
            Top             =   45
            Width           =   2800
            _ExtentX        =   4948
            _ExtentY        =   2170
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            OLEDropMode     =   1
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            OLEDropMode     =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Groups"
               Object.Width           =   4234
            EndProperty
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Custom Groups:"
            Height          =   315
            Index           =   1
            Left            =   45
            TabIndex        =   195
            Top             =   90
            Width           =   1125
         End
      End
      Begin VB.PictureBox picBranch 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   1050
         Left            =   8310
         ScaleHeight     =   1050
         ScaleWidth      =   3315
         TabIndex        =   140
         Top             =   2490
         Width           =   3315
         Begin VB.CheckBox chkContactByEmail 
            Caption         =   "Contact By Email"
            Height          =   255
            Left            =   75
            TabIndex        =   160
            Top             =   780
            Width           =   1935
         End
         Begin VB.TextBox txtWebPassword 
            Height          =   285
            Left            =   1320
            MaxLength       =   25
            TabIndex        =   159
            Top             =   465
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cboBranch 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Web Password:"
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Branch:"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.PictureBox picPhoneEmail 
         BackColor       =   &H00FFCCCC&
         BorderStyle     =   0  'None
         Height          =   1380
         Left            =   8295
         ScaleHeight     =   1380
         ScaleWidth      =   3300
         TabIndex        =   135
         Top             =   60
         Width           =   3300
         Begin VB.TextBox txtPhone1 
            DataField       =   "email"
            Height          =   315
            Left            =   840
            MaxLength       =   25
            TabIndex        =   39
            Tag             =   "1"
            Top             =   50
            Width           =   2385
         End
         Begin VB.TextBox txtPhone2 
            DataField       =   "email"
            Height          =   315
            Left            =   840
            MaxLength       =   25
            TabIndex        =   40
            Tag             =   "1"
            Top             =   380
            Width           =   2385
         End
         Begin VB.TextBox txtFax 
            DataField       =   "email"
            Height          =   315
            Left            =   840
            MaxLength       =   25
            TabIndex        =   41
            Tag             =   "1"
            Top             =   710
            Width           =   2385
         End
         Begin VB.TextBox txtEmail 
            DataField       =   "email"
            Height          =   285
            Left            =   840
            MaxLength       =   100
            TabIndex        =   42
            Tag             =   "1"
            Top             =   1040
            Width           =   2385
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone1:"
            Height          =   255
            Index           =   11
            Left            =   90
            TabIndex        =   139
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
            TabIndex        =   138
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
            TabIndex        =   137
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
            TabIndex        =   136
            Top             =   1065
            Width           =   615
         End
      End
      Begin VB.PictureBox picGeneral 
         BackColor       =   &H00FFCCCC&
         BorderStyle     =   0  'None
         Height          =   2130
         Left            =   45
         ScaleHeight     =   2130
         ScaleWidth      =   4065
         TabIndex        =   23
         Top             =   45
         Width           =   4065
         Begin VB.ComboBox cboStatus 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   1050
            Width           =   2775
         End
         Begin VB.ComboBox cboContactType 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   690
            Width           =   2775
         End
         Begin VB.ComboBox cboSalutation 
            Height          =   315
            ItemData        =   "FContact.frx":1F08
            Left            =   1230
            List            =   "FContact.frx":1F12
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   765
         End
         Begin VB.TextBox txtTitle 
            DataField       =   "Title"
            Height          =   285
            Left            =   2430
            MaxLength       =   50
            TabIndex        =   19
            Tag             =   "1"
            Top             =   390
            Width           =   1545
         End
         Begin VB.TextBox txtFirstName 
            DataField       =   "FirstName"
            Height          =   285
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   16
            Tag             =   "1"
            Top             =   60
            Width           =   1185
         End
         Begin VB.TextBox txtLastName 
            DataField       =   "LastName"
            Height          =   285
            Left            =   2430
            MaxLength       =   50
            TabIndex        =   17
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
            Left            =   30
            TabIndex        =   22
            Tag             =   "1"
            Top             =   1380
            Visible         =   0   'False
            Width           =   945
         End
         Begin SSDataWidgets_B.SSDBCombo cboSource 
            DataField       =   "Source"
            Height          =   315
            Left            =   1770
            TabIndex        =   24
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
            TabIndex        =   25
            Top             =   1750
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   556
            Calculator      =   "FContact.frx":1F20
            Caption         =   "FContact.frx":1F40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FContact.frx":1FAC
            Keys            =   "FContact.frx":1FCA
            Spin            =   "FContact.frx":2014
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
            ValueVT         =   1
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin GTMaskDate.GTMaskDate mskRateExpDate 
            Height          =   315
            Left            =   2730
            TabIndex        =   26
            Tag             =   "1"
            Top             =   1750
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
         Begin TDBDate6Ctl.TDBDate mskCreated 
            Height          =   315
            Left            =   3015
            TabIndex        =   21
            Top             =   1035
            Visible         =   0   'False
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   556
            Calendar        =   "FContact.frx":203C
            Caption         =   "FContact.frx":2154
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FContact.frx":21C0
            Keys            =   "FContact.frx":21DE
            Spin            =   "FContact.frx":223C
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
         Begin VB.Label Label10 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   134
            Top             =   1100
            Width           =   495
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "First/Last Name:"
            Height          =   255
            Index           =   2
            Left            =   30
            TabIndex        =   133
            Top             =   90
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Mr./Ms.:"
            Height          =   255
            Index           =   4
            Left            =   45
            TabIndex        =   132
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
            TabIndex        =   131
            Top             =   420
            Width           =   375
         End
         Begin VB.Label Label21 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Type:"
            Height          =   285
            Left            =   30
            TabIndex        =   130
            Top             =   750
            Width           =   1065
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Source:"
            Height          =   285
            Left            =   1095
            TabIndex        =   129
            Top             =   1425
            Width           =   615
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Rate:"
            Height          =   285
            Index           =   3
            Left            =   75
            TabIndex        =   128
            Top             =   1800
            Width           =   1005
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Exp Date:"
            Height          =   255
            Left            =   1920
            TabIndex        =   127
            Top             =   1800
            Width           =   825
         End
      End
      Begin VB.PictureBox picShipping 
         BackColor       =   &H00ABDDF8&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   4200
         ScaleHeight     =   1575
         ScaleWidth      =   3975
         TabIndex        =   120
         Top             =   90
         Width           =   3975
         Begin VB.CheckBox chkPreferredAddress 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFCCCC&
            Caption         =   "Preferred Address:"
            DataField       =   "BetaTester"
            Height          =   285
            Index           =   0
            Left            =   2205
            TabIndex        =   32
            Tag             =   "1"
            Top             =   1230
            Width           =   1695
         End
         Begin VB.TextBox txtZIP 
            Height          =   315
            Left            =   855
            MaxLength       =   10
            TabIndex        =   31
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox txtCity 
            DataField       =   "City"
            Height          =   285
            Left            =   855
            MaxLength       =   30
            TabIndex        =   29
            Tag             =   "1"
            Top             =   870
            Width           =   1995
         End
         Begin VB.TextBox txtAddress2 
            DataField       =   "Address2"
            Height          =   285
            Left            =   855
            MaxLength       =   30
            TabIndex        =   28
            Tag             =   "1"
            Top             =   570
            Width           =   3075
         End
         Begin VB.TextBox txtAddress1 
            DataField       =   "Address1"
            Height          =   285
            Left            =   855
            MaxLength       =   30
            TabIndex        =   27
            Tag             =   "1"
            Top             =   270
            Width           =   3075
         End
         Begin VB.TextBox txtState 
            DataField       =   "State"
            Height          =   315
            Left            =   3405
            MaxLength       =   20
            TabIndex        =   30
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
            TabIndex        =   126
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
            TabIndex        =   125
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
            TabIndex        =   124
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
            TabIndex        =   123
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
            TabIndex        =   122
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
            TabIndex        =   121
            Top             =   1230
            Width           =   375
         End
      End
      Begin VB.PictureBox picMailing 
         BackColor       =   &H00FFCCCC&
         BorderStyle     =   0  'None
         Height          =   1530
         Left            =   4215
         ScaleHeight     =   1530
         ScaleWidth      =   3990
         TabIndex        =   113
         Top             =   1740
         Width           =   3990
         Begin VB.TextBox txtMailState 
            DataField       =   "PermMailState"
            Height          =   315
            Left            =   3390
            MaxLength       =   50
            TabIndex        =   36
            Tag             =   "1"
            Top             =   840
            Width           =   525
         End
         Begin VB.TextBox txtMailAddress1 
            DataField       =   "PermMailAddress1"
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   33
            Tag             =   "1"
            Top             =   240
            Width           =   3075
         End
         Begin VB.TextBox txtMailAddress2 
            DataField       =   "PermMailAddress2"
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   34
            Tag             =   "1"
            Top             =   540
            Width           =   3075
         End
         Begin VB.TextBox txtMailCity 
            DataField       =   "PermMailCity"
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   35
            Tag             =   "1"
            Top             =   840
            Width           =   1995
         End
         Begin VB.TextBox txtMailZIP 
            Height          =   315
            Left            =   840
            MaxLength       =   50
            TabIndex        =   37
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
            TabIndex        =   38
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
            TabIndex        =   119
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
            TabIndex        =   118
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
            TabIndex        =   117
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
            TabIndex        =   116
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
            TabIndex        =   115
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
            TabIndex        =   114
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.PictureBox picPCEmail 
         BackColor       =   &H00FFCCCC&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   8295
         ScaleHeight     =   960
         ScaleWidth      =   3315
         TabIndex        =   109
         Top             =   1485
         Visible         =   0   'False
         Width           =   3315
         Begin VB.TextBox txtPCEmail 
            Height          =   285
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   43
            Top             =   285
            Width           =   2055
         End
         Begin VB.TextBox txtPCEmailPassword 
            Height          =   285
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   44
            Top             =   585
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   285
            Left            =   45
            TabIndex        =   112
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFCCCC&
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Height          =   255
            Left            =   45
            TabIndex        =   111
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
            TabIndex        =   110
            Top             =   30
            Width           =   2775
         End
      End
      Begin VB.PictureBox fmeAuthStatus 
         BorderStyle     =   0  'None
         Height          =   1995
         Left            =   75
         ScaleHeight     =   1995
         ScaleWidth      =   5685
         TabIndex        =   102
         Top             =   3600
         Width           =   5685
         Begin VB.TextBox txtSaleDays 
            Height          =   315
            Left            =   3240
            MaxLength       =   4
            TabIndex        =   161
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtGraceDays 
            DataField       =   "VersionShipped"
            Height          =   315
            Left            =   5085
            MaxLength       =   5
            TabIndex        =   152
            Tag             =   "1"
            Top             =   90
            Width           =   555
         End
         Begin VB.TextBox txtPendingDays 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5070
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   151
            Top             =   510
            Width           =   555
         End
         Begin VB.TextBox txtAuths 
            Height          =   315
            Left            =   3240
            MaxLength       =   1
            TabIndex        =   150
            Top             =   75
            Width           =   375
         End
         Begin VB.TextBox txtVersionShipped 
            DataField       =   "VersionShipped"
            Height          =   315
            Left            =   4920
            TabIndex        =   55
            Tag             =   "1"
            Top             =   1245
            Width           =   675
         End
         Begin VB.TextBox txtAuthDays 
            DataField       =   "AuthDays"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4200
            TabIndex        =   56
            Tag             =   "1"
            Top             =   1560
            Width           =   675
         End
         Begin SSDataWidgets_B.SSDBCombo cboAuthStatus 
            DataField       =   "AuthStatus"
            Height          =   315
            Left            =   810
            TabIndex        =   48
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
         Begin SSDataWidgets_B.SSDBCombo cboShipStatus 
            DataField       =   "ShipStatus"
            Height          =   315
            Left            =   810
            TabIndex        =   47
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
         Begin TDBDate6Ctl.TDBDate mskShipDate 
            Height          =   315
            Left            =   2700
            TabIndex        =   51
            Top             =   1245
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            Calendar        =   "FContact.frx":2264
            Caption         =   "FContact.frx":237C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FContact.frx":23E8
            Keys            =   "FContact.frx":2406
            Spin            =   "FContact.frx":2464
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
            TabIndex        =   53
            Top             =   1560
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            Calendar        =   "FContact.frx":248C
            Caption         =   "FContact.frx":25A4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FContact.frx":2610
            Keys            =   "FContact.frx":262E
            Spin            =   "FContact.frx":268C
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
         Begin Threed.SSCommand cmdShipDate 
            Height          =   315
            Left            =   3780
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1245
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            _Version        =   196610
            PictureFrames   =   1
            Picture         =   "FContact.frx":26B4
         End
         Begin Threed.SSCommand cmdAuthDate 
            Height          =   315
            Left            =   3780
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1575
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            _Version        =   196610
            PictureFrames   =   1
            Picture         =   "FContact.frx":2C4E
         End
         Begin SSDataWidgets_B.SSDBCombo cboDownloadStatus 
            DataField       =   "ShipStatus"
            Height          =   315
            Left            =   810
            TabIndex        =   46
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
         Begin TDBDate6Ctl.TDBDate mskDownloadDate 
            Height          =   315
            Left            =   2700
            TabIndex        =   49
            Top             =   900
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            Calendar        =   "FContact.frx":31E8
            Caption         =   "FContact.frx":3300
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FContact.frx":336C
            Keys            =   "FContact.frx":338A
            Spin            =   "FContact.frx":33E8
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
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   900
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            _Version        =   196610
            PictureFrames   =   1
            Picture         =   "FContact.frx":3410
         End
         Begin TDBDate6Ctl.TDBDate mskSaleDate 
            Height          =   315
            Left            =   885
            TabIndex        =   153
            Top             =   465
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            Calendar        =   "FContact.frx":39AA
            Caption         =   "FContact.frx":3AC2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FContact.frx":3B2E
            Keys            =   "FContact.frx":3B4C
            Spin            =   "FContact.frx":3BAA
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
         Begin Threed.SSCommand cmdSalesDate 
            Height          =   315
            Left            =   1965
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   465
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            _Version        =   196610
            PictureFrames   =   1
            Picture         =   "FContact.frx":3BD2
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Sale Days:"
            Height          =   285
            Index           =   12
            Left            =   2400
            TabIndex        =   162
            Top             =   525
            Width           =   1035
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Grace Period:"
            Height          =   285
            Index           =   9
            Left            =   3660
            TabIndex        =   157
            Top             =   120
            Width           =   1485
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Pending Days:"
            Height          =   285
            Index           =   8
            Left            =   3960
            TabIndex        =   156
            Top             =   525
            Width           =   1035
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Sale Date:"
            Height          =   285
            Left            =   45
            TabIndex        =   155
            Top             =   495
            Width           =   975
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Available Online Auths:"
            Height          =   255
            Left            =   1560
            TabIndex        =   149
            Top             =   120
            Width           =   1695
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
            TabIndex        =   108
            Top             =   90
            Width           =   1815
         End
         Begin VB.Label lblExpires 
            BackStyle       =   0  'Transparent
            Caption         =   "Expires:"
            Height          =   285
            Left            =   4215
            TabIndex        =   107
            Top             =   945
            Width           =   1590
         End
         Begin VB.Label lblAuthRemaining 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            DataField       =   "AuthRemaining"
            Height          =   315
            Left            =   4920
            TabIndex        =   57
            Tag             =   "1"
            Top             =   1575
            Width           =   675
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Auth:"
            Height          =   285
            Index           =   2
            Left            =   45
            TabIndex        =   106
            Top             =   1605
            Width           =   1005
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ver.:"
            Height          =   285
            Index           =   6
            Left            =   4230
            TabIndex        =   105
            Top             =   1275
            Width           =   645
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Shipping:"
            Height          =   285
            Index           =   1
            Left            =   45
            TabIndex        =   104
            Top             =   1275
            Width           =   765
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Download:"
            Height          =   285
            Index           =   5
            Left            =   45
            TabIndex        =   103
            Top             =   930
            Width           =   765
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   135
      Left            =   12120
      TabIndex        =   98
      Top             =   5760
      Width           =   15
   End
   Begin MSComctlLib.ListView lvwHistory 
      Height          =   1695
      Left            =   4920
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   5760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilContact"
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Result"
         Object.Width           =   11465
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Frame frmBelowContactList 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   120
      TabIndex        =   93
      Top             =   5160
      Width           =   3135
      Begin VB.TextBox txtCompanyNote 
         BackColor       =   &H8000000B&
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   99
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdClipBoard 
         Caption         =   "Copy to Clipboard"
         Height          =   255
         Left            =   105
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   45
         Width           =   1695
      End
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "Print List"
         Height          =   255
         Left            =   1920
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   45
         Width           =   975
      End
   End
   Begin VB.Frame frmNavigation 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   0
      TabIndex        =   89
      Top             =   0
      Width           =   15255
      Begin Threed.SSCommand cmdRand3 
         Height          =   255
         Left            =   14640
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   196610
      End
      Begin Threed.SSCommand cmdRand2 
         Height          =   255
         Left            =   14400
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   196610
      End
      Begin VB.CommandButton cmdHistoryType 
         Caption         =   "Change Events List"
         Height          =   315
         Left            =   12435
         TabIndex        =   3
         Top             =   60
         Width           =   1575
      End
      Begin VB.ComboBox cboSearchType 
         Height          =   315
         Left            =   9780
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   60
         Width           =   1695
      End
      Begin VB.CommandButton cmdBack 
         Enabled         =   0   'False
         Height          =   315
         Left            =   45
         Picture         =   "FContact.frx":416C
         Style           =   1  'Graphical
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdForward 
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         Picture         =   "FContact.frx":42B6
         Style           =   1  'Graphical
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   60
         Width           =   495
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   6105
         TabIndex        =   0
         Top             =   60
         Width           =   3615
      End
      Begin Threed.SSCommand cmdRefresh 
         Height          =   315
         Left            =   11505
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "Refresh"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdRand1 
         Height          =   255
         Left            =   14160
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   196610
         PictureAlignment=   9
      End
      Begin VB.Label lblCompany 
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
         Left            =   1200
         TabIndex        =   97
         Top             =   90
         Width           =   4035
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
         Left            =   5280
         TabIndex        =   92
         Top             =   75
         Width           =   825
      End
   End
   Begin VB.Frame frmContactActions 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   120
      TabIndex        =   88
      Top             =   480
      Width           =   11535
      Begin Threed.SSCommand cmdNewContact 
         Height          =   315
         Left            =   780
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "New Contact"
         Top             =   0
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "New Contact"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdEditContact 
         Height          =   315
         Left            =   1845
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Edit Contact"
         Top             =   0
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "Edit Contact"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdSaveContact 
         Height          =   315
         Left            =   2910
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "Save Changes"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdCancelContact 
         Height          =   315
         Left            =   4170
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "Cancel Changes"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdDeleteContact 
         Height          =   315
         Left            =   5460
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "Delete Contact"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdPrintLabel 
         Height          =   315
         Left            =   6735
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Print Label"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "Print Label"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdSetAppt 
         Height          =   315
         Left            =   7665
         TabIndex        =   12
         ToolTipText     =   "Set App. "
         Top             =   0
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "Set Appointment"
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   315
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Action"
      End
      Begin Threed.SSCommand cmdNewCompany 
         Height          =   315
         Left            =   9015
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "New Company"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmdEditCompany 
         Height          =   315
         Left            =   10365
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   196610
         Caption         =   "Edit Company"
         PictureAlignment=   9
      End
   End
   Begin TabDlg.SSTab tbContacts 
      Height          =   4215
      Left            =   120
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Company"
      TabPicture(0)   =   "FContact.frx":4400
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstBranch"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwContact"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmBranchControls"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Groups"
      TabPicture(1)   =   "FContact.frx":441C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwPMContacts"
      Tab(1).Control(1)=   "frmGroupList"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Search"
      TabPicture(2)   =   "FContact.frx":4438
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwSearchContacts"
      Tab(2).Control(1)=   "frmSearch"
      Tab(2).ControlCount=   2
      Begin VB.PictureBox frmBranchControls 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   142
         Top             =   1200
         Width           =   1935
         Begin VB.CommandButton cmdAddBranch 
            Caption         =   "Add"
            Height          =   255
            Left            =   0
            TabIndex        =   145
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton cmdDeleteBranch 
            Caption         =   "Delete"
            Height          =   255
            Left            =   600
            TabIndex        =   144
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton cmdEditBranch 
            Caption         =   "Edit"
            Height          =   255
            Left            =   1320
            TabIndex        =   143
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Frame frmGroupList 
         BorderStyle     =   0  'None
         Caption         =   "frmGroupList"
         Height          =   2040
         Left            =   -74880
         TabIndex        =   81
         Top             =   500
         Width           =   3015
         Begin VB.CommandButton cmdClearCustGroup 
            Caption         =   "Clear"
            Height          =   255
            Left            =   1080
            TabIndex        =   198
            Top             =   1500
            Width           =   855
         End
         Begin VB.CommandButton cmdDelGroup 
            Caption         =   "Delete"
            Height          =   255
            Left            =   2115
            TabIndex        =   197
            Top             =   1500
            Width           =   855
         End
         Begin VB.CommandButton cmdAddCustGroup 
            Caption         =   "Add"
            Height          =   255
            Left            =   0
            TabIndex        =   196
            Top             =   1500
            Width           =   855
         End
         Begin VB.ListBox lstGroups 
            Height          =   1425
            Left            =   0
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   0
            Width           =   3015
         End
         Begin VB.CheckBox chkAlpha 
            Caption         =   "Sort Alphabetically"
            Height          =   255
            Left            =   0
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.Frame frmSearch 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2000
         Left            =   -74880
         TabIndex        =   59
         Top             =   500
         Width           =   3015
         Begin VB.CommandButton cmdSearch 
            Appearance      =   0  'Flat
            Caption         =   "Search"
            Height          =   300
            Left            =   2235
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1650
            Width           =   735
         End
         Begin ActiveScroll.SSScroll scrollSearchFields 
            Height          =   1215
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   2143
            _Version        =   196610
            BorderStyle     =   2
            HScrollType     =   0
            ScrollingHeight =   3425
            TagVariant      =   ""
            Begin VB.TextBox txtSearchNotes 
               Height          =   285
               Left            =   600
               MaxLength       =   100
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   2970
               Width           =   2055
            End
            Begin VB.TextBox txtSearchCity 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   1560
               Width           =   2055
            End
            Begin VB.TextBox txtSearchZip 
               Height          =   315
               Left            =   600
               MaxLength       =   25
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   2280
               Width           =   2055
            End
            Begin VB.TextBox txtSearchSource 
               Height          =   285
               Left            =   600
               MaxLength       =   100
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   2640
               Width           =   2070
            End
            Begin VB.TextBox txtSearchFirstName 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtSearchLastName 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   840
               Width           =   2055
            End
            Begin VB.TextBox txtSearchCompany 
               Height          =   315
               Left            =   600
               MaxLength       =   100
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   1200
               Width           =   2055
            End
            Begin VB.ComboBox cboSearchStatus 
               Height          =   315
               Left            =   600
               TabIndex        =   63
               TabStop         =   0   'False
               Text            =   "Customer"
               Top             =   120
               Width           =   2055
            End
            Begin VB.TextBox txtSearchState 
               Height          =   315
               Left            =   600
               MaxLength       =   5
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   1920
               Width           =   2055
            End
            Begin VB.Label Label31 
               Caption         =   "Notes:"
               Height          =   255
               Left            =   15
               TabIndex        =   79
               Top             =   3000
               Width           =   495
            End
            Begin VB.Label Label4 
               Caption         =   "City:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   78
               Top             =   1590
               Width           =   375
            End
            Begin VB.Label Label5 
               Caption         =   "State:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   77
               Top             =   1950
               Width           =   495
            End
            Begin VB.Label Label24 
               Caption         =   "Zip:"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   15
               TabIndex        =   76
               Top             =   2310
               Width           =   345
            End
            Begin VB.Label Label12 
               Caption         =   "Source:"
               Height          =   255
               Left            =   15
               TabIndex        =   75
               Top             =   2655
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "First:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   15
               TabIndex        =   74
               Top             =   510
               Width           =   825
            End
            Begin VB.Label Label8 
               Caption         =   "Last:"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   15
               TabIndex        =   73
               Top             =   870
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Comp.:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   15
               TabIndex        =   72
               Top             =   1230
               Width           =   525
            End
            Begin VB.Label Label7 
               Caption         =   "Status:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   15
               TabIndex        =   71
               Top             =   150
               Width           =   525
            End
         End
         Begin VB.Label lblResults 
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1710
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lvwContact 
         Height          =   2325
         Left            =   75
         TabIndex        =   84
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
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2640
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
         Left            =   -74880
         TabIndex        =   86
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
         Height          =   645
         Left            =   120
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1138
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
            Object.Width           =   4586
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilContact 
      Left            =   3720
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":4454
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":4D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":5608
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":5EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":67BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":7096
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":7970
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":7ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":7C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":7F3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdHistory 
      Bindings        =   "FContact.frx":8098
      Height          =   1725
      Left            =   3480
      TabIndex        =   100
      Top             =   6240
      Visible         =   0   'False
      Width           =   8460
      ScrollBars      =   2
      _Version        =   196617
      DataMode        =   2
      ColumnHeaders   =   0   'False
      Col.Count       =   9
      UseGroups       =   -1  'True
      AllowUpdate     =   0   'False
      AllowColumnShrinking=   0   'False
      ForeColorEven   =   0
      BackColorEven   =   -2147483633
      BackColorOdd    =   -2147483633
      Levels          =   2
      RowHeight       =   847
      Groups(0).Width =   13917
      Groups(0).Caption=   "Date                Time                   Type                         User                          Subject"
      Groups(0).CaptionAlignment=   0
      Groups(0).Columns.Count=   9
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
      Groups(0).Columns(2).Width=   1931
      Groups(0).Columns(2).Caption=   "Date"
      Groups(0).Columns(2).Name=   "Date"
      Groups(0).Columns(2).CaptionAlignment=   1
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   7
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(3).Width=   2117
      Groups(0).Columns(3).Caption=   "Time"
      Groups(0).Columns(3).Name=   "Time"
      Groups(0).Columns(3).CaptionAlignment=   1
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   7
      Groups(0).Columns(3).FieldLen=   256
      Groups(0).Columns(4).Width=   2619
      Groups(0).Columns(4).Caption=   "Type"
      Groups(0).Columns(4).Name=   "Type"
      Groups(0).Columns(4).CaptionAlignment=   0
      Groups(0).Columns(4).DataField=   "Column 4"
      Groups(0).Columns(4).DataType=   8
      Groups(0).Columns(4).FieldLen=   256
      Groups(0).Columns(5).Width=   2646
      Groups(0).Columns(5).Caption=   "User"
      Groups(0).Columns(5).Name=   "User"
      Groups(0).Columns(5).CaptionAlignment=   0
      Groups(0).Columns(5).DataField=   "Column 5"
      Groups(0).Columns(5).DataType=   8
      Groups(0).Columns(5).FieldLen=   256
      Groups(0).Columns(6).Width=   4604
      Groups(0).Columns(6).Caption=   "Subject"
      Groups(0).Columns(6).Name=   "Subject"
      Groups(0).Columns(6).CaptionAlignment=   0
      Groups(0).Columns(6).DataField=   "Column 6"
      Groups(0).Columns(6).DataType=   8
      Groups(0).Columns(6).FieldLen=   256
      Groups(0).Columns(7).Width=   13917
      Groups(0).Columns(7).Caption=   "Results"
      Groups(0).Columns(7).Name=   "Results"
      Groups(0).Columns(7).CaptionAlignment=   0
      Groups(0).Columns(7).DataField=   "Column 7"
      Groups(0).Columns(7).DataType=   8
      Groups(0).Columns(7).Level=   1
      Groups(0).Columns(7).FieldLen=   256
      Groups(0).Columns(7).HasHeadForeColor=   -1  'True
      Groups(0).Columns(7).HasForeColor=   -1  'True
      Groups(0).Columns(7).HasBackColor=   -1  'True
      Groups(0).Columns(7).HeadForeColor=   16776960
      Groups(0).Columns(7).ForeColor=   -2147483640
      Groups(0).Columns(7).BackColor=   -2147483643
      Groups(0).Columns(8).Width=   3572
      Groups(0).Columns(8).Visible=   0   'False
      Groups(0).Columns(8).Caption=   "ProductID"
      Groups(0).Columns(8).Name=   "ProductID"
      Groups(0).Columns(8).DataField=   "Column 8"
      Groups(0).Columns(8).DataType=   3
      Groups(0).Columns(8).Level=   1
      Groups(0).Columns(8).FieldLen=   256
      _ExtentX        =   14922
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   960
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":80B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FContact.frx":8120
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuContact 
      Caption         =   "Contact"
      Visible         =   0   'False
      Begin VB.Menu mnuChangeCompany 
         Caption         =   "Change Company"
      End
   End
End
Attribute VB_Name = "FContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormControl   As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public ContactStack As New CContactStack
'
Dim Branchs As CBranchs
'
Private CboEvents                As New CComboSearch
Private TextSearch              As New CTextSearch
'
Private rsGroupCategories   As ADODB.Recordset
'
Private rsSearchCompany         As ADODB.Recordset
Private rsSearchContact         As ADODB.Recordset
'
Private CompanyData As CCompanyData
Private ContactData As CContactData
'
Private Company As New CCompany
Private Contact As New CContact
'
Private bEnabled As Boolean
Private bEnteringNewContact As Boolean
'
Private bUseNewNoteList As Boolean

Private Function GetSortField(ByRef plvwList As ListView) As String
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If plvwList.ColumnHeaders.Count >= (plvwList.SortKey + 1) Then
110     GetSortField = plvwList.ColumnHeaders(plvwList.SortKey + 1).Text
120   Else
130     GetSortField = vbNullString
140   End If
      '<EhFooter>
      '
      Exit Function
      '
EH:
      ErrorMgr.Raise "FContact", "GetSortField", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Private Function GetSortDirection(ByRef plvwList As ListView) As Integer
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   With plvwList
110     If .SortOrder = lvwAscending Then
120       GetSortDirection = 0
130     Else
140       GetSortDirection = 1
150     End If
160   End With
      '<EhFooter>
      '
      Exit Function
      '
EH:
      ErrorMgr.Raise "FContact", "GetSortDirection", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Private Sub cboSearchType_Change()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   RefeshTextSearch
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cboSearchType_Change", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub RefeshTextSearch()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100 On Error GoTo ErrCall
        '
110     Set rsSearchCompany = New Recordset
        '
120     rsSearchCompany.CursorLocation = adUseClient
130     rsSearchCompany.Open "UpSelectCompanyList", cnMain, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        '
140     Set rsSearchContact = New Recordset
        '
150     rsSearchContact.CursorLocation = adUseClient
160     rsSearchContact.Open "UpSelContactList", cnMain, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        '
170     Select Case cboSearchType.Text
          Case "Company"
180         TextSearch.Setup txtSearch, 1, rsSearchCompany
190       Case "Last Name"
200         TextSearch.Setup txtSearch, 4, rsSearchContact
            '
210         rsSearchContact.Sort = "LastFirst"
220       Case "First Name"
230         TextSearch.Setup txtSearch, 3, rsSearchContact
            '
240         rsSearchContact.Sort = "FirstLast"
250       Case "Phone 1"
260         TextSearch.Setup txtSearch, 5, rsSearchContact
            '
270         rsSearchContact.Sort = "Phone1"
280       Case "Phone 2"
290         TextSearch.Setup txtSearch, 8, rsSearchContact
            '
300         rsSearchContact.Sort = "Phone2"
310       Case "Email"
320         TextSearch.Setup txtSearch, 6, rsSearchContact
            '
330         rsSearchContact.Sort = "Email"
340     End Select
        '
350     SaveSetting App.Title, "Misc", "SearchType", cboSearchType.Text
        '
360     txtSearch.Text = vbNullString
370     If txtSearch.Enabled And txtSearch.Visible Then
380       txtSearch.SetFocus
390     End If
        '
400   Exit Sub
ErrCall:
410   MsgBox Err.Description & " in Search by Click."
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "RefeshTextSearch", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cboSearchType_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   RefeshTextSearch
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cboSearchType_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cboSource_InitColumnProps()
  On Error GoTo EH
  '
  cboSource.AddItem "Download"
  cboSource.AddItem "Web Request"
  cboSource.AddItem "Word of Mouth"
  '
  Exit Sub
EH:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FCustomer.cboSource_InitColumnProps.", vbCritical, "Error"
End Sub

Private Sub chkAlpha_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If chkAlpha.Value = vbChecked Then
110     SaveSetting App.Title, "ProspectMgt", "SortColumn", 1
120   Else
130     SaveSetting App.Title, "ProspectMgt", "SortColumn", 0
140   End If
      '
150   SetupPMGroups
160   LoadCustGroupsList
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "chkAlpha_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub chkSelected_Click()
  'this was used to print lables but it was changed when custom groups were added.
  '<EhHeader>
  On Error GoTo EH
  '
  '</EhHeader>
  '<EhFooter>
  '
  Exit Sub
  '
EH:
  ErrorMgr.Raise "FContact", "chkSelected_Click", Err.Number, Err.Description, Erl
  '</EhFooter>
End Sub

Private Sub cmdAction_Click()
     ' If FNote.NewNote(ContactData.ID) Then
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100  If FNote.NewNote(ContactData.ID) Then
105   FillCompanyContactList GetIDFromKey(lstBranch.SelectedItem.Key)
110   LoadContact ContactData.ID, False
120  End If
     ' Else
       ' LoadHistory ContactData.ID
      'End If
      '
  
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdAction_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdAddBranch_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim lResultBranchID As Long
      '
110   If CompanyData.ID > 0 Then
120     lResultBranchID = FBranch.NewBranch(CompanyData.ID)
        '
130     If lResultBranchID <> 0 Then
140       If ContactData.ID <> 0 Then
150         LoadContact ContactData.ID, True
160       Else
170         LoadOnlyCompany CompanyData.ID, True
180       End If
190     End If
200   Else
210     MsgBox "No Company Loaded"
220   End If
  
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdAddBranch_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdAddCustGroup_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim GroupList As CGroupList
110   Dim GroupListData As CGroupListData
120   Dim Employee As New CEmployee
130   Dim sTemp As String
      '
140   Set GroupList = New CGroupList
150   Set GroupListData = New CGroupListData
      '
160   sTemp = Left(InputBox("Enter the name of the new group.", "Add Custom Group"), 25)
      '
170   If Trim(sTemp) <> "" Then
180     GroupListData.ListName = sTemp
190     GroupListData.EmployeeID = Employee.GetEmployeeID(User.Name)
        '
200     GroupList.Save GroupListData, True
        '
210     LoadCustGroups
220     SetupPMGroups
230     LoadCustGroupsList
240   End If
      '
250   Set GroupList = Nothing
260   Set GroupListData = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdAddCustGroup_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdClearCustGroup_Click()
  If MsgBox("Remove all contacts from custom group?", vbYesNo, "Clear Custom Group") = vbYes Then
    Dim GroupList As New CGroupList
    Dim GroupListData As New CGroupListData
    '
    GroupList.Load GroupListData, lstGroups.ItemData(lstGroups.ListIndex) 'GetIDFromKey(lstCustGroups.SelectedItem.Key)
    '
    GroupList.Clear lstGroups.ItemData(lstGroups.ListIndex)   'GetIDFromKey(lstCustGroups.SelectedItem.Key)
    '
    LoadCustGroups
    SetupPMGroups
    LoadCustGroupsList
  
    '
    Set GroupList = Nothing
    Set GroupListData = Nothing
  End If
End Sub

Private Sub cmdDelGroup_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim GroupList As CGroupList
110   Dim GroupListData As CGroupListData
      '
120   Dim Employee As New CEmployee
      '
130   Set GroupList = New CGroupList
140   Set GroupListData = New CGroupListData
      '
150   GroupList.Load GroupListData, lstGroups.ItemData(lstGroups.ListIndex) 'GetIDFromKey(lstCustGroups.SelectedItem.Key)
      '
160   If GroupListData.EmployeeID = Employee.GetEmployeeID(User.Name) Or _
          Employee.InGroup(User.Name, "Management") = True Or _
          Employee.InGroup(User.Name, "Development") = True Then
        '
170     GroupList.Delete lstGroups.ItemData(lstGroups.ListIndex)   'GetIDFromKey(lstCustGroups.SelectedItem.Key)
        '
180     LoadCustGroups
190     SetupPMGroups
200     LoadCustGroupsList
210   Else
220     MsgBox "Removing someone elses group can be hazardous to your health!", vbCritical, "Ethical Violation!!"
230   End If
      '
240   Set GroupList = Nothing
250   Set GroupListData = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdDelGroup_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdAuthDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskAuthDate.Value = FDatePick.DateText(mskAuthDate.Value)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdAuthDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdCancelContact_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If bEnabled = True Then
110     If MsgBox("Are you sure you want to cancel and lose all changes to this contact?", vbQuestion + vbYesNo, "Cancel Changes") = vbYes Then
          'bEnteringNewContact = False
          '
120       If ContactData.ID = 0 Then
130         LoadOnlyCompany CompanyData.ID, True
140       Else
150         LoadContact ContactData.ID, False
160       End If
'170       LoadCustGroups
180     End If
190   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdCancelContact_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdDeleteBranch_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim Branch As New CBranch
      '
110   If GetIDFromKey(lstBranch.SelectedItem.Key) <> 0 Then
120     If (MsgBox("HEY! Are you sure you want to delete this branch?", vbYesNo) = vbYes) Then
130       If Not Branch.Delete(GetIDFromKey(lstBranch.SelectedItem.Key)) Then
140         MsgBox "Cannot delete branch. Make sure no contacts are assigned to it before deleting."
150       Else
160         If ContactData.ID <> 0 Then
170           LoadContact ContactData.ID, True
180         Else
190           LoadOnlyCompany CompanyData.ID, True
200         End If
210       End If
220     Else
230       Exit Sub
240     End If
250   Else
260     MsgBox "No branch is selected"
270   End If
      '
280   Set Branch = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdDeleteBranch_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdDeleteCompany_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If CompanyData.ID <> 0 Then
110     Company.Delete CompanyData.ID
        '
120     LoadOnlyCompany 0, False
130   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdDeleteCompany_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdDeleteContact_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If ContactData.ID <> 0 Then
110     If Contact.Delete(ContactData.ID) Then
120       LoadOnlyCompany CompanyData.ID, True
130     End If
140   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdDeleteContact_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdDownloadDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskDownloadDate.Value = FDatePick.DateText(mskDownloadDate.Value)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdDownloadDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdEditBranch_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim lResultBranchID As Long
      '
110   If GetIDFromKey(lstBranch.SelectedItem.Key) <> 0 Then
120     lResultBranchID = FBranch.EditBranch(GetIDFromKey(lstBranch.SelectedItem.Key))
        '
130     If lResultBranchID <> 0 Then
140       If ContactData.ID <> 0 Then
150         LoadContact ContactData.ID, True
160       Else
170         LoadOnlyCompany CompanyData.ID, True
180       End If
190     End If
200   Else
210     MsgBox "No Branch Selected Loaded"
220   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdEditBranch_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdEditCompany_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If CompanyData.ID <> 0 Then
110     FCompany2.LoadCompany (CompanyData.ID)
        '
120     If ContactData.ID <> 0 Then
130       LoadContact ContactData.ID, True
140     Else
150       LoadOnlyCompany CompanyData.ID, True
160     End If
170   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdEditCompany_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdEditContact_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   EnableEditContactControls True
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdEditContact_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdHistoryType_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   bUseNewNoteList = Not bUseNewNoteList
110   lvwHistory.Visible = Not lvwHistory.Visible
120   grdHistory.Visible = Not grdHistory.Visible
      '
130   If Not ContactData Is Nothing Then
140     If ContactData.ID <> 0 Then
150       LoadHistory ContactData.ID
160     End If
170   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdHistoryType_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdNewCompany_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100     Dim lResultCompanyID As Long
        '
110     lResultCompanyID = FCompany2.NewCompany
        '
120     If lResultCompanyID <> 0 Then
          'RefeshTextSearch
130       LoadOnlyCompany lResultCompanyID, True
140       txtSearch.Text = vbNullString
150     End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdNewCompany_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdNewContact_Click()
  Set ContactData = New CContactData
  '
  LoadOnlyCompany CompanyData.ID, False
  '
  ContactData.CompanyID = CompanyData.ID
  '
  bEnteringNewContact = True
  '
  EnableEditContactControls True
  '
  cboStatus.Text = "Prospect"
  '
  cboDownloadStatus.Text = "None"
  '
  cboShipStatus.Text = "Not Shipped"
  '
  cboAuthStatus.Text = "Not Authorized"
  '
  cboPVDownloadStatus.Text = "None"
  '
  cboPVShipStatus.Text = "Not Shipped"
  '
  cboPVAuthStatus.Text = "Not Authorized"
End Sub

Private Sub cmdPrintLabel_Click()
      'FSingleLabel.Show vbModal, FMain
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   FSingleLabel.PrintSingle ContactData.ID
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdPrintLabel_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdPVAuthDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskPVAuthDate.Value = FDatePick.DateText(mskPVAuthDate.Value)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdPVAuthDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdPVDownloadDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskPVDownloadDate.Value = FDatePick.DateText(mskPVDownloadDate.Value)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdPVDownloadDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdPVSaleDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskPVSaleDate.Value = FDatePick.DateText(mskPVSaleDate.Value)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdPVSaleDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdPVSalesDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskPVSaleDate.Value = FDatePick.DateText(mskPVSaleDate.Value)
110   Me.txtPVPendingDays = CalculatePendingDays(mskPVSaleDate.Value, nnNum(txtPVGraceDays.Text), nnNum(txtPVSaleDays.Text))
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdPVSalesDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdPVShipDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskPVShipDate.Value = FDatePick.DateText(mskPVShipDate.Value)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdPVShipDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

'Private Sub cmdRand_Click()
'  Randomize
'  SaveSetting App.Title, "Settings", "LightColor", RGB(Rnd * 255, Rnd * 255, Rnd * 255)
'  SaveSetting App.Title, "Settings", "DarkColor", RGB(Rnd * 255, Rnd * 255, Rnd * 255)
'  SaveSetting App.Title, "Settings", "ListColor", RGB(Rnd * 255, Rnd * 255, Rnd * 255)
'  RefreshColors
'End Sub

Private Sub cmdRand1_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Randomize
110   SaveSetting App.Title, "Settings", "LightColor", RGB(Rnd * 255, Rnd * 255, Rnd * 255)
120   RefreshColors
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdRand1_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdRand2_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Randomize
110   SaveSetting App.Title, "Settings", "DarkColor", RGB(Rnd * 255, Rnd * 255, Rnd * 255)
120   RefreshColors
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdRand2_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdRand3_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Randomize
110   SaveSetting App.Title, "Settings", "ListColor", RGB(Rnd * 255, Rnd * 255, Rnd * 255)
120   RefreshColors
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdRand3_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdSaleDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskSaleDate.Value = FDatePick.DateText(mskSaleDate.Value)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdSaleDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdSalesDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskSaleDate.Value = FDatePick.DateText(mskSaleDate.Value)
110   Me.txtPendingDays = CalculatePendingDays(mskSaleDate.Value, nnNum(txtGraceDays.Text), nnNum(txtSaleDays.Text))
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdSalesDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdSaveContact_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SaveContact
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdSaveContact_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdSearch_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SideSearch
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdSearch_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdSetAppt_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   FOutlookAppt.ContactName = txtFirstName & " " & txtLastName
110   FOutlookAppt.Show vbModal
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdSetAppt_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdShipDate_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   mskShipDate.Value = FDatePick.DateText(mskShipDate.Value)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdShipDate_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdAddGroup_Click()
      'Dim AllContacts As New CContacts
      '
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim rsContacts As New Recordset
110   Dim EventData As CEventData
120   Dim Event1 As New CEvent
      '
130   rsContacts.Open "SELECT ID, Notes FROM TContact WHERE NOT Notes LIKE ''", cnMain, adOpenKeyset, adLockOptimistic
      '*
140   While Not rsContacts.eof
150     If rsContacts!Notes & vbNullString <> vbNullString Then
160       Set EventData = New CEventData
170       EventData.CustRecID = rsContacts!ID
180       EventData.EventDate = Format(Now, "Short Date")
190       EventData.EventTime = Format(Now, "hh:nn AM/PM")
200       EventData.EventResults = rsContacts!Notes
210       EventData.EventSubject = "NOTE:"
220       EventData.EventType = "Note"
230       EventData.EventUser = "System"
240       EventData.OpenCall = False
250       EventData.ProductID = 1
260       EventData.Sticky = True
270       Event1.Save EventData, True
280     End If
290     rsContacts.MoveNext
300   Wend
      '
310   rsContacts.Close
      '
320   Set rsContacts = Nothing
330   Set EventData = Nothing
340   Set Event1 = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdAddGroup_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

'Private Sub cmdAddCustGroup_Click()
'  Dim AllCompanies As New CCompanys
'  Dim Contacts As New CContacts
'  '
'  Company.LoadCollection AllCompanies
'  '
'  Dim iPos As Integer
'  '
'  While iPos < AllCompanies.Count
'    iPos = iPos + 1
'    '
'    Set Contacts = New CContacts
'    '
'    Contact.LoadCollection AllCompanies.Item(iPos).ID, Contacts, 0
'    '
'    If Contacts.Count = 0 Then
'      Company.Delete AllCompanies.Item(iPos).ID
'    End If
'  Wend
'  '
'  Set AllCompanies = Nothing
'  Set Contacts = Nothing
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
  Set Branchs = Nothing
End Sub

Private Sub FormControl_SwitchFrom(bCancel As Boolean)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If bEnabled Then
110     If MsgBox("Save Changes?", vbYesNo) = vbYes Then
120       SaveContact
130     Else
140       LoadContact ContactData.ID, False
150     End If
160   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "FormControl_SwitchFrom", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub lstBranch_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   LoadOnlyCompany CompanyData.ID, True
    '  Dim Contacts As New CContacts
    '  '
    '  If Not lstBranch.SelectedItem Is Nothing Then
    '    Contact.LoadCollection CompanyData.ID, Contacts, GetIDFromKey(lstBranch.SelectedItem.Key)
    '
    '    'FillContactList lvwContact, Contacts
    '    Set ContactData = New CContactData
    '  '
    '    FillCompanyControls
    '    FillContactControls
    '    'FillControls
    '    '
    '    If CompanyData.ID <> 0 Then
    '      SetControlsCompanyOnlyLoaded
    '    Else
    '      SetControlsNothingLoaded
    '    End If
    '      'FillContactList GetIDFromKey(lstBranch.SelectedItem.Key)
    '    End If
    '  '
    '  Set Contacts = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "lstBranch_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub lvwContact_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo EH
  If Not lvwContact.SelectedItem Is Nothing Then
    LoadContact CLng(Mid$(lvwContact.SelectedItem.Key, 3)), False
  End If
  '
 ' LoadCustGroups
  Exit Sub
EH:
  MsgBox "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & " in FContact: lstContacts_KeyUp."
End Sub

Private Sub lvwContact_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo EH
  If Not lvwContact.SelectedItem Is Nothing Then
    LoadContact CLng(Mid$(lvwContact.SelectedItem.Key, 3)), False
  End If
  '
  If Button = vbRightButton Then
    If Not lvwContact.SelectedItem Is Nothing Then
      PopupMenu Me.mnuContact
    End If
  End If
  '
  'LoadCustGroups
  Exit Sub
EH:
  MsgBox "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & " in FContact: lstContacts_MouseUp."
End Sub

Private Sub lvwHistory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   With lvwHistory
      '
110   .SortKey = ColumnHeader.Index - 1
      '
120   sOrder = Not (sOrder)
      '
130   Select Case ColumnHeader.Index - 1 'piColumIndex - 1
       Case 3, 4:
          'Use sort routine to sort by date
140       .Sorted = False
150       SendMessage .hWnd, _
                      LVM_SORTITEMS, _
                      .hWnd, _
                      ByVal FARPROC(AddressOf CompareDates)
    '      Case 2:
    '        'Use sort routine to sort by value
    '        pListView.Sorted = False
    '        SendMessage pListView.hWnd, _
    '                   LVM_SORTITEMS, _
    '                   pListView.hWnd, _
    '                   ByVal FARPROC(AddressOf CompareValues)
          Case Else:
            'Use default sorting to sort the items in the list
            'lvwContact.SortKey = 0
160         .SortOrder = Abs(Not .SortOrder = 1)
170         .Sorted = True
180    End Select
       '
190    End With
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "lvwHistory_ColumnClick", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub lvwHistory_DblClick()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If Not lvwHistory.SelectedItem Is Nothing Then
110     If FNote.LoadNote(ContactData.ID, GetIDFromKey(lvwHistory.SelectedItem.Key)) Then
120       LoadContact ContactData.ID, False
          '
130     End If
140   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "lvwHistory_DblClick", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub lvwPMContacts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      '
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SortListView lvwPMContacts, ColumnHeader.Index
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "lvwPMContacts_ColumnClick", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub lvwPMContacts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error GoTo EH
  If Not lvwPMContacts.SelectedItem Is Nothing Then
    LoadContact CLng(Mid$(lvwPMContacts.SelectedItem.Key, 3)), False
  End If
  '
  'LoadCustGroups
  Exit Sub
EH:
  MsgBox "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & " in FContact: lstContacts_KeyUp."
End Sub

Private Sub lvwSearchContacts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SortListView lvwSearchContacts, ColumnHeader.Index
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "lvwSearchContacts_ColumnClick", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub lvwSearchContacts_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error GoTo EH
  If Not lvwSearchContacts.SelectedItem Is Nothing Then
    LoadContact CLng(Mid$(lvwSearchContacts.SelectedItem.Key, 3)), False
  End If
  '
 ' LoadCustGroups
  Exit Sub
EH:
  MsgBox "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & " in FContact: lvwSearchContacts_KeyUp."
End Sub

Private Sub lvwSearchContacts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error GoTo EH
  If Not lvwSearchContacts.SelectedItem Is Nothing Then
    LoadContact CLng(Mid$(lvwSearchContacts.SelectedItem.Key, 3)), False
  End If
  '
 ' LoadCustGroups
  Exit Sub
EH:
  MsgBox "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & " in FContact: lvwSearchContacts_KeyUp."
End Sub

Private Sub mnuChangeCompany_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim lSelectedContactID As Long
      '
110   lSelectedContactID = CLng(Mid$(lvwContact.SelectedItem.Key, 3))
      '
120   If Not lvwContact.SelectedItem Is Nothing Then
130     If FChangeCompany.ChangeCompany(lSelectedContactID) Then
140       LoadContact lSelectedContactID, True
150     End If
160   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "mnuChangeCompany_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtAuthDays_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100  RefreshAuthInfo
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtAuthDays_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtAuths_GotFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   InputNumber.Setup txtAuths, NumberTypeInteger
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtAuths_GotFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtPhone1_GotFocus()
  With txtPhone1
    .SelStart = 0
    .SelLength = Len(txtPhone1.Text)
  End With
End Sub

Private Sub txtPhone1_Validate(Cancel As Boolean)
  txtPhone1.Text = FormatPhoneNumber(txtPhone1.Text)
End Sub

Private Sub txtPhone2_GotFocus()
  With txtPhone2
    .SelStart = 0
    .SelLength = Len(txtPhone2.Text)
  End With
End Sub

Private Sub txtPhone2_Validate(Cancel As Boolean)
    txtPhone2.Text = FormatPhoneNumber(txtPhone2.Text)
End Sub

Private Sub txtFax_GotFocus()
  With txtFax
    .SelStart = 0
    .SelLength = Len(txtFax.Text)
  End With
End Sub

Private Sub txtFax_Validate(Cancel As Boolean)
    txtFax.Text = FormatPhoneNumber(txtFax.Text)
End Sub

Private Sub txtPVAuths_GotFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   InputNumber.Setup txtPVAuths, NumberTypeInteger
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtPVAuths_GotFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
'
Private Sub mskPVSaleDate_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Me.txtPVPendingDays = CalculatePendingDays(nnNum(mskPVSaleDate.Value), nnNum(txtPVGraceDays.Text), nnNum(txtPVSaleDays.Text))
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "mskPVSaleDate_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
'
Private Sub mskSaleDate_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Me.txtPendingDays = CalculatePendingDays(mskSaleDate.Value, nnNum(txtGraceDays.Text), nnNum(txtSaleDays.Text))
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "mskSaleDate_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtGraceDays_LostFocus()
  If Not IsNull(mskSaleDate.Value) Then
  '  MsgBox "Please enter a sale date"
 ' Else
    Me.txtPendingDays = CalculatePendingDays(mskSaleDate.Value, nnNum(txtGraceDays.Text), nnNum(txtSaleDays.Text))
  End If
End Sub

Private Sub txtPVGraceDays_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Me.txtPVPendingDays = CalculatePendingDays(nnNum(mskPVSaleDate.Value), nnNum(txtPVGraceDays.Text), nnNum(txtPVSaleDays.Text))
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtPVGraceDays_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
'
Private Sub txtPVSaleDays_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Me.txtPVPendingDays = CalculatePendingDays(mskPVSaleDate.Value, nnNum(txtPVGraceDays.Text), nnNum(txtPVSaleDays.Text))
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtPVSaleDays_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
'
Private Sub txtSaleDays_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If Not IsNull(mskSaleDate.Value) Then
110     'MsgBox "Please enter a sale date"
120   'Else
130     Me.txtPendingDays = CalculatePendingDays(mskSaleDate.Value, nnNum(txtGraceDays.Text), nnNum(txtSaleDays.Text))
140   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtSaleDays_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtSaleDays_GotFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   InputNumber.Setup txtSaleDays, NumberTypeInteger
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtSaleDays_GotFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtPVAuthDays_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   RefreshAuthInfo
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtPVAuthDays_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtGraceDays_GotFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   InputNumber.Setup txtGraceDays, NumberTypeInteger
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtGraceDays_GotFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtPVGraceDays_GotFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   InputNumber.Setup txtPVGraceDays, NumberTypeInteger
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtPVGraceDays_GotFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtPVSaleDays_GotFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   InputNumber.Setup txtPVSaleDays, NumberTypeInteger
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtPVSaleDays_GotFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtSearch_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   txtSearch.SelStart = 0
110   txtSearch.SelLength = Len(txtSearch.Text)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtSearch_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

'Private Sub lvwHistory_DblClick()
'  If Not lvwHistory.SelectedItem Is Nothing Then
'    FNote.LoadNote ContactData.ID, GetIDFromKey(lvwHistory.SelectedItem.Key)
'  End If
'End Sub

Private Sub txtSearch_GotFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   txtSearch.SelStart = 0
110   txtSearch.SelLength = Len(txtSearch.Text)
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "txtSearch_GotFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
  On Error GoTo EH
  '
  If KeyAscii = vbKeyReturn Then
    LoadSearchResults
    '
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch.Text)
    '
  End If
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in Text Search KeyPress."
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error GoTo EH
  '
  If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
    LoadSearchResults
  End If
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in Form Contact: Text Search: Key Up."
End Sub

Private Sub mskAuthDate_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   RefreshAuthInfo
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "mskAuthDate_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

'
Public Sub RefreshAuthInfo()
On Error GoTo EH
  '
  If IsDate(mskPVAuthDate) Then
    lblPVAuthRemaining.Caption = DateDiff("d", Now, DateAdd("d", CDbl(txtPVAuthDays.Text), mskPVAuthDate))
  Else
    lblPVAuthRemaining.Caption = 0
  End If
  '
  If CLng(lblPVAuthRemaining.Caption) > 0 Then
    lblPVExpires.Caption = "Expires: " & (Date + CLng(lblPVAuthRemaining.Caption))
  Else
    lblPVExpires.Caption = "Expires: NA"
  End If
  '
  'lblExpires.Caption = "Expires: " & (Date + lng(lblAuthRemaining.Caption))
  '
  Exit Sub
EH:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in RefeshAuthInfo.", vbCritical, "Error"
End Sub

'
Private Sub mskPVAuthDate_LostFocus()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   RefreshAuthInfo
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "mskPVAuthDate_LostFocus", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub mskFAX_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeySpace Then KeyAscii = 0
End Sub

Private Sub mskPhone1_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeySpace Then KeyAscii = 0
End Sub

Private Sub mskPhone2_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeySpace Then KeyAscii = 0
End Sub


Private Sub cmdPrintList_Click()
  Dim lstCustGroups As ListView
  'Private WithEvents objLvPrint As clsPrintLV
  Dim objLvPrint As clsPrintLV
  '
  Dim iCopies As Integer
  On Error GoTo EH
  '
  Select Case tbContacts.Tab
  Case 0
    Set lstCustGroups = lvwContact
  Case 1
    Set lstCustGroups = lvwPMContacts
  Case 2
    Set lstCustGroups = lvwSearchContacts
  End Select
  '
  FPrinterSelect.Show vbModal
  '
  If FPrinterSelect.bPrintCancel = True Then Exit Sub
  '
  'Report.FillList Report.rsReport, lstCustGroups
  '

  Set objLvPrint = New clsPrintLV
  '
  For iCopies = 1 To iNumofCopies 'iNumCopies is some global variable.

    objLvPrint.PrintListView lstCustGroups, 0.1, 8, "ListView Report", Landscape, True, False
  Next iCopies
  '
  'Destroy the object
  Set objLvPrint = Nothing
  '
  Set lstCustGroups = Nothing
  Exit Sub
EH:
 MsgBox Err.Description & " in FContact.cmdPrintList_Click."
End Sub

Private Sub cmdClipBoard_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim lstCustGroups As ListView
110   Dim sText As String
120   Dim lItemCount As Long
130   Dim lSubItemCount As Long
140   Dim lColumnCount As Long
      '
150   Clipboard.Clear
      '
160   Screen.MousePointer = vbHourglass
      '
170   Select Case tbContacts.Tab
      Case 0
180     Set lstCustGroups = lvwContact
190   Case 1
200     Set lstCustGroups = lvwPMContacts
210   Case 2
220     Set lstCustGroups = lvwSearchContacts
230   End Select
      '
240   For lColumnCount = 1 To lstCustGroups.ColumnHeaders.Count
250     sText = sText & lstCustGroups.ColumnHeaders(lColumnCount).Text & vbTab
260   Next
      '
270   sText = sText & vbCrLf
      '
280   For lItemCount = 1 To lstCustGroups.ListItems.Count
290     sText = sText & lstCustGroups.ListItems(lItemCount).Text & vbTab
        '
300     For lSubItemCount = 1 To lstCustGroups.ListItems(lItemCount).ListSubItems.Count
310       sText = sText & lstCustGroups.ListItems(lItemCount).ListSubItems(lSubItemCount) & vbTab
320     Next
        '
330     sText = sText & vbCrLf
340   Next
      '
      'MsgBox sText
350   Clipboard.SetText sText, vbCFText
      '
360   Screen.MousePointer = vbDefault
      '
370   Set lstCustGroups = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdClipBoard_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdRefresh_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   RefeshTextSearch
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdRefresh_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub Form_Load()
  On Error GoTo EH
  '
  'Set CboEvents = New CComboSearch
  Set TextSearch = New CTextSearch
  '
  Set ContactStack = New CContactStack
  '
  RefreshColors
  '
  Set FormControl = New CFormControl
  '
  FormControl.MinHeight = 5475
  FormControl.MinWidth = 10590
  FormControl.DataForm = True
  '
  grdHistory.StyleSets.Add "PV"
  grdHistory.StyleSets.Add "XML"
  '
  grdHistory.StyleSets("PV").BackColor = Product.GetColor(2)
  grdHistory.StyleSets("XML").BackColor = Product.GetColor(1)
  '
  If GetSetting(App.Title, "ProspectMgt", "SortColumn", 0) = 0 Then
    chkAlpha.Value = vbUnchecked
  Else
    chkAlpha.Value = vbChecked
  End If
  '
  FillTypesBox
  FillStatusBox
  FillSearchTypes
  FillDownloadStatusBox
  FillAuthStatusBox
  FillShipStatusBox
  '
  SetupPMGroups
  LoadCustGroupsList
  '
  tbContacts.Visible = True
  '
  SetControlsNothingLoaded
  '
  Me.lblCompany.Caption = "Company:"
  '
  Set CompanyData = New CCompanyData
  Set ContactData = New CContactData
  '
  bUseNewNoteList = True
  '
  Exit Sub
EH:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FCustomer.Form_Load.", vbCritical, "Error"

End Sub

Private Sub Form_Activate()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Me.Show
110   'SetStar 0
120   DoEvents
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "Form_Activate", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
  On Error Resume Next
  On Error Resume Next
  '
  Dim iCurrentTab As Integer
  Dim lBigHeight As Integer
  '
  Me.Enabled = False
  '
  '
  picCanvas.Visible = False
  'tbContacts.Visible = False
  '
  iCurrentTab = tbContacts.Tab
  '
  picCanvas.Top = 0
  picCanvas.Left = 0
  '
  'scrollFrame.ScrollingHeight = picCanvas.Height
  'scrollFrame.VPosition = 0
  '
  'scrollFrame.Move 3300, 850, 8205 + 255, Me.Height - 850 '3000
  '
  picCanvas.Left = 3300
  picCanvas.Top = 850
  '
  grdHistory.Move 3300, picCanvas.Top + picCanvas.Height, Me.Width - 3300, Me.Height - (picCanvas.Top + picCanvas.Height)
  lvwHistory.Move 3300, picCanvas.Top + picCanvas.Height, Me.Width - 3300, Me.Height - (picCanvas.Top + picCanvas.Height)
  '
  lvwHistory.ColumnHeaders(1).Width = lvwHistory.Width - 5500
  '
  tbContacts.Tab = 0
  '
  tbContacts.Move 15, 850, 3200, Me.Height - frmBelowContactList.Height - 850 '4300
  '
  lBigHeight = tbContacts.Height - (tbContacts.Height * 0.11)
  '
  lstBranch.Move 60, 325, tbContacts.Width - 120, 1025 + 100 '500
  '
  frmBranchControls.Left = 100
  frmBranchControls.Top = lstBranch.Top + lstBranch.Height + 50
  '
  lvwContact.Move 60, 1800, tbContacts.Width - 120, lBigHeight - 1900 + 325 '1900  '- 2600  ', 1325
  '
  tbContacts.Tab = 1
  '
  frmGroupList.Move 60, 325, tbContacts.Width - 120, 2100 '1025 + 500
  '
  lstGroups.Move 0, 0, frmGroupList.Width, 1425
  '
  'lstCustGroupSelect.Move 0, 1250, frmGroupList.Width, 510
  '
  lvwPMContacts.Move 60, 2400, tbContacts.Width - 120, lBigHeight - 2500 + 325 '1900  '- 2600  ', 1325
  '
  tbContacts.Tab = 2
  '
  frmSearch.Move 60, 325, tbContacts.Width - 120, 2000
  '
  scrollSearchFields.Move 0, 0, frmGroupList.Width, 1600
  '
  lvwSearchContacts.Move 60, 2325, tbContacts.Width - 120, lBigHeight - 2000  ' - 2325  '- 3000  ', 1325
  '
  tbContacts.Tab = iCurrentTab
  tbContacts.Move 15, 850, 3200, lBigHeight + 400
  '
  frmBelowContactList.Left = 0
  frmBelowContactList.Top = tbContacts.Top + tbContacts.Height + 50
  frmBelowContactList.Width = tbContacts.Width
  '
  'tbContacts.Visible = True
  picCanvas.Visible = True
  '
  Enabled = True
End Sub

'Private Sub Form_Resize2()
'  On Error Resume Next
'  On Error Resume Next
'  '
'  '
'  Me.Enabled = False
'  '
''  If Me.Width <> 11800 Then
''    Me.Width = 11800
''  End If
''  '
''  If Me.Height <> FMain.Height - 1000 Then
''    Me.Height = FMain.Height - 1500 '7300
''  End If
'  '
'  tbContacts.Visible = False
'  '
'  iCurrentTab = tbContacts.Tab
'  '
'
'  '
'  picCanvas.Top = 0
'  picCanvas.Left = 0
'  '
'  scrollFrame.ScrollingHeight = picCanvas.Height
'  '
'  scrollFrame.Move 3300, 850, 8205 + 255, 3000 ' Me.Height - 900 - 3105 + 1000 ' 2700 + 375
'  '
'  tbContacts.Tab = 0
'  'lBigHeight = Me.Height - 3000 - (Me.Height * 0.11)
'
'  tbContacts.Move 15, 850, 3200, 4300 'lBigHeight + 400 ',1500, 3200, Me.Height - 2000
'  'lBigHeight = Me.Height - 3000 - (Me.Height * 0.11)
'  lBigHeight = tbContacts.Height - (tbContacts.Height * 0.11)
'  lstBranch.Move 60, 325, tbContacts.Width - 120, 1025 + 500
'  lvwContact.Move 60, 1900, tbContacts.Width - 120, lBigHeight - 1900 + 325 '1900  '- 2600  ', 1325
'  '
'  tbContacts.Tab = 1
'  frmGroupList.Move 60, 325, tbContacts.Width - 120, 1025 + 500
'  lstGroups.Move 0, 0, frmGroupList.Width, 1350
'  lvwPMContacts.Move 60, 1900, tbContacts.Width - 120, lBigHeight - 1900 + 325 '1900  '- 2600  ', 1325
'  '
'  tbContacts.Tab = 2
'  '
'  frmSearch.Move 60, 325, tbContacts.Width - 120, 2000
'  '
'  scrollSearchFields.Move 0, 0, frmGroupList.Width, 1600
'  '
'  lvwSearchContacts.Move 60, 2325, tbContacts.Width - 120, lBigHeight - 2000  ' - 2325  '- 3000  ', 1325
'  '
'  tbContacts.Tab = iCurrentTab
'  tbContacts.Move 15, 850, 3200, lBigHeight + 400 ',1500, 3200, Me.Height - 2000
'  '
'
'  '
'  frmBelowContactList.Left = 0
'  frmBelowContactList.Top = tbContacts.Top + tbContacts.Height + 50
'  frmBelowContactList.Width = tbContacts.Width
'  '
'  '
'
'  'lvwHistory.Move 3300, scrollFrame.Top + scrollFrame.Height + 50, 8205 + 255, Top + scrollFrame.Height - tbContacts.Height - 50
'  'txtCompanyNote.Move 3300, scrollFrame.Top + scrollFrame.Height + 50, 2500, 1575
'  '
'  'fmeAuthStatus.Left = 3300 ' txtCompanyNote.Left + txtCompanyNote.Width + 50 '3300
'  'fmeAuthStatus.Top = scrollFrame.Top + scrollFrame.Height + 50 ''tbContacts.Top + tbContacts.Height + 50
'  '
'  'txtCompanyNote.Move 3300, scrollFrame.Top + scrollFrame.Height + 50, 2500, 1575
'  '
'  'lvwHistory.Move 0, fmeAuthStatus.Top + fmeAuthStatus.Height + 50, 11900, Me.Height - (fmeAuthStatus.Top + fmeAuthStatus.Height + 50)
'  'lvwHistory.Move 0, fmeAuthStatus.Top + fmeAuthStatus.Height + 50, Me.Width, Me.Height - (fmeAuthStatus.Top + fmeAuthStatus.Height + 50)
'
'  '  fmeAuthStatus.Top = conHistory.Height - fmeAuthStatus.Height
'   ' grdHistory.Height = conHistory.Height - fmeAuthStatus.Height - 100
'
' ' conHistory.Top = scrollFrame.Height
'  '
'  tbContacts.Visible = True
'  Me.Enabled = True
'
'End Sub

Private Sub FillSearchTypes()
      '
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   cboSearchType.Clear
      '
110   cboSearchType.AddItem "Company"
120   cboSearchType.AddItem "Last Name"
130   cboSearchType.AddItem "First Name"
140   cboSearchType.AddItem "Phone 1"
150   cboSearchType.AddItem "Phone 2"
160   cboSearchType.AddItem "Email"
      '
170   cboSearchType.Text = GetSetting(App.Title, "Misc", "SearchType", "Company")
      '
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "FillSearchTypes", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub FillTypesBox()
  'rs
  On Error GoTo EH
  '
  Dim rsTypes As New ADODB.Recordset
  '
  Set rsTypes = New ADODB.Recordset
  rsTypes.Open "SELECT * FROM Ttype", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  cboContactType.Clear
  '
  Do While Not rsTypes.eof
    '
    cboContactType.AddItem rsTypes!Description & vbNullString
    cboContactType.ItemData(cboContactType.NewIndex) = rsTypes!TypeID
    '
    rsTypes.MoveNext
  Loop
  '
  DBOps.ZapRS rsTypes
  '
  Exit Sub
EH:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FCustomer.InitCboContactType.", vbCritical, "Error"
End Sub

Private Sub FillStatusBox()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim rsStatus As Recordset
        '
        'Set rsStatus = dbMain.OpenRecordset("SELECT * FROM tblStatus", dbOpenForwardOnly)
110     Set rsStatus = New ADODB.Recordset
120     rsStatus.Open "SELECT * FROM tblStatus", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
        '
130     Do While Not rsStatus.eof
140       cboStatus.AddItem "" & rsStatus!Status
150       cboSearchStatus.AddItem "" & rsStatus!Status
160       rsStatus.MoveNext
170     Loop
        '
180     DBOps.ZapRS rsStatus
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "FillStatusBox", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub RefreshColors()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim lDarkColor As Long
110   Dim lLightColor As Long
120   Dim lListColor As Long
      '
130   lDarkColor = GetSetting(App.Title, "Settings", "DarkColor", &HFFCCCC)
140   lLightColor = GetSetting(App.Title, "Settings", "LightColor", &HFFE1E1)
150   lListColor = GetSetting(App.Title, "Settings", "ListColor", vbWhite)
      '
160   lvwHistory.BackColor = lListColor
170   picCanvas.BackColor = lLightColor
180   picShipping.BackColor = lDarkColor
190   picPhoneEmail.BackColor = lDarkColor
200   picGeneral.BackColor = lDarkColor
210   picPCEmail.BackColor = lDarkColor
220   picMailing.BackColor = lDarkColor
230   picBranch.BackColor = lDarkColor
240   picCustGroups.BackColor = lDarkColor
250   chkSelected.BackColor = lDarkColor
260   chkPreferredAddress(0).BackColor = lDarkColor
270   chkPreferredAddress(1).BackColor = lDarkColor
280   Me.chkContactByEmail.BackColor = lDarkColor
290   fmeAuthStatus.BackColor = lDarkColor
300   fmePVAuthStatus.BackColor = lDarkColor
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "RefreshColors", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdBack_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim lBack As Long
110   lBack = ContactStack.Back(ContactData.ID)
      '
120   If lBack > 0 Then
130     LoadContact lBack, False
140   End If
      '
150   cmdForward.Enabled = ContactStack.EnableForward
160   cmdBack.Enabled = ContactStack.EnableBack
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdBack_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub cmdForward_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim lForward As Long
110   lForward = ContactStack.Forward(ContactData.ID)
120   If lForward > 0 Then
130     LoadContact lForward, False
140   End If
150   cmdForward.Enabled = ContactStack.EnableForward
160   cmdBack.Enabled = ContactStack.EnableBack
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "cmdForward_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
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
  sFormula = Replace(sFormula, "PVAuthRemaining", "([PVAuthDays] - DateDiff(Day, [PVAuthDate], GETDATE()))")
  sFormula = Replace(sFormula, "isnull(shipdate)", "(ShipDate = Null)")
  '
  ConvertFormula = sFormula
  '
  Exit Function
EH:
  MsgBox Err.Description & " in Convert Formula.)"
End Function

Private Sub LoadSearchResults()
  On Error GoTo EH
  '
  If Me.cboSearchType = "Company" Then '* Company
    If Not (rsSearchCompany.BOF Or rsSearchCompany.eof) Then
        LoadOnlyCompany rsSearchCompany!ID, True
    End If
  Else '* Contact, Phone
    If Not (rsSearchContact.BOF Or rsSearchContact.eof) Then
      LoadContact rsSearchContact.Fields("ContactID").Value, False
    End If
  End If
  '
  Exit Sub
EH:
  MsgBox Err.Description & " in Load Search Results."
End Sub

Private Sub SaveContact()
  On Error GoTo EH
  '
  Dim bListPropertiesChanged As Boolean
  '
  If txtFirstName.Text = vbNullString Then
    MsgBox "First Name is a required field for a contact record. You must complete" & vbCrLf & "this field or cancel your changes before you can continue."
    Exit Sub
  End If
  '
  'If Not ContactData Is Nothing Then
    With ContactData
      If .FirstName <> Me.txtFirstName.Text Then
        bListPropertiesChanged = True
      End If
      '
      If .LastName <> Me.txtLastName.Text Then
        bListPropertiesChanged = True
      End If
      '
      If .AuthDays <> Me.txtAuthDays.Text Then
        bListPropertiesChanged = True
      End If
      '
      If .AuthRemaining <> Me.lblAuthRemaining.Caption Then
        bListPropertiesChanged = True
      End If
      '
      If .AuthDate <> nnNum(Me.mskAuthDate.Value) Then
        bListPropertiesChanged = True
      End If
      '
      'If .ContactType <> Me.cboContactType.Text Then
       ' bListPropertiesChanged = True
      'End If
      '
      If .Status <> Me.cboStatus.Text Then
        bListPropertiesChanged = True
      End If
      '
      .FirstName = Me.txtFirstName.Text
      .LastName = Me.txtLastName.Text
      .Salutation = Me.cboSalutation.Text
      .Title = Me.txtTitle.Text
      .Address1 = Me.txtAddress1.Text
      .Address2 = Me.txtAddress2.Text
      .City = Me.txtCity.Text
      .State = Me.txtState.Text
      .Zip = Me.txtZIP.Text
      '
      .MailAddress1 = Me.txtMailAddress1.Text
      .MailAddress2 = Me.txtMailAddress2.Text
      .MailCity = Me.txtMailCity.Text
      .MailState = Me.txtMailState.Text
      .MailZip = Me.txtMailZIP.Text
      '
      .PCEmail = Me.txtPCEmail.Text
      .PCEmailPassword = Me.txtPCEmailPassword.Text
      '
      .Phone1 = StripChars(Me.txtPhone1.Text) 'Me.mskPhone1.Value
      .Phone2 = StripChars(Me.txtPhone2.Text) 'Me.mskPhone2.Value
      .Fax = StripChars(Me.txtFax.Text) 'Me.mskFAX.Value
      .Email = Me.txtEmail.Text
      .Source = Me.cboSource.Text
      .Selected = Me.chkSelected.Value
      '
      '.WebPassword = Me.txtWebPassword
      .ContactByEmail = IIf((Me.chkContactByEmail = 1), True, False)
      '
      If Me.chkPreferredAddress(0) = 1 Then
        .PreferredAddress = 0
      Else
        .PreferredAddress = 1
      End If
      '
      '.Notes = Me.txtNotes.Text
      '
      .Status = Me.cboStatus.Text
      '
      '.DateEntered = nnNum(Me.mskCreated.Value)
      '
      .Rate = Me.tnmRate.Value
      .RateExpDate = nnNum(Me.mskRateExpDate.DateValue)
      '
      If cboContactType.ListIndex > -1 Then
        .ContactType = cboContactType.ItemData(cboContactType.ListIndex)
      Else
        .ContactType = 0
      End If
      '
      .ShipStatus = Me.cboShipStatus.Text
      .AuthStatus = Me.cboAuthStatus.Text
      .DownloadStatus = Me.cboDownloadStatus.Text
      .DownloadDate = nnNum(Me.mskDownloadDate.Value)
      .ShipDate = nnNum(Me.mskShipDate.Value)
      .AuthDate = nnNum(Me.mskAuthDate.Value)
      .AuthDays = Me.txtAuthDays.Text
      .VersionShipped = Me.txtVersionShipped.Text
      '.AuthRemaining = Me.lblAuthRemaining.Caption
      .GraceDays = nnNum(Me.txtGraceDays)
      .OnlineAuths = nnNum(Me.txtAuths)
      .SaleDate = nnNum(Me.mskSaleDate.Value)
      .SaleDays = nnNum(Me.txtSaleDays)
      '
      .PVShipStatus = Me.cboPVShipStatus.Text
      .PVAuthStatus = Me.cboPVAuthStatus.Text
      .PVDownloadStatus = Me.cboPVDownloadStatus.Text
      .PVDownloadDate = nnNum(Me.mskPVDownloadDate.Value)
      .PVShipDate = nnNum(Me.mskPVShipDate.Value)
      .PVAuthDate = nnNum(Me.mskPVAuthDate.Value)
      .PVAuthDays = Me.txtPVAuthDays.Text
      .PVVersionShipped = Me.txtPVVersionShipped.Text
      '.PVAuthRemaining = Me.lblPVAuthRemaining.Caption
      .PVGraceDays = nnNum(Me.txtPVGraceDays)
      .PVOnlineAuths = nnNum(Me.txtPVAuths)
      .PVSaleDate = nnNum(Me.mskPVSaleDate.Value)
      .PVSaleDays = nnNum(Me.txtPVSaleDays)
      '
      Dim lBranchTempID As Long
      '
      If cboBranch.Text <> "" Then
        If cboBranch.Text = "None" Then
          lBranchTempID = 0
        Else
          lBranchTempID = Branchs(cboBranch.ListIndex).BranchID
        End If
         '
        If .BranchID <> lBranchTempID Then
          bListPropertiesChanged = True
        End If
         '
        .BranchID = lBranchTempID
      End If
       '
      If Not (Contact.Save(ContactData, bEnteringNewContact)) Then
        GoTo EH
      End If
       '
       SaveCustGroups
       '
      LoadContact .ID, (bEnteringNewContact Or bListPropertiesChanged)
       '
      bEnteringNewContact = False
       '
    '  SaveCustGroups
    End With
  ' End If
   '
  Exit Sub
EH:
  MsgBox Err.Description & " in Save Contact."
End Sub

'Private Function GetBranchIDFromIndex(plIndex As Long) As Long
'  GetBranchIDFromIndex = branches(plIndex)
'End Function


Private Sub LoadOnlyCompany(plCompanyID As Long, pbTryToLoadSelected As Boolean)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Set ContactData = New CContactData
      '
110   Set CompanyData = New CCompanyData
      '
120   Company.Load CompanyData, plCompanyID
      '
130   FillCompanyControls
      '
140   If Not lvwContact.SelectedItem Is Nothing And pbTryToLoadSelected Then
150     If CLng(Mid$(lvwContact.SelectedItem.Key, 3)) > 0 Then
160       LoadContact CLng(Mid$(lvwContact.SelectedItem.Key, 3)), False
170     Else
180       FillContactControls
190     End If
200   Else
210     FillContactControls
220   End If
      '
230   If CompanyData.ID <> 0 Then
240     If ContactData.ID <= 0 Then
250       SetControlsCompanyOnlyLoaded
260     End If
270   Else
280     SetControlsNothingLoaded
290   End If
      '
300   tbContacts.Tab = 0
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "LoadOnlyCompany", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Sub LoadContact(plContactID As Long, pbForceReloadCompany As Boolean)
      '
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Set ContactData = New CContactData
      '
110   If plContactID <> 0 Then '
120     Contact.Load ContactData, plContactID
130   End If
      '
140   If (CompanyData.ID <> ContactData.CompanyID) Or pbForceReloadCompany Then
150     Set CompanyData = New CCompanyData
        '
160     If ContactData.CompanyID <> 0 Then
170       Company.Load CompanyData, ContactData.CompanyID
          '
180       FillCompanyControls
190     End If
200   Else
202     If lvwContact.SortKey <> 2 Then
210       SetContactSelected lvwContact, ContactData.ID
212     End If
220   End If
      '
230   FillContactControls
      '
      'tbContacts.Tab = 0
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "LoadContact", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub


Private Sub FillCompanyControls()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   lblCompany.Caption = "Company: " & CompanyData.Name '& " " & ContactData.FirstName & " " & ContactData.LastName
      '
110   txtCompanyNote.Text = CompanyData.Note
      '
120   If CompanyData.DoNotContact Then
130      txtCompanyNote.Text = "ONLY CONTACT COMPANY ADMIN'S " & vbCrLf & CompanyData.Note
140   End If
      '
145   SetStar CompanyData.InterestRank
150   FillBranches
      '
160   If Not lstBranch.SelectedItem Is Nothing Then
170     FillCompanyContactList GetIDFromKey(lstBranch.SelectedItem.Key)
180   Else
190     FillCompanyContactList 0
200   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "FillCompanyControls", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
Private Sub FillContactControls()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   bEnteringNewContact = False
      '
110   If ContactData.ID > 0 Then
120     ContactStack.Current ContactData.ID
130     cmdBack.Enabled = ContactStack.EnableBack
140     cmdForward.Enabled = ContactStack.EnableForward
150   End If
      '
160   On Error Resume Next
      '
170   With ContactData
180     Me.txtFirstName.Text = .FirstName
190     Me.txtLastName.Text = .LastName
        '
        'If .Loaded Then
        '  If .Salutation = "Ms." Then
        '    Me.cboSalutation.ListIndex = 1
        '  Else
        '    Me.cboSalutation.ListIndex = 0
        '  End If
        'Else
        '  Me.cboSalutation.ListIndex = -1
        'End If
200     Select Case .Salutation
          Case "Mr."
210         Me.cboSalutation.ListIndex = 0
220       Case "Ms."
230         Me.cboSalutation.ListIndex = 1
240       Case Else
250         Me.cboSalutation.ListIndex = -1
260     End Select
        '
270     Me.txtTitle.Text = .Title
280     Me.txtAddress1.Text = .Address1
290     Me.txtAddress2.Text = .Address2
300     Me.txtCity.Text = .City
310     Me.txtState.Text = .State
320     Me.txtZIP.Text = .Zip
        '
330     Me.txtMailAddress1.Text = .MailAddress1
340     Me.txtMailAddress2.Text = .MailAddress2
350     Me.txtMailCity.Text = .MailCity
360     Me.txtMailState.Text = .MailState
370     Me.txtMailZIP.Text = .MailZip
        '
380     Me.txtPCEmail.Text = .PCEmail
390     Me.txtPCEmailPassword.Text = .PCEmailPassword
        '
400     Me.txtPhone1.Text = FormatPhoneNumber(.Phone1) 'Me.mskPhone1.Value = .Phone1
410     Me.txtPhone2.Text = FormatPhoneNumber(.Phone2) 'Me.mskPhone2.Value = .Phone2
420     Me.txtFax.Text = FormatPhoneNumber(.Fax) 'Me.mskFAX.Value = .Fax
430     Me.txtEmail.Text = .Email
440     Me.cboSource.Text = .Source
450     Me.chkSelected.Value = .Selected
        '
460     If .PreferredAddress = 0 Then
470       Me.chkPreferredAddress(0) = 1
480       Me.chkPreferredAddress(1) = 0
490     Else
500       Me.chkPreferredAddress(0) = 0
510       Me.chkPreferredAddress(1) = 1
520     End If
        '
530     Me.chkContactByEmail = IIf(.ContactByEmail, 1, 0)
       ' Me.txtWebPassword = .WebPassword
        '
        'Me.txtNotes.Text = .Notes
        '
    '570     On Error Resume Next
540       Me.cboStatus.Text = .Status
    '590     On Error GoTo EH
    
        'Me.mskCreated.value = IIf(.DateEntered = 0, Null, .DateEntered)
        '
550     Me.tnmRate.Value = .Rate
560     SetCboContactType (.ContactType)
570     Me.mskRateExpDate.DateValue = IIf(.RateExpDate = 0, Null, .RateExpDate)
        '
580      Me.cboShipStatus.Text = .ShipStatus
590      Me.cboDownloadStatus.Text = .DownloadStatus
600      Me.cboAuthStatus.Text = .AuthStatus
610      Me.mskDownloadDate.Value = IIf(.DownloadDate = 0, Null, .DownloadDate)
620      Me.mskShipDate.Value = IIf(.ShipDate = 0, Null, .ShipDate)
630      Me.mskAuthDate.Value = IIf(.AuthDate = 0, Null, .AuthDate)
640      Me.txtAuthDays.Text = .AuthDays
650      Me.txtVersionShipped.Text = .VersionShipped
660      Me.txtGraceDays.Text = .GraceDays
670      Me.txtSaleDays.Text = .SaleDays
680      Me.txtAuths.Text = .OnlineAuths
690      Me.mskSaleDate.Value = IIf(.SaleDate = 0, Null, .SaleDate)
         '
700      Me.txtPendingDays.Text = .DaysPending 'CalculatePendingDays(mskSaleDate.Value, .GraceDays, .SaleDays)
         '
710      Me.lblAuthRemaining.Caption = .AuthRemaining
         '
720       If CLng(lblAuthRemaining.Caption) > 0 Then
730         lblExpires.Caption = "Expires: " & (Date + CLng(lblAuthRemaining.Caption))
740       Else
750         lblExpires.Caption = "Expires: NA"
760       End If
         '
770      If .ShipStatus = "" Then
780        .ShipStatus = GetSetting(App.Title, "Preferences", "InitShipStatus", "Not Shipped")
790      End If
         '
800      If .AuthStatus = "" Then
810        .AuthStatus = GetSetting(App.Title, "Preferences", "InitAuthStatus", "Not Authorized")
820      End If
         '
830      Me.cboPVShipStatus.Text = .PVShipStatus
840      Me.cboPVDownloadStatus.Text = .PVDownloadStatus
850      Me.cboPVAuthStatus.Text = .PVAuthStatus
860      Me.mskPVDownloadDate.Value = IIf(.PVDownloadDate = 0, Null, .PVDownloadDate)
870      Me.mskPVShipDate.Value = IIf(.PVShipDate = 0, Null, .PVShipDate)
880      Me.mskPVAuthDate.Value = IIf(.PVAuthDate = 0, Null, .PVAuthDate)
890      Me.txtPVAuthDays.Text = .PVAuthDays
        ' Me.mskCopies.value = .PVCopies
900      Me.txtPVVersionShipped.Text = .PVVersionShipped
910      Me.txtPVGraceDays = .PVGraceDays
         'Me.txtPVPendingDays = .PVSaleDays
920      Me.txtPVAuths = .PVOnlineAuths
930      Me.mskPVSaleDate = IIf(.PVSaleDate = 0, Null, .PVSaleDate)
940      Me.txtPVSaleDays.Text = .PVSaleDays
         '
950      Me.txtPVPendingDays.Text = .PVDaysPending 'CalculatePendingDays(mskPVSaleDate.Value, .PVGraceDays, .PVSaleDays)
         '
960      Me.lblPVAuthRemaining.Caption = .PVAuthRemaining
            '
970       If CLng(lblPVAuthRemaining.Caption) > 0 Then
980         lblPVExpires.Caption = "Expires: " & (Date + CLng(lblPVAuthRemaining.Caption))
990       Else
1000         lblPVExpires.Caption = "Expires: NA"
1010       End If
              '
1020      If .PVShipStatus = "" Then
1030        .PVShipStatus = GetSetting(App.Title, "Preferences", "InitShipStatus", "Not Shipped")
1040      End If
              '
1050      If .PVAuthStatus = "" Then
1060        .PVAuthStatus = GetSetting(App.Title, "Preferences", "InitAuthStatus", "Not Authorized")
1070      End If
               '
1080      FillBranches
1090      LoadCustGroups
     
               '
1100      EnableEditContactControls False
               '
               'If .ID > 0 Then
1110      LoadHistory .ID
              'End If
1120   End With
       '<EhFooter>
       '
       Exit Sub
       '
EH:
       ErrorMgr.Raise "FContact", "FillContactControls", Err.Number, Err.Description, Erl
       '</EhFooter>
End Sub

'Private Function CalculatePendingDays(pdSaleDate As Date, plGraceDays As Long, plSaleDays As Long) As Long
'  Dim lTempPending As Long
'  Dim lDaysPassed As Long
'  '
'  lDaysPassed = Abs(DateDiff("y", pdSaleDate, Now))
'  '
'  If plGraceDays < 0 Then
'    CalculatePendingDays = plSaleDays
'  Else
'    If lDaysPassed < plGraceDays Then
'      CalculatePendingDays = plSaleDays
'    Else
'      lTempPending = plSaleDays - (lDaysPassed - plGraceDays)
'      '
'      If lTempPending >= 0 Then
'        CalculatePendingDays = lTempPending
'      Else
'        CalculatePendingDays = 0
'      End If
'    End If
'  End If
'End Function

Private Sub FillCompanyContactList(plBranchID As Long)
  On Error GoTo EH
    Dim Contacts As New CContacts
    '
    If CompanyData.ID <> 0 Then
      Contact.LoadCollection CompanyData.ID, Contacts, plBranchID
      '
      FillContactList lvwContact, Contacts
      '
      On Error Resume Next
        Set lvwContact.SelectedItem = lvwContact.ListItems("ID" & Format(ContactData.ID))
      On Error GoTo EH
      '
    End If
    Set Contacts = Nothing
    '
    Exit Sub
EH:
    MsgBox Err.Description & " in FillContactList."
End Sub

Public Function FillContactList(ByRef pContactList As ListView, pContacts As CContacts) As Long
  On Error GoTo EH
  '
  Dim iPos As Integer
  Dim strKey As String
  Dim iIcon As Integer
  Dim lColor As Long
  '
  Dim sPreviousText As String
  '
  If Not (pContactList.SelectedItem Is Nothing) Then
    sPreviousText = pContactList.SelectedItem.Key
  End If
  '
  pContactList.ListItems.Clear
  pContactList.ColumnHeaders.Clear
  '
  pContactList.ColumnHeaders.Add , "a", "First Name", 1000
  pContactList.ColumnHeaders.Add , "b", "Last Name", 1000
  pContactList.ColumnHeaders.Add , "c", "Days", 1000
  '
  With pContacts
    For iPos = 1 To .Count
       Select Case .Item(iPos).ContactType
        Case 0  ' unknown
          iIcon = 1
        Case 1 'adjuster
          iIcon = 2
        Case 2 'admin
          iIcon = 4
        Case 3 'tech
          iIcon = 3
        Case 4 'sec
          iIcon = 5
        Case 5 'unknown
          iIcon = 1
        Case Else
         iIcon = 2
      End Select
      '
      Select Case .Item(iPos).Status
        '
        Case "Customer"
          lColor = &HC00000
        Case "Prospect"
          lColor = &HC000&
        Case "Future Prospect"
          lColor = &HC000&
        Case "Inactive"
          lColor = &H404040
        Case "Contact"
          lColor = &HC000C0
        Case Else
          lColor = &H0
      End Select
      '
      strKey = "ID" & .Item(iPos).ID
      '
      pContactList.ListItems.Add , strKey, .Item(iPos).FirstName, , iIcon
      pContactList.ListItems.Item(strKey).ForeColor = lColor
      pContactList.ListItems.Item(strKey).ListSubItems.Add(, , .Item(iPos).LastName & vbNullString).ForeColor = lColor
      pContactList.ListItems.Item(strKey).ListSubItems.Add(, , .Item(iPos).AuthDays - DateDiff("d", .Item(iPos).AuthDate, Date) & vbNullString).ForeColor = lColor
    Next
    FillContactList = .Count
  End With
  '
  If sPreviousText <> "" Then
   On Error Resume Next
   Set pContactList.SelectedItem = pContactList.ListItems(sPreviousText)
   pContactList.SelectedItem.EnsureVisible
   On Error GoTo EH
  End If
  '
Exit Function
EH:
 MsgBox Err.Description & " in FillList."
End Function

Private Sub SetContactSelected(ByRef pContactList As ListView, plContactID As Long)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   On Error Resume Next
110   If plContactID > 0 Then
120    Set pContactList.SelectedItem = pContactList.ListItems("ID" & plContactID)
130    pContactList.SelectedItem.EnsureVisible
140   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "SetContactSelected", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub SetCboContactType(NewValue As Integer)
  On Error GoTo EH
  Dim rsTypes As New ADODB.Recordset
  Dim iType As Integer
  '
  rsTypes.Open "SELECT * FROM Ttype", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  rsTypes.MoveFirst
  '
  rsTypes.Find "TypeID = " & NewValue
  '
  If Not rsTypes.eof Then
   iType = rsTypes!TypeID
  Else
   iType = 1
  End If
  '
  Dim i As Long
  '
  While iType <> cboContactType.ItemData(i) And i <= cboContactType.ListCount
    i = i + 1
  Wend
  cboContactType.ListIndex = i
  '
  DBOps.ZapRS rsTypes
  '
  Exit Sub
EH:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FCustomer.SetCboContactType.", vbCritical, "Error"

End Sub

Private Sub SetControlsCompanyOnlyLoaded()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   EnableEditContactControls False
      '
110   cmdEditContact.Enabled = False
      '
120   cmdDeleteContact.Enabled = False
      '
130   cmdAction.Enabled = False
      '
140   cmdPrintLabel.Enabled = False
      '
150   cmdSetAppt.Enabled = False
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "SetControlsCompanyOnlyLoaded", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub SetControlsNothingLoaded()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SetControlsCompanyOnlyLoaded
      '
110   cmdEditCompany.Enabled = False
      '
120   cmdDeleteCompany.Enabled = False
      '
130   cmdNewContact.Enabled = False
      '
140   tbContacts.Enabled = False
      '
150   cmdClipBoard.Enabled = False
      '
160   cmdPrintList.Enabled = False
      '
170   cmdAddBranch.Enabled = False
      '
180   cmdDeleteBranch.Enabled = False
      '
190   cmdEditBranch.Enabled = False
      '
200   cmdAddCustGroup.Enabled = False
      '
205   cmdClearCustGroup.Enabled = False
      '
210   cmdDelGroup.Enabled = False
      '
'220   lstCustGroups.Enabled = False
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "SetControlsNothingLoaded", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub EnableEditContactControls(pfEnable As Boolean)
  On Error GoTo EH
  '
  bEnabled = pfEnable
  '
  Dim lBackColor As Long
  Dim lForeColor As Long
  Dim fLocked As Boolean
  '
  lvwHistory.Enabled = Not pfEnable
  grdHistory.Enabled = Not pfEnable
  '
  cmdRand1.Enabled = Not pfEnable
  cmdRand2.Enabled = Not pfEnable
  cmdRand3.Enabled = Not pfEnable
  cmdHistoryType.Enabled = Not pfEnable
  tbContacts.Enabled = Not pfEnable
  chkAlpha.Enabled = Not pfEnable
  lstGroups.Enabled = Not pfEnable
  lstBranch.Enabled = Not pfEnable
  lvwContact.Enabled = Not pfEnable
  cmdClipBoard.Enabled = Not pfEnable
  cmdPrintList.Enabled = Not pfEnable
  lvwPMContacts.Enabled = Not pfEnable
  scrollSearchFields.Enabled = Not pfEnable
  cmdSearch.Enabled = Not pfEnable
  lvwSearchContacts.Enabled = Not pfEnable
  cmdNewCompany.Enabled = Not pfEnable
  cmdEditCompany.Enabled = Not pfEnable
  cmdDeleteCompany.Enabled = Not pfEnable
  cmdSetAppt.Enabled = Not pfEnable
  cmdPrintLabel.Enabled = Not pfEnable
  cmdDeleteContact.Enabled = Not pfEnable
  cmdEditContact.Enabled = Not pfEnable
  cmdNewContact.Enabled = Not pfEnable
  'cmdForward.Enabled = Not pfEnable
  'cmdBack.Enabled = Not pfEnable
  txtSearch.Enabled = Not pfEnable
  cboSearchType.Enabled = Not pfEnable
  cmdRefresh.Enabled = Not pfEnable
  cmdAction.Enabled = Not pfEnable
  cmdAddBranch.Enabled = Not pfEnable
  cmdDeleteBranch.Enabled = Not pfEnable
  cmdEditBranch.Enabled = Not pfEnable
  cmdAddCustGroup.Enabled = Not pfEnable
  cmdClearCustGroup.Enabled = Not pfEnable
  cmdDelGroup.Enabled = Not pfEnable
  '
  fLocked = Not pfEnable
  '
  If fLocked Then
    lForeColor = &H80000011
  Else
    lForeColor = &H80000008
  End If
  '
  If pfEnable Then
    lBackColor = &H80000005
  Else
    lBackColor = &H8000000F
  End If
  '
  'Me.mnuAssignContact.Enabled = pfEnable
  'Me.mnuCombineContact.Enabled = pfEnable
  'Me.mnuCtxTagEMail.Enabled = pfEnable
  '
  Me.txtFirstName.Locked = fLocked
  Me.txtLastName.Locked = fLocked
  Me.cboSalutation.Locked = fLocked
  Me.txtTitle.Locked = fLocked
  Me.txtAddress1.Locked = fLocked
  Me.txtAddress2.Locked = fLocked
  Me.txtCity.Locked = fLocked
  Me.txtState.Locked = fLocked
  Me.txtZIP.Locked = fLocked
  '
  Me.txtMailAddress1.Locked = fLocked
  Me.txtMailAddress2.Locked = fLocked
  Me.txtMailCity.Locked = fLocked
  Me.txtMailState.Locked = fLocked
  Me.txtMailZIP.Locked = fLocked
  '
  Me.txtPCEmail.Locked = fLocked
  Me.txtPCEmailPassword.Locked = fLocked
  '
  Me.txtPhone1.Locked = fLocked 'Me.mskPhone1.ReadOnly = fLocked
  Me.txtPhone2.Locked = fLocked  'Me.mskPhone2.ReadOnly = fLocked
  Me.txtFax.Locked = fLocked  'Me.mskFAX.ReadOnly = fLocked
  '
  Me.txtEmail.Locked = fLocked
  Me.cboSource.Enabled = pfEnable
  Me.chkSelected.Enabled = pfEnable
  Me.chkPreferredAddress(0).Enabled = pfEnable
  Me.chkPreferredAddress(1).Enabled = pfEnable
  'Me.txtNotes.Locked = fLocked
  'Me.txtNotes.ForeColor = lForeColor
  '
  Me.txtWebPassword.Locked = fLocked
  Me.chkContactByEmail.Enabled = pfEnable
  '
  Me.cboStatus.Enabled = pfEnable
  Me.cboShipStatus.Enabled = pfEnable
  Me.cboDownloadStatus.Enabled = pfEnable
 ' Me.cboProduct.Enabled = pfEnable
  Me.cboAuthStatus.Enabled = pfEnable
  Me.mskCreated.Enabled = pfEnable
  Me.mskShipDate.Enabled = pfEnable
  Me.mskDownloadDate.Enabled = pfEnable
  'Me.mskAuthDate.Enabled = pfEnable
  Me.cmdShipDate.Enabled = pfEnable
  Me.cmdDownloadDate.Enabled = pfEnable
  'Me.cmdAuthDate.Enabled = pfEnable
  'Me.txtAuthDays.Enabled = pfEnable
 ' Me.mskCopies.Enabled = pfEnable
  Me.txtVersionShipped.Enabled = pfEnable
  Me.txtPendingDays.Enabled = pfEnable
  Me.mskSaleDate.Enabled = pfEnable
  Me.cmdSalesDate.Enabled = pfEnable
  Me.txtSaleDays.Enabled = pfEnable
  Me.txtGraceDays.Enabled = pfEnable
  Me.txtAuths.Enabled = pfEnable
  Me.tnmRate.Enabled = pfEnable
   '
  Me.txtPVPendingDays.Enabled = pfEnable
  Me.mskPVSaleDate.Enabled = pfEnable
  Me.cmdPVSalesDate.Enabled = pfEnable
  Me.txtPVSaleDays.Enabled = pfEnable
  Me.txtPVGraceDays.Enabled = pfEnable
  Me.txtPVAuths.Enabled = pfEnable
  Me.cboPVShipStatus.Enabled = pfEnable
  Me.cboPVDownloadStatus.Enabled = pfEnable
  Me.cboPVAuthStatus.Enabled = pfEnable
  Me.mskPVShipDate.Enabled = pfEnable
  Me.mskPVDownloadDate.Enabled = pfEnable
   'Me.mskPVAuthDate.Enabled = pfEnable
  Me.cmdPVShipDate.Enabled = pfEnable
  Me.cmdPVDownloadDate.Enabled = pfEnable
  Me.txtPVVersionShipped.Enabled = pfEnable
   'Me.cmdPVAuthDate.Enabled = pfEnable
   'Me.txtPVAuthDays.Enabled = pfEnable
   '
  Me.cboContactType.Enabled = pfEnable
  Me.mskRateExpDate.Enabled = pfEnable
   '
  Me.cboBranch.Enabled = pfEnable
   '
  lstCustGroups.Enabled = pfEnable
  lstCustGroups.Refresh
 '
 '  Me.cmdAction(0).Enabled = pfEnable
 '  Me.cmdAction(1).Enabled = pfEnable
 '  Me.cmdAction(3).Enabled = pfEnable
 '  Me.cmdAction(4).Enabled = pfEnable
 '  Me.cmdAction(5).Enabled = pfEnable
  Me.cmdSaveContact.Enabled = pfEnable
  Me.cmdCancelContact.Enabled = pfEnable
   'Me.cmdDeleteContact.Enabled = pfEnable
  ' Me.cmdPrintLabel.Enabled = pfEnable
   'Me.cmdTagForEMail.Enabled = pfEnable
   '
  Me.txtFirstName.ForeColor = lForeColor
  Me.txtLastName.ForeColor = lForeColor
  Me.cboSalutation.ForeColor = lForeColor
  Me.txtTitle.ForeColor = lForeColor
  Me.txtAddress1.ForeColor = lForeColor
  Me.txtAddress2.ForeColor = lForeColor
  Me.txtCity.ForeColor = lForeColor
  Me.txtState.ForeColor = lForeColor
  Me.txtZIP.ForeColor = lForeColor
   '
  Me.txtMailAddress1.ForeColor = lForeColor
  Me.txtMailAddress2.ForeColor = lForeColor
  Me.txtMailCity.ForeColor = lForeColor
  Me.txtMailState.ForeColor = lForeColor
  Me.txtMailZIP.ForeColor = lForeColor
   '
  Me.txtPCEmail.ForeColor = lForeColor
  Me.txtPCEmailPassword.ForeColor = lForeColor
   '
'  Me.mskPhone1.ForeColor = lForeColor
'  Me.mskPhone2.ForeColor = lForeColor
'  Me.mskFAX.ForeColor = lForeColor
  Me.txtPhone1.ForeColor = lForeColor
  Me.txtPhone2.ForeColor = lForeColor
  Me.txtFax.ForeColor = lForeColor
  Me.txtEmail.ForeColor = lForeColor
   'Me.txtNotes.ForeColor = lForeColor
  Me.cboContactType.ForeColor = lForeColor
 '  '
 '  Me.cboStatus.ForeColor = lForeColor
 '  Me.cboShipStatus.ForeColor = lForeColor
 '  Me.cboAuthStatus.ForeColor = lForeColor
 '  Me.mskCreated.ForeColor = lForeColor
 '  Me.mskShipDate.ForeColor = lForeColor
 '  Me.mskAuthDate.ForeColor = lForeColor
 '  Me.txtAuthDays.ForeColor = lForeColor
 '  Me.mskCopies.ForeColor = lForeColor
 '  Me.txtVersionShipped.ForeColor = lForeColor
 '  Me.lblAuthRemaining.ForeColor = lForeColor
   '
  Exit Sub
EH:
  MsgBox Err.Description & " in FContact: DataContact_Enable."
End Sub

Private Sub LoadHistory(ByVal lcustomerID As Long)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If bUseNewNoteList = True Then
110     LoadHistoryListView lcustomerID
      'End If
      ''
      'If grdHistory.Visible = True Then
120   Else
130     LoadHistoryGrid lcustomerID
140   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "LoadHistory", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub LoadHistoryGrid(ByVal lcustomerID As Long)
  grdHistory.Redraw = False
  '
  Me.grdHistory.RemoveAll
  '
  If lcustomerID > 0 Then
    Dim Event1 As New CEvent
    Dim Events As New CEvents
    Dim iPos As Integer
    '
    Event1.LoadCollection lcustomerID, Events
    '
    For iPos = 1 To Events.Count
      With Events.Item(iPos)
        grdHistory.AddItem .RecID & vbTab & .CustRecID & vbTab & .EventDate & vbTab & .EventTime & vbTab & .EventType & vbTab & .EventUser & vbTab & .EventSubject & vbTab & .EventResults & vbTab & .ProductID
      End With
    Next
    '
    Set Events = Nothing
    Set Event1 = Nothing
  End If
  '
  grdHistory.Redraw = True
  '
  Exit Sub
EH:
  grdHistory.Redraw = True
  MsgBox Err.Description
End Sub

Private Sub grdHistory_RowLoaded(ByVal Bookmark As Variant)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim i As Integer
      '
110   If grdHistory.Columns(8).Value = 2 Then
120     For i = 0 To grdHistory.Cols - 3
130       grdHistory.Columns(i).CellStyleSet "PV"
140     Next i
150   End If
      '

      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "grdHistory_RowLoaded", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
Private Sub LoadHistoryListView(ByVal lcustomerID As Long)
  On Error GoTo EH
  '
  Dim sKey As String
  Dim lColor As Long
  '
  lvwHistory.Sorted = False
  '
  lvwHistory.Enabled = False
  '
  lvwHistory.ListItems.Clear
  '
  If lcustomerID > 0 Then
    Dim Event1 As New CEvent
    Dim Events As New CEvents
    Dim iPos As Integer
    Dim sSubject As String
    Dim lIcon As Long
    Dim lStickyIndex As Long
    '
    Event1.LoadCollection lcustomerID, Events
    '
    For iPos = 1 To Events.Count
      lColor = vbBlack
      '
      sKey = "I" & Events.Item(iPos).RecID
      '
      If Events.Item(iPos).EventSubject <> vbNullString Then
        sSubject = Events.Item(iPos).EventSubject & ", "
      Else
        sSubject = vbNullString
      End If
      '
      If Events.Item(iPos).Sticky Then
        lIcon = 10 'Pin Icon
        '
        lStickyIndex = lStickyIndex + 1
        '
        lvwHistory.ListItems.Add lStickyIndex, sKey, sSubject & Events.Item(iPos).EventResults, , lIcon
      Else
        lIcon = 9 'BNB Icon
        '
        lvwHistory.ListItems.Add , sKey, sSubject & Events.Item(iPos).EventResults, , lIcon
      End If
      '
      If Events.Item(iPos).ProductID = 2 Then
        lvwHistory.ListItems.Item(sKey).ForeColor = &HC00000   'Product.GetColor(2)
      End If
      '
      'lvwHistory.ListItems.Item(sKey).ForeColor = Product.GetColor(Events.Item(iPos).ProductID) 'lColor
      'lvwHistory.ListItems.Item(sKey).ToolTipText = Events.Item(iPos).EventResults
      'lvwHistory.ListItems.Item(sKey).ListSubItems.Add(, , Events.Item(iPos).EventSubject).ForeColor = lColor
      lvwHistory.ListItems.Item(sKey).ListSubItems.Add(, , Events.Item(iPos).EventType).ForeColor = lColor
      lvwHistory.ListItems.Item(sKey).ListSubItems.Add(, , Events.Item(iPos).EventUser).ForeColor = lColor
      lvwHistory.ListItems.Item(sKey).ListSubItems.Add(, , Events.Item(iPos).EventDate).ForeColor = lColor
      lvwHistory.ListItems.Item(sKey).ListSubItems.Add(, , Events.Item(iPos).EventTime).ForeColor = lColor
      
    Next iPos
    

   Set Event1 = Nothing
   
  End If
  '
  lvwHistory.Enabled = True
  '
  Exit Sub
EH:
  lvwHistory.Enabled = True
  '
  MsgBox Err.Description
End Sub

Private Sub lvwContact_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      '
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   SortListView lvwContact, ColumnHeader.Index
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "lvwContact_ColumnClick", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub FillDownloadStatusBox()
  On Error GoTo EH
  '
  Dim rsDownloadStatus As ADODB.Recordset
  '
  Set rsDownloadStatus = New ADODB.Recordset
  rsDownloadStatus.Open "SELECT * FROM tblDownloadStatus", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rsDownloadStatus.eof
    cboDownloadStatus.AddItem rsDownloadStatus!Status & vbNullString
    cboPVDownloadStatus.AddItem rsDownloadStatus!Status & vbNullString
    rsDownloadStatus.MoveNext
  Loop
  '
  DBOps.ZapRS rsDownloadStatus
  '
  Exit Sub
EH:
  MsgBox "Error " & Err.Number & ": " & "FillDownloadStatus", vbCritical, "Error"
End Sub

Private Sub FillShipStatusBox()
  On Error GoTo EH
  '
  Dim rsShipStatus As ADODB.Recordset
  '
  Set rsShipStatus = New ADODB.Recordset
  rsShipStatus.Open "SELECT * FROM tblShipStatus", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rsShipStatus.eof
    cboShipStatus.AddItem rsShipStatus!Status & vbNullString
    cboPVShipStatus.AddItem rsShipStatus!Status & vbNullString
    rsShipStatus.MoveNext
  Loop
  '
  DBOps.ZapRS rsShipStatus
  '
  Exit Sub
EH:
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FCustomer.FillDownloadStatusBox.", vbCritical, "Error"
End Sub

Private Sub FillAuthStatusBox()
  On Error GoTo EH
  '
  Dim rs As ADODB.Recordset
  '
  'Set rs = dbMain.OpenRecordset("SELECT * FROM tblAuthStatus ORDER BY RecID", dbOpenForwardOnly)
  Set rs = New ADODB.Recordset
  rs.Open "SELECT * FROM tblAuthStatus ORDER BY RecID", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rs.eof
    cboAuthStatus.AddItem rs!Status
    cboPVAuthStatus.AddItem rs!Status
    rs.MoveNext
  Loop
  '
  DBOps.ZapRS rs
  '
  Exit Sub
EH:
  MsgBox Err.Description
End Sub


Private Sub grdHistory_DblClick()
  On Error GoTo EH
  '
  If grdHistory.Columns(0).Value <> "" Then
    grdHistory.Redraw = False
    '
    If FNote.LoadNote(ContactData.ID, grdHistory.Columns(0).Value) Then
      LoadContact ContactData.ID, False
    End If
    'LoadHistory ContactData.ID
    grdHistory.Redraw = True
  End If
  '
  Exit Sub
EH:
  Screen.MousePointer = vbDefault
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FCustomer.grdHistory_DblClick.", vbCritical, "Error"
End Sub

Private Sub chkPreferredAddress_Click(Index As Integer)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If Index = 0 Then
110     If Me.chkPreferredAddress(0) = 1 Then
120       Me.chkPreferredAddress(1) = 0
130     Else
140       Me.chkPreferredAddress(1) = 1
150     End If
160   Else
170     If Me.chkPreferredAddress(1) = 1 Then
180       Me.chkPreferredAddress(0) = 0
190     Else
200       Me.chkPreferredAddress(0) = 1
210     End If
220   End If
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "chkPreferredAddress_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Sub SetupPMGroups()
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
  Dim sOrderField As String
  '
  If chkAlpha.Value = vbChecked Then
    sOrderField = "Label"
  Else
    sOrderField = "Priority"
  End If
'  If Filter = FilterAll Then
'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] ORDER BY [" & sOrderField & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
'  ElseIf Filter = FilterStandard Then
'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] NOT LIKE 'AM Best%' ORDER BY [" & sOrderField & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
'  ElseIf Filter = FilterAMBest Then
'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] LIKE 'AM Best%' ORDER BY [" & sOrderField & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
'  ElseIf Filter = FilterNone Then
'    rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] = 'Dummy'", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
'  End If
  '
  '
  rsGroupCategories.Open "SELECT * FROM [tblGroupCategories] WHERE [Label] NOT LIKE 'AM Best%' ORDER BY [" & sOrderField & "]", cnMain, adOpenKeyset, adLockReadOnly, adCmdText
  '
  Set rsGroups = New ADODB.Recordset
  '
  With rsGroupCategories
    Do While .eof = False
      If rsGroups.State = adStateOpen Then rsGroups.Close
      Dim sSQL As String
      '
      sSQL = "SELECT Count(*) as GroupCount " & _
        "FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
        "WHERE " & ConvertFormula(.Fields("Formula").Value)
      '
      rsGroups.Open sSQL, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
      '
      lstGroups.AddItem .Fields("Label").Value & " (" & rsGroups.Fields("GroupCount").Value & ")"
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
  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in frmContact.SetupGroups.", vbCritical, "Error"
End Sub

Private Sub SideSearch()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim SearchReport As New CReport
110   Dim CList        As New CContactList
      '
120   SearchReport.SortDirection = GetSortDirection(lvwSearchContacts)
130   SearchReport.SortField = GetSortField(lvwSearchContacts)
      '
      'SearchReport.ProductID = Product.GetProductID(cboProduct.Text)
140   SearchReport.FirstName = Me.txtSearchFirstName.Text
150   SearchReport.LastName = Me.txtSearchLastName.Text
160   SearchReport.Company = Me.txtSearchCompany.Text
170   SearchReport.Status = Me.cboSearchStatus.Text
180   SearchReport.City = Me.txtSearchCity.Text
190   SearchReport.Zip = Me.txtSearchZip.Text
200   SearchReport.State = Me.txtSearchState.Text
210   SearchReport.Source = Me.txtSearchSource.Text
220   SearchReport.Notes = Me.txtSearchNotes.Text
      '
230   SearchReport.Rtype = SimpleContact2
      '
240   lvwSearchContacts.ListItems.Clear
250   lvwSearchContacts.ColumnHeaders.Clear
      '
260   CList.FillList lvwSearchContacts, SearchReport.rsReport
270   lblResults.Caption = "Results: " & SearchReport.rsReport.RecordCount
      '
280   Set SearchReport = Nothing
290   Set CList = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "SideSearch", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub FillBranches()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim sPreviousKey As String
110   Dim lCurrentCboIndex As Long
120   Dim Branch As New CBranch
    '120   Dim BranchData As New CBranchData
      '
130   Set Branchs = New CBranchs
      '
140   Dim lBranchCount As Long
      '
150   If Not lstBranch.SelectedItem Is Nothing Then
160     sPreviousKey = lstBranch.SelectedItem.Key
170   End If
      '
180   lstBranch.ListItems.Clear
      '
190   lstBranch.ListItems.Add , , "Show All", , 8
      '
200   cboBranch.Clear
      '
210   cboBranch.AddItem "None"
      '
220   Branch.LoadCollection CompanyData.ID, Branchs
      '
230   lCurrentCboIndex = -1
      '
240   For lBranchCount = 1 To Branchs.Count
        '
250     lstBranch.ListItems.Add , "A" & Branchs(lBranchCount).BranchID, Branchs(lBranchCount).Name, , 7   ', BranchData.BranchID
        '
     '   sBranches(lBranchCount) = BranchData.Name
        '
260     If ContactData.BranchID = Branchs(lBranchCount).BranchID Then lCurrentCboIndex = lBranchCount
        '
270     cboBranch.AddItem Branchs(lBranchCount).Name
280   Next
      '
290   Set Branch = Nothing
      '
300   On Error Resume Next
310     If lCurrentCboIndex <> -1 Then
320       cboBranch.ListIndex = lCurrentCboIndex
330     End If
          '
340     If ContactData.ID > 0 Then
350       If GetIDFromKey(sPreviousKey) <> ContactData.BranchID Then Exit Sub
360     End If
        '
370     If sPreviousKey <> "" Then
380       Set lstBranch.SelectedItem = lstBranch.ListItems(sPreviousKey)
390       lstBranch.SelectedItem.EnsureVisible
400     End If
        '
410   On Error GoTo EH
      '
  
    '270   Set BranchData = Nothing
      'Set Branchs = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "FillBranches", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub


Private Function GetGroupQuery() As String
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim sSQL As String
      '
110   DoEvents
      '
120   rsGroupCategories.MoveFirst
130   rsGroupCategories.Find "RecID = " & lstGroups.ItemData(lstGroups.ListIndex), , adSearchForward
      '
140   If Not rsGroupCategories.eof Then
      '
150     sSQL = "SELECT   TContact.ID, TContact.ContactType, TContact.Status," & _
        " TContact.FirstName" & _
        ",TContact.LastName" & _
        ",TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & _
        ",TCompany.Name AS Company " & _
        ",TContact.State, TContact.Status, TContact.AuthDate, TContact.Phone1 " & _
        ",TContact.VersionShipped, TContact.PVVersionShipped " & _
        ", TContact.Source, DATEADD(day,([AuthDays] - DateDiff(Day, [AuthDate], GETDATE())),TContact.AuthDate) AS ExpDate" & _
        " FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID " & _
        "WHERE " & ConvertFormula(rsGroupCategories!Formula) '& " ORDER BY TContact.State"
        '
160    GetGroupQuery = sSQL
170   End If
      '<EhFooter>
      '
      Exit Function
      '
EH:
      ErrorMgr.Raise "FContact", "GetGroupQuery", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Private Sub lvwPMContacts_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo EH
  If Not lvwPMContacts.SelectedItem Is Nothing Then
    LoadContact CLng(Mid$(lvwPMContacts.SelectedItem.Key, 3)), False
  End If
  '
 ' LoadCustGroups
  Exit Sub
EH:
  MsgBox "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & " in FContact: lvwPMContacts_KeyUp."
End Sub

Private Sub SortListView(ByRef pListView As ListView, ByVal piColumIndex As Integer)
On Error GoTo EH
  '
  sOrder = Not sOrder
  '
  pListView.SortKey = piColumIndex - 1
  '
  Select Case piColumIndex - 1
  ' Case 1:
      'Use sort routine to sort by date
      'lstCustGroups.Sorted = False
      'SendMessage lstCustGroups.hWnd, _
      '            LVM_SORTITEMS, _
      '            lstCustGroups.hWnd, _
      '            ByVal FARPROC(AddressOf CompareDates)
      Case 2:
        'Use sort routine to sort by value
        pListView.Sorted = False
        SendMessage pListView.hWnd, _
                   LVM_SORTITEMS, _
                   pListView.hWnd, _
                   ByVal FARPROC(AddressOf CompareValues)
      Case Else:
        'Use default sorting to sort the items in the list
        'lvwContact.SortKey = 0
        pListView.SortOrder = Abs(sOrder) '=Abs(Not lstCustGroups.SortOrder = 1)
        pListView.Sorted = True
   End Select
   '
   Exit Sub
EH:
 MsgBox Err.Description & " in pListView_ColumnClick."
End Sub

Private Sub LoadCustGroups()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim GroupList As CGroupList
      'Dim GroupListData As CGroupListData
110   Dim GroupLists As CGroupListDatas
120   Dim GroupListLink As CGroupListLink
130   Dim x As Integer
140   Dim sKey As String
      '
150   Set GroupList = New CGroupList
      'Set GroupListData = New CGroupListData
160   Set GroupLists = New CGroupListDatas
170   Set GroupListLink = New CGroupListLink
      '
180   GroupList.LoadCollection GroupLists
      '
190   lstCustGroups.ListItems.Clear
      '
200   For x = 1 To GroupLists.Count
210     sKey = "A" & GroupLists.Item(x).ID
220     lstCustGroups.ListItems.Add , sKey, GroupLists.Item(x).ListName
230     lstCustGroups.ListItems.Item(sKey).Checked = GroupListLink.CheckContact(GroupLists.Item(x).ID, ContactData.ID)
240   Next
      '
250   SelectNone
      '
260   Set GroupList = Nothing
      'Set GroupListData = Nothing
270   Set GroupLists = Nothing
280   Set GroupListLink = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "LoadCustGroups", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub LoadCustGroupsList()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim GroupList As CGroupList
      'Dim GroupListData As CGroupListData
110   Dim GroupLists As CGroupListDatas
120   Dim GroupListLink As CGroupListLink
130   Dim x As Integer
140   Dim y As Integer
      '
150   Set GroupList = New CGroupList
      'Set GroupListData = New CGroupListData
160   Set GroupLists = New CGroupListDatas
170   Set GroupListLink = New CGroupListLink
      '
180   GroupList.LoadCollection GroupLists
      '
190   y = lstGroups.ListCount
      '
200   For x = 1 To GroupLists.Count
210     lstGroups.AddItem ("** " & GroupLists.Item(x).ListName)
220     lstGroups.ItemData(x + y - 1) = (GroupLists.Item(x).ID)
230   Next
      '
240   Set GroupList = Nothing
      'Set GroupListData = Nothing
250   Set GroupLists = Nothing
260   Set GroupListLink = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "LoadCustGroupsList", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Function GetCustGroupQuery(plCustID As Long) As String
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim sSQL As String
      '
110   sSQL = "SELECT   TContact.ID, TContact.ContactType, TContact.Status," & _
      " TContact.FirstName" & _
      ",TContact.LastName" & _
      ",TContact.AuthDays-DateDiff(d,TContact.AuthDate, GETDATE()) as Days " & _
      ",TCompany.Name AS Company " & _
      ",TContact.State, TContact.Status, TContact.AuthDate, TContact.Phone1 " & _
      ",TContact.VersionShipped, TContact.PVVersionShipped " & _
      ", TContact.Source, DATEADD(day,([AuthDays] - DateDiff(Day, [AuthDate], GETDATE())),TContact.AuthDate) AS ExpDate " & _
      "FROM TGroupListLink LEFT OUTER JOIN TContact ON TGroupListLink.ContactID = TContact.ID " & _
      "RIGHT OUTER JOIN TCompany ON TContact.CompanyID = TCompany.ID " & _
      "WHERE (TGroupListLink.ListID = " & plCustID & ")"
      '
120    GetCustGroupQuery = sSQL
      '
      '<EhFooter>
      '
      Exit Function
      '
EH:
      ErrorMgr.Raise "FContact", "GetCustGroupQuery", Err.Number, Err.Description, Erl
      '</EhFooter>
End Function

Private Sub lstGroups_Click()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim Report As New CReport
110   Dim CList As New CContactList
      '
120   If lstGroups.ListIndex <> -1 Then
130     If Left(lstGroups.Text, 3) = "** " Then
140       Report.ListQuery = GetCustGroupQuery(lstGroups.ItemData(lstGroups.ListIndex))
150     Else
160       Report.ListQuery = GetGroupQuery
170     End If
180       Report.Rtype = FromList
          '
190       lvwPMContacts.ListItems.Clear
200       lvwPMContacts.ColumnHeaders.Clear
210       CList.FillList lvwPMContacts, Report.rsReport
    
220   End If
      '
230   Set Report = Nothing
240   Set CList = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "lstGroups_Click", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub SaveCustGroups()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim GroupList As CGroupList
110   Dim GroupListData As CGroupListData
120   Dim GroupListLink As CGroupListLink
130   Dim x As Integer
      '
140   Set GroupList = New CGroupList
150   Set GroupListData = New CGroupListData
160   Set GroupListLink = New CGroupListLink
      '
170   For x = 1 To lstCustGroups.ListItems.Count
180     If GroupList.Load(GroupListData, GetIDFromKey(lstCustGroups.ListItems(x).Key)) Then
190       If lstCustGroups.ListItems(x).Checked = True Then
200         If GroupListLink.CheckContact(GroupListData.ID, ContactData.ID) = False Then
210           GroupListLink.AddContact ContactData.ID, GroupListData.ID
220         End If
230       Else
240         If GroupListLink.CheckContact(GroupListData.ID, ContactData.ID) = True Then
250           GroupListLink.DelContact ContactData.ID, GroupListData.ID
260         End If
270       End If
280     End If
290   Next
      '
300   SelectNone
      '
310   Set GroupList = Nothing
320   Set GroupListData = Nothing
330   Set GroupListLink = Nothing
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "SaveCustGroups", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub SelectNone()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim x As Integer
      '
110   For x = 1 To lstCustGroups.ListItems.Count
120     If lstCustGroups.ListItems(x).Selected = True Then lstCustGroups.ListItems(x).Selected = False
130   Next
      '<EhFooter>
      '
      Exit Sub
      '
EH:
      ErrorMgr.Raise "FContact", "SelectNone", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub SetStar(Index As Integer)
  Dim x As Integer
    For x = 0 To 5
      StarPicture(x).Picture = Me.ImageList2.ListImages(1).Picture
    Next x
    '
    For x = 1 To Index
      StarPicture(x).Picture = Me.ImageList2.ListImages(2).Picture
    Next x
End Sub



