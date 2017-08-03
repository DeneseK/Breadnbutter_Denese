VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{08769121-33BD-11D3-BD95-B44CFE3A3C4B}#1.0#0"; "VSSPELL6.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Note"
   ClientHeight    =   6780
   ClientLeft      =   4305
   ClientTop       =   2850
   ClientWidth     =   7785
   Icon            =   "FNote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7785
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeAuthRevise 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   2040
      TabIndex        =   67
      Top             =   6120
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtReviseAuthDays 
         Height          =   285
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   9
         Top             =   210
         Width           =   495
      End
      Begin MSComCtl2.DTPicker mskReviseAuthDate 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   61472769
         CurrentDate     =   38426
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Auth Days:"
         Height          =   255
         Left            =   2400
         TabIndex        =   69
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Auth Date:"
         Height          =   255
         Left            =   0
         TabIndex        =   68
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fmeSale 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   8160
      TabIndex        =   64
      Top             =   240
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox txtGraceDays 
         DataField       =   "VersionShipped"
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Tag             =   "1"
         Top             =   120
         Width           =   555
      End
      Begin VB.TextBox txtPendingDays 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Grace Period:"
         Height          =   285
         Index           =   9
         Left            =   2040
         TabIndex        =   66
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pending Days:"
         Height          =   375
         Left            =   0
         TabIndex        =   65
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkSticky 
      Caption         =   "Sticky"
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cboProduct 
      Height          =   315
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   450
      Width           =   1815
   End
   Begin VB.Frame frmCurrentDays 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   3000
      TabIndex        =   61
      Top             =   3960
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Days:"
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lblCurrentDays 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   1200
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame fmeUserNum 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Date Restoration"
      Height          =   855
      Left            =   6480
      TabIndex        =   45
      Top             =   3000
      Visible         =   0   'False
      Width           =   6105
      Begin VB.TextBox txtUserNum 
         Height          =   285
         Left            =   5160
         TabIndex        =   18
         Top             =   90
         Width           =   615
      End
      Begin VB.TextBox txtUserNumCode 
         Height          =   285
         Left            =   870
         MaxLength       =   31
         TabIndex        =   17
         Top             =   90
         Width           =   2880
      End
      Begin VB.TextBox txtUserNumKey 
         Height          =   285
         Left            =   870
         Locked          =   -1  'True
         MaxLength       =   31
         TabIndex        =   19
         Top             =   450
         Width           =   2895
      End
      Begin Threed.SSCommand cmdUserNumGen 
         Height          =   315
         Left            =   3870
         TabIndex        =   20
         Top             =   450
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FNote.frx":030A
         Caption         =   "Generate "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Users Allowed:"
         DataField       =   "HRI"
         Height          =   225
         Left            =   3960
         TabIndex        =   48
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Code:"
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Key:"
         Height          =   195
         Left            =   0
         TabIndex        =   46
         Top             =   480
         Width           =   705
      End
   End
   Begin VB.Frame fmeDeathorization 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Caption         =   "Deauthorization"
      Height          =   1155
      Left            =   7800
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   5325
      Begin VB.TextBox txtConfCode 
         Height          =   285
         Left            =   1020
         MaxLength       =   45
         TabIndex        =   12
         Top             =   30
         Width           =   2715
      End
      Begin VB.TextBox txtSiteDays 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   15
         Top             =   750
         Width           =   450
      End
      Begin Threed.SSCommand cmdDecrypt 
         Height          =   315
         Left            =   3840
         TabIndex        =   16
         Top             =   30
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FNote.frx":0464
         Caption         =   "Decrypt "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin TDBTime6Ctl.TDBTime txtClientTime 
         Height          =   285
         Left            =   2910
         TabIndex        =   14
         Top             =   390
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   503
         Caption         =   "FNote.frx":05BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "FNote.frx":062A
         Spin            =   "FNote.frx":067A
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn:ss"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn:ss"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__:__:__"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37189
      End
      Begin TDBDate6Ctl.TDBDate txtClientDate 
         Height          =   285
         Left            =   1020
         TabIndex        =   13
         Top             =   390
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         Calendar        =   "FNote.frx":06A2
         Caption         =   "FNote.frx":07BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FNote.frx":0826
         Keys            =   "FNote.frx":0844
         Spin            =   "FNote.frx":08A2
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37189
         CenturyMode     =   0
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Date:"
         Height          =   240
         Left            =   0
         TabIndex        =   40
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Conf. Code:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   39
         Top             =   60
         Width           =   975
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Time:"
         Height          =   240
         Left            =   2160
         TabIndex        =   38
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Days:"
         Height          =   240
         Left            =   0
         TabIndex        =   37
         Top             =   780
         Width           =   735
      End
   End
   Begin VB.Frame fmeDateRestore 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Caption         =   "Date Restoration"
      Height          =   855
      Left            =   6600
      TabIndex        =   41
      Top             =   4440
      Visible         =   0   'False
      Width           =   6105
      Begin VB.TextBox txtRestKey 
         Height          =   285
         Left            =   870
         Locked          =   -1  'True
         MaxLength       =   31
         TabIndex        =   23
         Top             =   450
         Width           =   2895
      End
      Begin VB.TextBox txtRestCode 
         Height          =   285
         Left            =   870
         MaxLength       =   31
         TabIndex        =   21
         Top             =   90
         Width           =   2880
      End
      Begin Threed.SSCommand cmdRestGen 
         Height          =   315
         Left            =   3870
         TabIndex        =   24
         Top             =   450
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FNote.frx":08CA
         Caption         =   "Generate "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin TDBDate6Ctl.TDBDate txtRestDate 
         Height          =   285
         Left            =   4740
         TabIndex        =   22
         Top             =   90
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         Calendar        =   "FNote.frx":0A24
         Caption         =   "FNote.frx":0B3C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FNote.frx":0BA8
         Keys            =   "FNote.frx":0BC6
         Spin            =   "FNote.frx":0C24
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37189
         CenturyMode     =   0
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Key:"
         Height          =   195
         Left            =   0
         TabIndex        =   44
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Code:"
         Height          =   255
         Left            =   0
         TabIndex        =   43
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Date:"
         DataField       =   "HRI"
         Height          =   225
         Left            =   3960
         TabIndex        =   42
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.Frame fmeAuth 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Authorization"
      Height          =   1155
      Left            =   6840
      TabIndex        =   49
      Top             =   5520
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton optWrite 
         Caption         =   "Extension"
         Height          =   225
         Index           =   1
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   60
         Width           =   1005
      End
      Begin VB.OptionButton optWrite 
         Caption         =   "New"
         Height          =   225
         Index           =   0
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   60
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.TextBox txtSiteCode 
         Height          =   285
         Left            =   870
         MaxLength       =   31
         TabIndex        =   28
         Top             =   30
         Width           =   2880
      End
      Begin VB.TextBox txtDays 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   870
         MaxLength       =   4
         TabIndex        =   27
         Text            =   "30"
         Top             =   360
         Width           =   450
      End
      Begin VB.TextBox txtSiteKey 
         Height          =   285
         Left            =   870
         Locked          =   -1  'True
         MaxLength       =   31
         TabIndex        =   32
         Top             =   720
         Width           =   2895
      End
      Begin Threed.SSCommand cmdGen 
         Height          =   315
         Left            =   3870
         TabIndex        =   33
         Top             =   720
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FNote.frx":0C4C
         Caption         =   "Generate "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin TDBDate6Ctl.TDBDate txtSiteDate 
         Height          =   285
         Left            =   4620
         TabIndex        =   31
         Top             =   360
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         Calendar        =   "FNote.frx":0DA6
         Caption         =   "FNote.frx":0EBE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FNote.frx":0F2A
         Keys            =   "FNote.frx":0F48
         Spin            =   "FNote.frx":0FA6
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37189
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate txtExpDate 
         Height          =   285
         Left            =   2220
         TabIndex        =   29
         Top             =   360
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         Calendar        =   "FNote.frx":0FCE
         Caption         =   "FNote.frx":10E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FNote.frx":1152
         Keys            =   "FNote.frx":1170
         Spin            =   "FNote.frx":11CE
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37189
         CenturyMode     =   0
      End
      Begin Threed.SSCommand cmdExpDate 
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   360
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FNote.frx":11F6
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp. Date:"
         DataField       =   "HRI"
         Height          =   225
         Left            =   1440
         TabIndex        =   54
         Top             =   390
         Width           =   795
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Date:"
         DataField       =   "HRI"
         Height          =   225
         Index           =   4
         Left            =   3840
         TabIndex        =   53
         Top             =   390
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Code:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   52
         Top             =   60
         Width           =   795
      End
      Begin VB.Label lblUnlockLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Days:"
         Height          =   240
         Left            =   0
         TabIndex        =   51
         Top             =   390
         Width           =   525
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Key:"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   50
         Top             =   750
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdSpell 
      Height          =   375
      Left            =   2040
      Picture         =   "FNote.frx":1790
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   1560
      Width           =   435
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Tag             =   "3"
      Top             =   810
      Width           =   2685
   End
   Begin VB.TextBox txtResults 
      DataField       =   "Results"
      Height          =   5505
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Tag             =   "3"
      Top             =   1200
      Width           =   5085
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6465
      Picture         =   "FNote.frx":18DA
      TabIndex        =   34
      Top             =   15
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   6465
      Picture         =   "FNote.frx":1D1C
      TabIndex        =   35
      Top             =   600
      Width           =   1275
   End
   Begin VB.ListBox lstTypes 
      Height          =   6690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker mskDate 
      Height          =   330
      Left            =   2640
      TabIndex        =   1
      Top             =   45
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   61472769
      CurrentDate     =   38231
   End
   Begin MSComCtl2.DTPicker mskTime 
      Height          =   330
      Left            =   4920
      TabIndex        =   2
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   61472770
      CurrentDate     =   38231
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   225
      Left            =   1920
      TabIndex        =   60
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      Height          =   285
      Index           =   7
      Left            =   4200
      TabIndex        =   59
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   58
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   57
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   225
      Index           =   1
      Left            =   1920
      TabIndex        =   56
      Top             =   1200
      Width           =   645
   End
   Begin VB.Label lblUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   450
      Width           =   1755
   End
   Begin VSSPELL6LibCtl.VSSpell vspNote 
      Left            =   2160
      Top             =   1770
      _ConvInfo       =   1
      BeginProperty DialogFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MainDictFile    =   ""
      Suggest         =   -1  'True
      CustomDictFile  =   "vssp_ae.dct"
      CommonWordCache =   1
      BadWordDialog   =   1
      DialogTitle     =   ""
      OptionBtnCaption=   ""
      OptionBtnVisible=   0   'False
      DialogTop       =   0
      DialogLeft      =   0
      IgnoreWithNumbers=   0   'False
      IgnoreInUpperCase=   0   'False
      IgnoreInMixedCase=   0   'False
      AddBtnVisible   =   -1  'True
      DontCorrectText =   0   'False
      CheckSpelling   =   -1  'True
      CustomDictFile2 =   ""
      CustomDictFile3 =   ""
      CustomDictFile4 =   ""
      CustomDictFile5 =   ""
      HelpBtnVisible  =   0   'False
      WhichCustomDict =   1
      TypingErrorAction=   3
      UnderlineColor  =   4
      UnderlineStyle  =   255
   End
End
Attribute VB_Name = "FNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Enum eReportType
'  DateRestoration
'  Reauthorization
'  Deathorization
'  SecondAuthorization
'  PaidAuthorization
'  UserLimitChanged
'  Sale
'  EvalAddition
'  EvalAuthorized
'  Normal
'End Enum
'
'Private NoteType As eReportType
Dim bNewNote As Boolean
Dim lProductID As Integer
Dim Contact    As New CContact
Dim ContactData As CContactData
Dim Event1 As New CEvent
Dim EventData As CEventData
Dim bSaved As Boolean
Dim bAuthEvent As Boolean
Dim ProductData As New CProductData
Dim Product As New CProduct
'
Private Declare Function pp_ctcodes Lib "TrgLib32.dll" (ByVal code As Long, ByVal cenum As Long, ByVal computer As Long, ByVal seed As Long) As Long
Private Declare Function pp_nencrypt Lib "KeyLib32.DLL" (ByVal Number As Long, ByVal seed As Long) As Long
Private Declare Sub pp_cedate Lib "TrgLib32.dll" (ByVal cenum As Long, ByRef month As Long, ByRef day As Long, ByRef year As Long)

Private Sub cboProduct_Change()
  SwitchProduct
  '
  If txtDays.Visible = True Then SuggestDays
End Sub

Private Sub SwitchProduct()
  Select Case cboProduct.Text
    Case "PowerClaim PV"
      Product.Load ProductData, 2
      '
      Me.BackColor = &HFF8080
    Case "PowerClaim XML"
      Product.Load ProductData, 1
      '
      Me.BackColor = vbButtonFace
  End Select
  '
End Sub

Private Sub cboProduct_Click()
  SwitchProduct
End Sub

Private Sub cmdCancel_Click(Index As Integer)
  Unload Me
End Sub

Private Sub cmdExpDate_Click()
  txtExpDate.Value = FDatePick.DateText(txtExpDate.Value)
End Sub

Private Sub cmdSave_Click()
  Dim Employee As New CEmployee
  '
  If lstTypes.Text = "" Then
    MsgBox "Please select type."
    Exit Sub
  End If
  '
  If bAuthEvent Then
    If txtSubject.Text <> vbNullString Then
      CommitAuth
    Else
      MsgBox "Please complete authorization before saving"
      Exit Sub
    End If
  End If
  '
  If lstTypes.Text = "Authorization Revision" Then
    CommitAuthorizationRevision
  End If
  '
  If lstTypes.Text = "Sale" Or lstTypes.Text = "Sale Revision" Then
    CommitSale
  End If

  'If Not bOk Then
 '   Exit Sub
 ' End If
  '
  EventData.CustRecID = ContactData.ID
  EventData.EventDate = mskDate.Value
  EventData.EventSubject = txtSubject.Text
  EventData.EventResults = txtResults.Text
  EventData.EventUser = lblUser.Caption
  EventData.EventTime = mskTime.Value
  EventData.EventType = lstTypes.Text
  EventData.Sticky = IIf((chkSticky.Value = 1), True, False)
  '
  Select Case cboProduct.Text
    Case "PowerClaim XML"
      EventData.ProductID = 1
    Case "PowerClaim PV"
      EventData.ProductID = 2
  End Select
  '
  If Employee.InGroup(StrUser, "Support") = True Then
    EventData.OpenCall = True
  Else
    EventData.OpenCall = False
  End If
  '
  If bNewNote = True Then
    Event1.Save EventData, True
  Else
    Event1.Save EventData, False
  End If
  '
  bSaved = True
  '
  Set Employee = Nothing
  '
  Unload Me
End Sub

Private Function CommitAuthorizationRevision() As Boolean
  If Not ContactData Is Nothing Then
      If ContactData.ID > 0 Then
        txtSubject.Text = txtReviseAuthDays.Text
        '
        Select Case cboProduct.Text
          Case "PowerClaim XML"
            ContactData.AuthDate = mskReviseAuthDate.Value
            ContactData.AuthDays = txtReviseAuthDays.Text
          Case "PowerClaim PV"
            ContactData.PVAuthDate = mskReviseAuthDate.Value
            ContactData.PVAuthDays = txtReviseAuthDays.Text
        End Select
        '
        Contact.Save ContactData, False
      End If
    End If
End Function

Private Function CommitSale() As Boolean
  If Not ContactData Is Nothing Then
      If ContactData.ID > 0 Then
        If UCase(ContactData.MailState) = "KY" Then
          MsgBox "Remember to charge Sales Tax!", vbInformation, "State Sales Tax"
        End If
        txtSubject.Text = txtPendingDays.Text
        '
        ContactData.Status = "Customer"
        '
        Select Case cboProduct.Text
          Case "PowerClaim XML"
            ContactData.SaleDate = mskDate.Value
            ContactData.SaleDays = CInt(nnNum(Me.txtPendingDays.Text))
            '
            If ContactData.AuthRemaining > 0 Then
              ContactData.GraceDays = ContactData.AuthRemaining + CInt(nnNum(Me.txtGraceDays.Text))
            Else
              ContactData.GraceDays = CInt(nnNum(Me.txtGraceDays.Text))
            End If
            '
            ContactData.OnlineAuths = 2
          Case "PowerClaim PV"
            ContactData.PVSaleDate = mskDate.Value
            ContactData.PVSaleDays = CInt(nnNum(Me.txtPendingDays.Text))
            '
            If ContactData.PVAuthRemaining > 0 Then
              ContactData.PVGraceDays = ContactData.PVAuthRemaining + CInt(nnNum(Me.txtGraceDays.Text))
            Else
              ContactData.PVGraceDays = CInt(nnNum(Me.txtGraceDays.Text))
            End If
            '
            ContactData.PVOnlineAuths = 2
        End Select
        '
        Contact.Save ContactData, False
      End If
  End If
  '
  
End Function

Private Function CommitAuth() As Boolean
  'If iAuthDays <= 0 Then
   ' MsgBox "Please enter the number of days before continuing.", vbInformation, "Authorization"
    'Exit Sub
  'Else
    If Not ContactData Is Nothing Then
      If ContactData.ID > 0 Then
        '
        Select Case lstTypes.Text
        Case "Paid Authorization"
          ContactData.Status = "Customer"
          ContactData.AuthStatus = "Purchase"
        Case "Eval Authorized"
          ContactData.AuthStatus = "Evaluation"
        Case "Eval Addition"
          ContactData.AuthStatus = "Extended Evaluation"
        End Select
        '
        Select Case cboProduct.Text
        Case "PowerClaim XML"
          Select Case lstTypes.Text
          Case "Paid Authorization"
            ContactData.Status = "Customer"
            ContactData.AuthStatus = "Purchase"
          Case "Eval Authorized"
            ContactData.AuthStatus = "Evaluation"
          Case "Eval Addition"
            ContactData.AuthStatus = "Extended Evaluation"
          End Select
          '
          ContactData.GraceDays = 0
          ContactData.SaleDate = 0
          ContactData.SaleDays = 0
          ContactData.AuthDays = CInt(nnNum(txtSubject.Text))
          ContactData.AuthRemaining = DateDiff("d", Now, DateAdd("d", nnNum(txtSubject.Text), mskDate.Value))
          ContactData.AuthDate = mskDate.Value
        Case "PowerClaim PV"
          Select Case lstTypes.Text
          Case "Paid Authorization"
            ContactData.Status = "Customer"
            ContactData.PVAuthStatus = "Purchase"
          Case "Eval Authorized"
            ContactData.PVAuthStatus = "Evaluation"
          Case "Eval Addition"
            ContactData.PVAuthStatus = "Extended Evaluation"
          End Select
          '
          ContactData.PVGraceDays = 0
          ContactData.PVSaleDate = 0
          ContactData.PVSaleDays = 0
          ContactData.PVAuthDays = CInt(nnNum(txtSubject.Text))
          ContactData.PVAuthRemaining = DateDiff("d", Now, DateAdd("d", CDbl(txtSubject.Text), mskDate.Value))
          ContactData.PVAuthDate = mskDate.Value
        End Select
        '
        'ContactData.Status = "Customer"
        '
        Contact.Save ContactData, False
      End If
    End If
  'End If
End Function

Private Function CommitAddition() As Boolean

End Function


Private Sub cmdSpell_Click()
On Error GoTo cmdSpell_Click_EH
  '
  cmdSpell.Enabled = False 'avoid reentrancy
  Screen.MousePointer = vbHourglass
  '
  With vspNote
    'lSpellStart = 0
    .Clear
    .Text = txtResults.Text
    .CheckText
    '
    txtResults.Text = .Text
  End With
   '
  Screen.MousePointer = vbDefault
  cmdSpell.Enabled = True
   '
  Exit Sub
cmdSpell_Click_EH:
  Screen.MousePointer = vbDefault
  cmdSpell.Enabled = True
End Sub

'Private Sub cmdSupport_Click2(Index As Integer)
'  Dim iSupportActID As Long
'  Dim iCaseID As Long
'  Dim Employee As New CEmployee
'  On Error GoTo EH
'  '
'  If Index = 0 Then 'Submit
'    If fAuthorizing Then
'      Dim iAuthDays As Integer
'      iAuthDays = CInt(nnNum(txtSubject.Text))
'      '
'      If iAuthDays <= 0 Then
'        MsgBox "Please enter the number of days before continuing.", vbInformation, "Authorization"
'        Exit Sub
'      Else
'        mskAuthDate.Text = mskDate.Text
'        '
'        Select Case cboType.Text
'        Case "Eval Authorized"
'          cboAuthStatus.Text = "Evaluation"
'        Case "Eval Addition"
'          cboAuthStatus.Text = "Extended Evaluation"
'        Case "Sale"
'          cboStatus.Text = "Customer"
'          cboAuthStatus.Text = "Purchase"
'        End Select
'        '
'        txtAuthDays.Text = txtSubject.Text
'        lblAuthRemaining.Caption = DateDiff("d", Now, DateAdd("d", CDbl(txtSubject.Text), mskDate.DateValue))
'        lblExpires.Caption = Date + CLng(lblAuthRemaining.Caption)
'      End If
'    End If
'    '
'    'Dim test As File
'    'test.
'
'  End If
'  '
'  fAuthorizing = False
'  fSupport = False
'  Me.DataContact.Save
'  '
'  Set Employee = Nothing
'  Exit Sub
'EH:
'  MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FCustomer.cmdOther_Click.", vbCritical, "Error"
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set ContactData = Nothing
  Set EventData = Nothing
  Set Contact = Nothing
  Set Event1 = Nothing
End Sub

Private Sub txtExpDate_Change()
  If txtExpDate.Value = Null Then
    txtDays.Text = "0"
  Else
    txtDays.Text = DateDiff("d", Now, nnNum(txtExpDate.Value))
  End If
End Sub

Private Sub txtReviseAuthDays_Gotfocus()
  InputNumber.Setup txtReviseAuthDays, NumberTypeInteger
End Sub

Private Sub txtSiteKey_GotFocus()
  On Error GoTo ErrorHandler
  '
  SelectText txtSiteKey
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtSiteKey.GotFocus"
End Sub

Private Sub txtSiteDays_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrorHandler
  '
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> 32 Then
      If KeyAscii <> vbKeyBack Then KeyAscii = 0
    End If
  End If
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtSiteDays.KeyPress"
End Sub

Private Sub txtSiteCode_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrorHandler
  '
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> 32 Then '\\ Space
      If KeyAscii <> vbKeyBack Then '\\ Backspace
        If KeyAscii <> 22 Then '\\ CTRL-V: Paste
          KeyAscii = 0
        End If
      End If
    End If
  End If
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtSiteCode.KeyPress"
End Sub

Private Sub txtConfCode_Change()
  On Error GoTo ErrorHandler
  '
  txtSiteDays.Text = vbNullString
  txtClientDate.Text = "__/__/____"
  txtClientTime.Text = "__:__:____"
  txtClientDate.BackColor = vbWhite
  txtClientDate.ForeColor = vbBlack
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtConfCode.Change"
End Sub

Private Sub txtConfCode_GotFocus()
  On Error GoTo ErrorHandler
  '
  SelectText txtConfCode
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtConfCode.GotFocus"
End Sub

Private Sub txtConfCode_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrorHandler
  '
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> 32 Then '\\ Space
      If KeyAscii <> vbKeyBack Then '\\ Backspace
        If KeyAscii <> 22 Then '\\ CTRL-V: Paste
          KeyAscii = 0
        End If
      End If
    End If
  End If
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtConfCode.KeyPress"
End Sub

Private Sub txtDays_Change()
  On Error GoTo ErrorHandler
  '
  txtSiteKey.Text = vbNullString
  'txtExpDate.Text = "__/__/____"
  txtExpDate.Value = DateAdd("d", Val(txtDays.Text), Now)
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtDays.Change"
End Sub

Private Sub txtDays_GotFocus()
  On Error GoTo ErrorHandler
  '
  SelectText txtDays
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtDays.GotFocus"
End Sub

Private Sub txtRestCode_Change()
  On Error GoTo ErrorHandler
  '
  txtRestKey.Text = vbNullString
  txtRestDate.BackColor = vbWhite
  txtRestDate.ForeColor = vbBlack
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtRestCode.Change"
End Sub

Private Sub txtRestCode_GotFocus()
  On Error GoTo ErrorHandler
  '
  SelectText txtRestCode
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtRestCode.GotFocus"
End Sub

Private Sub txtRestCode_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrorHandler
  '
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> 32 Then '\\ Space
      If KeyAscii <> vbKeyBack Then '\\ Backspace
        If KeyAscii <> 22 Then '\\ CTRL-V: Paste
          KeyAscii = 0
        End If
      End If
    End If
  End If
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtRestCode.KeyPress"
End Sub

Private Sub txtSiteCode_Change()
  On Error GoTo ErrorHandler
  '
  txtDays_Change
  txtSiteDate.Text = "__/__/____"
 ' txtExpDate.Text = "__/__/____"
  txtSiteKey.Text = vbNullString
  txtSiteDate.BackColor = vbWhite
  txtSiteDate.ForeColor = vbBlack
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtSiteCode.Change"
End Sub

Private Sub txtSiteCode_GotFocus()
  On Error GoTo ErrorHandler
  '
  SelectText txtSiteCode
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtSiteCode.GotFocus"
End Sub

Public Sub SelectText(pctrlCur As Control)
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim fAlt As Boolean
  '
  With pctrlCur
    .SelStart = 0
    .SelLength = Len(pctrlCur.DisplayText)
    If fAlt = True Then .SelLength = Len(pctrlCur)
  End With
  '
  Exit Sub
  '
ErrorHandler:
  If Err.Number = 438 Then '\\ Object Doesn't Support Property Or Method
    fAlt = True
    Resume Next
  End If
  '
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.General.SelectText"
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
  On Error GoTo ErrorHandler
  '
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> 32 Then
      If KeyAscii <> vbKeyBack Then KeyAscii = 0
    End If
  End If
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtDays.KeyPress"
End Sub

Private Sub cmdUserNumGen_Click()
On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iDelPos             As Integer
  Dim lKeyCode            As Long
  Dim lUsers               As Long
  Dim lDateDay            As Long
  Dim lDateMonth          As Long
  Dim lDateYear           As Long
  Dim lSiteKey1           As Long
  Dim lSiteKey2           As Long
  Dim sSiteCode           As String
  Dim sSiteCodeCompacted  As String
  Dim sSiteCode1          As String
  Dim sSiteCode2          As String
  '
  Dim rsLog As New ADODB.Recordset
  '
  If ValidateLicense <> CDbl(Date - dSecVar) Then Exit Sub
  '
  'If GeneralDataSpecified = False Then Exit Sub
  '
  sSiteCode = Trim(txtUserNumCode)
  '
  If InStr(1, sSiteCode, " ", vbBinaryCompare) <= 0 Then
    MsgBox "The site code you entered does not contain a space.", vbCritical + vbOKOnly, "Error: Site Code Does Not Contain Space"
    txtUserNumCode.SetFocus
    Exit Sub
  End If
  '
  sSiteCodeCompacted = Replace(sSiteCode, " ", vbNullString)
  lUsers = nnNum(txtUserNum.Text)
  'lDays = CLng(txtDays.Text)
  lKeyCode = 3 'IIf(optWrite(0) = True, 1, 2)
  '
  If sSiteCodeCompacted = vbNullString Then
    MsgBox "Please enter a valid site code.", vbCritical + vbOKOnly, "Error: Site Code Not Specified"
    txtUserNumCode.SetFocus
    Exit Sub
  End If
  '
  If IsNumeric(sSiteCodeCompacted) = False Then
    MsgBox "Site codes cannot contain non-numeric characters. Please enter a valid site code.", vbCritical + vbOKOnly, "Error: Non-Numeric Site Code Specified"
    txtUserNumCode.SetFocus
    Exit Sub
  End If
  '
  If lUsers < 1 Then
    MsgBox "Please enter the number of users to authorize.", vbCritical + vbOKOnly, "Error: Number of Days Not Specified"
    txtUserNum.SetFocus
    Exit Sub
  End If
  '
  iDelPos = InStr(1, sSiteCode, " ", vbBinaryCompare)
  If iDelPos > 0 Then
    sSiteCode2 = Left$(sSiteCode, iDelPos - 1)
    sSiteCode1 = Trim(Right$(sSiteCode, Len(sSiteCode) - iDelPos))
  Else
    sSiteCode1 = sSiteCode
  End If
  '
  pp_cedate CLng(sSiteCode1), lDateMonth, lDateDay, lDateYear
  txtSiteDate = CStr(lDateMonth) & "/" & CStr(lDateDay) & "/" & CStr(lDateYear)
  lSiteKey1 = pp_ctcodes(lKeyCode, Val(sSiteCode1), Val(sSiteCode2), ProductData.Seed1) '173)
  lSiteKey2 = pp_nencrypt(lUsers, ProductData.Seed2) '236
  '
  txtUserNumKey.Text = CStr(lSiteKey1) & " " & CStr(lSiteKey2)
 ' If lKeyCode = 1 Then txtExpDate.Value = Date + lDays
  '
  If CDate(txtSiteDate.Text) <> Date Then
    txtSiteDate.BackColor = vbRed
    txtSiteDate.ForeColor = vbYellow
  '  MsgBox "The client's system date (" & txtSiteDate.Text & ") does not match HRI's date (" & txtHRIDate.Text & ")." & vbCrLf & vbCrLf & _
  '  "Please request the client to verify and -- if necessary -- correct his system's date. After a date correction, PowerClaim must be shut down and restarted in order to generate a correct site code." & vbCrLf & vbCrLf & _
  '  "If the client verifies that his date is set correctly, he must refresh the site code in the advanced licensing area.", vbCritical + vbOKOnly, "Error: Client System Date Does Not Match HRI Date"
  End If
  '
  '\\ Log Action
  rsLog.LockType = adLockPessimistic
  rsLog.CursorType = adOpenForwardOnly
  rsLog.Open "SELECT * from tbllog", cnMain
 ' rsLog.RecordCount
  With rsLog
    .AddNew
    .Fields("ID").Value = rsLog.RecordCount + 1
    .Fields("Company").Value = ContactData.ID & vbNullString
    .Fields("User").Value = ContactData.FirstName & " " & ContactData.LastName & vbNullString
    .Fields("Employee").Value = User.Name & vbNullString
    .Fields("ActionDateTime").Value = Date + Time & vbNullString
    .Fields("ActionType").Value = "User Count Change"
    .Fields("ActionSubType").Value = "None"
'    If optWrite(0) = True Then
'      .Fields("SiteExpirationDate").Value = txtExpDate.Value & vbNullString
'    End If
    '.Fields("SiteDays").Value = txtDays.Text & vbNullString
    .Fields("SiteCompID").Value = CLng(sSiteCode2) & vbNullString
    .Fields("SiteSessionID").Value = CLng(sSiteCode1) & vbNullString
    .Fields("SiteKey").Value = txtSiteKey.Text & vbNullString
    .Fields("SiteConfCode").Value = "N/A"
    .Fields("SiteDateTime").Value = txtSiteDate.Value & vbNullString
    .Fields("ProductID").Value = ProductData.ProductID
    .Fields("UsersAllowed").Value = nnNum(Me.txtUserNum.Text)
    .UpdateBatch
  End With
  '
  txtSubject.Text = txtUserNum.Text
  '
  rsLog.Close
  Set rsLog = Nothing
  '
  txtUserNumCode.SetFocus
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.cmdUserNumGen.Click"
End Sub

Private Sub cmdDecrypt_Click()
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim lDelPos   As Long
  Dim sCodes()  As String
  Dim sConfCode As String
  Dim sDateCode As String
  Dim sDaysCode As String
  '
  Dim rsLog As New ADODB.Recordset
  '
  If ValidateLicense <> CDbl(Date - dSecVar) Then Exit Sub
  '
  'If GeneralDataSpecified = False Then Exit Sub
  '
  sConfCode = Trim(txtConfCode.Text)
  '
  If sConfCode = vbNullString Then
    MsgBox "Please enter a valid confirmation code.", vbCritical + vbOKOnly, "Error: Confirmation Code Not Specified"
    txtConfCode.SetFocus
    Exit Sub
  End If
  '
  sCodes = Split(sConfCode)
  '
  If sCodes(0) = vbNullString Or sCodes(1) = vbNullString Or sCodes(2) = vbNullString Or sCodes(3) = vbNullString Then
    MsgBox "The confirmation code you entered is not valid. A valid confirmation code consists of four numbers separated by spaces.", vbCritical + vbOKOnly, "Error: Site Code Does Not Contain Space"
    txtConfCode.SetFocus
    Exit Sub
  End If
  
  sDateCode = sCodes(0) & "." & sCodes(1)
  sDaysCode = sCodes(2) & "." & sCodes(3)
  '
  If IsNumeric(sDateCode) = False Or IsNumeric(sDaysCode) = False Then
    MsgBox "Confirmation codes cannot contain non-numeric characters. Please enter a valid confirmation code.", vbCritical + vbOKOnly, "Error: Non-Numeric Confirmation Code Specified"
    txtConfCode.SetFocus
    Exit Sub
  End If
  '
  txtClientDate.Value = CDate(CDbl(sDateCode))
  txtClientTime.Value = CDate(CDbl(sDateCode))
  txtSiteDays.Text = Round(CDbl(sDaysCode) * 1.27)
  '
  If CDate(txtClientDate.Text) <> Date Then
    txtClientDate.BackColor = vbRed
    txtClientDate.ForeColor = vbYellow
    'MsgBox "The client's system date (" & txtClientDate.Text & ") does not match HRI's date (" & txtHRIDate.Text & ") Please request the client to verify and -- if necessary -- correct his system's date and deauthorize his license again. It is not necessary to shut down and restart PowerClaim in order to generate a correct confirmation code after the client corrects his system's date.", vbCritical + vbOKOnly, "Error: Client System Date Does Not Match HRI Date"
    'txtConfCode.SetFocus
  End If
  '
  '\\ Log Action
  rsLog.LockType = adLockPessimistic
  rsLog.CursorType = adOpenDynamic
  rsLog.Open "SELECT * from tbllog", cnMain
  With rsLog
    .AddNew
    .Fields("ID") = .RecordCount + 1
    .Fields("Company").Value = ContactData.ID & vbNullString
    .Fields("User").Value = ContactData.FirstName & " " & ContactData.LastName & vbNullString
    .Fields("Employee").Value = User.Name & vbNullString
    .Fields("ActionDateTime").Value = Date
    .Fields("ActionType").Value = "Deauthorization"
    .Fields("ActionSubType").Value = "N/A"
   ' .Fields("SiteExpirationDate").Value = "N/A" 'Date '+ txtSiteDays.Text
    .Fields("SiteDays").Value = txtSiteDays.Text & vbNullString
    .Fields("SiteCompID").Value = 0
    .Fields("SiteSessionID").Value = 0
    .Fields("SiteKey").Value = "N/A"
    .Fields("SiteConfCode").Value = txtConfCode.Text & vbNullString
    .Fields("SiteDateTime").Value = CDate(txtClientDate.Value) ' + CDate(txtClientTime.Value)
    .UpdateBatch
  End With
  '
  txtSubject.Text = 0
  '
  rsLog.Close
  Set rsLog = Nothing
  '
  Exit Sub
  '
ErrorHandler:
  If Err.Number = 9 Then '\\ Subscript out of Range
    MsgBox "The confirmation code you entered is not valid. A valid confirmation code consists of four numbers separated by spaces.", vbCritical + vbOKOnly, "Error: Site Code Does Not Contain Space"
    txtConfCode.SetFocus
    Exit Sub
  End If
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.cmdDecrypt.Click"
End Sub

Private Sub cmdRestGen_Click()
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iDelPos             As Integer
  Dim lSiteKey1           As Long
  Dim lDateDay            As Long
  Dim lDateMonth          As Long
  Dim lDateYear           As Long
  Dim sSiteCode           As String
  Dim sSiteCodeCompacted  As String
  Dim sSiteCode1          As String
  Dim sSiteCode2          As String
  '
  Dim rsLog As New ADODB.Recordset
  '
  If ValidateLicense <> CDbl(Date - dSecVar) Then Exit Sub
  '
  'If GeneralDataSpecified = False Then Exit Sub
  '
  sSiteCode = Trim(txtRestCode.Text)
  '
  If InStr(1, sSiteCode, " ", vbBinaryCompare) <= 0 Then
    MsgBox "The site code you entered does not contain a space.", vbCritical + vbOKOnly, "Error: Site Code Does Not Contain Space"
    txtSiteCode.SetFocus
    Exit Sub
  End If
  '
  sSiteCodeCompacted = Replace(sSiteCode, " ", vbNullString)
  '
  If sSiteCodeCompacted = vbNullString Then
    MsgBox "Please enter a valid site code.", vbCritical + vbOKOnly, "Error: Site Code Not Specified"
    txtRestCode.SetFocus
    Exit Sub
  End If
  '
  If IsNumeric(sSiteCodeCompacted) = False Then
    MsgBox "Site codes cannot contain non-numeric characters. Please enter a valid site code.", vbCritical + vbOKOnly, "Error: Non-Numeric Site Code Specified"
    txtRestCode.SetFocus
    Exit Sub
  End If
  '
  iDelPos = InStr(1, sSiteCode, " ", vbBinaryCompare)
  If iDelPos > 0 Then
    sSiteCode2 = Left$(sSiteCode, iDelPos - 1)
    sSiteCode1 = Trim(Right$(sSiteCode, Len(sSiteCode) - iDelPos))
  Else
    sSiteCode1 = sSiteCode
  End If
  '
  pp_cedate CLng(sSiteCode1), lDateMonth, lDateDay, lDateYear
  txtRestDate = lDateMonth & "/" & lDateDay & "/" & lDateYear
  lSiteKey1 = pp_ctcodes(7, CLng(sSiteCode1), CLng(sSiteCode2), ProductData.Seed1) '173)
  txtRestKey.Text = CStr(lSiteKey1)
  '
  If CDate(txtRestDate.Text) <> Date Then
    txtRestDate.BackColor = vbRed
    txtRestDate.ForeColor = vbYellow
  '  MsgBox "The client's system date (" & txtSiteDate.Text & ") does not match HRI's date (" & txtHRIDate.Text & ")." & vbCrLf & vbCrLf & _
  '  "Please request the client to verify and -- if necessary -- correct his system's date. After a date correction, PowerClaim must be shut down and restarted in order to generate a correct site code." & vbCrLf & vbCrLf & _
  '  "If the client verifies that his date is set correctly, he must refresh the site code in the advanced licensing area.", vbCritical + vbOKOnly, "Error: Client System Date Does Not Match HRI Date"
  End If
  '
  txtRestKey.SetFocus
  '
  rsLog.LockType = adLockPessimistic
  rsLog.CursorType = adOpenForwardOnly
  rsLog.Open "SELECT * from tbllog", cnMain
 '
  With rsLog
    .AddNew
    .Fields("ID") = .RecordCount + 1
    .Fields("Company").Value = ContactData.ID & vbNullString
    .Fields("User").Value = ContactData.FirstName & " " & ContactData.LastName & vbNullString
    .Fields("Employee").Value = User.Name
    .Fields("ActionDateTime").Value = Date + Time
    .Fields("ActionType").Value = "Restoration"
    .Fields("ActionSubType").Value = "N/A"
   ' .Fields("SiteExpirationDate").Value = vbNullString
   ' .Fields("SiteDays").Value = vbNull
    .Fields("SiteCompID").Value = CLng(sSiteCode2) & vbNullString
    .Fields("SiteSessionID").Value = CLng(sSiteCode1) & vbNullString
    .Fields("SiteKey").Value = txtRestKey.Text & vbNullString
    .Fields("SiteConfCode").Value = "N/A"
    .Fields("SiteDateTime").Value = txtRestDate.Text
    .UpdateBatch
  End With
  '
  txtSubject.Text = ContactData.AuthRemaining
  '
  rsLog.Close
  Set rsLog = Nothing
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.cmdRestGen.Click"
End Sub

Private Sub txtRestKey_GotFocus()
  On Error GoTo ErrorHandler
  '
  SelectText txtRestKey
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.txtRestKey.GotFocus"
End Sub

Private Sub cmdGen_Click()
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iDelPos             As Integer
  Dim lKeyCode            As Long
  Dim lDays               As Long
  Dim lDateDay            As Long
  Dim lDateMonth          As Long
  Dim lDateYear           As Long
  Dim lSiteKey1           As Long
  Dim lSiteKey2           As Long
  Dim sSiteCode           As String
  Dim sSiteCodeCompacted  As String
  Dim sSiteCode1          As String
  Dim sSiteCode2          As String
  '
  Dim rsLog As New ADODB.Recordset
  '
  If ValidateLicense <> CDbl(Date - dSecVar) Then Exit Sub
  '
  'If GeneralDataSpecified = False Then Exit Sub
  '
  sSiteCode = Trim(txtSiteCode.Text)
  '
  If InStr(1, sSiteCode, " ", vbBinaryCompare) <= 0 Then
    MsgBox "The site code you entered does not contain a space.", vbCritical + vbOKOnly, "Error: Site Code Does Not Contain Space"
    txtSiteCode.SetFocus
    Exit Sub
  End If
  '
  sSiteCodeCompacted = Replace(sSiteCode, " ", vbNullString)
  lDays = CLng(txtDays.Text)
  lKeyCode = IIf(optWrite(0) = True, 1, 2)
  '
  If sSiteCodeCompacted = vbNullString Then
    MsgBox "Please enter a valid site code.", vbCritical + vbOKOnly, "Error: Site Code Not Specified"
    txtSiteCode.SetFocus
    Exit Sub
  End If
  '
  If IsNumeric(sSiteCodeCompacted) = False Then
    MsgBox "Site codes cannot contain non-numeric characters. Please enter a valid site code.", vbCritical + vbOKOnly, "Error: Non-Numeric Site Code Specified"
    txtSiteCode.SetFocus
    Exit Sub
  End If
  '
  If lDays < 1 Then
    MsgBox "Please enter the number of days to authorize.", vbCritical + vbOKOnly, "Error: Number of Days Not Specified"
    txtDays.SetFocus
    Exit Sub
  End If
  '
  iDelPos = InStr(1, sSiteCode, " ", vbBinaryCompare)
  If iDelPos > 0 Then
    sSiteCode2 = Left$(sSiteCode, iDelPos - 1)
    sSiteCode1 = Trim(Right$(sSiteCode, Len(sSiteCode) - iDelPos))
  Else
    sSiteCode1 = sSiteCode
  End If
  '
  pp_cedate CLng(sSiteCode1), lDateMonth, lDateDay, lDateYear
  txtSiteDate = CStr(lDateMonth) & "/" & CStr(lDateDay) & "/" & CStr(lDateYear)
  lSiteKey1 = pp_ctcodes(lKeyCode, Val(sSiteCode1), Val(sSiteCode2), ProductData.Seed1) '173)
  lSiteKey2 = pp_nencrypt(lDays, ProductData.Seed2) '236
  '
  txtSiteKey.Text = CStr(lSiteKey1) & " " & CStr(lSiteKey2)
  If lKeyCode = 1 Then txtExpDate.Value = Date + lDays
  '
  If CDate(txtSiteDate.Text) <> Date Then
    txtSiteDate.BackColor = vbRed
    txtSiteDate.ForeColor = vbYellow
  '  MsgBox "The client's system date (" & txtSiteDate.Text & ") does not match HRI's date (" & txtHRIDate.Text & ")." & vbCrLf & vbCrLf & _
  '  "Please request the client to verify and -- if necessary -- correct his system's date. After a date correction, PowerClaim must be shut down and restarted in order to generate a correct site code." & vbCrLf & vbCrLf & _
  '  "If the client verifies that his date is set correctly, he must refresh the site code in the advanced licensing area.", vbCritical + vbOKOnly, "Error: Client System Date Does Not Match HRI Date"
  End If
  '
  '\\ Log Action
  rsLog.LockType = adLockPessimistic
  rsLog.CursorType = adOpenForwardOnly
  rsLog.Open "SELECT * from tbllog", cnMain
 ' rsLog.RecordCount
  With rsLog
    .AddNew
    .Fields("ID").Value = rsLog.RecordCount + 1
    .Fields("Company").Value = ContactData.ID & vbNullString
    .Fields("User").Value = ContactData.FirstName & " " & ContactData.LastName & vbNullString
    .Fields("Employee").Value = User.Name & vbNullString
    .Fields("ActionDateTime").Value = Date + Time & vbNullString
    .Fields("ActionType").Value = "Authorization"
    .Fields("ActionSubType").Value = IIf(lKeyCode = 1, "New", "Extension")
    If optWrite(0) = True Then
      .Fields("SiteExpirationDate").Value = txtExpDate.Value & vbNullString
    End If
    .Fields("SiteDays").Value = txtDays.Text & vbNullString
    .Fields("SiteCompID").Value = CLng(sSiteCode2) & vbNullString
    .Fields("SiteSessionID").Value = CLng(sSiteCode1) & vbNullString
    .Fields("SiteKey").Value = txtSiteKey.Text & vbNullString
    .Fields("SiteConfCode").Value = "N/A"
    .Fields("SiteDateTime").Value = txtSiteDate.Value & vbNullString
    .Fields("ProductID").Value = ProductData.ProductID
    .UpdateBatch
  End With
  '
  txtSubject.Text = txtDays.Text
  '
  rsLog.Close
  Set rsLog = Nothing
  '
  txtSiteKey.SetFocus
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.cmdGen.Click"
End Sub
'End Sub

Public Function LoadNote(plCustomerID As Long, plNoteID As Long) As Boolean
    Dim iPos As Integer
    '
    bNewNote = False
    '
    Set ContactData = New CContactData
    '
    Set EventData = New CEventData
    '
    Contact.Load ContactData, plCustomerID
    '
    Event1.Load EventData, plNoteID
    '
    Me.mskDate.Value = EventData.EventDate
    Me.mskDate.Enabled = False
    '
    Me.mskTime.Value = EventData.EventTime
    Me.mskTime.Enabled = False
    '
    Me.lblUser.Caption = EventData.EventUser
    '
    Me.txtSubject = EventData.EventSubject
    Me.txtResults = EventData.EventResults
    '
    Me.chkSticky.Value = IIf(EventData.Sticky, 1, 0)
    '
    Select Case EventData.ProductID
      Case 1
        cboProduct.Text = "PowerClaim XML"
      Case 2
        cboProduct.Text = "PowerClaim PV"
    End Select
    '
    Product.Load ProductData, EventData.ProductID
    '
    SwitchProduct
    '
    For iPos = 0 To lstTypes.ListCount - 1
      If lstTypes.list(iPos) = EventData.EventType Then
        lstTypes.ListIndex = iPos
      End If
    Next
    '
    Me.Show vbModal
    '
    LoadNote = bSaved
End Function

Public Function NewNote(plCustomerID As Long) As Boolean
    bNewNote = True
    '
    Set ContactData = New CContactData
    '
    Set EventData = New CEventData
    '
    Contact.Load ContactData, plCustomerID
    '
    cboProduct.Text = "PowerClaim XML"
    '
    lblUser.Caption = User.Name
    '
    mskDate.Value = CDate(Now)
    '
    mskTime.Value = Format(Now, "hh:nn AM/PM")
    '
    Me.Show vbModal
    '
    NewNote = bSaved
End Function

Private Sub Form_Load()
    mskDate = Date
    mskTime = Time
    '
    LoadTypes
    '
    cboProduct.AddItem "PowerClaim XML"
    cboProduct.AddItem "PowerClaim PV"
    '
    Product.Load ProductData, 1
    '
    bSaved = False
End Sub

Private Sub LoadTypes()
  Dim rsType As New Recordset
  '
  rsType.Open "SELECT * FROM tblActivities WHERE ActivityType = 1 ORDER BY Activity", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rsType.eof
    lstTypes.AddItem rsType!Activity & vbNullString
    rsType.MoveNext
  Loop
  '
  rsType.Close
  '
  lstTypes.AddItem vbNullString
  '
  rsType.Open "SELECT * FROM tblActivities WHERE ActivityType = 0 ORDER BY Activity", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  Do While Not rsType.eof
    lstTypes.AddItem rsType!Activity & vbNullString
    rsType.MoveNext
  Loop
  '
  rsType.Close
  '
  Set rsType = Nothing
End Sub

Private Sub SuggestAuthorizationRevisionData()
  If cboProduct.Text = "PowerClaim XML" Then
    'If ContactData.DaysPending > 0 Then
      Me.txtReviseAuthDays = ContactData.AuthDays
      Me.mskReviseAuthDate = ContactData.AuthDate
    'End If
  End If
  '
  If cboProduct.Text = "PowerClaim PV" Then
    'If ContactData.PVDaysPending > 0 Then
      Me.txtReviseAuthDays = ContactData.PVAuthDays
      Me.mskReviseAuthDate = ContactData.PVAuthDate
   ' End If
  End If
End Sub

Private Sub SuggestDays()
  If cboProduct.Text = "PowerClaim XML" Then
    If ContactData.DaysPending > 0 Then
      txtDays.Text = ContactData.DaysPending
    End If
  End If
  '
  If cboProduct.Text = "PowerClaim PV" Then
    If ContactData.PVDaysPending > 0 Then
      txtDays.Text = ContactData.PVDaysPending
    End If
  End If
End Sub

Private Sub lstTypes_Click()
  If bNewNote Then
    txtResults.Height = 5000
    fmeAuth.Visible = False
    fmeDeathorization.Visible = False
    fmeUserNum.Visible = False
    fmeAuthRevise.Visible = False
    fmeDateRestore.Visible = False
    fmeSale.Visible = False
    txtSubject.Enabled = True
    bAuthEvent = False
    '
    Select Case lstTypes.Text
      Case "Eval Authorized", "Eval Addition", "Second Authorization", "Reauthorization", "Paid Authorization"
        bAuthEvent = True
        txtResults.Height = 4000
        fmeAuth.Top = 5500
        fmeAuth.Left = 1920
        fmeAuth.Visible = True
        fmeAuth.BackColor = vbButtonFace
        txtSubject.Text = vbNullString
        txtSubject.Enabled = False
        '
        SuggestDays
      Case "Authorization Revision"
        txtResults.Height = 4000
        fmeAuthRevise.Top = 5500
        fmeAuthRevise.Left = 1920
        fmeAuthRevise.Visible = True
        fmeAuthRevise.BackColor = vbButtonFace
        txtSubject.Text = vbNullString
        txtSubject.Enabled = False
        '
        SuggestAuthorizationRevisionData
      Case "Deathorization"
        txtResults.Height = 4000
        fmeDeathorization.Top = 5500
        fmeDeathorization.Left = 1920
        fmeDeathorization.Visible = True
        fmeDeathorization.BackColor = vbButtonFace
        txtSubject.Text = vbNullString
        txtSubject.Enabled = False
      Case "User Limit Changed"
        txtResults.Height = 4000
        fmeUserNum.Top = 5500
        fmeUserNum.Left = 1920
        fmeUserNum.Visible = True
        fmeUserNum.BackColor = vbButtonFace
        txtSubject.Enabled = False
        txtSubject.Text = vbNullString
      Case "Date Restoration"
        txtResults.Height = 4000
        fmeDateRestore.Top = 5500
        fmeDateRestore.Left = 1920
        fmeDateRestore.Visible = True
        fmeDateRestore.BackColor = vbButtonFace
        txtSubject.Enabled = False
        txtSubject.Text = vbNullString
      Case "Sale", "Sale Revision"
        txtResults.Height = 4000
        fmeSale.Top = 5500
        fmeSale.Left = 1920
        fmeSale.Visible = True
        fmeSale.BackColor = vbButtonFace
        txtSubject.Enabled = False
        txtSubject.Text = vbNullString
        txtGraceDays.Text = 14
    End Select
  End If
End Sub

'Private Sub txtSubject_KeyPress(KeyAscii As Integer)
'  If fAuthorizing Then
'    Select Case KeyAscii
'    Case 8, 48 To 57 ' 48 to 57 0-9   8=backspace
'    Case Else
'      KeyAscii = 0
'    End Select
'  End If
'End Sub

Private Sub txtPendingDays_GotFocus()
  InputNumber.Setup txtPendingDays, NumberTypeInteger
End Sub

Private Sub txtGraceDays_GotFocus()
  InputNumber.Setup txtGraceDays, NumberTypeInteger
End Sub

