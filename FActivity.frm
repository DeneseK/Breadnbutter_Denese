VERSION 5.00
Begin VB.Form FActivity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activity"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "FActivity.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblSiteExpDate 
      Alignment       =   1  'Right Justify
      Caption         =   "N/A"
      DataField       =   "HRI"
      Height          =   225
      Left            =   8610
      TabIndex        =   25
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label lblSiteDays 
      Alignment       =   1  'Right Justify
      Caption         =   "N/A"
      Height          =   225
      Left            =   9150
      TabIndex        =   24
      Top             =   1710
      Width           =   405
   End
   Begin VB.Label lblConfCode 
      Caption         =   "N/A"
      Height          =   225
      Left            =   6090
      TabIndex        =   23
      Top             =   1140
      Width           =   3465
   End
   Begin VB.Label lblSiteKey 
      Caption         =   "N/A"
      Height          =   225
      Left            =   6090
      TabIndex        =   22
      Top             =   810
      Width           =   3465
   End
   Begin VB.Label lblSiteCode 
      Caption         =   "N/A"
      Height          =   225
      Left            =   6090
      TabIndex        =   21
      Top             =   480
      Width           =   3465
   End
   Begin VB.Label lblAction 
      Caption         =   "N/A"
      Height          =   225
      Left            =   6090
      TabIndex        =   20
      Top             =   150
      Width           =   3465
   End
   Begin VB.Label lblEmp 
      Caption         =   "N/A"
      Height          =   225
      Left            =   1020
      TabIndex        =   19
      Top             =   1050
      Width           =   3495
   End
   Begin VB.Label lblUser 
      Caption         =   "N/A"
      Height          =   225
      Left            =   1020
      TabIndex        =   18
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label lblCompany 
      Caption         =   "N/A"
      Height          =   225
      Left            =   1020
      TabIndex        =   17
      Top             =   1710
      Width           =   3495
   End
   Begin VB.Label lblSiteTime 
      Caption         =   "N/A"
      DataField       =   "HRI"
      Height          =   225
      Left            =   6090
      TabIndex        =   16
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lblSiteDate 
      Caption         =   "N/A"
      DataField       =   "HRI"
      Height          =   225
      Left            =   6090
      TabIndex        =   15
      Top             =   1710
      Width           =   945
   End
   Begin VB.Label lblHRITime 
      Caption         =   "N/A"
      DataField       =   "HRI"
      Height          =   225
      Left            =   1020
      TabIndex        =   14
      Top             =   480
      Width           =   795
   End
   Begin VB.Label lblHRIDate 
      Caption         =   "N/A"
      DataField       =   "HRI"
      Height          =   225
      Left            =   1020
      TabIndex        =   13
      Top             =   150
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Action:"
      Height          =   225
      Left            =   5100
      TabIndex        =   12
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label6 
      Caption         =   "HRI Date:"
      DataField       =   "HRI"
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   150
      Width           =   795
   End
   Begin VB.Label Label9 
      Caption         =   "HRI Time:"
      Height          =   225
      Left            =   150
      TabIndex        =   10
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label8 
      Caption         =   "Site Time:"
      Height          =   240
      Left            =   5100
      TabIndex        =   9
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label Label7 
      Caption         =   "Conf. Code:"
      Height          =   225
      Left            =   5100
      TabIndex        =   8
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Site Key:"
      Height          =   225
      Left            =   5100
      TabIndex        =   7
      Top             =   810
      Width           =   705
   End
   Begin VB.Label lblUnlockLevel 
      Caption         =   "Site Days:"
      Height          =   240
      Left            =   7410
      TabIndex        =   6
      Top             =   1710
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Site Code:"
      Height          =   225
      Left            =   5100
      TabIndex        =   5
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label10 
      Caption         =   "Site Date:"
      DataField       =   "HRI"
      Height          =   225
      Left            =   5100
      TabIndex        =   4
      Top             =   1710
      Width           =   795
   End
   Begin VB.Label Label16 
      Caption         =   "Site Exp. Date:"
      DataField       =   "HRI"
      Height          =   225
      Left            =   7410
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Employee:"
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "User:"
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Company:"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   1710
      Width           =   735
   End
End
Attribute VB_Name = "FActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  On Error GoTo ErrorHandler
  '
  Dim rsLog As New ADODB.Recordset
  Dim X As Long
  '
  '\\ Local Declarations
  Dim iEntryNo    As Integer
  '
  rsLog.Open "SELECT * FROM tblLog", cnMain
  '
  With rsLog
    '.FindFirst "[ID] = " & CLng(Right$(FAuthLog.lvwLog.SelectedItem.Key, Len(FAuthLog.lvwLog.SelectedItem.Key) - 1))
    'If .NoMatch = False Then
      X = CLng(Right$(FAuthLog.lvwLog.SelectedItem.Key, Len(FAuthLog.lvwLog.SelectedItem.Key) - 1))
    .MoveFirst
    .Find "ID = " & X
     If Not .EOF Then
      lblHRIDate.Caption = Format(.Fields("ActionDateTime").Value, "YYYY.Mm.Dd")
      lblHRITime.Caption = Format(.Fields("ActionDateTime").Value, "Hh:Nn:Ss")
      lblEmp.Caption = .Fields("Employee").Value
      lblCompany.Caption = .Fields("Company").Value
      lblUser.Caption = .Fields("User").Value
      lblSiteCode.Caption = IIf(.Fields("SiteCompID").Value <> 0, CStr(.Fields("SiteCompID").Value) & " " & CStr(.Fields("SiteSessionID").Value), "N/A")
      lblSiteKey.Caption = .Fields("SiteKey").Value
      lblConfCode.Caption = .Fields("SiteConfCode").Value
      lblSiteDate.Caption = Format(.Fields("SiteDateTime").Value, "YYYY.Mm.Dd")
      lblSiteTime.Caption = Format(.Fields("SiteDateTime").Value, "Hh:Nn:Ss")
      lblSiteDays.Caption = IIf(.Fields("SiteDays").Value <> vbNull, CStr(.Fields("SiteDays").Value), "N/A")
      lblSiteExpDate.Caption = IIf(.Fields("SiteExpirationDate").Value <> vbNull, Format(.Fields("SiteExpirationDate").Value, "YYYY.Mm.Dd"), "N/A")
      If .Fields("ActionType").Value = "Authorization" Then
        lblAction.Caption = "Authorization (" & .Fields("ActionSubType").Value & ")"
      Else
        lblAction.Caption = .Fields("ActionType").Value
      End If
    End If
  End With
  '
  rsLog.Close
  Set rsLog = Nothing
  '
  Exit Sub
  '
ErrorHandler:
  MsgBox "(" & Err.number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: FPrimary.General.RefreshLogDisplay"
End Sub
