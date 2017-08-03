VERSION 5.00
Begin VB.Form FChooseGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Groups"
   ClientHeight    =   2235
   ClientLeft      =   6300
   ClientTop       =   5115
   ClientWidth     =   4095
   Icon            =   "FChooseGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox lstCustGroups 
      Height          =   1035
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please choose the group you wish to load."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3630
   End
End
Attribute VB_Name = "FChooseGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lListID As Long

Private Sub cmdCancel_Click()
  lListID = 0
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If lstCustGroups.ListIndex <> -1 Then
    lListID = lstCustGroups.ItemData(lstCustGroups.ListIndex)
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  LoadCustGroupsList
End Sub

Private Sub LoadCustGroupsList()
  Dim GroupList As CGroupList
  Dim GroupLists As CGroupListDatas
  Dim GroupListLink As CGroupListLink
  Dim X As Integer
  Dim sKey As String
  '
  Set GroupList = New CGroupList
  Set GroupLists = New CGroupListDatas
  Set GroupListLink = New CGroupListLink
  '
  GroupList.LoadCollection GroupLists
  '
  'lstCustGroups.Clear
  '
  For X = 1 To GroupLists.Count
    sKey = "A" & GroupLists.Item(X).ID
    lstCustGroups.AddItem GroupLists.Item(X).ListName
    lstCustGroups.ItemData(X - 1) = GroupLists.Item(X).ID
  Next
  '
  Set GroupList = Nothing
  Set GroupLists = Nothing
  Set GroupListLink = Nothing
End Sub

Public Function GetGroup() As Long
  Me.Show vbModal
  GetGroup = lListID
End Function

