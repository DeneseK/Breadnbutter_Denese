VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FSelect 
   BorderStyle     =   0  'None
   Caption         =   "Select Customers for Processing"
   ClientHeight    =   5235
   ClientLeft      =   1020
   ClientTop       =   2430
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Deselect All"
      Height          =   315
      Index           =   1
      Left            =   7590
      TabIndex        =   2
      Top             =   270
      Width           =   1245
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select All"
      Height          =   315
      Index           =   0
      Left            =   6210
      TabIndex        =   1
      Top             =   270
      Width           =   1275
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Projects\Breadnbutter\Data\PCCustomers.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"FSelect.frx":0000
      Top             =   4890
      Visible         =   0   'False
      Width           =   1875
   End
   Begin SSDataWidgets_B.SSDBGrid grdSelect 
      Bindings        =   "FSelect.frx":00AB
      Height          =   4335
      Left            =   150
      TabIndex        =   0
      Top             =   630
      Width           =   8655
      _Version        =   196617
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   4445
      Columns(0).Caption=   "Company"
      Columns(0).Name =   "Company"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Company"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3625
      Columns(1).Caption=   "Contact"
      Columns(1).Name =   "Contact"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Contact"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4815
      Columns(2).Caption=   "Address1"
      Columns(2).Name =   "Address1"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Address1"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1217
      Columns(3).Caption=   "Select"
      Columns(3).Name =   "BetaTester"
      Columns(3).Alignment=   2
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "BetaTester"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      _ExtentX        =   15266
      _ExtentY        =   7646
      _StockProps     =   79
      Caption         =   "Select Customers for Printing"
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
End
Attribute VB_Name = "FSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents FormControl As CFormControl
Attribute FormControl.VB_VarHelpID = -1
Public WithEvents FormData As CFormData
Attribute FormData.VB_VarHelpID = -1

Private Sub cmdSelect_Click(Index As Integer)
  On Error GoTo ErrCall
  '
  grdSelect.Redraw = False
  With Data1.Recordset
  .MoveFirst
  Do While Not .EOF
    .Edit
    !betatester = (Index = 0)
    .Update
    .MoveNext
  Loop
  End With
  grdSelect.Redraw = True
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & "in frmSelect.cmdSelect_Click.", vbCritical, "Error"
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
  MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & "in frmSelect.Form_Initialize.", vbCritical, "Error"
End Sub

Private Sub Form_Load()
  On Error GoTo ErrCall
  '
  'DBOps.SetDatDB dbMain, Data1
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & "in frmSelect.Form_Load.", vbCritical, "Error"
End Sub
