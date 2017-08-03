VERSION 5.00
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form FOutlookAppt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appointment"
   ClientHeight    =   3390
   ClientLeft      =   4125
   ClientTop       =   3120
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSales 
      Caption         =   "Add to Sales Contact"
      Height          =   195
      Left            =   2640
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   1470
      TabIndex        =   2
      Top             =   1620
      Width           =   2715
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Width           =   1065
   End
   Begin VB.TextBox txtNote 
      Height          =   705
      Left            =   1470
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2010
      Width           =   2715
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   210
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin GTMaskDate.GTMaskDate mskTime 
      Height          =   345
      Left            =   1470
      TabIndex        =   1
      Top             =   1200
      Width           =   1395
      _Version        =   65537
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MaskType        =   1
      SpinButton      =   -1  'True
      SpinIncrement   =   6
      CalDropDown     =   0   'False
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
   Begin GTMaskDate.GTMaskDate mskDate 
      Height          =   345
      Left            =   1500
      TabIndex        =   0
      Top             =   780
      Width           =   1365
      _Version        =   65537
      _ExtentX        =   2408
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
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
   Begin GTMaskDate.GTMaskDate mskTime2 
      Height          =   345
      Left            =   1680
      TabIndex        =   14
      Top             =   1200
      Width           =   1395
      _Version        =   65537
      _ExtentX        =   2461
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MaskType        =   1
      SpinButton      =   -1  'True
      SpinKeys        =   0   'False
      SpinIncrement   =   5
      CalDropDown     =   0   'False
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
   Begin VB.Label lblCompany 
      Height          =   285
      Left            =   1530
      TabIndex        =   13
      Top             =   390
      Width           =   2865
   End
   Begin VB.Label Label1 
      Caption         =   "Company:"
      Height          =   285
      Index           =   1
      Left            =   570
      TabIndex        =   12
      Top             =   390
      Width           =   855
   End
   Begin VB.Label lblDate 
      Caption         =   "Subject:"
      Height          =   285
      Index           =   3
      Left            =   570
      TabIndex        =   11
      Top             =   1650
      Width           =   795
   End
   Begin VB.Label lblName 
      Height          =   285
      Left            =   1530
      TabIndex        =   10
      Top             =   30
      Width           =   2865
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   285
      Index           =   0
      Left            =   570
      TabIndex        =   9
      Top             =   30
      Width           =   855
   End
   Begin VB.Label lblDate 
      Caption         =   "Note:"
      Height          =   285
      Index           =   2
      Left            =   540
      TabIndex        =   8
      Top             =   2070
      Width           =   795
   End
   Begin VB.Label lblDate 
      Caption         =   "Time:"
      Height          =   285
      Index           =   1
      Left            =   570
      TabIndex        =   7
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label lblDate 
      Caption         =   "Date:"
      Height          =   285
      Index           =   0
      Left            =   570
      TabIndex        =   6
      Top             =   810
      Width           =   795
   End
End
Attribute VB_Name = "FOutlookAppt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrCall
  '
  ' Start Outlook.
  ' If it is already running, you'll use the same instance...
  Dim olApp As Outlook.Application
  Set olApp = CreateObject("Outlook.Application")
  ' Logon. Doesn't hurt if you are already running and logged on...
  Dim olNs As Outlook.NameSpace
  Set olNs = olApp.GetNamespace("MAPI")
  olNs.Logon
  ' Create and Open a new contact.
  'Dim olItem As Outlook.ContactItem
  'Set olItem = olApp.CreateItem(olContactItem)
  ' Setup Contact information...
  'With olItem
  '.FullName = "James Smith"
  '.Birthday = "9/15/1975"
  '.CompanyName = "Microsoft"
  '.HomeTelephoneNumber = "704-555-8888"
  '.Email1Address = "someone@microsoft.com"
  '.JobTitle = "Developer"
  '.HomeAddress = "111 Main St." & vbCr & "Charlotte, NC 28226"
  'End With
  ' Save Contact...
  'olItem.Save
  ' Create a new appointment.
  Dim olAppt As Outlook.AppointmentItem
  Set olAppt = olApp.CreateItem(olAppointmentItem)
  ' Set start time for 2-minutes from now...
  olAppt.Start = CDate(mskDate) + CDate(mskTime) 'Now() + (2# / 24# / 60#)
  ' Setup other appointment information...
  With olAppt
  .Duration = 5
  .Subject = lblName.Caption & " " & lblCompany.Caption & ": " & txtSubject.Text
  If txtNote.Text <> "" Then .Body = txtNote.Text
  '.Location = "Home Office"
  .ReminderMinutesBeforeStart = 1
  .ReminderSet = True
  End With
  ' Save Appointment...
  olAppt.Save
  ' Send a message to your new contact.
  'Dim olMail As Outlook.MailItem
  'Set olMail = olApp.CreateItem(olMailItem)
  ' Fill out & send message...
  'olMail.To = olItem.Email1Address
  'olMail.Subject = "About our meeting..."
  'olMail.Body = _
  '"Dear " & olItem.FirstName & ", " & vbCr & vbCr & vbTab & _
  '"I'll see you in 2 minutes for our meeting!" & vbCr & vbCr & _
  '"Btw: I've added you to my contact list."
  'olMail.Send
  ' Clean up...
  'MsgBox "Appointment Set.", vbInformation
  
  olNs.Logoff
  Set olNs = Nothing
  'Set olMail = Nothing
  Set olAppt = Nothing
  'Set olItem = Nothing
  Set olApp = Nothing
  '
  If chkSales.Visible = True And chkSales.Value = vbChecked Then
      SendIt2
  End If
  '
  Unload Me
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description
End Sub

Private Sub Form_Activate()
  mskDate = Now
  mskTime = Now
End Sub

Public Property Get ContactName() As String

End Property

Public Property Let ContactName(ByVal psContactName As String)
  lblName.Caption = psContactName
End Property

Public Property Let Company(ByVal psCompany As String)
  lblCompany.Caption = psCompany
End Property

Private Sub Form_Load()
  Dim Employee As New CEmployee
  '
  If Employee.InGroup(User.Name, "Sales") = True Then
      chkSales.Visible = True
      If GetSetting(App.Title, "OutlookAppt", "Sales_Contact", "true") = "true" Then
        chkSales.Value = vbChecked
      Else
        chkSales.Value = vbUnchecked
      End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If chkSales.Value = vbChecked Then
    SaveSetting App.Title, "OutlookAppt", "Sales_Contact", "true"
  Else
    SaveSetting App.Title, "OutlookAppt", "Sales_Contact", "false"
  End If
End Sub

Private Sub mskTime2_GotFocus()
  mskTime2 = mskTime
End Sub

Private Sub mskTime2_SpinClick(ByVal Direction As Integer)
  Dim iHour As Integer, iMinute As Integer, iQuarters As Integer
  If Direction = 1 Then
    iHour = Hour(mskTime)
    iMinute = Minute(mskTime)
    iQuarters = iMinute \ 15
    iMinute = 15 * (iQuarters + 1)
    mskTime = TimeSerial(iHour, iMinute, 0)
  Else
   If Direction = 0 Then
        iHour = Hour(mskTime)
        iMinute = Minute(mskTime)
        iQuarters = iMinute \ 15
        iMinute = 15 * (iQuarters - 1)
        mskTime = TimeSerial(iHour, iMinute, 0)
  End If
  End If
End Sub

Private Sub SendIt2()
  On Error GoTo ErrCall
  
  Dim ol As Outlook.Application
  Dim ns As Outlook.NameSpace
  Dim pubFolder As Outlook.MAPIFolder
  Dim allPubFolders As Outlook.MAPIFolder
  Dim lglCalendar As Outlook.MAPIFolder
  
  Set ol = GetObject(, "Outlook.Application")
  Set ns = ol.GetNamespace("MAPI")
  ns.Logon
  Set pubFolder = ns.Folders("Public Folders")
  Set allPubFolders = pubFolder.Folders("All Public Folders")
  Set lglCalendar = allPubFolders.Folders("Sales Contact")
  ' Create a new appointment.
  Dim olAppt As Outlook.AppointmentItem
  Set olAppt = lglCalendar.Items.Add(olAppointmentItem) 'ol.CreateItem(olAppointmentItem)
  ' Set start time for 2-minutes from now...
  olAppt.Start = CDate(mskDate) + CDate(mskTime) 'Now() + (2# / 24# / 60#)
  ' Setup other appointment information...
  With olAppt
  .Duration = 5
  .Subject = lblName.Caption & " " & lblCompany.Caption & ": " & txtSubject.Text
  If txtNote.Text <> "" Then .Body = txtNote.Text
  '.Location = "Home Office"
  .ReminderMinutesBeforeStart = 1
  .ReminderSet = True
  End With
  ' Save Appointment...
  olAppt.Save
  ns.Logoff
  Set ns = Nothing
  Set olAppt = Nothing
  Set ol = Nothing
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description, vbExclamation, "Add to Sales Contact"
  
End Sub


