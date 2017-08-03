Option Strict Off
Option Explicit On
Friend Class FOutlookAppt
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error GoTo ErrCall
		'
		' Start Outlook.
		' If it is already running, you'll use the same instance...
		Dim olApp As Microsoft.Office.Interop.Outlook.Application
		olApp = CreateObject("Outlook.Application")
		' Logon. Doesn't hurt if you are already running and logged on...
		Dim olNs As Microsoft.Office.Interop.Outlook.NameSpace
		olNs = olApp.GetNamespace("MAPI")
		olNs.Logon()
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
		Dim olAppt As Microsoft.Office.Interop.Outlook.AppointmentItem
		olAppt = olApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
		' Set start time for 2-minutes from now...
		olAppt.Start = System.Date.FromOADate(CDate(mskDate._Text).ToOADate + CDate(mskTime._Text).ToOADate) 'Now() + (2# / 24# / 60#)
		' Setup other appointment information...
		With olAppt
			.Duration = 5
			.Subject = lblName.Text & " " & lblCompany.Text & ": " & txtSubject.Text
			If txtNote.Text <> "" Then .Body = txtNote.Text
			'.Location = "Home Office"
			.ReminderMinutesBeforeStart = 1
			.ReminderSet = True
		End With
		' Save Appointment...
		olAppt.Save()
		' Send a message to your new contact.
		'Dim olMail As Outlook.MailItem
		'Set olMail = olApp.CreateItem(olMailItem)
		' Fill out & send message...
		'olMail.To = olItem.Email1Address
		'olMail.Subject = "About our meeting..."
		'olMail.Body = _
		''"Dear " & olItem.FirstName & ", " & vbCr & vbCr & vbTab & _
		''"I'll see you in 2 minutes for our meeting!" & vbCr & vbCr & _
		''"Btw: I've added you to my contact list."
		'olMail.Send
		' Clean up...
		'MsgBox "Appointment Set.", vbInformation
		
		olNs.Logoff()
		'UPGRADE_NOTE: Object olNs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		olNs = Nothing
		'Set olMail = Nothing
		'UPGRADE_NOTE: Object olAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		olAppt = Nothing
		'Set olItem = Nothing
		'UPGRADE_NOTE: Object olApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		olApp = Nothing
		'
		If chkSales.Visible = True And chkSales.CheckState = System.Windows.Forms.CheckState.Checked Then
			SendIt2()
		End If
		'
		Me.Close()
		'
		Exit Sub
ErrCall: 
		MsgBox(Err.Description)
	End Sub
	
	'UPGRADE_WARNING: Form event FOutlookAppt.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FOutlookAppt_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		mskDate._Text = CStr(Now)
		mskTime._Text = CStr(Now)
	End Sub
	
	
	Public Property ContactName() As String
		Get
			
		End Get
		Set(ByVal Value As String)
			lblName.Text = Value
		End Set
	End Property
	
	Public WriteOnly Property Company() As String
		Set(ByVal Value As String)
			lblCompany.Text = Value
		End Set
	End Property
	
	Private Sub FOutlookAppt_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim Employee As New CEmployee
		'
		If Employee.InGroup((User.Name), "Sales") = True Then
			chkSales.Visible = True
			If GetSetting(My.Application.Info.Title, "OutlookAppt", "Sales_Contact", "true") = "true" Then
				chkSales.CheckState = System.Windows.Forms.CheckState.Checked
			Else
				chkSales.CheckState = System.Windows.Forms.CheckState.Unchecked
			End If
		End If
	End Sub
	
	Private Sub FOutlookAppt_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		If chkSales.CheckState = System.Windows.Forms.CheckState.Checked Then
			SaveSetting(My.Application.Info.Title, "OutlookAppt", "Sales_Contact", "true")
		Else
			SaveSetting(My.Application.Info.Title, "OutlookAppt", "Sales_Contact", "false")
		End If
	End Sub
	
	Private Sub mskTime2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mskTime2.Enter
		mskTime2._Text = mskTime._Text
	End Sub
	
	Private Sub mskTime2_SpinClick(ByVal Direction As Short)
		Dim iMinute, iHour, iQuarters As Short
		If Direction = 1 Then
			iHour = Hour(CDate(mskTime._Text))
			iMinute = Minute(CDate(mskTime._Text))
			iQuarters = iMinute \ 15
			iMinute = 15 * (iQuarters + 1)
			mskTime._Text = CStr(TimeSerial(iHour, iMinute, 0))
		Else
			If Direction = 0 Then
				iHour = Hour(CDate(mskTime._Text))
				iMinute = Minute(CDate(mskTime._Text))
				iQuarters = iMinute \ 15
				iMinute = 15 * (iQuarters - 1)
				mskTime._Text = CStr(TimeSerial(iHour, iMinute, 0))
			End If
		End If
	End Sub
	
	Private Sub SendIt2()
		On Error GoTo ErrCall
		
		Dim ol As Microsoft.Office.Interop.Outlook.Application
		Dim ns As Microsoft.Office.Interop.Outlook.NameSpace
		Dim pubFolder As Microsoft.Office.Interop.Outlook.MAPIFolder
		Dim allPubFolders As Microsoft.Office.Interop.Outlook.MAPIFolder
		Dim lglCalendar As Microsoft.Office.Interop.Outlook.MAPIFolder
		
		ol = GetObject( , "Outlook.Application")
		ns = ol.GetNamespace("MAPI")
		ns.Logon()
		pubFolder = ns.Folders("Public Folders")
		allPubFolders = pubFolder.Folders("All Public Folders")
		lglCalendar = allPubFolders.Folders("Sales Contact")
		' Create a new appointment.
		Dim olAppt As Microsoft.Office.Interop.Outlook.AppointmentItem
		olAppt = lglCalendar.Items.Add(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem) 'ol.CreateItem(olAppointmentItem)
		' Set start time for 2-minutes from now...
		olAppt.Start = System.Date.FromOADate(CDate(mskDate._Text).ToOADate + CDate(mskTime._Text).ToOADate) 'Now() + (2# / 24# / 60#)
		' Setup other appointment information...
		With olAppt
			.Duration = 5
			.Subject = lblName.Text & " " & lblCompany.Text & ": " & txtSubject.Text
			If txtNote.Text <> "" Then .Body = txtNote.Text
			'.Location = "Home Office"
			.ReminderMinutesBeforeStart = 1
			.ReminderSet = True
		End With
		' Save Appointment...
		olAppt.Save()
		ns.Logoff()
		'UPGRADE_NOTE: Object ns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ns = Nothing
		'UPGRADE_NOTE: Object olAppt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		olAppt = Nothing
		'UPGRADE_NOTE: Object ol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ol = Nothing
		'
		Exit Sub
ErrCall: 
		MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Add to Sales Contact")
		
	End Sub
End Class