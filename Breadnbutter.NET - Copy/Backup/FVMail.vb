Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FVMail
	Inherits System.Windows.Forms.Form
	
	'Private rs As New ADODB.Recordset
	'
	Public WithEvents FormControl As CFormControl
	Public WithEvents FormData As CFormData
	'
	Private rsMessages As New ADODB.Recordset
	'
	
	Private Sub SaveCheckMarks(ByRef pRS As ADODB.Recordset)
		Dim iListCount As Short
		Dim i As Short
		iListCount = ListView1.Items.Count
		'
		For i = 1 To iListCount
			If Not pRS.eof Then
				pRS.MoveFirst()
				'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				pRS.Find("messageID=" & CInt(VB.Right(ListView1.Items.Item(i).Name, Len(ListView1.Items.Item(i).Name) - 1)))
				'
				If Not pRS.eof Then
					'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					pRS.Fields("Checked").Value = ListView1.Items.Item(i).Checked
				End If
			Else
				Exit For
			End If
		Next 
		pRS.UpdateBatch()
	End Sub
	
	Public Sub RefreshMessages()
		Dim iPosition As Short
		Dim rs As New ADODB.Recordset
		'On Error GoTo ErrorHandler
		'  If rs.State <> 0 Then
		'    rs.Close
		'  End If
		'
		rs = GetRS(choice)
		'
		SaveCheckMarks(rs)
		'
		If Not ListView1.FocusedItem Is Nothing Then
			SavedIndex = ListView1.FocusedItem.Name
			iPosition = ListView1.FocusedItem.Index
		End If
		'
		lblcount.Text = "Messages Shown: " & FillList(rs, ListView1)
		'
		On Error Resume Next
		ListView1.FocusedItem = ListView1.Items.Item(SavedIndex)
		'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		'UPGRADE_WARNING: MSComctlLib.IListItem method ListView1.ListItems.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		ListView1.Items.Item(iPosition).EnsureVisible()
		
		'On Error GoTo ErrorHandler
		'
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(Mid(GetLastUpdate, 1, 11)), TimeOfDay) > 30 Or DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(Mid(GetLastUpdate, 12, 11)), Today) > 0 Then
			lblLastServer.BackColor = System.Drawing.Color.Red
			lblLastServer.Text = "Last Server Update: " & GetLastUpdate
			lblLastServer.Text = lblLastServer.Text & " Server May Be Down!!! Tell Supervisor"
		Else
			lblLastServer.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
			lblLastServer.Text = "Last Server Update: " & GetLastUpdate
		End If
		'
		lblLastClient.Text = "Last Client Update: " & TimeOfDay & " " & Today
		'
		rs.Close()
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
		
		Exit Sub
		'ErrorHandler:
		' MsgBox "Error filling list"
	End Sub
	
	'UPGRADE_WARNING: Event cmbCaller.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbCaller.Change was upgraded to cmbCaller.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbCaller_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbCaller.TextChanged
		cmdContactInfo.Enabled = True
	End Sub
	
	Private Sub cmdAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAll.Click
		cmdNew.Font = VB6.FontChangeBold(cmdNew.Font, False)
		cmdAll.Font = VB6.FontChangeBold(cmdAll.Font, True)
		cmdOld.Font = VB6.FontChangeBold(cmdOld.Font, False)
		choice = ALLCALLS
		RefreshMessages()
		ListView1_Click(ListView1, New System.EventArgs())
	End Sub
	
	Private Sub LoadEmailBodyBody()
		Dim rsBody As New ADODB.Recordset
		'
		rsBody.Open("SELECT [Body] FROM TVMailMessages WHERE MessageID = '" & CStr(sMessageID) & "'", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rsBody
			If Not .eof Then
				If Not .Fields("Body").Value = vbNullString Then
					txtBody.Text = .Fields("Body").Value & vbNullString
					sBody = .Fields("Body").Value & vbNullString
				Else
					txtBody.Text = vbNullString
					sBody = vbNullString
				End If
			Else
				txtBody.Text = vbNullString
				sBody = vbNullString
			End If
		End With
		'
		'UPGRADE_NOTE: Object rsBody may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsBody = Nothing
	End Sub
	
	Private Sub DisplayEmailBody()
		Dim TSTemp As Scripting.TextStream
		Dim fso As New Scripting.FileSystemObject
		'
		TSTemp = fso.OpenTextFile(My.Application.Info.DirectoryPath & "\temp.html", Scripting.IOMode.ForWriting, True, Scripting.Tristate.TristateUseDefault)
		TSTemp.Write(txtBody.Text)
		TSTemp.Close()
		'UPGRADE_WARNING: Navigate2 was upgraded to Navigate and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		webBody.Navigate(New System.URI(My.Application.Info.DirectoryPath & "\temp.html"))
	End Sub
	
	Private Sub cmdBrowser_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowser.Click
		Dim TSTemp As Scripting.TextStream
		Dim fso As New Scripting.FileSystemObject
		'
		TSTemp = fso.OpenTextFile(My.Application.Info.DirectoryPath & "\temp.html", Scripting.IOMode.ForWriting, True, Scripting.Tristate.TristateUseDefault)
		TSTemp.Write(txtBody.Text)
		TSTemp.Close()
		PlayTextFile("temp.html")
		'webBody.Navigate2 App.Path & "\temp.html"
	End Sub
	
	'Private Sub SendToOld(oldRS As Recordset)
	'  With oldRS
	'      If Not !Completed Then
	'        !DateCompleted = Date
	'        !TimeCompleted = Time
	'        !User = StrUser
	'      End If
	'      '
	'      !Completed = True
	'      On Error GoTo ErrorHandler
	'    End With
	'  Exit Sub
	'ErrorHandler:
	'    MsgBox ("Error. It's possible somebody else has changed this record since this window was opened")
	'End Sub
	
	Private Sub cmdCompleted_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCompleted.Click
		Dim i As Short
		Dim oldRS As New ADODB.Recordset
		Dim iListCount As Short
		Dim lMessageID As Integer
		Dim sKey As String
		'adOpenKeyset
		oldRS.Open("SELECT messageID, Checked, Completed, TimeCompleted, DateCompleted, [User] FROM TVMailMessages WHERE Completed = " & "'False'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		iListCount = ListView1.Items.Count
		'
		For i = 1 To iListCount
			'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If ListView1.Items.Item(i).Checked Then
				'UPGRADE_WARNING: Lower bound of collection ListView1.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				sKey = Trim(ListView1.Items.Item(i).Name)
				lMessageID = CInt(VB.Right(sKey, Len(sKey) - 1))
				oldRS.MoveFirst()
				oldRS.Find("messageID=" & lMessageID)
				If Not oldRS.eof Then
					oldRS.Fields("Checked").Value = False
					'SendToOld oldRS
					If Not oldRS.Fields("Completed").Value Then
						oldRS.Fields("DateCompleted").Value = Today
						oldRS.Fields("TimeCompleted").Value = TimeOfDay
						oldRS.Fields("User").Value = StrUser
					End If
					'
					oldRS.Fields("Completed").Value = True
				End If
			End If
		Next 
		'
		oldRS.UpdateBatch()
		oldRS.Close()
		'
		RefreshMessages()
		ListView1_Click(ListView1, New System.EventArgs())
		FMain.tmrMessages_Tick(Nothing, New System.EventArgs())
	End Sub
	
	Private Sub cmdContactInfo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdContactInfo.Click
		'Set FormMgr = New CFormMgr
		'FormMgr.Setup FMain
		'  sPhone = txtPhone.Text
		'  bVMail = True
		'  'fmain.tbMain.ToolBars(1).Tools.item..ID = "ID_Lookup"
		'  FContact.Form_Load
		'  FContact.Show
		'  FContact.WindowState = 2
		'  FContact.txtSearch = sPhone
		'
		'  FMain.tbMain_Go (FMain.tbMain.ToolBars(3).Tools.Item(4).ID)
		'  Unload FVMail
		'Debug.Print "tools " & FMain.tbMain.ToolBars(3).Tools.Item(4).ID
		
		sContact = cmbCaller.Text
		sContact = Trim(sContact)
		If InStr(sContact, " ") <> 0 Then
			GetContactInfo()
			'Company.Fetch iCompany
			'Company.Contact.Fetch iContact
			cmbCaller.Text = ""
			FContact.LoadContact(iContact, True)
			FormMgr.ShowForm(Me, FContact)
			
		Else
			MsgBox("Please enter First and Last Name", MsgBoxStyle.Information, "Bread 'n' Butter")
		End If
	End Sub
	
	Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
		MsgBox("Are You Sure You want to Delete Selected record?", MsgBoxStyle.OKCancel)
	End Sub
	
	Private Sub cmdDetails_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDetails.Click
		'
		ListView1_Click(ListView1, New System.EventArgs())
	End Sub
	
	
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		Me.Close()
	End Sub
	
	Private Sub cmdForward_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdForward.Click
		Dim TSTemp As Scripting.TextStream
		Dim fso As New Scripting.FileSystemObject
		'
		TSTemp = fso.OpenTextFile(My.Application.Info.DirectoryPath & "\tempMessage.txt", Scripting.IOMode.ForWriting, True, Scripting.Tristate.TristateUseDefault)
		TSTemp.WriteLine(("Received: " & sReceived))
		TSTemp.WriteLine(("From: " & sCaller))
		TSTemp.WriteLine((sBody))
		TSTemp.Close()
		
		FSendTo.ShowDialog()
	End Sub
	
	
	Private Sub cmdGetNames_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGetNames.Click
		FindNames()
		If Trim(cmbCaller.Text) = "" Then
			cmdContactInfo.Enabled = False
		Else
			cmdContactInfo.Enabled = True
		End If
	End Sub
	
	Private Sub cmdNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNew.Click
		cmdNew.Font = VB6.FontChangeBold(cmdNew.Font, True)
		cmdAll.Font = VB6.FontChangeBold(cmdAll.Font, False)
		cmdOld.Font = VB6.FontChangeBold(cmdOld.Font, False)
		choice = NEWCALLS
		RefreshMessages()
		ListView1_Click(ListView1, New System.EventArgs())
	End Sub
	
	Private Sub cmdOld_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOld.Click
		cmdNew.Font = VB6.FontChangeBold(cmdNew.Font, False)
		cmdAll.Font = VB6.FontChangeBold(cmdAll.Font, False)
		cmdOld.Font = VB6.FontChangeBold(cmdOld.Font, True)
		choice = OLDCALLS
		RefreshMessages()
		ListView1_Click(ListView1, New System.EventArgs())
	End Sub
	
	Private Sub CmdPlay_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdPlay.Click
		Dim strStream As ADODB.Stream
		Dim FileSys As New Scripting.FileSystemObject
		Dim myStr As String
		Dim rsAttach As New ADODB.Recordset
		'
		If Not FileSys.FolderExists(My.Application.Info.DirectoryPath & "\Temp") Then
			FileSys.CreateFolder(My.Application.Info.DirectoryPath & "\Temp")
		End If
		'
		'UPGRADE_WARNING: Lower bound of collection ListView1.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		myStr = ListView1.FocusedItem.SubItems(1).Text & vbNullString
		rsAttach.Open("SELECT * FROM TVMailMessages WHERE MessageID like '" & CStr(sMessageID) & "'", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rsAttach
			.MoveFirst()
			.Find("MessageName = " & "'" & myStr & "'")
			If Not .eof Then
				If Not .Fields("MessageName").Value = vbNullString Then
					If VB.Right(.Fields("MessageName").Value, 3) = "WAV" Or VB.Right(.Fields("MessageName").Value, 3) = "wav" Then
						strStream = New ADODB.Stream
						strStream.Type = ADODB.StreamTypeEnum.adTypeBinary
						strStream.Open()
						strStream.Write(.Fields("Attachment"))
						If Not FileSys.FileExists(My.Application.Info.DirectoryPath & "\Temp\" & .Fields("MessageName").Value) Then
							strStream.SaveToFile(My.Application.Info.DirectoryPath & "\Temp\" & .Fields("MessageName").Value, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
						End If
						strStream.Close()
						'UPGRADE_NOTE: Object strStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						strStream = Nothing
						'
						PlaySound((.Fields("MessageName").Value))
					Else
						MsgBox("Invalid File Format", MsgBoxStyle.Exclamation, "Warning")
					End If
				Else
					MsgBox("There is no file attached to play!", MsgBoxStyle.Information)
				End If
			End If
		End With
		'
		'UPGRADE_NOTE: Object rsAttach may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsAttach = Nothing
		'UPGRADE_NOTE: Object strStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		strStream = Nothing
		'UPGRADE_NOTE: Object FileSys may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FileSys = Nothing
	End Sub
	
	Private Sub cmdEditGroups_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEditGroups.Click
		FUserGroups.ShowDialog()
		RefreshMessages()
		FMain.tmrMessages_Tick(Nothing, New System.EventArgs())
	End Sub
	
	Private Sub cmdRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRefresh.Click
		RefreshMessages()
		ListView1_Click(ListView1, New System.EventArgs())
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		With rsMessages
			If Not .Fields("Completed").Value Then
				.Fields("DateCompleted").Value = Today
				.Fields("TimeCompleted").Value = TimeOfDay
				.Fields("User").Value = StrUser
			End If
			'
			If cmbComment.Text <> vbNullString Then
				.Fields("Comments").Value = cmbComment.Text
			End If
			If cmbCaller.Text <> vbNullString Then
				.Fields("Caller").Value = cmbCaller.Text
			End If
			If chkComp.CheckState = 1 Then
				.Fields("Completed").Value = True
			Else
				.Fields("Completed").Value = False
			End If
			On Error GoTo ErrorHandler
			.UpdateBatch()
		End With
		'
		RefreshMessages()
		ListView1_Click(ListView1, New System.EventArgs())
		FMain.tmrMessages_Tick(Nothing, New System.EventArgs())
		Exit Sub
ErrorHandler: 
		MsgBox("Error. It's possible somebody else has changed this record since this window was opened")
		'
	End Sub
	
	Private Sub FVMail_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim rsComments As New ADODB.Recordset
		'
		FormData = New CFormData
		FormControl = New CFormControl
		'
		FormControl.MinHeight = 5475
		FormControl.MinWidth = 10590
		FormControl.DataForm = True
		'
		choice = NEWCALLS
		cmdNew.Font = VB6.FontChangeBold(cmdNew.Font, True)
		cmdAll.Font = VB6.FontChangeBold(cmdAll.Font, False)
		cmdOld.Font = VB6.FontChangeBold(cmdOld.Font, False)
		'
		cmbComment.Items.Clear()
		rsComments.Open("select * from TVMailComment", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		With rsComments
			Do While Not .eof
				cmbComment.Items.Add(.Fields("Comment").Value)
				.MoveNext()
			Loop 
			.Close()
		End With
		'
		'UPGRADE_NOTE: Object rsComments may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsComments = Nothing
		'
		cmbMessageGroup.Items.Add("All")
		cmbMessageGroup.Items.Add("Authorizations")
		cmbMessageGroup.Items.Add("Sales")
		cmbMessageGroup.Items.Add("Support")
		GetUserGroups()
		'
		cmbMessageGroup.SelectedIndex = 0
		'
		Me.Text = ""
		'
		InitializeVmail()
		'
		Timer1.Interval = 30000
		'
		GetColumnWidths()
		'
		ListView1.Checkboxes = True
		'
		RefreshMessages()
		'
		ListView1_Click(ListView1, New System.EventArgs())
		'
	End Sub
	
	'UPGRADE_WARNING: Event FVMail.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub FVMail_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If VB6.PixelsToTwipsY(Me.Height) > 5000 And VB6.PixelsToTwipsX(Me.Width) > 1000 Then
			ListView1.SetBounds(0, VB6.TwipsToPixelsY(400), VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 100), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 4000))
			fraDetails.SetBounds(VB6.TwipsToPixelsX(150), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 3450), VB6.TwipsToPixelsX(11535), VB6.TwipsToPixelsY(2700))
			Shape1.SetBounds(VB6.TwipsToPixelsX(100), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 3500), VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 300), VB6.TwipsToPixelsY(2800))
			cmdDetails.SetBounds(0, 0, VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(400))
			cmdPlay.SetBounds(0, 0, VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(400))
			cmdCompleted.SetBounds(VB6.TwipsToPixelsX(1010), 0, VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(400))
			cmdRefresh.SetBounds(VB6.TwipsToPixelsX(2020), 0, VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(400))
			cmdNew.SetBounds(VB6.TwipsToPixelsX(5000), 0, VB6.TwipsToPixelsX(600), VB6.TwipsToPixelsY(400))
			cmdOld.SetBounds(VB6.TwipsToPixelsX(5610), 0, VB6.TwipsToPixelsX(600), VB6.TwipsToPixelsY(400))
			cmdAll.SetBounds(VB6.TwipsToPixelsX(6220), 0, VB6.TwipsToPixelsX(600), VB6.TwipsToPixelsY(400))
			lblMessageGroup.SetBounds(VB6.TwipsToPixelsX(7000), VB6.TwipsToPixelsY(100), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			lblGroups.SetBounds(VB6.TwipsToPixelsX(8300), VB6.TwipsToPixelsY(100), VB6.TwipsToPixelsX(1750), 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
			cmbMessageGroup.SetBounds(VB6.TwipsToPixelsX(8250), VB6.TwipsToPixelsY(50), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			cmdEditGroups.SetBounds(VB6.TwipsToPixelsX(10100), 0, VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(400))
			lblShow.SetBounds(VB6.TwipsToPixelsX(4300), VB6.TwipsToPixelsY(100), VB6.TwipsToPixelsX(500), VB6.TwipsToPixelsY(400))
			lblcount.SetBounds(0, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 600), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			lblLastClient.SetBounds(VB6.TwipsToPixelsX(2000), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 600), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			lblLastServer.SetBounds(VB6.TwipsToPixelsX(5750), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 600), VB6.TwipsToPixelsX(8000), 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
		End If
	End Sub
	
	Private Sub FVMail_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Dim FileSys As New Scripting.FileSystemObject
		
		Dim fileCount As Short
		Dim Index As Short
		'
		Timer1_Tick(Timer1, New System.EventArgs())
		If FileSys.FolderExists(My.Application.Info.DirectoryPath & "\Temp") Then
			File1.Path = My.Application.Info.DirectoryPath & "\Temp"
			File1.Refresh()
			fileCount = File1.Items.Count
			If fileCount > 0 Then
				For Index = 0 To fileCount - 1
					
					On Error Resume Next
					FileSys.DeleteFile((My.Application.Info.DirectoryPath & "\Temp\" & File1.Items(Index)))
				Next 
			End If
		End If
		SaveColumnWidths()
		'rs.Close
	End Sub
	
	Private Sub ListView1_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ListView1.DoubleClick
		ListView1_Click(ListView1, New System.EventArgs())
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		On Error Resume Next
		
		RefreshMessages()
		'
		
		
		'ListView1.SetFocus
	End Sub
	
	Private Sub ListView1_ColumnClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
		Dim ColumnHeader As System.Windows.Forms.ColumnHeader = ListView1.Columns(eventArgs.Column)
		On Error GoTo ErrorHandler
		'
		SortListView(ListView1, ColumnHeader)
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.lvwLog.ColumnClick")
	End Sub
	
	Public Sub SortListView(ByVal lvwCur As System.Windows.Forms.ListView, ByVal colHdr As System.Windows.Forms.ColumnHeader, Optional ByVal sSortOrder As String = "")
		On Error GoTo ErrorHandler
		'
		With lvwCur
			'
			'UPGRADE_ISSUE: MSComctlLib.ListView property lvwCur.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Lower bound of collection lvwCur.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_ISSUE: MSComctlLib.ColumnHeader property ColumnHeaders.Item.Icon was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			If .SortKey > -1 Then .Columns.Item(.SortKey + 1).Icon = 0
			'
			'UPGRADE_ISSUE: MSComctlLib.ListView property lvwCur.SortKey was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.SortKey = colHdr.Index - 1
			'
			If sSortOrder <> vbNullString Then
				.Sorting = IIf(sSortOrder = "Ascending", System.Windows.Forms.SortOrder.Ascending, System.Windows.Forms.SortOrder.Descending)
			Else
				.Sorting = IIf(.Sorting = System.Windows.Forms.SortOrder.Ascending, System.Windows.Forms.SortOrder.Descending, System.Windows.Forms.SortOrder.Ascending)
			End If
			'
			.Sort()
			'
			'.ColumnHeaders.Item(colHdr.Index).Icon = IIf(.SortOrder = lvwAscending, "imgAscending", "imgDescending")
			'
		End With
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox(ErrorToString(Err.Number), MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error: FPrimary.General.SortListView")
	End Sub
	
	Public Sub GetUserGroups()
		Dim iTemp As Short
		Dim sTempMsg As String
		'
		'
		If iGroupNumber = 15 Then
			lblGroups.Text = "All"
		Else
			iTemp = iGroupNumber
			If iTemp >= 8 Then
				sTempMsg = "Authorizations"
				iTemp = iTemp - 8
			End If
			'
			If iTemp >= 4 Then
				If Len(sTempMsg) > 0 Then
					sTempMsg = sTempMsg & ", Sales"
				Else
					sTempMsg = "Sales"
				End If
				iTemp = iTemp - 4
			End If
			'
			If iTemp >= 2 Then
				If Len(sTempMsg) > 0 Then
					sTempMsg = sTempMsg & ", Support"
				Else
					sTempMsg = "Support"
				End If
				iTemp = iTemp - 2
			End If
			'
			If iTemp >= 1 Then
				If Len(sTempMsg) > 0 Then
					sTempMsg = sTempMsg & ", Operator"
				Else
					sTempMsg = "Operator"
				End If
			End If
			lblGroups.Text = sTempMsg
		End If
		'  Select Case iGroupNumber
		'    Case 1
		'      lblGroups.Caption = "Authorizations"
		'    Case 2
		'      lblGroups.Caption = "Sales"
		'    Case 3
		'      lblGroups.Caption = "Support"
		'    Case 4
		'      lblGroups.Caption = "All"
		'    Case 5
		'      lblGroups.Caption = "Authorizations, Sales"
		'    Case 6
		'      lblGroups.Caption = "Authorizations, Support"
		'    Case 7
		'      lblGroups.Caption = "Sales, Support"
		'    Case Else
		'  End Select
		'
	End Sub
	
	Public Sub ListView1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ListView1.Click
		
		Dim lMessageID As Integer
		'
		If Me.ListView1.Items.Count > 0 Then
			If rsMessages.State <> 0 Then
				rsMessages.Close()
			End If
			'
			lMessageID = CInt(VB.Right(Me.ListView1.FocusedItem.Name, Len(Me.ListView1.FocusedItem.Name) - 1))
			'
			rsMessages = GetMessageRecord(lMessageID)
			'
			With rsMessages
				.MoveFirst()
				'.Find "messageID = " & X
				If Not .eof Then
					sMessageID = lMessageID
					sMessageName = .Fields("MessageName").Value & vbNullString
					'
					txtPhone.Text = .Fields("PhoneNumber").Value & vbNullString
					'
					If .Fields("Caller").Value & vbNullString <> "" Then
						cmbCaller.Text = .Fields("Caller").Value & vbNullString
					Else
						cmbCaller.Text = ""
					End If
					'
					If txtPhone.Text = "" Then
						cmdGetNames.Enabled = False
					Else
						cmdGetNames.Enabled = True
					End If
					'
					sReceived = .Fields("TimeReceived").Value & " " & .Fields("DateReceived").Value & vbNullString
					'
					cmbComment.Text = .Fields("Comments").Value & vbNullString
					'
					If Not .Fields("Caller").Value = "" Then
						sCaller = .Fields("Caller").Value & vbNullString
					Else
						sCaller = .Fields("From").Value & vbNullString
					End If
					'
					txtsubject.Text = .Fields("Subject").Value & vbNullString
					' Body only used here!
					'txtBody.Text = !Body & vbNullString
					'sBody = !Body & vbNullString
					LoadEmailBodyBody()
					'
					If txtBody.Text <> "" Then
						DisplayEmailBody()
					Else
						'UPGRADE_WARNING: Navigate2 was upgraded to Navigate and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						webBody.Navigate(New System.URI("about:blank"))
					End If
					'
					If .Fields("PhoneNumber").Value <> "" Then
						sSubject = .Fields("PhoneNumber").Value & vbNullString
					Else
						sSubject = .Fields("Subject").Value & vbNullString
					End If
					'
					If .Fields("Completed").Value = True Then
						chkComp.CheckState = System.Windows.Forms.CheckState.Checked
					Else
						chkComp.CheckState = System.Windows.Forms.CheckState.Unchecked
					End If
					'
				End If
			End With
			'
			If rsMessages.Fields("MessageName").Value & vbNullString = "" Then
				cmdPlay.Enabled = False
			Else
				cmdPlay.Enabled = True
			End If
			'
			'If ListView1.Visible Then
			'  ListView1.SetFocus
			' End If
		End If
	End Sub
	
	Public Sub FindNames()
		Dim rs As New ADODB.Recordset
		Dim i As Short
		'
		cmbCaller.Items.Clear()
		'
		i = 0
		'
		rs.Open("SELECT [FirstName], [LastName],[Phone1], [Phone2],[fax], [ID]  FROM TContact", cnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockBatchOptimistic)
		
		With rs
			.MoveFirst()
			While Not .eof
				If Trim(.Fields("Phone1").Value) = txtPhone.Text Or Trim(.Fields("Phone2").Value) = txtPhone.Text Or Trim(.Fields("Fax").Value) = txtPhone.Text Then
					cmbCaller.Items.Add(.Fields("FirstName").Value & " " & .Fields("LastName").Value)
					i = i + 1
				End If
				.MoveNext()
			End While
		End With
		'
		If Not i = 0 Then
			cmbCaller.SelectedIndex = 0
		Else
			MsgBox("No Results Found For Phone Number", MsgBoxStyle.Information, "vMail")
		End If
		'
		rs.Close()
		'
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
	End Sub
	
	Public Sub GetColumnWidths()
		Dim rsColumns As New ADODB.Recordset
		'
		rsColumns.Open("Select * from TEmailAddresses where [Name] = '" & StrUser & "'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rsColumns
			If Not .eof Then
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenGroup = nnNum(.Fields("LenGroup"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenMessage = nnNum(.Fields("LenMessage"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenPhone = nnNum(.Fields("LenPhone"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenFrom = nnNum(.Fields("LenFrom"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenSubject = nnNum(.Fields("LenSubject"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenDateRec = nnNum(.Fields("LenDateRec"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenTimeRec = nnNum(.Fields("LenTimeRec"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenMessageNum = nnNum(.Fields("LenMessageNum"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenUser = nnNum(.Fields("LenUser"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenCaller = nnNum(.Fields("LenCaller"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenComments = nnNum(.Fields("LenComments"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenDateCom = nnNum(.Fields("LenDateCom"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iLenTimeCom = nnNum(.Fields("LenTimeCom"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				iFromAddress = nnNum(.Fields("LenFromAddress"))
			End If
		End With
		bLoad = True
		'
	End Sub
	
	Public Sub SaveColumnWidths()
		Dim rsColumns As New ADODB.Recordset
		'
		rsColumns.Open("Select * from TEmailAddresses where [Name] = '" & StrUser & "'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		'
		With rsColumns
			If Not .eof Then
				.Fields("LenGroup").Value = iLenGroup
				.Fields("LenMessage").Value = iLenMessage
				.Fields("LenPhone").Value = iLenPhone
				.Fields("LenFrom").Value = iLenFrom
				.Fields("LenSubject").Value = iLenSubject
				.Fields("LenDateRec").Value = iLenDateRec
				.Fields("LenTimeRec").Value = iLenTimeRec
				.Fields("LenMessageNum").Value = iLenMessageNum
				.Fields("LenUser").Value = iLenUser
				.Fields("LenCaller").Value = iLenCaller
				.Fields("LenComments").Value = iLenComments
				.Fields("LenDateCom").Value = iLenDateCom
				.Fields("LenTimeCom").Value = iLenTimeCom
				.Fields("LenFromAddress").Value = iFromAddress
				.UpdateBatch()
			End If
		End With
		'
	End Sub
End Class