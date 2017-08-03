Option Strict Off
Option Explicit On
Friend Class CContactList
	
	Public Function FillList(ByRef pContactList As System.Windows.Forms.ListView, ByRef rs As ADODB.Recordset) As Integer
		On Error GoTo EH
		'
		'list = pContactList
		Dim iStartField As Short
		Dim strKey As String
		Dim iLineCount As Short
		Dim iFieldPos As Short
		Dim iTotalCharacters As Integer
		Dim iIcon As Short
		Dim lColor As Integer
		Dim sColumnName As String
		'
		Dim sPreviousText As String
		'
		iStartField = 3
		'
		If Not (pContactList.FocusedItem Is Nothing) Then
			sPreviousText = pContactList.FocusedItem.Name
		End If
		'
		pContactList.Items.Clear()
		pContactList.Columns.Clear()
		'
		iTotalCharacters = 0
		iLineCount = 0
		iFieldPos = iStartField
		With rs
			If .RecordCount > 0 Then
				Do 
					'
					iTotalCharacters = 0
					.MoveFirst()
					Do 
						iTotalCharacters = Len(CStr(.Fields(iFieldPos).Value & vbNullString)) + iTotalCharacters
						.MoveNext()
					Loop Until .eof
					'
					'        Select Case .Fields(iFieldPos).Name
					'         Case "LastName"
					'          sColumnName = "Last"
					'         Case "FirstName"
					'          sColumnName = "First"
					'         Case "AuthRemaining"
					'          sColumnName = "Days"
					'        Case Else
					sColumnName = .Fields.Item(iFieldPos).Name
					'        End Select
					'
					pContactList.Columns.Add("w1" & iFieldPos, sColumnName, CInt(VB6.TwipsToPixelsX(400 + ((iTotalCharacters / .RecordCount) * 100))))
					iFieldPos = iFieldPos + 1
					'
				Loop Until iFieldPos = .Fields.Count
				'
				.MoveFirst()
				'
				iFieldPos = iStartField
				'
				iLineCount = 0
				'
				Do Until .eof
					Select Case .Fields("ContactType").Value
						Case 0 ' unknown
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
					Select Case .Fields("Status").Value
						'
						Case "Customer"
							lColor = &HC00000
						Case "Prospect"
							lColor = &HC000
						Case "Future Prospect"
							lColor = &HC000
						Case "Inactive"
							lColor = &H404040
						Case "Contact"
							lColor = &HC000C0
						Case Else
							lColor = &H0s
					End Select
					'
					strKey = "ID" & VB6.Format(.Fields("ID").Value)
					'UPGRADE_WARNING: Lower bound of collection pContactList.ListItems.ImageList has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					'UPGRADE_WARNING: Image property was not upgraded Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="D94970AE-02E7-43BF-93EF-DCFCD10D27B5"'
					pContactList.Items.Add(strKey, .Fields(iFieldPos).Value & vbNullString, iIcon)
					pContactList.Items.Item(strKey).ForeColor = System.Drawing.ColorTranslator.FromOle(lColor)
					iFieldPos = iFieldPos + 1
					Do 
						pContactList.Items.Item(strKey).SubItems.Add(.Fields(iFieldPos).Value & vbNullString).ForeColor = System.Drawing.ColorTranslator.FromOle(lColor)
						iFieldPos = iFieldPos + 1
					Loop Until iFieldPos = .Fields.Count
					iFieldPos = iStartField
					.MoveNext()
					iLineCount = iLineCount + 1
				Loop 
			End If
		End With
		FillList = iLineCount
		'
		Dim itmFound As System.Windows.Forms.ListViewItem ' FoundItem variable.
		If sPreviousText <> "" Then
			'Debug.Print pContactList.ListItems(sPreviousText).
			If VerifyKeyInList(pContactList, sPreviousText) Then
				itmFound = pContactList.Items.Item(sPreviousText)
				'  '
				If Not (itmFound Is Nothing) Then
					'UPGRADE_WARNING: MSComctlLib.ListItem method itmFound.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					itmFound.EnsureVisible()
					itmFound.Selected = True
					
					' Set pContactList.SelectedItem = pContactList.ListItems(sPreviousText)
				End If
			End If
		End If
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in FillList.")
	End Function
	
	Private Function VerifyKeyInList(ByRef pContactList As System.Windows.Forms.ListView, ByVal psKey As String) As Boolean
		Dim i As Short
		'
		VerifyKeyInList = False
		'
		For i = 1 To pContactList.Items.Count
			'UPGRADE_WARNING: Lower bound of collection pContactList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If pContactList.Items.Item(i).Name = psKey Then VerifyKeyInList = True
		Next i
	End Function
End Class