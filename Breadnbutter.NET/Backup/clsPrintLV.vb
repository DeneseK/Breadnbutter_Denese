Option Strict Off
Option Explicit On
Friend Class clsPrintLV
	'---------------------------------------------------------------
	' Copyright ©2002 Veign, All rights reserved
	'---------------------------------------------------------------
	'Type: Class Module
	'Name: clsPrintListView
	'Purpose: Print the contents of a listview control
	'Limitations: See the Full Printing OCX in the Download ActiveX section
	'             of Veign's Website
	'Author: Chris Hanscom
	'Arguments: Reference main method of class (PrintListView)
	'Return Value: none
	'Useage:Dim objPrintLV As clsPrintLV
	'       Set objPrintLV = New clsPrintLV
	'       objPrintLV.PrintListView ListView1, 0.1, 8, "Sample ListView Report", _
	'Portrait, True
	'       Set objPrintLV = Nothing
	'Notes:
	
	
	'Paper Orientation
	Enum Orientation
		Landscape
		Portrait
	End Enum
	
	'Setup constants for printing
	Private Const MIN_COL_SPACING As Single = 0.25
	Private Const GRID_LINE_WIDTH As Short = 1
	Private Const GRID_LINE_COLHEADER_WIDTH As Short = 6
	
	'Store a value for the text height
	Private msngTextHeight As Single
	Private sngHeaderHeight As Single
	
	'Printer Object
	'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Private objPrinter As Printer
	
	'Store the Column Widths
	Private sngColumnWidth() As Single
	
	'Events
	Public Event PrintComplete()
	Public Event PrintError()
	
	Public Sub PrintListView(ByRef lvToPrint As System.Windows.Forms.ListView, ByRef sngRowSpacing As Single, ByRef intItemFontSize As Short, Optional ByRef strHeader As String = "ListView Report", Optional ByRef Orientation As Orientation = Orientation.Portrait, Optional ByRef ShowGrid As Boolean = False, Optional ByRef AutoWidth As Boolean = True)
		
		On Error GoTo Hell
		
		'This sub Prints the data from a listview
		Dim ListItemNo As Short
		Dim ColumnCount As Short
		Dim ItemPerPage As Short
		'
		Dim sText As String
		'
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim prt As Printer
		'UPGRADE_ISSUE: Printers object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		For	Each prt In Printers
			'UPGRADE_ISSUE: Printer property prt.DeviceName was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			If prt.DeviceName = sPrinterName Then
				'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
				Printer = prt
				Exit For
			End If
		Next prt
		
		'Create the Printer object
		'UPGRADE_ISSUE: Printer object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		objPrinter = Printer
		
		
		'Setting paper orientation
		If Orientation = Orientation.Landscape Then
			'UPGRADE_ISSUE: Constant vbPRORLandscape was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_ISSUE: Printer property objPrinter.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			objPrinter.Orientation = vbPRORLandscape
		Else
			'UPGRADE_ISSUE: Constant vbPRORPortrait was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			'UPGRADE_ISSUE: Printer property objPrinter.Orientation was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			objPrinter.Orientation = vbPRORPortrait
		End If
		
		'Setup the printer
		'UPGRADE_ISSUE: Constant vbInches was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Printer property objPrinter.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		objPrinter.ScaleMode = vbInches
		
		'Retrieve the Column Widths
		ColumnWidth(lvToPrint, intItemFontSize)
		
		'Print the Header
		PageHeader(lvToPrint)
		
		'Set the Items font size
		'UPGRADE_ISSUE: Printer property objPrinter.FontSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		objPrinter.FontSize = intItemFontSize
		
		'Grab the height of the font
		'UPGRADE_ISSUE: Printer method objPrinter.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		msngTextHeight = objPrinter.TextHeight("V")
		
		'Used to shift each line item down (set to first line)
		ItemPerPage = 1
		
		'Print Listview Items
		Dim lvListItem As System.Windows.Forms.ListViewItem
		With objPrinter
			'UPGRADE_ISSUE: Printer property objPrinter.FontBold was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			.FontBold = False
			
			
			For ListItemNo = 1 To lvToPrint.Items.Count
				'Display Grid Lines
				If ShowGrid Then
					'UPGRADE_ISSUE: Printer property objPrinter.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					.CurrentX = 0
					'UPGRADE_ISSUE: Printer property objPrinter.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					.CurrentY = (sngHeaderHeight + (msngTextHeight * ItemPerPage)) + (sngRowSpacing * ItemPerPage) - (sngRowSpacing / 2)
					
					'Set the draw width of the line
					'UPGRADE_ISSUE: Printer property objPrinter.DrawWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					.DrawWidth = IIf(ListItemNo = 1, GRID_LINE_COLHEADER_WIDTH, GRID_LINE_WIDTH)
					
					'UPGRADE_ISSUE: Printer property objPrinter.ScaleWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					'UPGRADE_ISSUE: Printer method objPrinter.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					objPrinter.Line((0, 0) - (.ScaleWidth, 0))
				End If
				
				'Set to start point
				'UPGRADE_ISSUE: Printer property objPrinter.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				.CurrentX = 0 '0.25
				
				'Set the current Listitem
				'UPGRADE_WARNING: Lower bound of collection lvToPrint.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				lvListItem = lvToPrint.Items.Item(ListItemNo)
				
				'Print line of data (all columns)
				For ColumnCount = 1 To lvToPrint.Columns.Count
					'.CurrentY = sngRowSpacing * 3 + (sngRowSpacing * ItemPerPage)
					'UPGRADE_ISSUE: Printer property objPrinter.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					.CurrentY = sngHeaderHeight + (msngTextHeight * ItemPerPage) + (sngRowSpacing * ItemPerPage) ' Added TextHeight value
					
					If ColumnCount = 1 Then
						sText = lvListItem.Text
						'objPrinter.Print lvListItem.Text
					Else
						'UPGRADE_WARNING: Lower bound of collection lvListItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						sText = lvListItem.SubItems(ColumnCount - 1).Text
						'Print the line as is
						'objPrinter.Print lvListItem.SubItems(ColumnCount - 1)
					End If
					'
					'UPGRADE_ISSUE: Printer method objPrinter.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					'UPGRADE_ISSUE: Printer property objPrinter.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					While .CurrentX + .TextWidth(sText) + 0.1 > sngColumnWidth(ColumnCount)
						sText = Left(sText, Len(sText) - 1)
					End While
					'
					'UPGRADE_ISSUE: Printer method objPrinter.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					objPrinter.Print(sText)
					'
					'UPGRADE_ISSUE: Printer property objPrinter.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					.CurrentX = sngColumnWidth(ColumnCount)
				Next 
				
				'Check if exceeding a page
				'UPGRADE_ISSUE: Printer property objPrinter.ScaleHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				'UPGRADE_ISSUE: Printer method objPrinter.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				'UPGRADE_ISSUE: Printer property objPrinter.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				If (.CurrentY + .TextHeight("Test") + 1) > .ScaleHeight Then
					
					'Print Grid Lines if requested
					If ShowGrid Then
						DrawVerticalGridLine()
					End If
					
					'Increment the page
					'UPGRADE_ISSUE: Printer method objPrinter.NewPage was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					.NewPage()
					
					'Print the Header
					PageHeader(lvToPrint)
					
					'ReSet the Items font size
					'UPGRADE_ISSUE: Printer property objPrinter.FontSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					objPrinter.FontSize = intItemFontSize
					
					'Reset the ItemPerPage counter
					ItemPerPage = 1
				Else
					'Increment the item per page
					ItemPerPage = 1 + ItemPerPage
				End If
			Next 
			
			'Clear the listitem object
			'UPGRADE_NOTE: Object lvListItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			lvListItem = Nothing
			
		End With
		
		'Print Grid Lines if requested
		If ShowGrid Then
			DrawVerticalGridLine()
		End If
		
		'Send to the printer
		'UPGRADE_ISSUE: Printer method objPrinter.EndDoc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		objPrinter.EndDoc()
		
		'Destroy the print object
		'UPGRADE_NOTE: Object objPrinter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPrinter = Nothing
		
		'Raise event that the print completed
		RaiseEvent PrintComplete()
		
		Exit Sub
		
Hell: 
		'Destroy print job
		'UPGRADE_ISSUE: Printer method objPrinter.KillDoc was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		objPrinter.KillDoc()
		'UPGRADE_NOTE: Object objPrinter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objPrinter = Nothing
		
		'Raise error event
		RaiseEvent PrintError()
		
	End Sub
	
	Private Function PageHeader(ByRef lvPrint As System.Windows.Forms.ListView) As Single
		Dim sText As String
		'
		'Print Listview Column Headers
		Dim ColumnCount As Short
		With objPrinter
			'UPGRADE_ISSUE: Printer method objPrinter.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			'UPGRADE_ISSUE: Printer property objPrinter.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			.CurrentY = (.TextHeight("A") + 0.5)
			'UPGRADE_ISSUE: Printer property objPrinter.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			.CurrentX = 0 '0.25
			'UPGRADE_ISSUE: Printer property objPrinter.FontSize was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			.FontSize = 8
			
			For ColumnCount = 1 To lvPrint.Columns.Count
				'UPGRADE_WARNING: Lower bound of collection lvPrint.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				sText = lvPrint.Columns.Item(ColumnCount).Text
				'
				'UPGRADE_ISSUE: Printer method objPrinter.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				'UPGRADE_ISSUE: Printer property objPrinter.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				While .CurrentX + .TextWidth(sText) + 0.1 > sngColumnWidth(ColumnCount)
					sText = Left(sText, Len(sText) - 1)
				End While
				'
				'UPGRADE_ISSUE: Printer method objPrinter.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				objPrinter.Print(sText) 'lvPrint.ColumnHeaders.Item(ColumnCount)
				'UPGRADE_ISSUE: Printer property objPrinter.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				.CurrentX = sngColumnWidth(ColumnCount)
				'UPGRADE_ISSUE: Printer method objPrinter.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				'UPGRADE_ISSUE: Printer property objPrinter.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				.CurrentY = (.TextHeight("A") + 0.5)
			Next 
		End With
		
		'Return and set the Header Height
		'UPGRADE_ISSUE: Printer property objPrinter.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		sngHeaderHeight = objPrinter.CurrentY + 0.1
		PageHeader = sngHeaderHeight
		
	End Function
	
	Private Sub ColumnWidth(ByRef objListView As System.Windows.Forms.ListView, ByRef intFontSize As Short)
		
		Dim ColumnCount As Short
		
		'Set the array to match column header qty
		ReDim sngColumnWidth(objListView.Columns.Count)
		
		'Set the column widths (evenly spaced) based on # of columns
		For ColumnCount = 1 To objListView.Columns.Count
			'sngColumnWidth(ColumnCount) = (objPrinter.ScaleWidth / objListView.ColumnHeaders.Count) * ColumnCount
			If ColumnCount = 1 Then
				'UPGRADE_WARNING: Lower bound of collection objListView.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				sngColumnWidth(ColumnCount) = (VB6.PixelsToTwipsX(objListView.Columns.Item(ColumnCount).Width) / 1440) '+ 0.25
			Else
				'UPGRADE_WARNING: Lower bound of collection objListView.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				sngColumnWidth(ColumnCount) = (VB6.PixelsToTwipsX(objListView.Columns.Item(ColumnCount).Width) / 1440) + sngColumnWidth(ColumnCount - 1)
			End If
		Next 
		
	End Sub
	
	Private Sub DrawVerticalGridLine()
		
		'Draws the vertical grid lines if requested
		'  Horizontal lines are drawn in-line with the items being printed
		
		'Skip the last column
		Dim ColumnCount As Short
		For ColumnCount = 1 To UBound(sngColumnWidth) - 1
			'UPGRADE_ISSUE: Printer property objPrinter.ScaleHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			'UPGRADE_ISSUE: Printer method objPrinter.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			objPrinter.Line((sngColumnWidth(ColumnCount) - MIN_COL_SPACING / 4, sngHeaderHeight) - (sngColumnWidth(ColumnCount) - MIN_COL_SPACING / 4, objPrinter.ScaleHeight))
		Next 
		
	End Sub
End Class