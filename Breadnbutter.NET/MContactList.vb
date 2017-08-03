Option Strict Off
Option Explicit On
Module MContactList
	
	Public bAscDesc As Boolean
	Public iColNum As Short
	Public objFind As LV_FINDINFO
	Public objItem As LV_ITEM
	'
	'variable to hold the sort order (ascending or descending)
	Public sOrder As Boolean
	'
	Public Structure POINTAPI
		Dim X As Integer
		Dim y As Integer
	End Structure
	'
	Public Structure LV_FINDINFO
		Dim FLAGS As Integer
		Dim psz As String
		Dim lParam As Integer
		Dim pt As POINTAPI
		Dim vkDirection As Integer
	End Structure
	'
	Public Structure LV_ITEM
		Dim mask As Integer
		Dim iItem As Integer
		Dim iSubItem As Integer
		Dim State As Integer
		Dim stateMask As Integer
		Dim pszText As String
		Dim cchTextMax As Integer
		Dim iImage As Integer
		Dim lParam As Integer
		Dim iIndent As Integer
	End Structure
	'
	'Constants
	Public Const LVFI_PARAM As Integer = &H1s
	Public Const LVIF_TEXT As Integer = &H1s
	Public Const LVM_FIRST As Integer = &H1000s
	Public Const LVM_FINDITEM As Integer = (LVM_FIRST + 13)
	Public Const LVM_GETITEMTEXT As Integer = (LVM_FIRST + 45)
	Public Const LVM_SORTITEMS As Integer = (LVM_FIRST + 48)
	'
	'API declarations
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function SendMessage Lib "USER32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	'
	Public Function CompareDates(ByVal lParam1 As Integer, ByVal lParam2 As Integer, ByVal hWnd As Integer) As Integer
		On Error GoTo EH
		'CompareDates: This is the sorting routine that gets passed to the
		'ListView control to provide the comparison test for date values.
		
		'Compare returns:
		' 0 = Less Than
		' 1 = Equal
		' 2 = Greater Than
		
		Dim dDate1 As Date
		Dim dDate2 As Date
		
		'Obtain the item names and dates corresponding to the
		'input parameters
		dDate1 = ListView_GetItemDate(hWnd, lParam1)
		dDate2 = ListView_GetItemDate(hWnd, lParam2)
		
		'based on the Public variable sOrder set in the
		'columnheader click sub, sort the dates appropriately:
		Select Case sOrder
			Case True 'sort descending
				
				If dDate1 < dDate2 Then
					CompareDates = 0
				ElseIf dDate1 = dDate2 Then 
					CompareDates = 1
				Else : CompareDates = 2
				End If
				
			Case Else 'sort ascending
				
				If dDate1 > dDate2 Then
					CompareDates = 0
				ElseIf dDate1 = dDate2 Then 
					CompareDates = 1
				Else : CompareDates = 2
				End If
				
		End Select
		Exit Function
EH: 
		MsgBox(Err.Description & " in MUtility.CompareDates.")
	End Function
	
	
	Public Function CompareValues(ByVal lParam1 As Integer, ByVal lParam2 As Integer, ByVal hWnd As Integer) As Integer
		On Error GoTo EH
		'CompareValues: This is the sorting routine that gets passed to the
		'ListView control to provide the comparison test for numeric values.
		
		'Compare returns:
		' 0 = Less Than
		' 1 = Equal
		' 2 = Greater Than
		
		Dim val1 As Integer
		Dim val2 As Integer
		
		'Obtain the item names and values corresponding
		'to the input parameters
		val1 = ListView_GetItemValueStr(hWnd, lParam1)
		val2 = ListView_GetItemValueStr(hWnd, lParam2)
		
		'based on the Public variable sOrder set in the
		'columnheader click sub, sort the values appropriately:
		Select Case sOrder
			Case True 'sort descending
				
				If val1 < val2 Then
					CompareValues = 0
				ElseIf val1 = val2 Then 
					CompareValues = 1
				Else : CompareValues = 2
				End If
				
			Case Else 'sort ascending
				
				If val1 > val2 Then
					CompareValues = 0
				ElseIf val1 = val2 Then 
					CompareValues = 1
				Else : CompareValues = 2
				End If
				
		End Select
		Exit Function
EH: 
		MsgBox(Err.Description & " in MUtility.CompareValues.")
	End Function
	
	Public Function ListView_GetItemDate(ByRef hWnd As Integer, ByRef lParam As Integer) As Date
		On Error GoTo EH
		Dim hIndex As Integer
		Dim r As Integer
		
		'Convert the input parameter to an index in the list view
		objFind.FLAGS = LVFI_PARAM
		objFind.lParam = lParam
		'UPGRADE_WARNING: Couldn't resolve default property of object objFind. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
		
		'Obtain the value of the specified list view item.
		'The objItem.iSubItem member is set to the index
		'of the column that is being retrieved.
		objItem.mask = LVIF_TEXT
		objItem.iSubItem = 3
		objItem.pszText = Space(32)
		objItem.cchTextMax = Len(objItem.pszText)
		
		'get the string at subitem 1
		'and convert it into a date and exit
		'UPGRADE_WARNING: Couldn't resolve default property of object objItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
		If r > 0 Then
			ListView_GetItemDate = CDate(Left(objItem.pszText, r))
		End If
		Exit Function
EH: 
		MsgBox(Err.Description & " in MUtility.ListView_GetItemDate.")
	End Function
	
	
	Public Function ListView_GetItemValueStr(ByRef hWnd As Integer, ByRef lParam As Integer) As Integer
		On Error GoTo EH
		Dim hIndex As Integer
		Dim r As Integer
		
		'Convert the input parameter to an index in the list view
		objFind.FLAGS = LVFI_PARAM
		objFind.lParam = lParam
		'UPGRADE_WARNING: Couldn't resolve default property of object objFind. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
		
		'Obtain the value of the specified list view item.
		'The objItem.iSubItem member is set to the index
		'of the column that is being retrieved.
		objItem.mask = LVIF_TEXT
		objItem.iSubItem = 2
		objItem.pszText = Space(32)
		objItem.cchTextMax = Len(objItem.pszText)
		
		'get the string at subitem 2
		'and convert it into a long
		'UPGRADE_WARNING: Couldn't resolve default property of object objItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
		If r > 0 Then
			ListView_GetItemValueStr = CInt(Left(objItem.pszText, r))
		End If
		Exit Function
EH: 
		MsgBox(Err.Description & " in MUtility.ListView_GetItemValueStr.")
	End Function
	
	Public Function FARPROC(ByVal pfn As Integer) As Integer
		On Error GoTo EH
		'
		FARPROC = pfn
		'
		Exit Function
EH: 
		MsgBox(Err.Description & " in MUtility.FarProc.")
	End Function
	
	'Private Sub Contactlist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
	'  '
	'  On Error GoTo EH
	'  '
	'  sOrder = Not sOrder
	'  '
	'  ContactList.SortKey = ColumnHeader.Index - 1
	'  '
	'  Select Case ColumnHeader.Index - 1
	'      Case DateColumn:
	'      'Use sort routine to sort by date
	'      ContactList.Sorted = False
	'      SendMessage ContactList.hWnd, _
	''                  LVM_SORTITEMS, _
	''                  ContactList.hWnd, _
	''                  ByVal FARPROC(AddressOf CompareDates)
	'      Case DaysColumn:
	'        'Use sort routine to sort by value
	'        ContactList.Sorted = False
	'        SendMessage ContactList.hWnd, _
	''                   LVM_SORTITEMS, _
	''                   ContactList.hWnd, _
	''                   ByVal FARPROC(AddressOf CompareValues)
	'      Case Else:
	'        'Use default sorting to sort the items in the list
	'        ContactList.SortKey = 0
	'        ContactList.SortOrder = Abs(sOrder) '=Abs(Not ListView1.SortOrder = 1)
	'        ContactList.Sorted = True
	'   End Select
	'   '
	'   Exit Function
	'EH:
	' MsgBox Err.Description & " in ContactList_ColumnClick."
	'End Function
	'
	'
End Module