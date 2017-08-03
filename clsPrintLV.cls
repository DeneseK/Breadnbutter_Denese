VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintLV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
                    Portrait, True
'       Set objPrintLV = Nothing
'Notes:

Option Explicit

'Paper Orientation
Enum Orientation
    Landscape
    Portrait
End Enum

'Setup constants for printing
Private Const MIN_COL_SPACING As Single = 0.25
Private Const GRID_LINE_WIDTH As Integer = 1
Private Const GRID_LINE_COLHEADER_WIDTH As Integer = 6

'Store a value for the text height
Private msngTextHeight As Single
Private sngHeaderHeight As Single

'Printer Object
Private objPrinter As Printer

'Store the Column Widths
Private sngColumnWidth() As Single

'Events
Public Event PrintComplete()
Public Event PrintError()

Public Sub PrintListView(lvToPrint As ListView, sngRowSpacing As Single, _
                    intItemFontSize As Integer, _
                    Optional strHeader As String = "ListView Report", _
                    Optional Orientation As Orientation = Portrait, _
                    Optional ShowGrid As Boolean = False, _
                    Optional AutoWidth As Boolean = True)

On Error GoTo Hell

'This sub Prints the data from a listview
Dim ListItemNo As Integer
Dim ColumnCount As Integer
Dim ItemPerPage As Integer
'
Dim sText As String
'
Dim prt As Printer
For Each prt In Printers
    If prt.DeviceName = sPrinterName Then
        Set Printer = prt
        Exit For
    End If
Next
    
'Create the Printer object
Set objPrinter = Printer


'Setting paper orientation
If Orientation = Landscape Then
    objPrinter.Orientation = vbPRORLandscape
Else
    objPrinter.Orientation = vbPRORPortrait
End If

'Setup the printer
objPrinter.ScaleMode = vbInches

'Retrieve the Column Widths
ColumnWidth lvToPrint, intItemFontSize

'Print the Header
PageHeader lvToPrint

'Set the Items font size
objPrinter.FontSize = intItemFontSize

'Grab the height of the font
msngTextHeight = objPrinter.TextHeight("V")

'Used to shift each line item down (set to first line)
ItemPerPage = 1

'Print Listview Items
With objPrinter
    .FontBold = False
    
    Dim lvListItem As MSComctlLib.ListItem
    
    For ListItemNo = 1 To lvToPrint.ListItems.Count
        'Display Grid Lines
        If ShowGrid Then
            .CurrentX = 0
            .CurrentY = (sngHeaderHeight + (msngTextHeight * ItemPerPage)) + _
                        (sngRowSpacing * ItemPerPage) - (sngRowSpacing / 2)

            'Set the draw width of the line
            .DrawWidth = IIf(ListItemNo = 1, GRID_LINE_COLHEADER_WIDTH, _
                                       GRID_LINE_WIDTH)

            objPrinter.Line Step(0, 0)-Step(.ScaleWidth, 0)
        End If
        
        'Set to start point
        .CurrentX = 0 '0.25
        
        'Set the current Listitem
        Set lvListItem = lvToPrint.ListItems(ListItemNo)
        
        'Print line of data (all columns)
        For ColumnCount = 1 To lvToPrint.ColumnHeaders.Count
            '.CurrentY = sngRowSpacing * 3 + (sngRowSpacing * ItemPerPage)
            .CurrentY = sngHeaderHeight + (msngTextHeight * ItemPerPage) + _
                        (sngRowSpacing * ItemPerPage) ' Added TextHeight value
        
            If ColumnCount = 1 Then
                sText = lvListItem.Text
                'objPrinter.Print lvListItem.Text
            Else
                sText = lvListItem.SubItems(ColumnCount - 1)
                'Print the line as is
                'objPrinter.Print lvListItem.SubItems(ColumnCount - 1)
            End If
            '
            While .CurrentX + .TextWidth(sText) + 0.1 > sngColumnWidth(ColumnCount)
              sText = Left(sText, Len(sText) - 1)
            Wend
            '
            objPrinter.Print sText
            '
            .CurrentX = sngColumnWidth(ColumnCount)
        Next
        
        'Check if exceeding a page
        If (.CurrentY + .TextHeight("Test") + 1) > .ScaleHeight Then
            
            'Print Grid Lines if requested
            If ShowGrid Then
                DrawVerticalGridLine
            End If
        
            'Increment the page
            .NewPage
            
            'Print the Header
            PageHeader lvToPrint
            
            'ReSet the Items font size
            objPrinter.FontSize = intItemFontSize
            
            'Reset the ItemPerPage counter
            ItemPerPage = 1
        Else
            'Increment the item per page
            ItemPerPage = 1 + ItemPerPage
        End If
    Next
    
    'Clear the listitem object
    Set lvListItem = Nothing
    
End With

'Print Grid Lines if requested
If ShowGrid Then
    DrawVerticalGridLine
End If

'Send to the printer
objPrinter.EndDoc

'Destroy the print object
Set objPrinter = Nothing

'Raise event that the print completed
RaiseEvent PrintComplete

Exit Sub

Hell:
    'Destroy print job
    objPrinter.KillDoc
    Set objPrinter = Nothing
    
    'Raise error event
    RaiseEvent PrintError

End Sub

Private Function PageHeader(lvPrint As ListView) As Single
  Dim sText As String
  '
  'Print Listview Column Headers
  With objPrinter
      .CurrentY = (.TextHeight("A") + 0.5)
      .CurrentX = 0 '0.25
      .FontSize = 8
      
      Dim ColumnCount As Integer
      For ColumnCount = 1 To lvPrint.ColumnHeaders.Count
          sText = lvPrint.ColumnHeaders.Item(ColumnCount)
          '
          While .CurrentX + .TextWidth(sText) + 0.1 > sngColumnWidth(ColumnCount)
            sText = Left(sText, Len(sText) - 1)
          Wend
          '
          objPrinter.Print sText 'lvPrint.ColumnHeaders.Item(ColumnCount)
          .CurrentX = sngColumnWidth(ColumnCount)
          .CurrentY = (.TextHeight("A") + 0.5)
      Next
  End With
  
  'Return and set the Header Height
  sngHeaderHeight = objPrinter.CurrentY + 0.1
  PageHeader = sngHeaderHeight

End Function

Private Sub ColumnWidth(objListView As ListView, intFontSize As Integer)

Dim ColumnCount As Integer

'Set the array to match column header qty
ReDim sngColumnWidth(objListView.ColumnHeaders.Count)

'Set the column widths (evenly spaced) based on # of columns
For ColumnCount = 1 To objListView.ColumnHeaders.Count
    'sngColumnWidth(ColumnCount) = (objPrinter.ScaleWidth / objListView.ColumnHeaders.Count) * ColumnCount
    If ColumnCount = 1 Then
      sngColumnWidth(ColumnCount) = (objListView.ColumnHeaders.Item(ColumnCount).Width / 1440) '+ 0.25
    Else
      sngColumnWidth(ColumnCount) = (objListView.ColumnHeaders.Item(ColumnCount).Width / 1440) + sngColumnWidth(ColumnCount - 1)
    End If
Next

End Sub

Private Sub DrawVerticalGridLine()

'Draws the vertical grid lines if requested
'  Horizontal lines are drawn in-line with the items being printed

'Skip the last column
Dim ColumnCount As Integer
For ColumnCount = 1 To UBound(sngColumnWidth) - 1
   objPrinter.Line (sngColumnWidth(ColumnCount) - MIN_COL_SPACING / 4, sngHeaderHeight)- _
            (sngColumnWidth(ColumnCount) - MIN_COL_SPACING / 4, objPrinter.ScaleHeight)
Next

End Sub



