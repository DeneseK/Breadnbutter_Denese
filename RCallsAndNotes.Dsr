VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RCallsAndNotes 
   Caption         =   "Breadnbutter - RCallsAndNotes (ActiveReport)"
   ClientHeight    =   8235
   ClientLeft      =   2040
   ClientTop       =   1635
   ClientWidth     =   12795
   _ExtentX        =   22569
   _ExtentY        =   14526
   SectionData     =   "RCallsAndNotes.dsx":0000
End
Attribute VB_Name = "RCallsAndNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iRow As Integer
Private sName() As String
Private iNumofCalls() As Integer
Private iBNBNotes() As Integer
Private sCallTime() As String
Private iFollowups() As Integer
Private iWalkThroughs() As Integer
Private iSales() As Integer
Private sAvgCallTime() As String
Private iNumofUsers As Integer

Public Sub GetData(Name() As String, Calls() As Integer, Notes() As Integer, CallTime() As String, Followup() As Integer, WalkThroughs() As Integer, Sales() As Integer, AvgCallTime() As String, Users As Integer, Date1 As String, Date2 As String)
  sName = Name
  iNumofCalls = Calls
  iBNBNotes = Notes
  sCallTime = CallTime
  iFollowups = Followup
  iWalkThroughs = WalkThroughs
  iSales = Sales
  sAvgCallTime = AvgCallTime
  iNumofUsers = Users
  lblDate1.Caption = Date1
  lblDate2.Caption = Date2
  '
  'ActiveReport_ReportStart
  '
  Me.Show
End Sub

Private Sub ActiveReport_DataInitialize()
    Fields.Add "txtName"
    Fields.Add "txtCalls"
    Fields.Add "txtNotes"
    Fields.Add "Percent"
    Fields.Add "txtTotalCallTime"
    Fields.Add "txtFollowup"
    Fields.Add "txtWalkThroughs"
    Fields.Add "txtSales"
    Fields.Add "txtAvgCallTime"
    ' iRow is the current record pointer
    iRow = LBound(sName)
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim iHolder As Integer
  Dim sHolder As String
     ' If we processed all element then we exit the event
     ' after setting the eof parameter to True
     ' this will promp AR to end the report
     If iRow > UBound(sName) Then
      eof = True
      Exit Sub
    End If
    '
    If iNumofCalls(iRow) = 0 Or iBNBNotes(iRow) = 0 Then
      If iBNBNotes(iRow) < 0 Then
        iHolder = 100
      Else
        iHolder = 0
      End If
    Else
      iHolder = (iBNBNotes(iRow) / iNumofCalls(iRow)) * 100
    End If
    sHolder = iHolder & "%"
    '
    Fields("txtName") = sName(iRow)
    Fields("txtCalls") = iNumofCalls(iRow)
    Fields("txtNotes") = iBNBNotes(iRow)
    Fields("Percent") = sHolder
    Fields("txtTotalCallTime") = sCallTime(iRow)
    Fields("txtFollowup") = iFollowups(iRow)
    Fields("txtWalkThroughs") = iWalkThroughs(iRow)
    Fields("txtSales") = iSales(iRow)
    Fields("txtAvgCallTime") = sAvgCallTime(iRow)
    iRow = iRow + 1
    ' We must set the eof parameter to True as
    ' long as there is more data to be processed
    eof = False
End Sub

