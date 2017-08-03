VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RCallLog 
   Caption         =   "Breadnbutter - RCallLog (ActiveReport)"
   ClientHeight    =   8235
   ClientLeft      =   2010
   ClientTop       =   2055
   ClientWidth     =   11565
   _ExtentX        =   20399
   _ExtentY        =   14526
   SectionData     =   "RCallLog.dsx":0000
End
Attribute VB_Name = "RCallLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rslabels As ADODB.Recordset
Dim IntCount As Integer
Dim IntSum As Double
Dim IntAvg As Integer
Dim DTStart As Date
Dim DTEnd As Date
Dim intWorkgroup As Integer
Dim strSQL As String
Dim strDirection As String

Public Sub GetData(pStart As Date, pEnd As Date, pExt As Integer, pSQL As String, pDir As String)
  DTStart = pStart
  DTEnd = pEnd
  intWorkgroup = pExt
  strSQL = pSQL
  strDirection = pDir
  '
  Me.Show
End Sub

Private Sub ActiveReport_ReportEnd()
    rslabels.Close
    Set rslabels = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
  Dim sGroup As String
  '
  Set rslabels = New Recordset
  lblHead.Caption = strDirection & "calls to Ext# " & intWorkgroup & " between " & DTStart & " and " & DTEnd
  rslabels.Open strSQL, cnMain, adOpenForwardOnly, adLockReadOnly
  
  Set DataControl1.Recordset = rslabels
  
End Sub

Private Sub Detail_Format()
  Select Case rslabels!Direction
    Case "2"
      fldCallDir.Text = "Incoming"
    Case "4"
      fldCallDir.Text = "Outgoing"
    Case Else
      fldCallDir.Text = "Other"
  End Select
    Select Case rslabels!Phone
    Case "P"
      fldPhoneNum.Text = "Private"
    Case "O"
      fldPhoneNum.Text = "Unknown"
  End Select
Dim mySeconds As Variant
Dim myHours As Variant
Dim myMinutes As Variant
Dim temp As ADODB.Recordset
    ' //duration '
    Set temp = DataControl1.Recordset
    mySeconds = temp!Duration
    
        'average
        IntCount = IntCount + 1
        IntSum = IntSum + mySeconds
        IntAvg = IntSum / IntCount

    'if mySeconds is greater or equal to 3,600 seconds
    If mySeconds >= 3600 Then
     'get hours which is equal to seconds divided by 3600
        myHours = mySeconds / 3600
      'set the seconds to the numbers after the decimal sign
      'thats what mod does
        mySeconds = mySeconds Mod 3600
    Else
    'if not greater than 3600, just set it to 0
        myHours = 0
    End If

    If mySeconds >= 60 Then
    'greater than or equal to 60
    'set the minutes equal to the value of (seconds divided by 60).
    'and get the remaining numbers after the decimal
    'which will be the seconds
     'using the mod sign
        myMinutes = mySeconds \ 60
        mySeconds = mySeconds Mod 60
    Else
    'if not set to 0
        myMinutes = 0
    End If
    
    fldCallDur.Text = Int(myHours) & ":" & Format$(myMinutes, "00") & ":" & Format$(mySeconds, "00")
    
End Sub

Private Sub ReportFooter_Format()
Dim mySeconds As Variant
Dim myHours As Variant
Dim myMinutes As Variant
    
    mySeconds = IntAvg
    
    If mySeconds >= 3600 Then

        myHours = mySeconds / 3600

        mySeconds = mySeconds Mod 3600
    Else

        myHours = 0
    End If

    If mySeconds >= 60 Then

        myMinutes = mySeconds \ 60
        mySeconds = mySeconds Mod 60
    Else
  
        myMinutes = 0
    End If
    
    lblAvgTime.Caption = "The average call duration between " & DTStart & " and " & DTEnd & " was " & Int(myHours) & ":" & Format(myMinutes, "00") & ":" & Format(mySeconds, "00")

End Sub

