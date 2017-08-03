VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RLabels 
   Caption         =   "Mailing Labels"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   _ExtentX        =   18336
   _ExtentY        =   10001
   SectionData     =   "RLabels.dsx":0000
End
Attribute VB_Name = "RLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sheet(20) As Integer
Private OneOrAll As Integer
'
Dim Count As Integer
Dim rslabels As New ADODB.Recordset
'
Public Sub SetPages(pValue As Integer)
  OneOrAll = pValue
End Sub
'
Public Sub SetSheet(pValue As Variant)
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   Dim i As Integer
110   For i = 1 To 20
120     Sheet(i) = pValue(i)
130   Next i
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "RLabels", "SetSheet", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub ActiveReport_ReportEnd()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   rslabels.Close
110   Set rslabels = Nothing
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "RLabels", "ActiveReport_ReportEnd", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Private Sub ActiveReport_ReportStart()
  Count = 1
End Sub

'
Private Sub Detail_Format()
    '
      'If FPrintLabels.chkPageNum.value = 1 Then
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   If OneOrAll = 1 Then
110     OneDetail
120   Else
130     AllDetail
140   End If
     '
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "RLabels", "Detail_Format", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
Private Sub OneDetail()
  
   If Count <= 20 Then
        'If FPrintLabels.cmdLabel(Count).Caption = "Print" Then
        If Sheet(Count) = True Then
          With rslabels
            If Not .eof Then
              If nnNum(!PreferredAddress) = 0 Then
                EnterShipping
              Else
                EnterMailing
              End If
              Detail.PrintSection
              .MoveNext
            End If
          End With 'rsLabels
        Else
          EnterBlank
          Detail.PrintSection
        End If
    Else
     
    End If
  Count = Count + 1
 
End Sub

Private Sub CurrentDetail()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
  
100    If Count <= 20 Then
            'If FPrintLabels.cmdLabel(Count).Caption = "Print" Then
110         If Sheet(Count) = True Then
120           With rslabels
130             If Not .eof Then
140               If nnNum(!PreferredAddress) = 0 Then
150                 EnterShipping
160               Else
170                 EnterMailing
180               End If
190               Detail.PrintSection
200               .MoveNext
210             End If
220           End With 'rsLabels
230         Else
240           EnterBlank
250           Detail.PrintSection
260         End If
270     Else
     
280     End If
290   Count = Count + 1
 
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "RLabels", "CurrentDetail", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub
Private Sub AllDetail()
      '
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100    With rslabels
110        If Not .eof And Count < 20 Then
120           If nnNum(!PreferredAddress) <> 1 Then
130             EnterShipping
140           Else
150             EnterMailing
160           End If
170           Detail.PrintSection
180           .MoveNext
190         End If
200       End With
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "RLabels", "AllDetail", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Sub SetDB()
  Clear
  rslabels.Open "SELECT * FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID WHERE BetaTester = 1", cnMain, adOpenForwardOnly, adLockReadOnly
End Sub

Public Sub SetDBCurrent(plContactID As Long)
  Clear
  rslabels.Open "SELECT * FROM TCompany LEFT JOIN TContact ON TCompany.ID = TContact.CompanyID WHERE TContact.ID = " & plContactID, cnMain, adOpenForwardOnly, adLockReadOnly
End Sub

Public Sub SetDBGroup(plListID As Long)
  Clear
  rslabels.Open "SELECT * FROM TGroupListLink LEFT OUTER JOIN TContact ON TGroupListLink.ContactID = TContact.ID RIGHT OUTER JOIN TCompany ON TContact.CompanyID = TCompany.ID " & _
    "WHERE (TGroupListLink.ListID = " & plListID & ")", cnMain, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub EnterBlank()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   txtName = ""
110   txtAddress = ""
120   txtCSZ = ""
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "RLabels", "EnterBlank", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Sub EnterShipping()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   With rslabels
110     txtName = FormatName(!Name, IIf((!FirstName & !LastName & "") <> "", !FirstName & " " & !LastName, ""))
120     txtAddress = FormatAddress("" & !Address1, "" & !Address2, 0)
130     txtCSZ = FormatCSZ("" & !City, "" & !State, "" & !Zip)
140   End With
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "RLabels", "EnterShipping", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Sub EnterMailing()
      '<EhHeader>
      On Error GoTo EH
      '
      '</EhHeader>
100   With rslabels
110     txtName = FormatName(!Name, IIf((!FirstName & !LastName & "") <> "", !FirstName & " " & !LastName, ""))
120     txtAddress = FormatAddress("" & !PermMailAddress1, "" & !PermMailAddress2, 0)
130     txtCSZ = FormatCSZ("" & !PermMailCity, "" & !PermMailState, "" & !PermMailZip)
140   End With
      '<EhFooter>
      '
      Exit Sub
EH:
      ErrorMgr.Raise "RLabels", "EnterMailing", Err.Number, Err.Description, Erl
      '</EhFooter>
End Sub

Public Sub Clear()
  If rslabels.State <> adStateClosed Then
    rslabels.Close
  End If
End Sub
