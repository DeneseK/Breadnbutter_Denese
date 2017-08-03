Attribute VB_Name = "MReporter"
Option Explicit

'Public dbmInt As New DBMgr
'Public dbmMats As New DBMgr

Public Enum HeaderType
  htFull
  htPartial
  htNone
End Enum

Public Function FormatCSZ(City, State, Zip) As String
  On Error GoTo ErrCall
  '
  Dim CSZTemp As String
  
  If Not IsNull(City) Then CSZTemp = City
  If Not IsNull(State) Then CSZTemp = CSZTemp & ", " & Trim(State)
  If Not IsNull(Zip) Then
    If Len(Trim(Zip)) > 5 Then
      CSZTemp = CSZTemp & " " & Format(Zip, "&&&&&" & "-" & "&&&&")
    Else
      CSZTemp = CSZTemp & " " & Trim(Zip)
    End If
  End If
  
  FormatCSZ = CSZTemp
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & "in Module1.FormatCSZ", vbCritical, "Error"
End Function


Public Function FormatPhone(sPhone As String) As String
  On Error GoTo ErrCall
  '
  Dim PhoneStripped As String
  Dim PhoneTemp As String
  Dim i As Integer
  '
  PhoneStripped = Trim$(sPhone)
  PhoneTemp = Format(Left$(PhoneStripped, 10), "!&&&-&&&-&&&&")
  If Len(PhoneStripped) > 10 Then
      PhoneTemp = PhoneTemp & " x " & Mid$(PhoneStripped, 11)
  End If
  '
  FormatPhone = PhoneTemp
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & "in Module1.FormatPhone", vbCritical, "Error"
End Function


Public Function FormatAddress(sAddress1 As String, sAddress2 As String, Optional iOrientation As Integer) As String
  On Error GoTo ErrCall
  '
  If sAddress2 = "" Then
    FormatAddress = sAddress1
  Else
    If sAddress1 = "" Then
      FormatAddress = sAddress2
    Else
      FormatAddress = sAddress1 & IIf(iOrientation, ", ", vbCrLf) & sAddress2
    End If
  End If
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & "in Module1.FormatAddress", vbCritical, "Error"
End Function

Public Function FormatName(sName1 As String, sName2 As String) As String
  On Error GoTo ErrCall
  '
  If sName2 = "" Then
    FormatName = sName1
  Else
    FormatName = sName1 & vbCrLf & "Attention: " & sName2
  End If
  '
  Exit Function
ErrCall:
  MsgBox "Error " & Err.number & ": " & Err.Description & vbCrLf & "in Module1.FormatName", vbCritical, "Error"
End Function


