Attribute VB_Name = "MLicense"
Option Explicit

'\\ Security
Public bLicChecked         As Boolean
Public bLicTimer           As Boolean
Public lSecValCode(1 To 3) As Long
Public lSecValRslt(1 To 3) As Long
Public dSecVar             As Double
Public bLicError            As Boolean
Public bSecDisp             As Boolean
Public lSecCompID           As Long
'
Public Sub InitLicense()
  '
  '\\ Security
  If bLicChecked = False Then
    With FMain.License
      lSecValCode(1) = 1514044385
      lSecValRslt(1) = 6035375
      lSecValCode(2) = 2067020449
      lSecValRslt(2) = 4912112
      lSecValCode(3) = 1463088053
      lSecValRslt(3) = 8960596
      dSecVar = Int((3000 - 1 + 1) * Rnd + 1)
      .Enabled = False
      '.CPAlgorithm = 1 + 2
      .CPAlgorithm = 65536
      .CPAlgorithmDrive = Left$(App.Path, 1)
      .TCSeed = 192
      .TCRegKey2Seed = 48
      .LFPassword = "D" & "uct" & "Ta" & "p" & "e"
      .LFName = App.Path & "\PowerKey.lf"
      .ExpireDateHard = "12/31/2040"
      .Enabled = True
      lSecCompID = .CPCompNo
      .ForceStatusChanged
      bLicChecked = True
    End With
  End If
End Sub
'
Public Function CPCheck() As Double
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iCt As Integer
  '
  CPCheck = CDbl(Date - (dSecVar / 3.14))
  '
  With FMain.License
    For iCt = 1 To .LicensedComputers
      If lSecCompID = .LicensedComputer(iCt) Then
        CPCheck = CDbl(Date - dSecVar)
        Exit For
      End If
    Next
  End With
  '
  Exit Function
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.General.CPCheck"
End Function

Public Function ValidateLicense() As Double
  On Error GoTo ErrorHandler
  '
  '\\ Local Declarations
  Dim iSecValPair As Integer
  Dim sDlgMsg     As String
  '
  With FMain.License
    '
    iSecValPair = Int((3 - 1 + 1) * Rnd + 1)
    '
    If .LibTest(lSecValCode(iSecValPair)) <> lSecValRslt(iSecValPair) Then
      ValidateLicense = CDbl(Date - (dSecVar / 3.14))
    ElseIf CPCheck = (Date - dSecVar) Then
      If .IsExpired Then
        ValidateLicense = CDbl(Date - (dSecVar / 3.14))
      Else
        If .IsClockTurnedBack Then
          ValidateLicense = CDbl(Date - (dSecVar / 3.14))
        Else
          ValidateLicense = CDbl(Date - dSecVar)
        End If
      End If
    Else
      If .ExpireMode = "D" Then
        If .IsClockTurnedBack Then
          ValidateLicense = CDbl(Date - (dSecVar / 3.14))
        Else
          ValidateLicense = CDbl(Date - (dSecVar / 3.14))
        End If
      Else
        ValidateLicense = CDbl(Date - (dSecVar / 3.14))
      End If
    End If
    '
  End With
  '
  If ValidateLicense = CDbl(Date - (dSecVar / 3.14)) Then
    sDlgMsg = "Your license is either damaged or has expired. You will not be" & vbCrLf & _
              "able to authorize licenses until you contact Jason or Eric." & vbCrLf & _
              "Click on the License Facility button for more information."
    MsgBox sDlgMsg, vbCritical + vbOKOnly, "ERROR: Unauthorized License"
  End If
  '
  Exit Function
  '
ErrorHandler:
  MsgBox "(" & Err.Number & ") " & Err.Description, vbCritical + vbOKOnly, "ERROR: Fcontact.General.ValidateLicense"
End Function



