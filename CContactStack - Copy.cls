VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContactStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private bBack As Boolean
Private bForward As Boolean
Private lForward() As Long
Private lBack(100) As Long
Public Function EnableBack() As Boolean
  If lBack(2) <= 0 Then
    EnableBack = False
  Else
    EnableBack = True
  End If
End Function

Public Function EnableForward() As Boolean
  If lForward(1) <= 0 Then
    EnableForward = False
  Else
    EnableForward = True
  End If
End Function


Public Function Back(plCurrentContactID As Long) As Long
  Dim i As Integer
  '
  For i = 2 To 100
    lBack(i - 1) = lBack(i)
  Next i
  '
  If lBack(1) = 0 Then
    Back = 0
  Else
    Back = lBack(1)
  End If
  '
  lBack(1) = plCurrentContactID
  '
  
  '
  If lBack(1) > 0 Then
    bBack = True
    '
    For i = 99 To 1 Step -1
      lForward(i + 1) = lForward(i)
    Next i
    '
    lForward(1) = plCurrentContactID
  End If
End Function

Public Function Forward(plCurrentContactID As Long) As Long
  Dim i As Integer
  '
  Forward = lForward(1)
  '
  For i = 2 To 100
    lForward(i - 1) = lForward(i)
  Next i
  '
  If lForward(1) > 0 Then
    For i = 99 To 1 Step -1
      lBack(i + 1) = lBack(i)
    Next i
    
    '
    'lForward(1) = plCurrentContactID
    '
    bBack = True
  End If
  lBack(1) = plCurrentContactID
End Function

Public Sub Current(plCurrentContactID As Long)
  'ReDim lForward(100)
  Dim i As Integer
  '
  If bBack = False Then
    For i = 99 To 1 Step -1
      lBack(i + 1) = lBack(i)
    Next i
    lBack(1) = plCurrentContactID
    '
    ReDim lForward(100)
  Else
    
    bBack = False
  End If
  
End Sub


