VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDailyItemSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Theses should be made into properties
'
Public StatusType As String
Public SupportType As String
Public InitialDays As Integer
Public RegularDays As Integer
Public DaysDivider As Integer
Public DaysOperator As String
Public DaysToSearch As Integer
Public Name As String
