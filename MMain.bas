Attribute VB_Name = "MMain"
Option Explicit

Public Const REFRESH_DELAY As Single = 0.5

Public cnMain     As ADODB.Connection
'Public Const sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Password=CTRALEQDISMC;User ID=Eric;Data Source=D:\Projects\Breadnbutter\Data\BNB_DATA.mdb;Persist Security Info=True;Jet OLEDB:System database=C:\WINNT\System32\hr.mdw"


Public ErrorMgr   As CErrorMgr
Public FileOps    As CFileOps
Public FormMgr    As CFormMgr

Public Product As CProduct

Public Company    As CCompany

Public sUserName  As String
Public sCaseName

'Public ContactStack As CContactStack

Public Type DataItem
  DataName        As String
  DataValue       As Variant
End Type

Public DBOps      As CDBOps
Public User       As CUser

Public Enum ConnectionTypeEnum
  SQL
  Access
End Enum

Public sPrinterName As String
Public iNumofCopies As Integer

Public ConnType As ConnectionTypeEnum

Public InputNumber As New CInputNumber

Public sSQLServerName   As String
Public sSQLServerDB     As String
Public sAccessDB        As String
Public sLogin           As String
Public sLoginName       As String
Public sPassword        As String
Public bCases           As Boolean
Public bFromCases       As Boolean

Private Const LOWER_LIMIT As Long = 48   'ascii for 0
Private Const UPPER_LIMIT As Long = 125  'ascii for {
Private Const CHARMAP     As Long = 39

'Public Sub Main1()
'  On Error GoTo ErrCall
'  '
'  App.Title = "Bread 'n' Butter"
'  '
'  InitializeObjects
'  '
'  FSelectDB.Show vbModal
'  '
'  If Not FSelectDB.Cancelled Then
'    sSQLServerName = FSelectDB.cboServer
'    sSQLServerDB = FSelectDB.cboDatabase
'    sAccessDB = FSelectDB.txtDatabase
'    '
'    Unload FSelectDB
'    '
'    If ConnType = SQL Then
'      cnMain.CursorLocation = adUseClient
'      '
'      If sSQLServerName = "GALE_LAPTOP" Then
'        cnMain.Open "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=" _
'          & sSQLServerDB & ";Data Source=" & sSQLServerName & ";User ID=sa"
'      Else
'        If sLogin = "1" Then
'          cnMain.Provider = "SQLOLEDB"
'          cnMain.Properties("Data Source").Value = sSQLServerName
'          cnMain.Properties("Initial Catalog").Value = sSQLServerDB
'          cnMain.Properties("User ID").Value = sLoginName
'          cnMain.Properties("Password").Value = sPassword
'          cnMain.Properties("Persist Security Info") = False
'          cnMain.CursorLocation = adUseClient
'          cnMain.Open
'        Else
'          cnMain.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" _
'          & sSQLServerDB & ";Data Source=" & sSQLServerName
'        End If
'      End If
'    Else
'      cnMain.CursorLocation = adUseClient
'      cnMain.Open "Provider=Microsoft.Jet.OLEDB.4.0;Password=CTRALEQDISMC;User ID=Eric;Data Source=" & sAccessDB & ";Persist Security Info=True;Jet OLEDB:System database=" & FileOps.SystemPath & "hr.mdw"
'    End If
'    '
'    If cnMain.State = adStateOpen Then
'      FEmployeeLog.Show vbModal
'      '
'      If User.LogResults = True Then
'        Load FMain
'      End If
'    Else
'      MsgBox "Could not connect to database. Application will exit."
'    End If
'    '
'    'Company.LoadCompanyList
'  End If
'  '
'  Exit Sub
'ErrCall:
'  MsgBox "Error " & Err.Description & " in Main."
'End Sub

Public Sub Main()
  On Error GoTo ErrCall
  '
  App.Title = "Bread 'n' Butter"
'  '
  InitializeObjects
'  '
  FLogon.Show vbModal
'  '
'  If Not FSelectDB.Cancelled Then
'    sSQLServerName = FLogon.cboServer
'    sSQLServerDB = FLogon.cboDatabase
'    'sAccessDB = FSelectDB.txtDatabase
'    '
'    'Unload FSelectDB
'    '
'    'If ConnType = SQL Then
'      cnMain.CursorLocation = adUseClient
'      '
'      'If sLogin = "1" Then
'        cnMain.Provider = "SQLOLEDB"
'        cnMain.Properties("Data Source").Value = sSQLServerName
'        cnMain.Properties("Initial Catalog").Value = sSQLServerDB
'        cnMain.Properties("User ID").Value = sLoginName
'        cnMain.Properties("Password").Value = sPassword
'        cnMain.Properties("Persist Security Info") = False
'        cnMain.CursorLocation = adUseClient
'        cnMain.Open
'      'Else
'      '  cnMain.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" _
'      '  & sSQLServerDB & ";Data Source=" & sSQLServerName
'      'End If
'    'Else
'    '  cnMain.CursorLocation = adUseClient
'    '  cnMain.Open "Provider=Microsoft.Jet.OLEDB.4.0;Password=CTRALEQDISMC;User ID=Eric;Data Source=" & sAccessDB & ";Persist Security Info=True;Jet OLEDB:System database=" & FileOps.SystemPath & "hr.mdw"
'    'End If
'    '
    If cnMain.State = adStateOpen Then
'      'FEmployeeLog.Show vbModal
'      '
      If User.LogResults = True Then
        Load FMain
      End If
    Else
      MsgBox "Could not connect to database. Application will exit."
      Exit Sub
    End If
    '
    'Company.LoadCompanyList
  'End If
  '
  Exit Sub
ErrCall:
  MsgBox "Error " & Err.Description & " in Main."
End Sub

Public Sub InitializeObjects()
  On Error GoTo ErrCall
  '
  Set ErrorMgr = New CErrorMgr
  Set FormMgr = New CFormMgr
  Set FileOps = New CFileOps
  '
  FormMgr.Setup FMain
  '
  
  '
  Set DBOps = New CDBOps
  Set User = New CUser
  '
  Set Company = New CCompany
  '
  Set Product = New CProduct
  '
  Set InputNumber = New CInputNumber
  '
  Set cnMain = New ADODB.Connection
  '
  'Company.Clear
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description
End Sub

Public Function Rot39(ByVal sData As String) As String

  'ROT39 (a variation of the ROT13 function) by Dag Sunde

   Dim sReturn As String
   Dim nCode As Long
   Dim nData As Long
   Dim bData() As Byte
   
   On Error GoTo Rot39_error
   
  'initialize the byte array to the
  'size of the string passed.
   ReDim bData(Len(sData)) As Byte
    
  'cast string into the byte array
   bData = StrConv(sData, vbFromUnicode)
    
   For nData = 0 To UBound(bData)
    
     'with the ASCII value of the character
      nCode = bData(nData)
        
     'assure the ASCII value is between
     'the lower and upper limits
      If ((nCode >= LOWER_LIMIT) And (nCode <= UPPER_LIMIT)) Then
         
        'shift the ASCII value by the
        'CHARMAP const value
         nCode = nCode + CHARMAP
         
        'perform a check against the upper
        'limit. If the new value exceeds the
        'upper limit, rotate the value to offset
        'from the beginning of the character set.
         If nCode > UPPER_LIMIT Then
            nCode = nCode - UPPER_LIMIT + LOWER_LIMIT - 1
         End If
      End If
        
     'reassign the new shifted value to
     'the current byte
      bData(nData) = nCode
        
   Next nData
    
  'convert the byte array back
  'to a string and exit
   sReturn = StrConv(bData, vbUnicode)

Rot39_exit:
   
  'assign the return string
   Rot39 = sReturn
   Exit Function
    
Rot39_error:
 
  'error! Return an empty string
   sReturn = ""
   Resume Rot39_exit:
    
End Function
