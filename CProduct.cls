VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub LoadCollection(ByRef pProducts As CProducts)

  Dim rslist As New ADODB.Recordset
  Dim ProductData As CProductData
  '
  rslist.Open "SELECT * FROM TProduct ORDER BY ProductID", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText

  'Set rsList = dbPropertyValuation.OpenRecordset("SELECT * FROM TAttachedGarage" & _
    " ORDER BY SquareFoot", dbOpenForwardOnly)
  '
    While Not rslist.EOF
      With rslist
        Set ProductData = New CProductData
        '
        ProductData.ProductID = nnNum(!ProductID)
        ProductData.Product = !ProductName & vbNullString
        ProductData.Color = nnNum(!Color)
        ProductData.Seed1 = nnNum(!Seed1)
        ProductData.Seed2 = nnNum(!Seed2)
        '
        pProducts.Add ProductData
        '
        rslist.MoveNext
      End With
    Wend
    '
    rslist.Close
  '
  Set rslist = Nothing
  Set ProductData = Nothing
  '
End Sub

Public Function Load(ByRef pProductData As CProductData, plProductID As Long)
  Dim rs As New ADODB.Recordset
  '
  rs.Open "SELECT * FROM TProduct WHERE ProductID = " & plProductID, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  If Not rs.EOF Then
    With rs
      pProductData.ProductID = nnNum(!ProductID)
      pProductData.Product = !ProductName & vbNullString
      pProductData.Color = nnNum(!Color)
      pProductData.Seed1 = nnNum(!Seed1)
      pProductData.Seed2 = nnNum(!Seed2)
    End With
  End If
  '
  rs.Close
  Set rs = Nothing
  '
End Function

Public Function GetProduct(plProductID As Long) As String
  Dim rs As New ADODB.Recordset
  Dim ProductData As CProductData
  '
  rs.Open "SELECT ProductName FROM TProduct WHERE ProductID = " & plProductID, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  If Not rs.EOF Then
    GetProduct = rs!ProductName & vbNullString
  End If
  '
  rs.Close
  Set rs = Nothing
End Function

Public Function GetProductID(psProduct As String) As Long
  Dim rs As New ADODB.Recordset
  Dim ProductData As CProductData
  '
  rs.Open "SELECT ProductID FROM TProduct WHERE ProductName = '" & psProduct & "'", cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  If Not rs.EOF Then
    GetProductID = nnNum(rs!ProductID)
  Else
    GetProductID = 0
  End If
  '
  rs.Close
  Set rs = Nothing
End Function

Public Function GetColor(plProductID As Long) As Long
  Dim rs As New ADODB.Recordset
  Dim ProductData As CProductData
  '
  rs.Open "SELECT Color FROM TProduct WHERE ProductID = " & plProductID, cnMain, adOpenForwardOnly, adLockReadOnly, adCmdText
  '
  If Not rs.EOF Then
    GetColor = nnNum(rs!Color)
  End If
  '
  rs.Close
  Set rs = Nothing
End Function
