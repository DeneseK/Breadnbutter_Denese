Option Strict Off
Option Explicit On
Friend Class CProduct
	
	
	Public Sub LoadCollection(ByRef pProducts As CProducts)
		
		Dim rslist As New ADODB.Recordset
		Dim ProductData As CProductData
		'
		rslist.Open("SELECT * FROM TProduct ORDER BY ProductID", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		
		'Set rsList = dbPropertyValuation.OpenRecordset("SELECT * FROM TAttachedGarage" & _
		'" ORDER BY SquareFoot", dbOpenForwardOnly)
		'
		While Not rslist.EOF
			With rslist
				ProductData = New CProductData
				'
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ProductData.ProductID = nnNum(.Fields("ProductID"))
				ProductData.Product = .Fields("ProductName").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ProductData.Color = nnNum(.Fields("Color"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ProductData.Seed1 = nnNum(.Fields("Seed1"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ProductData.Seed2 = nnNum(.Fields("Seed2"))
				'
				pProducts.Add(ProductData)
				'
				rslist.MoveNext()
			End With
		End While
		'
		rslist.Close()
		'
		'UPGRADE_NOTE: Object rslist may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rslist = Nothing
		'UPGRADE_NOTE: Object ProductData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ProductData = Nothing
		'
	End Sub
	
	Public Function Load(ByRef pProductData As CProductData, ByRef plProductID As Integer) As Object
		Dim rs As New ADODB.Recordset
		'
		rs.Open("SELECT * FROM TProduct WHERE ProductID = " & plProductID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		If Not rs.EOF Then
			With rs
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pProductData.ProductID = nnNum(.Fields("ProductID"))
				pProductData.Product = .Fields("ProductName").Value & vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pProductData.Color = nnNum(.Fields("Color"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pProductData.Seed1 = nnNum(.Fields("Seed1"))
				'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pProductData.Seed2 = nnNum(.Fields("Seed2"))
			End With
		End If
		'
		rs.Close()
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
		'
	End Function
	
	Public Function GetProduct(ByRef plProductID As Integer) As String
		Dim rs As New ADODB.Recordset
		Dim ProductData As CProductData
		'
		rs.Open("SELECT ProductName FROM TProduct WHERE ProductID = " & plProductID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		If Not rs.EOF Then
			GetProduct = rs.Fields("ProductName").Value & vbNullString
		End If
		'
		rs.Close()
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
	End Function
	
	Public Function GetProductID(ByRef psProduct As String) As Integer
		Dim rs As New ADODB.Recordset
		Dim ProductData As CProductData
		'
		rs.Open("SELECT ProductID FROM TProduct WHERE ProductName = '" & psProduct & "'", cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		If Not rs.EOF Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetProductID = nnNum(rs.Fields("ProductID"))
		Else
			GetProductID = 0
		End If
		'
		rs.Close()
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
	End Function
	
	Public Function GetColor(ByRef plProductID As Integer) As Integer
		Dim rs As New ADODB.Recordset
		Dim ProductData As CProductData
		'
		rs.Open("SELECT Color FROM TProduct WHERE ProductID = " & plProductID, cnMain, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
		'
		If Not rs.EOF Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nnNum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetColor = nnNum(rs.Fields("Color"))
		End If
		'
		rs.Close()
		'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rs = Nothing
	End Function
End Class