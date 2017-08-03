Option Strict Off
Option Explicit On
Friend Class CContactStack
	Private bBack As Boolean
	Private bForward As Boolean
	Private lForward() As Integer
	Private lBack(100) As Integer
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
	
	
	Public Function Back(ByRef plCurrentContactID As Integer) As Integer
		Dim i As Short
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
	
	Public Function Forward(ByRef plCurrentContactID As Integer) As Integer
		Dim i As Short
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
	
	Public Sub Current(ByRef plCurrentContactID As Integer)
		'ReDim lForward(100)
		Dim i As Short
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
End Class