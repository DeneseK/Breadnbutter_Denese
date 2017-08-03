Option Strict Off
Option Explicit On
Friend Class CFormControl
	
	Private sngMinHeight As Single
	Private sngMinWidth As Single
	
	Private bDataForm As Boolean
	
	Private sDescription As String
	
	Public AlwaysVisible As Boolean
	
	Public Event SwitchFrom(ByRef bCancel As Boolean)
	Public Event SwitchTo(ByRef bCancel As Boolean)
	
	'UPGRADE_NOTE: SwitchFrom was upgraded to SwitchFrom_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function SwitchFrom_Renamed() As Boolean
		On Error GoTo ErrCall
		'
		Dim bCancel As Boolean
		'
		RaiseEvent SwitchFrom(bCancel)
		'
		SwitchFrom_Renamed = Not bCancel
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormControl.SwitchFrom.", MsgBoxStyle.Critical, "Error")
	End Function
	
	'UPGRADE_NOTE: SwitchTo was upgraded to SwitchTo_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function SwitchTo_Renamed() As Boolean
		On Error GoTo ErrCall
		'
		Dim bCancel As Boolean
		'
		RaiseEvent SwitchTo(bCancel)
		'
		SwitchTo_Renamed = Not bCancel
		'
		Exit Function
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormControl.SwitchTo.", MsgBoxStyle.Critical, "Error")
	End Function
	
	
	Public Property MinHeight() As Single
		Get
			MinHeight = sngMinHeight
		End Get
		Set(ByVal Value As Single)
			sngMinHeight = Value
		End Set
	End Property
	
	
	Public Property MinWidth() As Single
		Get
			MinWidth = sngMinWidth
		End Get
		Set(ByVal Value As Single)
			sngMinWidth = Value
		End Set
	End Property
	
	
	Public Property DataForm() As Boolean
		Get
			DataForm = bDataForm
		End Get
		Set(ByVal Value As Boolean)
			bDataForm = Value
		End Set
	End Property
	
	Public ReadOnly Property Description() As String
		Get
			Description = sDescription
		End Get
	End Property
	
	Public Sub Setup(ByRef pForm As System.Windows.Forms.Form, ByRef pbData As Boolean, Optional ByRef psngHeight As Single = 0, Optional ByRef psngWidth As Single = 0, Optional ByRef psDescription As String = "", Optional ByRef pfAlwaysVisible As Boolean = False)
		On Error GoTo ErrCall
		'
		sngMinHeight = IIf(psngHeight = 0, VB6.PixelsToTwipsY(pForm.Height), psngHeight)
		sngMinWidth = IIf(psngWidth = 0, VB6.PixelsToTwipsX(pForm.Width), psngWidth)
		bDataForm = pbData
		sDescription = psDescription
		AlwaysVisible = pfAlwaysVisible
		'
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in clsFormControl.Setup.", MsgBoxStyle.Critical, "Error")
	End Sub
End Class