Option Strict Off
Option Explicit On
Friend Class FResult
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		On Error GoTo ErrCall
		Me.Close()
		Exit Sub
ErrCall: 
		MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & "in FResult.cmdClose", MsgBoxStyle.Critical, "Error")
	End Sub
End Class