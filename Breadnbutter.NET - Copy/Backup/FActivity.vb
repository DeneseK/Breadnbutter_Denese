Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class FActivity
	Inherits System.Windows.Forms.Form
	
	'UPGRADE_WARNING: Form event FActivity.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub FActivity_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error GoTo ErrorHandler
		'
		Dim rsLog As New ADODB.Recordset
		Dim X As Integer
		'
		'\\ Local Declarations
		Dim iEntryNo As Short
		'
		rsLog.Open("SELECT * FROM tblLog", cnMain)
		'
		With rsLog
			'.FindFirst "[ID] = " & CLng(Right$(FAuthLog.lvwLog.SelectedItem.Key, Len(FAuthLog.lvwLog.SelectedItem.Key) - 1))
			'If .NoMatch = False Then
			X = CInt(VB.Right(FAuthLog.lvwLog.FocusedItem.Name, Len(FAuthLog.lvwLog.FocusedItem.Name) - 1))
			.MoveFirst()
			.Find("ID = " & X)
			If Not .EOF Then
				lblHRIDate.Text = VB6.Format(.Fields("ActionDateTime").Value, "YYYY.Mm.Dd")
				lblHRITime.Text = VB6.Format(.Fields("ActionDateTime").Value, "Hh:Nn:Ss")
				lblEmp.Text = .Fields("Employee").Value
				lblCompany.Text = .Fields("Company").Value
				lblUser.Text = .Fields("User").Value
				lblSiteCode.Text = IIf(.Fields("SiteCompID").Value <> 0, CStr(.Fields("SiteCompID").Value) & " " & CStr(.Fields("SiteSessionID").Value), "N/A")
				lblSiteKey.Text = .Fields("SiteKey").Value
				lblConfCode.Text = .Fields("SiteConfCode").Value
				lblSiteDate.Text = VB6.Format(.Fields("SiteDateTime").Value, "YYYY.Mm.Dd")
				lblSiteTime.Text = VB6.Format(.Fields("SiteDateTime").Value, "Hh:Nn:Ss")
				lblSiteDays.Text = IIf(.Fields("SiteDays").Value <> VariantType.Null, CStr(.Fields("SiteDays").Value), "N/A")
				lblSiteExpDate.Text = IIf(.Fields("SiteExpirationDate").Value <> VariantType.Null, VB6.Format(.Fields("SiteExpirationDate").Value, "YYYY.Mm.Dd"), "N/A")
				If .Fields("ActionType").Value = "Authorization" Then
					lblAction.Text = "Authorization (" & .Fields("ActionSubType").Value & ")"
				Else
					lblAction.Text = .Fields("ActionType").Value
				End If
			End If
		End With
		'
		rsLog.Close()
		'UPGRADE_NOTE: Object rsLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsLog = Nothing
		'
		Exit Sub
		'
ErrorHandler: 
		MsgBox("(" & Err.Number & ") " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "ERROR: FPrimary.General.RefreshLogDisplay")
	End Sub
End Class