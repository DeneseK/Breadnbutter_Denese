Option Strict Off
Option Explicit On
Friend Class FDetails
	Inherits System.Windows.Forms.Form
	Dim i As Short
	Dim x As Short
	Dim n As Short
	Dim y As Short
	Dim l As Short
	Dim g As Short
	Dim q As Short
	
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		
		If Not i > (VB6.PixelsToTwipsX(Me.Width) - 500) Then
			Select Case x
				Case 1
					Image1.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
					Image2.Visible = False
					Image1.Visible = True
					x = 2
					i = i + 200
				Case 2
					Image1.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
					Image2.Visible = False
					Image1.Visible = True
					x = 3
				Case 3
					Image2.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
					Image2.Visible = True
					Image1.Visible = False
					x = 1
				Case Else
					x = 1
			End Select
		End If
		'
		If Not n > (VB6.PixelsToTwipsY(Me.Height) - 1000) Then
			Image3.SetBounds(VB6.TwipsToPixelsX(l), VB6.TwipsToPixelsY(n), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			n = n + 100
		Else
			Image3.Visible = False
			Image4.SetBounds(VB6.TwipsToPixelsX(l), VB6.TwipsToPixelsY(n), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			Image4.Visible = True
			l = l - 100
		End If
		
		If l = 3420 Then
			l = l - 20
			Image4.SetBounds(VB6.TwipsToPixelsX(l), VB6.TwipsToPixelsY(n), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			Timer1.Enabled = False
			Timer2.Enabled = True
			Exit Sub
		End If
		
		
	End Sub
	
	Private Sub FDetails_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		x = 1
		y = 1
		n = 720
		l = 6320
		g = 1
		i = 0
		Timer1_Tick(Timer1, New System.EventArgs())
		GetWavFiles()
		Timer4_Tick(Timer4, New System.EventArgs())
	End Sub
	
	Private Sub GameOver()
		Image4.Visible = False
	End Sub
	
	Private Sub Timer2_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer2.Tick
		Image4.Visible = False
		Timer2.Enabled = False
		Timer3.Enabled = True
	End Sub
	
	Private Sub Timer3_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer3.Tick
		Timer4.Enabled = False
		Select Case g
			Case 1
				Call sndPlaySound(My.Application.Info.DirectoryPath & "\killed.wav", &H1s)
				Image1.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				g = 2
			Case 2
				Image5.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Image1.Visible = False
				Image5.Visible = True
				g = 3
			Case 3
				Image6.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Image5.Visible = False
				Image6.Visible = True
				g = 4
			Case 4
				Image7.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Image6.Visible = False
				Image7.Visible = True
				g = 5
			Case 5
				Image8.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Image7.Visible = False
				Image8.Visible = True
				g = 6
			Case 6
				Image9.SetBounds(VB6.TwipsToPixelsX(i), VB6.TwipsToPixelsY(3000), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				Image8.Visible = False
				Image9.Visible = True
				g = 7
				Timer3.Interval = 500
			Case 7
				Image9.Visible = False
				g = 8
			Case 8
				Label1.Visible = True
				g = 9
			Case 9
				Label2.Visible = True
				g = 10
				Timer3.Interval = 1000
			Case 10
				Me.Close()
		End Select
	End Sub
	
	Private Sub Timer4_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer4.Tick
		Call sndPlaySound(My.Application.Info.DirectoryPath & "\pacchomp.wav", &H1s)
	End Sub
	
	Private Sub GetWavFiles()
		Dim rsWav As New ADODB.Recordset
		Dim strStream As ADODB.Stream
		rsWav.Open("Select * from TVMailMessages where [MessageName] = 'pacchomp.wav'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		If Not rsWav.eof Then
			If Not rsWav.BOF Then
				strStream = New ADODB.Stream
				strStream.Type = ADODB.StreamTypeEnum.adTypeBinary
				strStream.Open()
				strStream.Write(rsWav.Fields("Attachment"))
				strStream.SaveToFile(My.Application.Info.DirectoryPath & "\" & rsWav.Fields("MessageName").Value, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
				strStream.Close()
				'UPGRADE_NOTE: Object strStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				strStream = Nothing
			End If
		End If
		rsWav.Close()
		rsWav.Open("Select * from TVMailMessages where [MessageName] = 'killed.wav'", cnMain, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
		If Not rsWav.eof Then
			If Not rsWav.BOF Then
				strStream = New ADODB.Stream
				strStream.Type = ADODB.StreamTypeEnum.adTypeBinary
				strStream.Open()
				strStream.Write(rsWav.Fields("Attachment"))
				strStream.SaveToFile(My.Application.Info.DirectoryPath & "\" & rsWav.Fields("MessageName").Value, ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
				strStream.Close()
				'UPGRADE_NOTE: Object strStream may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				strStream = Nothing
			End If
		End If
		rsWav.Close()
		
		
	End Sub
End Class