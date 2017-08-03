Option Strict Off
Option Explicit On
Friend Class CTrayHook
	
	Private WithEvents frmHooked As System.Windows.Forms.Form
	Private WithEvents tmrTray As System.Windows.Forms.Timer
	Private WithEvents scTray As AxDwsbcLib.AxSubClass
	Private mPopup As System.Windows.Forms.ToolStripMenuItem
	'
	' User defined values
	'
	Private Const cbNotify As Integer = &H4000s
	Private Const uID As Integer = 61860
	'
	' Persistent objects
	'
	Private nid As NOTIFYICONDATA
	
	Private Structure NOTIFYICONDATA
		Dim cbSize As Integer
		Dim hwnd As Short
		Dim uID As Integer
		Dim uFlags As Integer
		Dim uCallbackMessage As Integer
		Dim hIcon As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(64),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=64)> Public szTip() As Char
	End Structure
	
	Private Const NIM_ADD As Short = &H0s
	Private Const NIM_MODIFY As Short = &H1s
	Private Const NIM_DELETE As Short = &H2s
	
	Private Const NIF_MESSAGE As Short = &H1s
	Private Const NIF_ICON As Short = &H2s
	Private Const NIF_TIP As Short = &H4s
	
	'UPGRADE_WARNING: Structure NOTIFYICONDATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function ShellNotifyIcon Lib "shell32.dll"  Alias "Shell_NotifyIconA"(ByVal dwMessage As Integer, ByRef lpData As NOTIFYICONDATA) As Integer
	
	Private Const WM_MOUSEMOVE As Short = &H200s
	Private Const WM_LBUTTONDOWN As Short = &H201s
	Private Const WM_LBUTTONUP As Short = &H202s
	Private Const WM_LBUTTONDBLCLK As Short = &H203s
	Private Const WM_RBUTTONDOWN As Short = &H204s
	Private Const WM_RBUTTONUP As Short = &H205s
	Private Const WM_RBUTTONDBLCLK As Short = &H206s
	Private Const WM_MBUTTONDOWN As Short = &H207s
	Private Const WM_MBUTTONUP As Short = &H208s
	Private Const WM_MBUTTONDBLCLK As Short = &H209s
	
	Private Structure OSVERSIONINFO
		Dim dwOSVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char '  Maintenance string for PSS usage
	End Structure
	
	Private Const VER_PLATFORM_WIN32s As Short = 0
	Private Const VER_PLATFORM_WIN32_WINDOWS As Short = 1
	Private Const VER_PLATFORM_WIN32_NT As Short = 2
	
	'UPGRADE_WARNING: Structure OSVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetVersionEx Lib "kernel32"  Alias "GetVersionExA"(ByRef lpVersionInformation As OSVERSIONINFO) As Integer
	
	Private Sub frmHooked_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles frmHooked.Load
		frmHooked.Visible = False
		frmHooked.Text = "Bread 'n' Butter"
		'UPGRADE_ISSUE: App property App.Title was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		App.Title = frmHooked.Text
		'
		' Setup MsgHook
		'
		scTray.HwndParam = frmHooked.Handle.ToInt32
		scTray.Messages = cbNotify
		
		'
		' Setup icon notification from shell
		'
		nid.cbSize = Len(nid)
		nid.hwnd = scTray.HwndParam
		nid.uID = uID
		nid.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
		nid.uCallbackMessage = cbNotify
		nid.hIcon = CInt(CObj(frmHooked.Icon))
		nid.szTip = frmHooked.Text & Chr(0)
	End Sub
	
	Private Sub frmHooked_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		If UnloadMode = System.Windows.Forms.CloseReason.UserClosing Then
			'
			' Just hide form if user presses Close button or closing from code
			'
			frmHooked.Visible = False
			Cancel = True
		End If
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub frmHooked_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'
		' Remove icon from tray
		'
		Call ShellNotifyIcon(NIM_DELETE, nid)
		End
	End Sub
	
	'UPGRADE_WARNING: Application will terminate when Sub Main() finishes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="E08DDC71-66BA-424F-A612-80AF11498FF8"'
	Public Sub Main()
		Dim os As OSVERSIONINFO
		'
		' Insure that newshell is running.
		'
		os.dwOSVersionInfoSize = Len(os)
		Call GetVersionEx(os)
		If os.dwMajorVersion < 4 Then
			MsgBox("This program requires NT4 or Win95", MsgBoxStyle.Critical, "Program Ending")
			End
		End If
		'
		' Go ahead and load.
		'
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmHooked)
	End Sub
	
	Public Sub Setup(ByRef pfrmHooked As System.Windows.Forms.Form, ByRef ptmrTray As System.Windows.Forms.Timer, ByRef pscTray As AxDwsbcLib.AxSubClass, ByRef pmPopup As System.Windows.Forms.ToolStripMenuItem)
		On Error GoTo ErrCall
		'
		frmHooked = pfrmHooked
		tmrTray = ptmrTray
		scTray = pscTray
		mPopup = pmPopup
		'
		Exit Sub
ErrCall: 
		MsgBox(Err.Description)
	End Sub
	
	Private Sub scTray_WndProc(ByRef msg As Integer, ByRef wParam As Integer, ByRef lParam As Integer, ByRef Result As Integer)
		Dim param As String
		param = "msg: " & msg & "    wp: " & wParam & "    lp: " & lParam
		
		If wParam = uID Then
			Select Case lParam
				Case WM_MOUSEMOVE
				Case WM_LBUTTONDOWN
				Case WM_LBUTTONUP
				Case WM_LBUTTONDBLCLK
					'
					' Show form
					'
					frmHooked.Visible = True
					AppActivate(frmHooked.Text)
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				Case WM_RBUTTONDOWN
				Case WM_RBUTTONUP
					'
					' Display context menu
					' Highlight default (Open)
					'
					'UPGRADE_ISSUE: MDIForm method frmHooked.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					frmHooked.PopupMenu(mPopup)
					', , , , mPop(0)
				Case WM_RBUTTONDBLCLK
				Case WM_MBUTTONDOWN
				Case WM_MBUTTONUP
				Case WM_MBUTTONDBLCLK
				Case Else
					Debug.Print("Message unknown!" & param)
			End Select
		End If
	End Sub
	
	Private Sub scTray_WndMessage(ByVal eventSender As System.Object, ByVal eventArgs As AxDwsbcLib._DDwsbcEvents_WndMessageEvent) Handles scTray.WndMessage
		Dim param As String
		param = "msg: " & eventArgs.msg & "    wp: " & eventArgs.wp & "    lp: " & eventArgs.lp
		
		If eventArgs.wp = uID Then
			Select Case eventArgs.lp
				Case WM_MOUSEMOVE
				Case WM_LBUTTONDOWN
				Case WM_LBUTTONUP
				Case WM_LBUTTONDBLCLK
					'
					' Show form
					'
					frmHooked.Visible = True
					AppActivate(frmHooked.Text)
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				Case WM_RBUTTONDOWN
				Case WM_RBUTTONUP
					'
					' Display context menu
					' Highlight default (Open)
					'
					'UPGRADE_ISSUE: MDIForm method frmHooked.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					frmHooked.PopupMenu(mPopup)
					', , , , mPop(0)
				Case WM_RBUTTONDBLCLK
				Case WM_MBUTTONDOWN
				Case WM_MBUTTONUP
				Case WM_MBUTTONDBLCLK
				Case Else
					Debug.Print("Message unknown!" & param)
			End Select
		End If
	End Sub
	
	Private Sub tmrTray_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrTray.Tick
		If ShellNotifyIcon(NIM_ADD, nid) Then
			'
			On Error GoTo errTimer
			'
			tmrTray.Enabled = False
			'frmHooked.Visible = True
		End If
		'
		Exit Sub
		
errTimer: 
		MsgBox(Err.Description & " " & Err.Number)
		Resume Next
	End Sub
End Class