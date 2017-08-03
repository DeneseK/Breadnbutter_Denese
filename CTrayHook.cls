VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrayHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmHooked As MDIForm
Attribute frmHooked.VB_VarHelpID = -1
Private WithEvents tmrTray As Timer
Attribute tmrTray.VB_VarHelpID = -1
Private WithEvents scTray As SubClass
Attribute scTray.VB_VarHelpID = -1
Private mPopup As Menu
'
' User defined values
'
Private Const cbNotify& = &H4000
Private Const uID& = 61860
'
' Persistent objects
'
Private nid As NOTIFYICONDATA

Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Stdole.OLE_HANDLE
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type
   
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Declare Function ShellNotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Sub frmHooked_Load()
   frmHooked.Visible = False
   frmHooked.Caption = "Bread 'n' Butter"
   App.Title = frmHooked.Caption
   '
   ' Setup MsgHook
   '
   scTray.HwndParam = frmHooked.hwnd
   scTray.Messages = cbNotify
   
   '
   ' Setup icon notification from shell
   '
   nid.cbSize = Len(nid)
   nid.hwnd = scTray.HwndParam
   nid.uID = uID
   nid.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
   nid.uCallbackMessage = cbNotify
   nid.hIcon = frmHooked.Icon
   nid.szTip = frmHooked.Caption & Chr(0)
End Sub

Private Sub frmHooked_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    '
    ' Just hide form if user presses Close button or closing from code
    '
    frmHooked.Visible = False
    Cancel = True
  End If
End Sub

Private Sub frmHooked_Unload(Cancel As Integer)
   '
   ' Remove icon from tray
   '
   Call ShellNotifyIcon(NIM_DELETE, nid)
   End
End Sub

Sub Main()
   Dim os As OSVERSIONINFO
   '
   ' Insure that newshell is running.
   '
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   If os.dwMajorVersion < 4 Then
      MsgBox "This program requires NT4 or Win95", _
         vbCritical, "Program Ending"
      End
   End If
   '
   ' Go ahead and load.
   '
   Load frmHooked
End Sub

Public Sub Setup(pfrmHooked As Form, ptmrTray As Timer, pscTray As SubClass, pmPopup As Menu)
  On Error GoTo ErrCall
  '
  Set frmHooked = pfrmHooked
  Set tmrTray = ptmrTray
  Set scTray = pscTray
  Set mPopup = pmPopup
  '
  Exit Sub
ErrCall:
  MsgBox Err.Description
End Sub

Private Sub scTray_WndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
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
          AppActivate frmHooked.Caption
          Screen.MousePointer = vbDefault
       Case WM_RBUTTONDOWN
       Case WM_RBUTTONUP
          '
          ' Display context menu
          ' Highlight default (Open)
          '
          frmHooked.PopupMenu mPopup
          ', , , , mPop(0)
       Case WM_RBUTTONDBLCLK
       Case WM_MBUTTONDOWN
       Case WM_MBUTTONUP
       Case WM_MBUTTONDBLCLK
       Case Else
          Debug.Print "Message unknown!" & param
    End Select
   End If
End Sub

Private Sub scTray_WndMessage(hwnd As Stdole.OLE_HANDLE, msg As Stdole.OLE_HANDLE, wp As Stdole.OLE_HANDLE, lp As Long, retval As Long, nodef As Integer)
  Dim param As String
  param = "msg: " & msg & "    wp: " & wp & "    lp: " & lp
      
  If wp = uID Then
    Select Case lp
       Case WM_MOUSEMOVE
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
          '
          ' Show form
          '
          frmHooked.Visible = True
          AppActivate frmHooked.Caption
          Screen.MousePointer = vbDefault
       Case WM_RBUTTONDOWN
       Case WM_RBUTTONUP
          '
          ' Display context menu
          ' Highlight default (Open)
          '
          frmHooked.PopupMenu mPopup
          ', , , , mPop(0)
       Case WM_RBUTTONDBLCLK
       Case WM_MBUTTONDOWN
       Case WM_MBUTTONUP
       Case WM_MBUTTONDBLCLK
       Case Else
          Debug.Print "Message unknown!" & param
    End Select
   End If
End Sub

Private Sub tmrTray_Timer()
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
  MsgBox Err.Description & " " & Err.Number
  Resume Next
End Sub
