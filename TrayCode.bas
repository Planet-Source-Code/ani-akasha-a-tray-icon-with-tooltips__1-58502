Attribute VB_Name = "TrayCode"
'This sample was downloaded from http://www.nekhbet.tk
'Made by Trambitas Sorin @ 19.01.2005
'For questions please contact me at TrimbitasSorin@Yahoo.com
'A small part of this code was taken from a sample
'found at http://www.vb-helper.com

Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private OldWindowProc As Long
Private TheForm As Form
Private TheMenu As Menu
Private TheData As NOTIFYICONDATA

Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&
Private Const WM_USER = &H400
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONUP = &H205
Private Const TRAY_CALLBACK = (WM_USER + 1001&)
Private Const GWL_WNDPROC = (-4)
Private Const GWL_USERDATA = (-21)
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIF_INFO = &H10
Private Const NIIF_INFO = &H1
Private Const NIF_MESSAGE = &H1
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_NULL = &H0
Private Const WM_MOUSEMOVE = &H200

Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Private IsTrayIconActive As Boolean

Public Function InitializeTrayModule()
  IsTrayIconActive = False
End Function


'The replacement window process
Private Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const WM_NCDESTROY = &H82
  If Msg = WM_NCDESTROY Then
    RemoveFromTray
  Else
    If Msg = TRAY_CALLBACK Then
      If lParam = WM_RBUTTONUP Then
        SetForegroundWindow TheForm.hwnd
        TheForm.PopupMenu TheMenu
        If Not (TheForm Is Nothing) Then
          PostMessage TheForm.hwnd, WM_NULL, ByVal 0&, ByVal 0&
        End If
        Exit Function
      End If
    End If
  End If
  NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
End Function

'Add the form's icon to the tray.
Public Sub AddToTray(frm As Form, mnu As Menu, Optional awa As String = "none", Optional aqa As String = "none", Optional TQS As Byte = 1)
  If IsTrayIconActive = False Then
    Set TheForm = frm
    Set TheMenu = mnu
    OldWindowProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
    With TheData
      If (aqa = "none") And (awa = "none") Then
        .uID = 0
      Else
        .uID = vbNull
      End If
      .hwnd = frm.hwnd
      .cbSize = Len(TheData)
      .hIcon = frm.Icon.Handle
      .uFlags = NIF_ICON Or NIF_MESSAGE
      .uCallbackMessage = TRAY_CALLBACK
      .cbSize = Len(TheData)
    End With
  
    Shell_NotifyIcon NIM_ADD, TheData
  
    If (awa <> "none") And (aqa <> "none") Then
      ShowPopUp awa, aqa
      Sleep TQS * 1000
    End If
    IsTrayIconActive = True
  End If
End Sub

'Show a tooltip attached at the icon from tray
Public Function AddToTrayToolTip(formName As Form, menuName As Menu, TipMsg As String, TipTitle As String, Optional TipTimeOutInSeconds As Byte = 1)
  If IsTrayIconActive = True Then
    RemoveFromTray
    AddToTray formName, menuName, TipMsg, TipTitle, TipTimeOutInSeconds
    RemoveFromTray
    AddToTray formName, menuName
  End If
End Function

'Remove the icon from the system tray.
Public Sub RemoveFromTray()
  If IsTrayIconActive = True Then
    With TheData
      .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData
    SetWindowLong TheForm.hwnd, GWL_WNDPROC, OldWindowProc
    Set TheForm = Nothing
    IsTrayIconActive = False
  End If
End Sub

'Set a tray tip.
Public Sub SetTrayTip(tip As String)
  If IsTrayIconActive = True Then
    With TheData
      .szTip = tip & vbNullChar
      .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
  End If
End Sub

'Show Tooltip
Private Function ShowPopUp(Message As String, Title As String)
  With TheData
    .cbSize = Len(TheData)
    .hwnd = Form1.hwnd
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = TheForm.Icon
    .szTip = Title & vbNullChar
    .dwState = 0
    .dwStateMask = 0
    .szInfo = Message & Chr(0)
    .szInfoTitle = Title & Chr(0)
    .dwInfoFlags = NIF_INFO
  End With
  Shell_NotifyIcon NIM_MODIFY, TheData
End Function



