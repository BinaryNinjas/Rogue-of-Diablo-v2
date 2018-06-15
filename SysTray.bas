Attribute VB_Name = "SysTray"
'######################################################################
'System Tray Declarations Starts
'######################################################################

Public INTRAY As Boolean 'Boolean to detect App Status[Max or Min]

' Declare Tray Icon
Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uID As Long
     uFlags As Long
     uCallbackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type

' tray Return values
Public Const trayLBUTTONDOWN = 7695
Public Const trayLBUTTONUP = 7710
Public Const trayLBUTTONDBLCLK = 7725

Public Const trayRBUTTONDOWN = 7740
Public Const trayRBUTTONUP = 7755
Public Const trayRBUTTONDBLCLK = 7770

Public Const trayMOUSEMOVE = 7680

Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_LBUTTONDBLCLK = &H203

Global Const NIM_ADD = &H0& 'constants & flags for NotifyIcons
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200

Global NI As NOTIFYICONDATA

'Systray API
Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public result As Long
