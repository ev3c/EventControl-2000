Attribute VB_Name = "NotIcon"
' estructura que usa la funcion  Shell_NotifyIcon
Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
' Flags para NOTIFYICONDATA
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' comandos que procesa  Shell_NotifyIcon
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
' funcion   Shell_NotifyIcon
Declare Sub Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA)
' funcion para copiar string
Declare Sub lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal size As Long)
' eventos del mouse
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209
' Global, solo usamos un icono
Dim NotIcon As NOTIFYICONDATA
Sub IconModify(ico As Long)
   NotIcon.hIcon = ico
   Shell_NotifyIcon NIM_MODIFY, NotIcon
End Sub

Sub IconDelete()
   Shell_NotifyIcon NIM_DELETE, NotIcon
End Sub

Sub IconAdd(hwnd As Long, ico As Long, tip As String)
   NotIcon.cbSize = 88
   NotIcon.hwnd = hwnd
   NotIcon.uID = 1
   NotIcon.uFlags = NIF_ICON + NIF_MESSAGE + NIF_TIP
   NotIcon.uCallbackMessage = WM_MOUSEMOVE
   NotIcon.hIcon = ico
   lstrcpyn NotIcon.szTip, tip, 63
   Shell_NotifyIcon NIM_ADD, NotIcon
End Sub


