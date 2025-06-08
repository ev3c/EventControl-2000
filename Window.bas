Attribute VB_Name = "Window"
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE

Declare Function SetWindowPos Lib "User32" _
(ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal x As Long, _
ByVal y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal wFlags As Long) As Long

' Para hacer la ventana visible en primer plano
 'If SetWindowPos(hwnd, -1, 0, 0, 0, 0, SWP_FLAGS) Then
 'End If
'Para deshabilitar la ventana en primer plano
 'If SetWindowPos(hwnd, -2, 0, 0, 0, 0, SWP_FLAGS) Then
 'End If



