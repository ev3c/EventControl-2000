Attribute VB_Name = "URLAdress"
Option Explicit

Public Type ProcData
    AppHwnd As Long
    title As String
    Placement As String
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lParam As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0

Private iURL As Integer

' ***********************************************
' If this window is of the Edit class, return
' its contents. Otherwise search its children
' for an Edit object.
' ***********************************************
Public Function EditInfo(window_hwnd As Long) As String
Dim txt As String
Dim buf As String
Dim buflen As Long
Dim child_hwnd As Long
Dim children() As Long
Dim num_children As Integer
Dim i As Integer

    ' Get the class name.
    buflen = 256
    buf = Space$(buflen - 1)
    buflen = GetClassName(window_hwnd, buf, buflen)
    buf = Left$(buf, buflen)
    
    ' See if we found an Edit object.
    If buf = "Edit" Then
        EditInfo = WindowText(window_hwnd)
        Exit Function
    End If
    
    ' It's not an Edit object. Search the children.
    ' Make a list of the child windows.
    num_children = 0
    child_hwnd = GetWindow(window_hwnd, GW_CHILD)
    Do While child_hwnd <> 0
        num_children = num_children + 1
        ReDim Preserve children(1 To num_children)
        children(num_children) = child_hwnd
        
        child_hwnd = GetWindow(child_hwnd, GW_HWNDNEXT)
    Loop
    
    ' Get information on the child windows.
    For i = 1 To num_children
        txt = EditInfo(children(i))
        If txt <> "" Then
          Exit For
        End If
    Next i

    EditInfo = txt
End Function
' ************************************************
' Return the text associated with the window.
' ************************************************
Public Function WindowText(window_hwnd As Long) As String
Dim txtlen As Long
Dim txt As String

    WindowText = ""
    If window_hwnd = 0 Then Exit Function
    
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
    If txtlen = 0 Then Exit Function
    
    txtlen = txtlen + 1
    txt = Space$(txtlen)
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
    WindowText = Left$(txt, txtlen)
End Function


Public Function EnumProc(ByVal app_hwnd As Long, ByVal lParam As Long) As Boolean
Dim buf As String * 1024
Dim title As String
Dim length As Long
Dim strURLAdress As String
Dim bSave As Boolean
Dim iFirstPoint, iSecondPoint As Integer

  ' Get the window's title.
  length = GetWindowText(app_hwnd, buf, Len(buf))
  title = Left$(buf, length)
  
  ' See if the title ends with " - Netscape".
  If InStr(title, "- Netscape") Or _
     InStr(title, "- Microsoft Internet Explorer") Then
      ' This is it. Find the ComboBox information.

        strURLAdress = EditInfo(app_hwnd)
        bSave = True
    
    iFirstPoint = InStr(strURLAdress, ".")
    If iFirstPoint <> 0 Then
       iSecondPoint = InStr(iFirstPoint + 1, strURLAdress, ".")
    End If
    
    If Left(strURLAdress, 7) = "http://" And _
       iFirstPoint <> 0 And _
       iSecondPoint <> 0 Then
    
        gaURL(iURL) = strURLAdress
        iURL = iURL + 1
    
'      Do While gaHistorial(x) <> ""
'        If gaHistorial(x) = strURLAdress Then
'          bSave = False
'          Exit Do
'        Else
'          x = x + 1
'        End If
'      Loop
'
'      If bSave Then
'        x = 0
'        Do While gaHistorial(x) <> ""
'           x = x + 1
'        Loop
'        gaHistorial(x) = strURLAdress
'        Call frmEventControl.Evento_Historial_Grabar( _
'             MyDate, Time, strURLAdress)
'      End If
    
    End If
    
  End If
  
  'Continue search
  EnumProc = 1
  
End Function
' Start the enumeration.
Public Sub GetURLAdress()
  Dim x, y As Integer
  Dim bDelete, bAdd As Boolean
  
  'Redimensiona el array de las URL activas
  ReDim gaURL(0 To 999)
  iURL = 0
  
  'Lee las URL activas en gaURL()
  Call EnumWindows(AddressOf EnumProc, 0)


  'Comprobamos si cada dirección de Historial está en
  'gaURL activas, si no está la borramos de Historial
  Do

    bDelete = True      'borrar de historial si no está
    For y = 0 To iURL
      If gaHistorial(x) = gaURL(y) Then
        bDelete = False
      End If
    Next
    
    If bDelete Then
      gaHistorial = Array_Delete(gaHistorial(), x)
    Else
      x = x + 1
    End If
  
  Loop While gaHistorial(x) <> ""
  
  
  'Comprueba si las URL de gaURL() están en gaHistorial
  'y si no están las Añade.
  x = 0
  Do
    bAdd = True
    y = 0
    Do
      If gaURL(x) = gaHistorial(y) Then
        bAdd = False
      End If
      y = y + 1
    Loop While gaHistorial(y) <> ""
  
    If bAdd Then
      gaHistorial(y) = gaURL(x)
      Call frmEventControl.Evento_Historial_Grabar( _
            MyDate, Time, gaURL(x))
    End If
    x = x + 1
  Loop While gaURL(x) <> ""

End Sub

Private Function Array_Delete(lArray() As String, iPos) As Variant
  Dim x As Integer
  
  x = iPos
  Do
    lArray(x) = lArray(x + 1)
    x = x + 1
  Loop While lArray(x) <> ""

  Array_Delete = lArray()

End Function
