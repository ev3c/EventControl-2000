Attribute VB_Name = "msdos"
'// Declaraciones
Private Declare Function GetShortPathName Lib "kernel32" _
        Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'// Funcion
Public Function NombreCorto(sNombreLargo As String) As String
    Dim sNombreCorto As String * 255
    GetShortPathName sNombreLargo, sNombreCorto, 255
    NombreCorto = sNombreCorto
End Function

