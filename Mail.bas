Attribute VB_Name = "Mail"
Public Declare Function ShellExecute Lib "shell32.dll" _
       Alias "ShellExecuteA" ( _
       ByVal hwnd As Long, _
       ByVal lpOperation As String, _
       ByVal lpFile As String, _
       ByVal lpParameters As String, _
       ByVal lpDirectory As String, _
       ByVal nShowCmd As Long) As Long



