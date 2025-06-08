Attribute VB_Name = "TimerMsgBox"
Public Function TimerMsgBox(Texto As String, Titulo As String, Tiempo As Byte) As String
Dim Sh As New IWshRuntimeLibrary.IWshShell_Class
   Sh.Popup Texto, Tiempo, Titulo, vbInformation + vbSystemModal + vbMsgBoxSetForeground
End Function

