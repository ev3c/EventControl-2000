Attribute VB_Name = "Modem"
Private Const ERROR_SUCCESS = 0&
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1 ' Unicode nul terminated string
Private Const REG_BINARY = 3 ' Free form binary

Private Declare Function RegCloseKey Lib "advapi32.dll" _
                (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" _
                Alias "RegOpenKeyA" (ByVal hKey As Long, _
                ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                ByVal lpValueName As String, ByVal lpReserved _
                As Long, lpType As Long, lpData As Any, _
                lpcbData As Long) As Long

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
  Dim lResult, lValueType, lDataBufSize  As Long
  Dim strBuf As String
  'retrieve nformation about the key
  lResult = RegQueryValueEx(hKey, strValueName, 0, _
            lValueType, ByVal 0, lDataBufSize)
  If lResult = 0 Then
    If lValueType = REG_SZ Then
      'Create a buffer
      strBuf = String(lDataBufSize, Chr$(0))
      'retrieve the key's content
      lResult = RegQueryValueEx(hKey, strValueName, 0, _
                        0, ByVal strBuf, lDataBufSize)
      If lResult = 0 Then
        'Remove the unnecessary chr$(0)'s
        RegQueryStringValue = Left$(strBuf, InStr(1, _
                              strBuf, Chr$(0)) - 1)
      End If
    ElseIf lValueType = REG_BINARY Then
      Dim strData As Integer
      'retrieve the key's value
      lResult = RegQueryValueEx(hKey, strValueName, 0, _
                0, strData, lDataBufSize)
      If lResult = 0 Then
        RegQueryStringValue = strData
      End If
    End If
  End If
End Function
Function GetString(hKey As Long, strPath As String, strValue As String)
  Dim Ret
  'Open the key
  RegOpenKey hKey, strPath, Ret
  'Get the key's content
  GetString = RegQueryStringValue(Ret, strValue)
  'Close the key
  RegCloseKey Ret
End Function

Public Function IsConnected() As Boolean
  Dim lpSubKey As String
  
  IsConnected = False
  lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
  
  Ret = GetString(HKEY_LOCAL_MACHINE, lpSubKey, _
                  "Remote Connection")
  If Ret = 0 Or Ret = "" Then
    IsConnected = False
  Else
    IsConnected = True
  End If
End Function



