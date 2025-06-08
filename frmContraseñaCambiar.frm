VERSION 5.00
Begin VB.Form frmContrase�aCambiar 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtNuevaContrase�a 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtContrase�a 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtNuevaContrase�a2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   3975
      Begin VB.Label lblNuevaContrase�a2 
         Caption         =   "Reentrar Nueva Contrase�a"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblNuevaContrase�a 
         Caption         =   "Nueva Contrase�a"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label lblContrase�a 
      Caption         =   "Contrase�a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmContrase�aCambiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  If txtContrase�a.Text = gstrContrase�a Then
    If txtNuevaContrase�a = txtNuevaContrase�a2 Then
      If EC_SHAREWARE Then
        Select Case frmEventControl.cboIdioma.ListIndex
          Case EC_ESPA�OL
            MsgBox "Solamente puede cambiar la contrase�a " & _
            "en la versi�n registrada de " & gstrPrograma, _
            vbInformation, "Cambiar Contrase�a Informaci�n"
          Case EC_ENGLISH
            MsgBox "Only can change the Password in the " & _
            "registered version of " & gstrPrograma, _
            vbInformation, "Change Password Information"
          Case EC_CATALA
            MsgBox "Solsament pot cambiar la contrasenya " & _
            "a la versi� registrada de " & gstrPrograma, _
            vbInformation, "Cambiar Contrasenya Informaci�"
        End Select
      Else
        
        gdb.NewPassword gstrContrase�a, txtNuevaContrase�a
        gdbHistorial.NewPassword gstrContrase�a, txtNuevaContrase�a
        
        gstrContrase�a = txtNuevaContrase�a
        SaveSetting "EventControl", "Inicio", _
            "gstrContrase�a", _
            Contrase�a_Encriptar(gstrContrase�a)
      
      End If
      Unload frmContrase�aCambiar
    Else
      Select Case frmEventControl.cboIdioma.ListIndex
        Case EC_ESPA�OL
          MsgBox "La nueva contrase�a no coincide", _
          vbInformation, "Cambiar Contrase�a - Error"
        Case EC_ENGLISH
          MsgBox "The new password don't coincides", _
          vbInformation, "Change Password - Error"
        Case EC_CATALA
          MsgBox "La nova contrasenya no coincideix", _
          vbInformation, "Cambiar Contrasenya - Error"
      End Select
      txtNuevaContrase�a.SetFocus
    End If
  Else
    
    Select Case frmEventControl.cboIdioma.ListIndex
      Case EC_ESPA�OL
        MsgBox "Contrase�a Erronea", vbInformation, _
               "Cambiar Contrase�a - Error"
      Case EC_ENGLISH
        MsgBox "Invalid Password", vbInformation, _
               "Change Password - Error"
      Case EC_CATALA
        MsgBox "Contrasenya Erronia", vbInformation, _
               "Cambiar Contrasenya - Error"
    End Select
    txtContrase�a.SetFocus

  End If
End Sub

Private Sub cmdCancelar_Click()
  Unload frmContrase�aCambiar
End Sub

Private Sub Form_Load()
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPA�OL
      frmContrase�aCambiar.Caption = "Cambiar Contrase�a"
      lblContrase�a = "Contrase�a"
      lblNuevaContrase�a.Caption = "Nueva Contrase�a"
      lblNuevaContrase�a2.Caption = "Reentrar nueva Contrase�a"
      cmdAceptar.Caption = "&Aceptar"
      cmdCancelar.Caption = "&Cancelar"
    Case EC_ENGLISH
      frmContrase�aCambiar.Caption = "Change Password"
      lblContrase�a = "Password"
      lblNuevaContrase�a.Caption = "New Password"
      lblNuevaContrase�a2.Caption = "Re-entrer new Password"
      cmdAceptar.Caption = "&Accept"
      cmdCancelar.Caption = "&Cancel"
    Case EC_CATALA
      frmContrase�aCambiar.Caption = "Cambiar Contrasenya"
      lblContrase�a = "Contrasenya"
      lblNuevaContrase�a.Caption = "Nova Contrasenya"
      lblNuevaContrase�a2.Caption = "Reentrar nova Contrasenya"
      cmdAceptar.Caption = "&Acceptar"
      cmdCancelar.Caption = "&Cancelar"
  End Select
End Sub


Private Sub txtContrase�a_()

End Sub

Private Sub txtContrase�a_KeyPress(KeyAscii As Integer)
  If InStr("0123456789" & _
           "abcdefghijklmnopqrstuvwxyz��" & _
           "ABCDEFGHIJKLMNOPQRSTUVWXYZ��", _
           Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
  End If
End Sub
Private Sub txtNuevaContrase�a_KeyPress(KeyAscii As Integer)
  If InStr("0123456789" & _
           "abcdefghijklmnopqrstuvwxyz��" & _
           "ABCDEFGHIJKLMNOPQRSTUVWXYZ��", _
           Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
  End If
End Sub
Private Sub txtNuevaContrase�a2_KeyPress(KeyAscii As Integer)
  If InStr("0123456789" & _
           "abcdefghijklmnopqrstuvwxyz��" & _
           "ABCDEFGHIJKLMNOPQRSTUVWXYZ��", _
           Chr(KeyAscii)) = 0 Then
    If KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
  End If
End Sub

