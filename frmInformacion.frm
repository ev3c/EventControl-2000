VERSION 5.00
Begin VB.Form frmInformacion 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picShareIt 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   720
      MouseIcon       =   "frmInformacion.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmInformacion.frx":030A
      ScaleHeight     =   435
      ScaleWidth      =   1500
      TabIndex        =   16
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   4200
      TabIndex        =   11
      Top             =   3600
      Width           =   3615
      Begin VB.TextBox txtNumeroRegistro 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblNumeroHD 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblNumeroRegistro 
         Caption         =   "Numero de Registro"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblIdentificacion 
         Caption         =   "Identificación"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton cmdFormulario 
      Caption         =   "&Editar Formulario de Registro"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   3855
      Begin VB.Label lblWeb 
         Caption         =   "http://perso.wanadoo.es/evalenti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         MouseIcon       =   "frmInformacion.frx":08FB
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label lblMailWanadoo 
         Caption         =   "evalenti@wanadoo.es"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         MouseIcon       =   "frmInformacion.frx":0C05
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Girona - Catalunya - España"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "17220 - Sant Feliu de Guíxols"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Grup Sot dels Canyers, esc-6 pis-60"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "© PsicoSoft && Esteve Valentí"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   2640
      Picture         =   "frmInformacion.frx":1047
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label lblShareWare 
      Height          =   615
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label lblPrograma 
      Caption         =   "Event Control 2.000 v1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   2640
      Picture         =   "frmInformacion.frx":1D10
      Top             =   2400
      Width           =   480
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   2400
      Top             =   120
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   2625
      Left            =   120
      Picture         =   "frmInformacion.frx":2152
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   2400
      Picture         =   "frmInformacion.frx":72D0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmInformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  
  gstrNumeroRegistro = Trim(UCase(CStr(txtNumeroRegistro)))
  
  If Encriptar(gstrNumeroHD) = gstrNumeroRegistro Then
    EC_SHAREWARE = False
  Else
    If gstrNumeroRegistro <> "" Then
      Select Case frmEventControl.cboIdioma.ListIndex
        Case EC_ESPAÑOL
          MsgBox "Número de Registro ERRONEO", _
          vbInformation, "Registrar " & gstrPrograma
        Case EC_ENGLISH
          MsgBox "WRONG Register Number", _
          vbInformation, "Register " & gstrPrograma
        Case EC_CATALA
          MsgBox "Número de Registre ERRONI", _
          vbInformation, "Registrar " & gstrPrograma
      End Select
    End If
    EC_SHAREWARE = True
  End If
  
  Unload frmInformacion
End Sub

Private Sub cmdFormulario_Click()
  Dim oeCrLf, strBody As String
  oeCrLf = "%0a"
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strBody = "Identificación: " & gstrNumeroHD & oeCrLf & _
                "Empresa:" & oeCrLf & _
                "Apellidos:" & oeCrLf & _
                "Nombre:" & oeCrLf & _
                "Dirección:" & oeCrLf & _
                "Código Postal:" & oeCrLf & _
                "Población:" & oeCrLf & _
                "País:" & oeCrLf & _
                "e-mail:" & oeCrLf & oeCrLf & _
                "REGISTRO 1 programa ....... 2000pts/12euros/10$" & oeCrLf & _
                "1)Ingreso/Transferencia en c/c La Caixa 2100-3701-11-2500016526" & oeCrLf & _
                "    (indicar Nombre o Identificación en la transferencia)" & oeCrLf & _
                "2)Envío cheque al portador" & oeCrLf & _
                "3)Envío en Metálico"
    Case EC_ENGLISH
      strBody = "Identification: " & gstrNumeroHD & oeCrLf & _
                "Company:" & oeCrLf & _
                "Surname:" & oeCrLf & _
                "Name:" & oeCrLf & _
                "Direction:" & oeCrLf & _
                "Postal Code:" & oeCrLf & _
                "City:" & oeCrLf & _
                "Country:" & oeCrLf & _
                "e-mail:" & oeCrLf & oeCrLf & _
                "REGISTER 1 program ....... 2000pts/12euros/10$" & oeCrLf & _
                "1))Entrance/Transference in c/c La Caixa 2700-3701-11-2500016526" & oeCrLf & _
                "    (indicate Name or Identification in the transference)" & oeCrLf & _
                "2)Postal shipment of check to the carrier " & oeCrLf & _
                "3)Postal shipment in Cash"
   Case EC_CATALA
      strBody = "Identificació: " & gstrNumeroHD & oeCrLf & _
                "Empresa:" & oeCrLf & _
                "Cognoms:" & oeCrLf & _
                "Nom:" & oeCrLf & _
                "Direcció:" & oeCrLf & _
                "Codi Postal:" & oeCrLf & _
                "Població:" & oeCrLf & _
                "País:" & oeCrLf & _
                "e-mail:" & oeCrLf & oeCrLf & _
                "REGISTRE 1 programa ....... 2000pts/12euros/10$" & oeCrLf & _
                "1)Ingrés/Transferència al c/c La Caixa 2700-3701-11-2500016526" & oeCrLf & _
                "    (indicar Nom o Identificació a la transferència)" & oeCrLf & _
                "2)Enviament txec al portador" & oeCrLf & _
                "3)Enviament en Metàl·lic"
 End Select
  Ret = ShellExecute(Me.hwnd, "Open", _
       "mailto:evalenti@wanadoo.es?subject=REGISTRAR " & gstrPrograma & _
       " - ID:" & gstrNumeroHD & _
       "&Body=" & strBody, "", "", 3)
End Sub


Private Sub Form_Load()
    
  txtNumeroRegistro = gstrNumeroRegistro
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      frmInformacion.Caption = "Información"
      cmdFormulario.Caption = "&Enviar Formulario de Registro"
      cmdAceptar.Caption = "&Aceptar"
      lblIdentificacion.Caption = "Identificación"
      lblNumeroRegistro.Caption = "Número de Registro"
      picShareIt.ToolTipText = "Registrar el programa con Tarjeta de Crédito"
    Case EC_ENGLISH
      frmInformacion.Caption = "Information"
      cmdFormulario.Caption = "&Send Register Form"
      cmdAceptar.Caption = "&Accept"
      lblIdentificacion.Caption = "Identification"
      lblNumeroRegistro.Caption = "Register Number"
      picShareIt.ToolTipText = "Register the program with Credit Card"
    Case EC_CATALA
      frmInformacion.Caption = "Informació"
      cmdFormulario.Caption = "&Enviar Formulari de Registre"
      cmdAceptar.Caption = "&Acceptar"
      lblIdentificacion.Caption = "Identificació"
      lblNumeroRegistro.Caption = "Número de Registre"
      picShareIt.ToolTipText = "Registrar el programa amb Tarjeta de Crèdit"
  End Select
  
  If EC_SHAREWARE Then
    Select Case frmEventControl.cboIdioma.ListIndex
      Case EC_ESPAÑOL
        lblShareWare.Caption = "Este programa es ShareWare. Por favor, REGISTRA el programa para colaborar en el desarrollo de nuevas versiones."
      Case EC_ENGLISH
        lblShareWare.Caption = "This program is ShareWare. Please, REGISTERS the program to collaborate in the development of new versions."
      Case EC_CATALA
        lblShareWare.Caption = "Aquest programa és ShareWare. Per favor, REGISTRA el programa per a col·laborar en el desenvopulament de noves versions."
    End Select
  Else
    Select Case frmEventControl.cboIdioma.ListIndex
      Case EC_ESPAÑOL
        lblShareWare.Caption = "Programa REGISTRADO."
      Case EC_ENGLISH
        lblShareWare.Caption = "REGISTERED Program."
      Case EC_CATALA
        lblShareWare.Caption = "Programa REGISTRAT."
    End Select
  
    txtNumeroRegistro.Enabled = False
  
  End If
  
  lblNumeroHD = gstrNumeroHD
  lblPrograma.Caption = gstrPrograma
End Sub

Private Sub lblMailWanadoo_Click()
  Call ShellExecute(Me.hwnd, "Open", "mailto:evalenti@wanadoo.es?subject=" & gstrPrograma & " - ID:" & gstrNumeroHD, "", "", 3)
End Sub

Private Sub lblWeb_Click()
  Call ShellExecute(Me.hwnd, "open", "http://perso.wanadoo.es/evalenti", "", "", 5)
End Sub

Private Sub picShareIt_Click()
Dim strShareIt As String
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strShareItWeb = "http://shareit1.element5.com/programs.html?productid=139357&language=Spanish"
    Case EC_ENGLISH
      strShareItWeb = "http://shareit1.element5.com/programs.html?productid=139357&language=English"
    Case EC_CATALA
      strShareItWeb = "http://shareit1.element5.com/programs.html?productid=139357&language=Spanish"
  End Select

  Call ShellExecute(Me.hwnd, "open", strShareItWeb, "", "", 5)

End Sub
