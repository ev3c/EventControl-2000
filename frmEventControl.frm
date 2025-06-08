VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEventControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EventControl 2000 v1.0"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8325
   Icon            =   "frmEventControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInternetHistorial 
      Caption         =   "Internet &Historial"
      Height          =   375
      Left            =   6840
      TabIndex        =   59
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5400
      TabIndex        =   56
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdPausar 
      Caption         =   "Pa&usar"
      Height          =   375
      Left            =   6840
      TabIndex        =   45
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdInformacion 
      Height          =   615
      Left            =   7200
      Picture         =   "frmEventControl.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdAnalizarFecha 
      Cancel          =   -1  'True
      Caption         =   "Analizar &Fecha"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   1560
      Width           =   1695
   End
   Begin VB.PictureBox picIconoBandeja 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3720
      Picture         =   "frmEventControl.frx":0884
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Timer tmrEventControl 
      Left            =   3960
      Top             =   840
   End
   Begin VB.Frame fraControlTiempo 
      Caption         =   "Control de Tiempo"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   8055
      Begin VB.CommandButton cmdCrono 
         Caption         =   "On Crono"
         Height          =   300
         Left            =   5280
         TabIndex        =   53
         Top             =   250
         Width           =   1215
      End
      Begin VB.ComboBox cboPrg 
         Height          =   315
         ItemData        =   "frmEventControl.frx":0CC6
         Left            =   6720
         List            =   "frmEventControl.frx":0CC8
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.UpDown updWinIdx 
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblWinIdx"
         BuddyDispid     =   196644
         OrigLeft        =   1200
         OrigTop         =   480
         OrigRight       =   1395
         OrigBottom      =   735
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updScrIdx 
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblScrIdx"
         BuddyDispid     =   196643
         OrigLeft        =   2640
         OrigTop         =   480
         OrigRight       =   2835
         OrigBottom      =   735
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updMdmIdx 
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblMdmIdx"
         BuddyDispid     =   196642
         OrigLeft        =   4200
         OrigTop         =   840
         OrigRight       =   4395
         OrigBottom      =   1125
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updPrgIdx 
         Height          =   255
         Left            =   6960
         TabIndex        =   40
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblPrgIdx"
         BuddyDispid     =   196629
         OrigLeft        =   4200
         OrigTop         =   840
         OrigRight       =   4395
         OrigBottom      =   1125
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updTmrIdx 
         Height          =   255
         Left            =   5520
         TabIndex        =   47
         Top             =   600
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblTmrIdx"
         BuddyDispid     =   196619
         OrigLeft        =   4200
         OrigTop         =   840
         OrigRight       =   4395
         OrigBottom      =   1125
         Max             =   1
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTmrIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   52
         Top             =   600
         Width           =   300
      End
      Begin VB.Label lblTmrSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   51
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblTmrDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblTmrMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   49
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblTmrAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5760
         TabIndex        =   48
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPrgAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   44
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblPrgMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   43
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblPrgDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   42
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblPrgSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   41
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblPrgIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6720
         TabIndex        =   39
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblMdmAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   37
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblMdmMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblMdmDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   35
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblMdmSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblScrAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblScrMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblScrDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblScrSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblWinAño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblWinMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblWinSesion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblMdmIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   600
         Width           =   450
      End
      Begin VB.Label lblScrIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblWinIdx 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblMdm 
         Caption         =   "Modem"
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
         Left            =   4320
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblScr 
         Caption         =   "ScrSaver"
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
         Left            =   2880
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblWindows 
         Caption         =   "Windows"
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
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblAño 
         Caption         =   "Año"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblMes 
         Caption         =   "Mes"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDia 
         Caption         =   "Día"
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
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblSesion 
         Caption         =   "Sesión"
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
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdProgramas 
      Caption         =   "&Programas"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCambiarContraseña 
      Caption         =   "&Contraseña"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   840
      Width           =   1335
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
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cboIdioma 
      Height          =   315
      ItemData        =   "frmEventControl.frx":0CCA
      Left            =   3600
      List            =   "frmEventControl.frx":0CD7
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.CheckBox chkAutoArranque 
      Caption         =   "Arrancar EventControl al iniciar Windows"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.CheckBox chkVerPantalla 
      Caption         =   "Ver esta pantalla antes de salir de Windows"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.CheckBox chkMostrarIcono 
      Caption         =   "Mostrar Icono en Barra de Tareas"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Frame fraConfiguracion 
      Caption         =   "Configuración"
      Height          =   1155
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      Begin VB.PictureBox picIconoBandejaPausa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4320
         Picture         =   "frmEventControl.frx":0CF5
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   58
         Top             =   600
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   54
      Top             =   1320
      Width           =   5175
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   1800
         TabIndex        =   57
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21233667
         CurrentDate     =   36545
      End
      Begin VB.Label lblIntroduceFecha 
         Caption         =   "Introduce fecha:"
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
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Menu mnuIconoBandeja 
      Caption         =   "IconoBandeja"
      Visible         =   0   'False
      Begin VB.Menu mnuVerControlTiempo 
         Caption         =   "&Ver Control de Tiempo"
      End
      Begin VB.Menu mnuNull0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCambiarContraseña 
         Caption         =   "&Cambiar Contraseña"
      End
      Begin VB.Menu mnuProgramas 
         Caption         =   "&Programas"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuInternetHistorial 
         Caption         =   "Internet &Historial"
      End
      Begin VB.Menu mnuNull1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPausar 
         Caption         =   "Pa&usar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmEventControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const EC_SAVESETTINGS = True
#Const EC_EXE = True

Private Sub cboIdioma_Click()
  giIdioma = cboIdioma.ListIndex
  Call Configuracion_Idioma
End Sub

Private Sub cboPrg_Click()
  Dim dtFecha As Date
  
  gnPrg = cboPrg.ListIndex + 1
  If giPrg(gnPrg) <= 1 Then
    giPrg(gnPrg) = 1
  Else
    If gaPrg(gnPrg, giPrg(gnPrg)).On = CDate("0") Then
      If gtPrg(gnPrg).On = CDate("0") Then
        giPrg(gnPrg) = giPrg(gnPrg) - 1
      End If
    End If
  End If
  
  dtFecha = CDate(lblFecha)
  ipDia = Day(dtFecha)
  ipMes = Month(dtFecha)
  ipAño = Year(dtFecha)
  
  iDia = Day(MyDate)
  iMes = Month(MyDate)
  iAño = Year(MyDate)

  If gtPrg(gnPrg).On <> CDate("0") And lblFecha = MyDate Then
    gdPrgDia(gnPrg) = gdPrgDia(gnPrg) + (Time - gtPrg(gnPrg).On)
  End If
  If gtPrg(gnPrg).On <> CDate("0") And _
     iMes = ipMes And iAño = ipAño Then
    gdPrgMes(gnPrg) = gdPrgMes(gnPrg) + (Time - gtPrg(gnPrg).On)
  End If
  If gtPrg(gnPrg).On <> CDate("0") And iAño = ipAño Then
    gdPrgAño(gnPrg) = gdPrgAño(gnPrg) + (Time - gtPrg(gnPrg).On)
  End If
  updPrgIdx.Max = giPrg(gnPrg)
  updPrgIdx.Value = giPrg(gnPrg)
  lblPrgDia = Hora_Suma(CDate(gdPrgDia(gnPrg)))
  lblPrgMes = Hora_Suma(CDate(gdPrgMes(gnPrg)))
  lblPrgAño = Hora_Suma(CDate(gdPrgAño(gnPrg)))
  
  Call updPrgIdx_Change
End Sub

Private Sub chkAutoArranque_Click()
  If chkAutoArranque.Value = vbUnchecked Then
    If Not bolPrimerArranque Then
      If Not Contraseña_Entrar Then
         chkAutoArranque.Value = vbChecked
      End If
    End If
  End If
End Sub

Private Sub chkMostrarIcono_Click()
  If chkMostrarIcono.Value = vbUnchecked Then
    Select Case cboIdioma.ListIndex
      Case EC_ESPAÑOL            'Español
        strMsg = "Cuando pulse Aceptar no verá ni el Icono " & _
        "ni el Programa en la Barra de Tareas " & vbCrLf & _
        "RECUERDE que puede acceder a EventControl pulsando   Ctrl Alt Mayus E"
        strMsgSW = "Solo puede ocultar el Icono en la versión Registrada"
      Case EC_ENGLISH
        strMsg = "When push Accept don't can see the Icon " & _
        "and the Program in the Task Bar" & vbCrLf & _
        "REMEMBER you can acced to EventControl pushing   Ctrl Alt Shift E"
        strMsgSW = "Only can hide the Icon in the Registered version"
      Case EC_CATALA            'Español
        strMsg = "Quan pulsi Acceptar no veurà ni la Icona " & _
        "ni el Programa a la Barra de Tasques " & vbCrLf & _
        "RECORDI que pot accedir a EventControl pulsant   Ctrl Alt Majus E"
        strMsgSW = "Sols pot ocultar la Icona en la versió Registrada"
      Case -1
        Exit Sub
    End Select
    
    If MsgBox(strMsg, vbYesNo + vbCritical, gstrPrograma) = vbYes Then
      If EC_SHAREWARE Then
        Call MsgBox(strMsgSW, vbInformation, gstrPrograma)
        chkMostrarIcono.Value = vbChecked
      Else
        Call IconDelete
      End If
    Else
      chkMostrarIcono.Value = vbChecked
    End If
  Else
    If gblnPausar = False Then
      IconAdd picIconoBandeja.hwnd, picIconoBandeja.Picture, gstrIconMsg
    Else
      IconAdd picIconoBandeja.hwnd, picIconoBandejaPausa.Picture, gstrIconMsg
    End If
  End If
End Sub


Private Sub cmdAceptar_Click()
  frmEventControl.Hide
  App.TaskVisible = False
End Sub

Private Sub cmdAnalizarFecha_Click()
    If IsDate(dtpFecha) Then
      lblFecha = Format(dtpFecha, "dd/MM/yyyy")
      Select Case cboIdioma.ListIndex
        Case EC_ESPAÑOL
          lblAño.Caption = "Año " & Year(dtpFecha)
          lblMes.Caption = "Mes " & Month(dtpFecha)
          lblDia.Caption = "Día " & Day(dtpFecha)
        Case EC_ENGLISH
          lblAño.Caption = "Year " & Year(dtpFecha)
          lblMes.Caption = "Month " & Month(dtpFecha)
          lblDia.Caption = "Day " & Day(dtpFecha)
        Case EC_CATALA
          lblAño.Caption = "Any " & Year(dtpFecha)
          lblMes.Caption = "Mes " & Month(dtpFecha)
          lblDia.Caption = "Dia " & Day(dtpFecha)
      End Select
      
      Call Fecha_Analizar(lblFecha)
    Else
      Select Case cboIdioma.ListIndex
        Case EC_ESPAÑOL
          strMsg = "Fecha Incorrecta " & dtpFecha
        Case EC_ENGLISH
          strMsg = "Incorrect Date " & dtpFecha
        Case EC_CATALA
          strMsg = "Data Incorrecta " & dtpFecha
        End Select
        MsgBox strMsg, vbInformation, gstrPrograma
    End If
End Sub

Private Sub cmdCambiarContraseña_Click()
  If Not frmContraseñaCambiar.Visible Then
    Load frmContraseñaCambiar
    frmContraseñaCambiar.Show vbModal
  End If
End Sub

Private Sub cmdCrono_Click()
  If gblnCronoOn Then
    cmdCrono.Caption = "On Crono"
    gblnCronoOn = False
  Else
    cmdCrono.Caption = "Off Crono"
    gblnCronoOn = True
  End If
End Sub

Private Sub cmdImprimir_Click()
  If Not frmImprimir.Visible Then
    Load frmImprimir
    frmImprimir.Show vbModal
  End If
End Sub

Private Sub cmdInternetHistorial_Click()
 If Not frmInternetHistorial.Visible Then
    Load frmInternetHistorial
    frmInternetHistorial.Show vbModal
  End If
End Sub

Private Sub cmdPausar_Click()
  
  If frmEventControl.Contraseña_Entrar Then
    
    Select Case cboIdioma.ListIndex
    Case EC_ESPAÑOL
      If gblnPausar Then
        strMenu = "Pa&usar"
      Else
        strMenu = "React&ivar"
      End If
    Case EC_ENGLISH
      If gblnPausar Then
        strMenu = "Pa&use"
      Else
        strMenu = "React&ivate"
      End If
    Case EC_CATALA
      If gblnPausar Then
        strMenu = "Pa&usar"
      Else
        strMenu = "React&ivar"
      End If
    End Select
    cmdPausar.Caption = strMenu
    mnuPausar.Caption = strMenu

    If gblnPausar Then
      gblnPausar = False
      tmrEventControl.Enabled = True
      If chkMostrarIcono.Value = vbChecked Then
        Call IconModify(picIconoBandeja)
      End If
    Else
      gblnPausar = True
      tmrEventControl.Enabled = False
      If chkMostrarIcono.Value = vbChecked Then
        Call IconModify(picIconoBandejaPausa)
      End If
    End If
    
  End If

End Sub

Private Sub cmdProgramas_Click()
  If Not frmPrograma.Visible Then
    Load frmPrograma
    frmPrograma.Show vbModal
  End If
  Call cmdAnalizarFecha_Click
End Sub

Private Sub cmdSalir_Click()
  
  If Salir_SiNo = vbYes Then
    Call Form_Unload(False)
  End If
  
End Sub
Public Function Contraseña_Entrar() As Boolean

If EC_SHAREWARE Then      'No ho entenc
  Contraseña_Entrar = True
  Select Case cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsg = "Esta opción está Protegida por Contraseña en la" & vbCrLf & _
        "versión registrada de " & gstrPrograma
    Case EC_ENGLISH
      strMsg = "This option it's Password Protected in the" & vbCrLf & _
        "registered version of " & gstrPrograma
    Case EC_CATALA
      strMsg = "Aquesta opció està Protegida per Contrasenya a la" & vbCrLf & _
        "versió registrada de " & gstrPrograma
  End Select
  MsgBox strMsg, vbInformation + vbSystemModal + vbMsgBoxSetForeground, gstrPrograma
Else
  Load frmContraseñaEntrar
  frmContraseñaEntrar.Show vbModal
  If gblnContraseña = True Then
    Contraseña_Entrar = True
  Else
    Contraseña_Entrar = False
  End If
End If
End Function
Public Function Contraseña_Entrar_Timer() As Boolean

If EC_SHAREWARE Then      'No ho entenc
  Contraseña_Entrar_Timer = True
  Select Case cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsg = "El acceso a este programa está Protegido por Contraseña" & vbCrLf & _
        "en la versión registrada de " & gstrPrograma
    Case EC_ENGLISH
      strMsg = "The access to this program it's Password Protected" & vbCrLf & _
        "in the registered version of " & gstrPrograma
    Case EC_CATALA
      strMsg = "L'accés a aquest programa està Protegit per Contrasenya" & vbCrLf & _
        "a la versió registrada de " & gstrPrograma
  End Select
  MsgBox strMsg, vbInformation + vbSystemModal + vbMsgBoxSetForeground, gstrPrograma
Else
  
  gblnContraseña = False
  
  frmContraseñaEntrar.tmrContraseña.Enabled = True
  frmContraseñaEntrar.tmrContraseña.Interval = 30000
  
  Load frmContraseñaEntrar
  SetWindowPos frmContraseñaEntrar.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
               SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
  frmContraseñaEntrar.Show vbModal
  
  If gblnContraseña = True Then
    Contraseña_Entrar_Timer = True
  Else
    Contraseña_Entrar_Timer = False
  End If
End If
End Function

Private Sub Form_Load()
   
  On Error Resume Next
  
  'Establece MyDate
  MyDate = Date
  
  'Comprueba si ya está activo
  If App.PrevInstance Then
    End
  End If
        
  SetAttr App.Path, vbHidden
    
    
#If EC_EXE Then
  'Subclassifica la finestra per captura Ctrl+Alt+Mays+E
  Dim lRet As Long
  lRet = RegisterHotKey(frmEventControl.hwnd, &HB000&, _
         MOD_ALT Or MOD_CONTROL Or MOD_SHIFT, vbKeyE)
  Call Subclasifica_Ventana(frmEventControl.hwnd)
#End If
  
  gstrPrograma = "EventControl 2000 v" & _
    App.Major & "." & App.Minor & "." & App.Revision

  ReDim gaWin(1 To 99) As ecOnOff
  ReDim gaScr(1 To 99) As ecOnOff
  ReDim gaMdm(1 To 99) As ecOnOff
  ReDim gaTmr(1 To 99) As ecOnOff
  ReDim giPrg(1 To 99) As Integer
  ReDim gaPrg(1 To 99, 1 To 99) As ecOnOff
  
  ReDim gtPrg(1 To 99) As ecOnOff
  
  ReDim gdPrgDia(1 To 99) As Date
  ReDim gdPrgMes(1 To 99) As Date
  ReDim gdPrgAño(1 To 99) As Date
  
  dtpFecha.Value = MyDate
  lblFecha = Format(MyDate, "dd/MM/yyyy")
  
  
#If EC_EXE Then
  If getVersion = 1 Then
    'Oculta EventControl de Ctrl+Alt+Supr
    RegisterServiceProcess GetCurrentProcessId, 1
    'Hide app
  Else
    
  End If
#End If

  Call Evento_Abrir
  grsPrograma.MoveFirst
  Do While Not grsPrograma.EOF
    If grsPrograma.Fields("ID") > 4 Then
      gaPrograma(iPrg, 0) = grsPrograma.Fields("Nombre")
      gaPrograma(iPrg, 1) = grsPrograma.Fields("Path")
      gaPrograma(iPrg, 2) = grsPrograma.Fields("Password")
      iPrg = iPrg + 1
    End If
    grsPrograma.MoveNext
  Loop
  iPrg = 0
  Do While gaPrograma(iPrg, 0) <> ""
    cboPrg.AddItem gaPrograma(iPrg, 0)
    iPrg = iPrg + 1
  Loop
  
  Call Configuracion_Settings_Leer
  Call Configuracion_Idioma

  giWin = 1: updWinIdx.Max = 1: updWinIdx.Value = 1
  giScr = 1: updScrIdx.Max = 1: updScrIdx.Value = 1
  giMdm = 1: updMdmIdx.Max = 1: updMdmIdx.Value = 1
  giTmr = 1: updTmrIdx.Max = 1: updTmrIdx.Value = 1
  If cboPrg.ListIndex = -1 Then
    gnPrg = 1
  Else
    gnPrg = cboPrg.ListIndex + 1
  End If
  For x = 0 To cboPrg.ListCount
    giPrg(x + 1) = 1
    updPrgIdx.Max = giPrg(x + 1)
    updPrgIdx.Value = giPrg(x + 1)
  Next x
  
  gtWin.On = Time
  gdFechaOn = MyDate

  'Lee el numero del HD y comprueba
  gstrNumeroHD = LeerNumeroHD(Left$(App.Path, 3))

  If gstrNumeroRegistro = Encriptar(gstrNumeroHD) Then
    EC_SHAREWARE = False
  Else
    EC_SHAREWARE = True
  End If

  'Esconde el formulario.
  frmEventControl.Hide
  
  If EC_SHAREWARE Then
    Load frmInformacion
    frmInformacion.Show vbModal
  End If
  
  Call gaHistorial_Leer
  
  'Comprueba eventos cada segundo
  tmrEventControl.Interval = 1000
  
  
'  Call tmrEventControl_Timer
  Call cmdAnalizarFecha_Click
  
End Sub

Private Sub EnConstruccion()
    Select Case cboIdioma.ListIndex
      Case EC_ESPAÑOL
        strMsg = "Opción en Construcción"
      Case EC_ENGLISH
        strMsg = "Option Under Construction"
      Case EC_CATALA
        strMsg = "Opció en Construcció"
    End Select
    MsgBox strMsg, vbInformation, gstrPrograma
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Select Case UnloadMode
    Case vbFormControlMenu, vbAppTaskManager
      If Salir_SiNo = vbNo Then
        Cancel = True
      End If
    Case vbAppWindows   ', vbAppTaskManager
      If chkVerPantalla.Value = vbChecked Then
        Call mnuVerControlTiempo_Click
        'Delay
        For x = 1 To 10000000
        Next x
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If Cancel = False Then
   
   tmrEventControl.Enabled = False
   Call Configuracion_Settings_Grabar
   Call IconDelete
   Call Evento_Cerrar
   
#If EC_EXE Then
   'Recupera la HotKey y desclasifica ventana
   Call UnregisterHotKey(frmEventControl.hwnd, &HB000&)
   Call Ventana_Normal(frmEventControl.hwnd)

 ' Vuelve a mostrar Programa en Ctrl+Alt+Supr
   If getVersion = 1 Then
     'Put the following code in Form_Unload()
     RegisterServiceProcess GetCurrentProcessId, 0
     'Remove service flag
   End If

#End If

   '***************
   'End the program
   End
  
  End If
  
End Sub


Private Sub mnuCambiarContraseña_Click()
  frmEventControl.Show
  Call cmdCambiarContraseña_Click
End Sub


Private Sub mnuPausar_Click()
  Call cmdPausar_Click
End Sub

Private Sub mnuProgramas_Click()
  frmEventControl.Show
  Call cmdProgramas_Click
End Sub

Private Sub mnuImprimir_Click()
  frmEventControl.Show
  Call cmdImprimir_Click
End Sub
Private Sub mnuInternetHistorial_Click()
  frmEventControl.Show
  Call cmdInternetHistorial_Click
End Sub

Private Sub mnuSalir_Click()
  frmEventControl.Show
  Call cmdSalir_Click
End Sub

Public Sub mnuVerControlTiempo_Click()
  On Error GoTo errHandler
    frmEventControl.Show
    Call VerControlTiempo_Actualizar
    
    Exit Sub

errHandler:
  Select Case cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsg = "Había una subventana abierta"
    Case EC_ENGLISH
      strMsg = "There was open a subwindow"
    Case EC_CATALA
      strMsg = "Hi havia una subfinestra oberta"
  End Select
  MsgBox strMsg, vbInformation + vbSystemModal + vbMsgBoxSetForeground, gstrPrograma
End Sub

Private Sub cmdInformacion_Click()
  If Not frmInformacion.Visible Then
    Load frmInformacion
    frmInformacion.Show vbModal
  End If
End Sub

Private Sub picIconoBandeja_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Select Case x
    Case WM_LBUTTONDBLCLK       'Doble click
      Call mnuVerControlTiempo_Click
      
      SetWindowPos Me.hwnd, HWND_TOPMOST, _
        Me.Left / 15, Me.Top / 15, _
        Me.Width / 15, Me.Height / 15, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW
        
      SetWindowPos Me.hwnd, HWND_NOTOPMOST, _
        Me.Left / 15, Me.Top / 15, _
        Me.Width / 15, Me.Height / 15, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW
     
    Case WM_RBUTTONDOWN         'Botón derecho
      PopupMenu mnuIconoBandeja
  End Select
End Sub


Private Sub tmrEventControl_Timer()
  Dim iPrograma, iPrg As Integer
  Dim IsPrgOn(0 To 99) As Boolean
  Dim IsScrSavOn As Boolean
  Dim dFecha As Date
  Static bApagado(0 To 99) As Boolean
  
  'Lee los programas que están activos i el Salva Pantallas
  Call ControlEventos(IsScrSavOn, IsPrgOn())
  
  'Controla Screen Saver
  If IsScrSavOn Then
    If gtScr.On = CDate("0") Then
      If gaScr(giScr).On <> CDate("0") Then
        giScr = giScr + 1
      End If
      gtScr.On = Time
      updScrIdx.Max = giScr
      updScrIdx.Value = giScr
    End If
  Else
    If gtScr.On <> CDate("0") Then
      gtScr.Off = Time
    End If
  End If
  
  'Controla Modem
  If IsConnected Then
    If gtMdm.On = CDate("0") Then
      If gaMdm(giMdm).On <> CDate("0") Then
        giMdm = giMdm + 1
      End If
      gtMdm.On = Time
      updMdmIdx.Max = giMdm
      updMdmIdx.Value = giMdm
    End If
  Else
    If gtMdm.On <> CDate("0") Then
      gtMdm.Off = Time
    End If
  End If
  
  'Controla Cronometro
  If gblnCronoOn Then
    If gtTmr.On = CDate("0") Then
      If gaTmr(giTmr).On <> CDate("0") Then
        giTmr = giTmr + 1
      End If
      gtTmr.On = Time
      updTmrIdx.Max = giTmr
      updTmrIdx.Value = giTmr
    End If
  Else
    If gtTmr.On <> CDate("0") Then
      gtTmr.Off = Time
    End If
  End If
    
  'Controla Programa
  iPrograma = 0
  Do While gaPrograma(iPrograma, 1) <> ""
    iPrg = iPrograma + 1
    If IsPrgOn(iPrograma) Then
      
      If gaPrograma(iPrograma, 2) = 1 Then
        If Not bApagado(iPrograma) Then
          If Not Contraseña_Entrar_Timer Then
            Call Programa_Apagar(gaPrograma(iPrograma, 1))
            bApagado(iPrograma) = False
            Exit Sub  'Sale de timer para no contar programa
          Else
            bApagado(iPrograma) = True
          End If
        End If
      End If
    
      If gtPrg(iPrg).On = CDate("0") Then
        gtPrg(iPrg).On = Time
        If gaPrg(iPrg, giPrg(iPrg)).On <> CDate("0") Then
           giPrg(iPrg) = giPrg(iPrg) + 1
        End If
        If iPrograma = cboPrg.ListIndex Then
          updPrgIdx.Max = giPrg(iPrg)
          updPrgIdx.Value = giPrg(iPrg)
        End If
      End If
    Else
      If gtPrg(iPrg).On <> CDate("0") Then
        gtPrg(iPrg).Off = Time
        bApagado(iPrograma) = False
      End If
    End If
    iPrograma = iPrograma + 1
  Loop
  
  If Time >= #11:59:59 PM# Then
    MyDate = DateAdd("d", 1, MyDate)
  End If
  
  'Controla Medianoche
  If gdFechaOn < MyDate Then
    If gtWin.On <> CDate("0") Then
      Call Evento_Grabar(1, gdFechaOn, gtWin.On, #11:59:59 PM#)
    End If
    gtWin.On = #12:00:01 AM#
    gtWin.Off = CDate("0")
    If gtScr.On <> CDate("0") Then
      Call Evento_Grabar(2, gdFechaOn, gtScr.On, #11:59:59 PM#)
      gtScr.On = #12:00:01 AM#
    End If
    gtScr.Off = CDate("0")
    If gtMdm.On <> CDate("0") Then
      Call Evento_Grabar(3, gdFechaOn, gtMdm.On, #11:59:59 PM#)
      gtMdm.On = #12:00:01 AM#
    End If
    gtMdm.Off = CDate("0")
    If gtTmr.On <> CDate("0") Then
      Call Evento_Grabar(4, gdFechaOn, gtTmr.On, #11:59:59 PM#)
      gtTmr.On = #12:00:01 AM#
    End If
    gtTmr.Off = CDate("0")
   
    iPrograma = 0
    Do While gaPrograma(iPrograma, 1) <> ""
      iPrg = iPrograma + 1
      If gtPrg(iPrg).On <> CDate("0") Then
        Call Evento_Grabar(iPrg + 4, gdFechaOn, _
          gtPrg(iPrg).On, #11:59:59 PM#)
        gtPrg(iPrg).On = #12:00:01 AM#
      End If
      gtPrg(iPrg).Off = CDate("0")
      iPrograma = iPrograma + 1
    Loop
        
    gdFechaOn = MyDate
       
    dtpFecha.Value = MyDate
    'Delay
    For x = 1 To 10000000
    Next x
    Call cmdAnalizarFecha_Click
    Call gaHistorial_Leer
    
  End If
  
  Call VerControlTiempo_Actualizar

  'Lee URL Adress y las graba si no existen en el día
  Call GetURLAdress
  
End Sub


Private Sub updWinIdx_Change()
  Dim iWin As Integer
  iWin = updWinIdx.Value
  lblWinSesion = Format(gaWin(iWin).Off - gaWin(iWin).On, "h:mm:ss")
  lblWinSesion.ToolTipText = gaWin(iWin).On & " - " & gaWin(iWin).Off
End Sub

Private Sub updMdmIdx_Change()
  Dim iMdm As Integer
  iMdm = updMdmIdx.Value
  lblMdmSesion = Format(gaMdm(iMdm).Off - gaMdm(iMdm).On, "h:mm:ss")
  lblMdmSesion.ToolTipText = gaMdm(iMdm).On & " - " & gaMdm(iMdm).Off
End Sub

Private Sub updTmrIdx_Change()
  Dim iTmr As Integer
  iTmr = updTmrIdx.Value
  lblTmrSesion = Format(gaTmr(iTmr).Off - gaTmr(iTmr).On, "h:mm:ss")
  lblTmrSesion.ToolTipText = gaTmr(iTmr).On & " - " & gaTmr(iTmr).Off
End Sub

Private Sub updPrgIdx_Change()
  Dim iPrg As Integer
  iPrg = updPrgIdx.Value
  lblPrgSesion = Format(gaPrg(gnPrg, iPrg).Off - gaPrg(gnPrg, iPrg).On, "h:mm:ss")
  lblPrgSesion.ToolTipText = gaPrg(gnPrg, iPrg).On & " - " & gaPrg(gnPrg, iPrg).Off
End Sub

Private Sub updScrIdx_Change()
  Dim iScr As Integer
  iScr = updScrIdx.Value
  lblScrSesion = Format(gaScr(iScr).Off - gaScr(iScr).On, "h:mm:ss")
  lblScrSesion.ToolTipText = gaScr(iScr).On & " - " & gaScr(iScr).Off
End Sub

Public Sub Evento_Abrir()
  
  On Error GoTo Evento_Error
  
  gstrContraseña = Contraseña_DesEncriptar( _
      GetSetting("EventControl", "Inicio", _
      "gstrContraseña", ""))
  
  
  Set gdb = DBEngine.Workspaces(0).OpenDatabase( _
       App.Path & "\EventControl.mdb", True, False, _
       ";PWD=" & gstrContraseña)

  Set grsEvento = gdb.OpenRecordset("SELECT * " & _
                  "FROM tblEvento " & _
                  "ORDER BY ProgramaID, Fecha, HoraOn")
  
  Set grsPrograma = gdb.OpenRecordset("SELECT * " & _
                    "FROM tblPrograma")
  
  
  Set gdbHistorial = DBEngine.Workspaces(0).OpenDatabase( _
       App.Path & "\Historial.mdb", True, False, _
       ";PWD=" & gstrContraseña)
  
  Set grsHistorial = gdbHistorial.OpenRecordset("SELECT * " & _
                  "FROM tblHistorial " & _
                  "ORDER BY url_fecha, url_hora")
  
  
  Exit Sub
  
Evento_Error:

  cboIdioma.ListIndex = GetSetting("EventControl", _
      "Inicio", "cboIdioma", 0)
  Select Case cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsg = "Error al Abrir la Base de Datos" & vbCrLf & _
                Err.Number & " " & Err.Description
    Case EC_ENGLISH
      strMsg = "Error Opening the Data Base" & vbCrLf & _
                Err.Number & " " & Err.Description
    Case EC_CATALA
      strMsg = "Error en Obrir la Base de Dades" & vbCrLf & _
                Err.Number & " " & Err.Description
  End Select
  MsgBox strMsg, vbCritical, gstrPrograma
  End
  
End Sub

Public Sub Evento_Cerrar()
  
  Dim dFecha As Date
  dFecha = MyDate
  If gtWin.On <> CDate("0") Then
    Call Evento_Grabar(1, dFecha, gtWin.On, Time)
  End If
  If gtScr.On <> CDate("0") Then
    Call Evento_Grabar(2, dFecha, gtScr.On, Time)
  End If
  If gtMdm.On <> CDate("0") Then
    Call Evento_Grabar(3, dFecha, gtMdm.On, Time)
  End If
  If gtTmr.On <> CDate("0") Then
    Call Evento_Grabar(4, dFecha, gtTmr.On, Time)
  End If
  iPrograma = 0
  Do While gaPrograma(iPrograma, 1) <> ""
    If gtPrg(iPrograma + 1).On <> CDate("0") Then
      Call Evento_Grabar(iPrograma + 5, dFecha, _
          gtPrg(iPrograma + 1).On, Time)
    End If
    iPrograma = iPrograma + 1
  Loop
  
  
  'Cierra EventControl.mdb
  grsEvento.Close
  Set grsEvento = Nothing
  grsPrograma.Close
  Set grsPrograma = Nothing
  gdb.Close
  Set gdb = Nothing
  
  'Cierra Historial.mdb
  grsHistorial.Close
  Set grsHistorial = Nothing
  gdbHistorial.Close
  Set gdbHistorial = Nothing
  
End Sub

Public Sub Evento_Grabar(id, fecha, HoraOn, HoraOff)

  grsEvento.AddNew
  grsEvento.Fields("ProgramaID") = id
  grsEvento.Fields("Fecha") = fecha
  grsEvento.Fields("HoraOn") = HoraOn
  grsEvento.Fields("HoraOff") = HoraOff
  grsEvento.Update
  
End Sub
Public Sub Evento_Historial_Grabar(fecha, Hora, Adress)

  On Error GoTo evento_historial_grabar_err

  Adress = Left$(Adress, 250)

  grsHistorial.AddNew
  grsHistorial.Fields("url_fecha") = fecha
  grsHistorial.Fields("url_hora") = Hora
  grsHistorial.Fields("url_adress") = Adress
  grsHistorial.Update
  
evento_historial_grabar_err:
  
End Sub

Public Sub Fecha_Analizar(dtFecha As Date)
  Dim iDia, iMes, iAño As Integer
  Dim ipDia, ipMes, ipAño As Integer
  Dim dDia, dMes, dAño As Date
    
  frmEventControl.MousePointer = vbArrowHourglass
    
  ipDia = Day(dtFecha)
  ipMes = Month(dtFecha)
  ipAño = Year(dtFecha)
  
  ReDim gaWin(1 To 99) As ecOnOff
  ReDim gaScr(1 To 99) As ecOnOff
  ReDim gaMdm(1 To 99) As ecOnOff
  ReDim gaTmr(1 To 99) As ecOnOff
  ReDim giPrg(1 To 99) As Integer
  ReDim gaPrg(1 To 99, 1 To 99) As ecOnOff
  giWin = 1: giScr = 1: giMdm = 1: giTmr = 1
  For x = 0 To cboPrg.ListCount
    giPrg(x + 1) = 1
  Next x
  
  ReDim gdPrgDia(1 To 99) As Date
  ReDim gdPrgMes(1 To 99) As Date
  ReDim gdPrgAño(1 To 99) As Date
  
  If grsEvento.RecordCount > 0 Then
    grsEvento.MoveFirst
  End If
  Do While Not grsEvento.EOF
    fecha = grsEvento.Fields("Fecha")
    iDia = Day(fecha)
    iMes = Month(fecha)
    iAño = Year(fecha)
    HoraOn = grsEvento.Fields("HoraOn")
    HoraOff = grsEvento.Fields("HoraOff")
    Evento = grsEvento.Fields("ProgramaID")
    Select Case Evento
      Case 1  'Windows
        If fecha = dtFecha Then
          dWinDia = dWinDia + (HoraOff - HoraOn)
          gaWin(giWin).On = HoraOn
          gaWin(giWin).Off = HoraOff
          giWin = giWin + 1
        End If
        If iMes = ipMes And iAño = ipAño Then
          dWinMes = dWinMes + (HoraOff - HoraOn)
        End If
        If iAño = ipAño Then
          dWinAño = dWinAño + (HoraOff - HoraOn)
        End If
      Case 2  'Salva Pantallas
        If fecha = dtFecha Then
          dScrDia = dScrDia + (HoraOff - HoraOn)
          gaScr(giScr).On = HoraOn
          gaScr(giScr).Off = HoraOff
          giScr = giScr + 1
        End If
        If iMes = ipMes And iAño = ipAño Then
          dScrMes = dScrMes + (HoraOff - HoraOn)
        End If
        If iAño = ipAño Then
          dScrAño = dScrAño + (HoraOff - HoraOn)
        End If
      Case 3  ' Modem
        If fecha = dtFecha Then
          dMdmDia = dMdmDia + (HoraOff - HoraOn)
          gaMdm(giMdm).On = HoraOn
          gaMdm(giMdm).Off = HoraOff
          giMdm = giMdm + 1
        End If
        If iMes = ipMes And iAño = ipAño Then
          dMdmMes = dMdmMes + (HoraOff - HoraOn)
        End If
        If iAño = ipAño Then
          dMdmAño = dMdmAño + (HoraOff - HoraOn)
        End If
      Case 4 ' Cronometro
        If fecha = dtFecha Then
          dTmrDia = dTmrDia + (HoraOff - HoraOn)
          gaTmr(giTmr).On = HoraOn
          gaTmr(giTmr).Off = HoraOff
          giTmr = giTmr + 1
        End If
        If iMes = ipMes And iAño = ipAño Then
          dTmrMes = dTmrMes + (HoraOff - HoraOn)
        End If
        If iAño = ipAño Then
          dTmrAño = dTmrAño + (HoraOff - HoraOn)
        End If
      Case Else ' Programas
        iPrg = Evento - 4
        If fecha = dtFecha Then
          gdPrgDia(iPrg) = gdPrgDia(iPrg) + (HoraOff - HoraOn)
          gaPrg(iPrg, giPrg(iPrg)).On = HoraOn
          gaPrg(iPrg, giPrg(iPrg)).Off = HoraOff
          giPrg(iPrg) = giPrg(iPrg) + 1
        End If
        If iMes = ipMes And iAño = ipAño Then
          gdPrgMes(iPrg) = gdPrgMes(iPrg) + (HoraOff - HoraOn)
        End If
        If iAño = ipAño Then
          gdPrgAño(iPrg) = gdPrgAño(iPrg) + (HoraOff - HoraOn)
        End If
    End Select
    grsEvento.MoveNext
  Loop
  
  iDia = Day(MyDate)
  iMes = Month(MyDate)
  iAño = Year(MyDate)

  ' Windows
  If gtWin.On <> CDate("0") And dtFecha = MyDate Then
    dWinDia = dWinDia + (Time - gtWin.On)
  Else
    If giWin > 1 Then giWin = giWin - 1
  End If
  If gtWin.On <> CDate("0") And _
     iMes = ipMes And iAño = ipAño Then
    dWinMes = dWinMes + (Time - gtWin.On)
  End If
  If gtWin.On <> CDate("0") And iAño = ipAño Then
    dWinAño = dWinAño + (Time - gtWin.On)
  End If
  updWinIdx.Max = giWin
  updWinIdx.Value = giWin
  lblWinDia = Hora_Suma(CDate(dWinDia))
  lblWinMes = Hora_Suma(CDate(dWinMes))
  lblWinAño = Hora_Suma(CDate(dWinAño))
  
  ' Salva Pantallas
  If gtScr.On <> CDate("0") And dtFecha = MyDate Then
    dScrDia = dScrDia + (Time - gtScr.On)
  Else
    If giScr > 1 Then giScr = giScr - 1
  End If
  If gtScr.On <> CDate("0") And _
     iMes = ipMes And iAño = ipAño Then
    dScrMes = dScrMes + (Time - gtScr.On)
  End If
  If gtScr.On <> CDate("0") And iAño = ipAño Then
    dScrAño = dScrAño + (Time - gtScr.On)
  End If
  updScrIdx.Max = giScr
  updScrIdx.Value = giScr
  lblScrDia = Hora_Suma(CDate(dScrDia))
  lblScrMes = Hora_Suma(CDate(dScrMes))
  lblScrAño = Hora_Suma(CDate(dScrAño))
  
  ' Modem
  If gtMdm.On <> CDate("0") And dtFecha = MyDate Then
    dMdmDia = dMdmDia + (Time - gtMdm.On)
  Else
    If giMdm > 1 Then giMdm = giMdm - 1
  End If
  If gtMdm.On <> CDate("0") And _
     iMes = ipMes And iAño = ipAño Then
    dMdmMes = dMdmMes + (Time - gtMdm.On)
  End If
  If gtMdm.On <> CDate("0") And iAño = ipAño Then
    dMdmAño = dMdmAño + (Time - gtMdm.On)
  End If
  updMdmIdx.Max = giMdm
  updMdmIdx.Value = giMdm
  lblMdmDia = Hora_Suma(CDate(dMdmDia))
  lblMdmMes = Hora_Suma(CDate(dMdmMes))
  lblMdmAño = Hora_Suma(CDate(dMdmAño))
  
  ' Cronometro
  If gtTmr.On <> CDate("0") And dtFecha = MyDate Then
    dTmrDia = dTmrDia + (Time - gtTmr.On)
  Else
    If giTmr > 1 Then giTmr = giTmr - 1
  End If
  If gtTmr.On <> CDate("0") And _
     iMes = ipMes And iAño = ipAño Then
    dTmrMes = dTmrMes + (Time - gtTmr.On)
  End If
  If gtTmr.On <> CDate("0") And iAño = ipAño Then
    dTmrAño = dTmrAño + (Time - gtTmr.On)
  End If
  updTmrIdx.Max = giTmr
  updTmrIdx.Value = giTmr
  lblTmrDia = Hora_Suma(CDate(dTmrDia))
  lblTmrMes = Hora_Suma(CDate(dTmrMes))
  lblTmrAño = Hora_Suma(CDate(dTmrAño))
  
  ' Programa
  iPrograma = 0
  Do While gaPrograma(iPrograma, 1) <> ""
    iPrg = iPrograma + 1
    If gtPrg(iPrg).On <> CDate("0") And dtFecha = MyDate Then
      gdPrgDia(iPrg) = gdPrgDia(iPrg) + (Time - gtPrg(iPrg).On)
    Else
      If giPrg(iPrg) > 1 Then giPrg(iPrg) = giPrg(iPrg) - 1
    End If
    If gtPrg(iPrg).On <> CDate("0") And _
       iMes = ipMes And iAño = ipAño Then
      gdPrgMes(iPrg) = gdPrgMes(iPrg) + (Time - gtPrg(iPrg).On)
    End If
    If gtPrg(iPrg).On <> CDate("0") And iAño = ipAño Then
      gdPrgAño(iPrg) = gdPrgAño(iPrg) + (Time - gtPrg(iPrg).On)
    End If
    iPrograma = iPrograma + 1
  Loop
  updPrgIdx.Max = giPrg(gnPrg)
  updPrgIdx.Value = giPrg(gnPrg)
  lblPrgDia = Hora_Suma(CDate(gdPrgDia(gnPrg)))
  lblPrgMes = Hora_Suma(CDate(gdPrgMes(gnPrg)))
  lblPrgAño = Hora_Suma(CDate(gdPrgAño(gnPrg)))
     
  frmEventControl.MousePointer = vbDefault
 
End Sub

Public Function Salir_SiNo()
  If Contraseña_Entrar Then
    Select Case cboIdioma.ListIndex
      Case EC_ESPAÑOL
        strMsg = "Seguro que desea salir de " & gstrPrograma
      Case EC_ENGLISH
        strMsg = "Sure you wants to leave " & gstrPrograma
      Case EC_CATALA
        strMsg = "Segur que vol sortir de " & gstrPrograma
    End Select
    Salir_SiNo = MsgBox(strMsg, vbYesNo + vbCritical, gstrPrograma)
  Else
    Salir_SiNo = vbNo
  End If
End Function
Public Sub Configuracion_Idioma()

  frmEventControl.Caption = gstrPrograma

  Select Case cboIdioma.ListIndex
    Case EC_ESPAÑOL
      fraControlTiempo.Caption = "Control de Tiempo"
      fraConfiguracion.Caption = "Configuración"
      lblSesion.Caption = "Sesión"
      lblAño.Caption = "Año " & Year(lblFecha)
      lblMes.Caption = "Mes " & Month(lblFecha)
      lblDia.Caption = "Día " & Day(lblFecha)
      cmdPausar.Caption = "Pa&usar"
      cmdSalir.Caption = "&Salir"
      cmdProgramas.Caption = "&Programas"
      cmdImprimir.Caption = "&Imprimir"
      cmdCambiarContraseña.Caption = "&Ca.Contraseña"
      cmdAceptar.Caption = "&Aceptar"
      cmdInternetHistorial.Caption = "Internet &Historial"
      cmdAnalizarFecha.Caption = "Analizar &Fecha"
      chkAutoArranque.Caption = "Arrancar EventControl al iniciar Windows"
      chkVerPantalla.Caption = "Ver esta pantalla antes de salir de Windows"
      chkMostrarIcono.Caption = "Mostrar Icono en Barra de Tareas"
      lblIntroduceFecha.Caption = "Introduce fecha:"
      lblFecha.ToolTipText = "Fecha Analizada"
      mnuVerControlTiempo.Caption = "&Ver Control de Tiempo"
      mnuCambiarContraseña.Caption = "&Cambiar Contraseña"
      mnuProgramas.Caption = "&Programas"
      mnuImprimir.Caption = "&Imprimir"
      mnuInternetHistorial.Caption = "Internet &Historial"
      mnuPausar.Caption = "Pa&usar"
      mnuSalir.Caption = "&Salir"
      dtpFecha.CustomFormat = "dd/MM/yyyy"
      cmdInformacion.ToolTipText = "Ver Información"
      cboPrg.ToolTipText = "Programas definidos por el Usuario"
      cboIdioma.ToolTipText = "Lenguaje usado en " & gstrPrograma
      lblWindows.ToolTipText = "Sistema Operativo Windows"
      lblScr.ToolTipText = "Salva Pantallas de Windows"
      lblMdm.ToolTipText = "Conexión Telefónica por Módem"
      
      gstrIconMsg = "Doble Click para ver Control de Tiempo"
      
      gstrFormatoFecha = "dd/MM/yyyy"
      
  Case EC_ENGLISH
      
      fraControlTiempo.Caption = "Control of Time"
      fraConfiguracion.Caption = "Configuration"
      lblSesion.Caption = "Session"
      lblAño.Caption = "Year " & Year(lblFecha)
      lblMes.Caption = "Month " & Month(lblFecha)
      lblDia.Caption = "Day " & Day(lblFecha)
      cmdPausar.Caption = "Pa&use"
      cmdSalir.Caption = "&Exit"
      cmdProgramas.Caption = "&Programs"
      cmdImprimir.Caption = "Pr&int"
      cmdCambiarContraseña.Caption = "&Ch.Password"
      cmdAceptar.Caption = "&Accept"
      cmdInternetHistorial.Caption = "Internet &History"
      cmdAnalizarFecha.Caption = "Analyze &Date"
      chkAutoArranque.Caption = "Start EventControl when Windows begin"
      chkVerPantalla.Caption = "See this Screen before Exit Windows"
      chkMostrarIcono.Caption = "Show Icon in Task Bar"
      lblIntroduceFecha.Caption = "Introduce Date:"
      lblFecha.ToolTipText = "Analyzed Date"
      mnuVerControlTiempo.Caption = "&See Control of Time"
      mnuCambiarContraseña.Caption = "&Change Password"
      mnuProgramas.Caption = "&Programs"
      mnuImprimir.Caption = "&Pr&int"
      mnuInternetHistorial.Caption = "Internet &History"
      mnuPausar.Caption = "Pa&use"
      mnuSalir.Caption = "&Exit"
      dtpFecha.CustomFormat = "MM/dd/yyyy"
      cmdInformacion.ToolTipText = "Show Information"
      cboPrg.ToolTipText = "User defined Programs"
      cboIdioma.ToolTipText = "Lenguage used in " & gstrPrograma
      lblWindows.ToolTipText = "Windows System Operating"
      lblScr.ToolTipText = "Windows Screen Saver"
      lblMdm.ToolTipText = "Telefonic Connection via Módem"
     
      gstrIconMsg = "Double Click to see Control of Time"
  
      gstrFormatoFecha = "MM/dd/yyyy"
  
    Case EC_CATALA
      fraControlTiempo.Caption = "Control de Temps"
      fraConfiguracion.Caption = "Configuració"
      lblSesion.Caption = "Sessió"
      lblAño.Caption = "Any " & Year(lblFecha)
      lblMes.Caption = "Mes " & Month(lblFecha)
      lblDia.Caption = "Dia " & Day(lblFecha)
      cmdPausar.Caption = "Pa&usar"
      cmdSalir.Caption = "&Sortir"
      cmdProgramas.Caption = "&Programes"
      cmdImprimir.Caption = "&Imprimir"
      cmdCambiarContraseña.Caption = "&Ca.Contrasenya"
      cmdAceptar.Caption = "&Acceptar"
      cmdInternetHistorial.Caption = "Internet &Historial"
      cmdAnalizarFecha.Caption = "Analitzar &Data"
      chkAutoArranque.Caption = "Engegar EventControl al iniciar Windows"
      chkVerPantalla.Caption = "Veure aquesta pantalla abans de sortir de Windows"
      chkMostrarIcono.Caption = "Mostrar Icona a la Barra de Tasques"
      lblIntroduceFecha.Caption = "Introduïr data:"
      lblFecha.ToolTipText = "Data Analitzada"
      mnuVerControlTiempo.Caption = "&Veure Control de Temps"
      mnuCambiarContraseña.Caption = "&Cambiar Contrasenya"
      mnuProgramas.Caption = "&Programes"
      mnuImprimir.Caption = "&Imprimir"
      mnuInternetHistorial.Caption = "Internet &Historial"
      mnuPausar.Caption = "Pa&usar"
      mnuSalir.Caption = "&Sortir"
      dtpFecha.CustomFormat = "dd/MM/yyyy"
      cmdInformacion.ToolTipText = "Veure Informació"
      cboPrg.ToolTipText = "Programes definits per l'Usuari"
      cboIdioma.ToolTipText = "Llenguatge usat a " & gstrPrograma
      lblWindows.ToolTipText = "Sistema Operatiu Windows"
      lblScr.ToolTipText = "Salva Pantalles de Windows"
      lblMdm.ToolTipText = "Conexió Telefònica per Mòdem"
      
      gstrIconMsg = "Doble Click per veure Control de Temps"
      
      gstrFormatoFecha = "dd/MM/yyyy"
      
  End Select
  
  If chkMostrarIcono.Value = vbChecked Then
      IconDelete
      IconAdd picIconoBandeja.hwnd, picIconoBandeja.Picture, gstrIconMsg
  Else
      IconAdd picIconoBandeja.hwnd, picIconoBandeja.Picture, gstrIconMsg
      IconDelete
  End If
End Sub
Public Sub Configuracion_Settings_Leer()
  On Error Resume Next
  
  bolPrimerArranque = True
  
  chkAutoArranque.Value = GetSetting("EventControl", _
      "Inicio", "chkAutoArranque", vbChecked)
  chkMostrarIcono.Value = GetSetting("EventControl", _
      "Inicio", "chkMostrarIcono", vbChecked)
  chkVerPantalla.Value = GetSetting("EventControl", _
      "Inicio", "chkVerPantalla", vbChecked)
  cboIdioma.ListIndex = GetSetting("EventControl", _
      "Inicio", "cboIdioma", 0)
  gstrContraseña = Contraseña_DesEncriptar( _
      GetSetting("EventControl", "Inicio", _
      "gstrContraseña", ""))
  cboPrg.ListIndex = GetSetting("EventControl", _
      "Inicio", "cboPrg", 0)
  gstrNumeroRegistro = GetSetting("EventControl", _
      "Inicio", "gstrNumeroRegistro", "")
 
 
  If chkAutoArranque.Value = vbUnchecked Then
    DeleteValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Run", _
    "EventControl"
  Else
    res = SetRegValue(HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Run", _
    "EventControl", _
    App.Path & "\EventControl.exe")
  End If
  
  bolPrimerArranque = False

End Sub

Public Sub Configuracion_Settings_Grabar()

#If EC_SAVESETTINGS Then        'No ho entenc
  
  SaveSetting "EventControl", "Inicio", _
      "chkAutoArranque", chkAutoArranque.Value
  SaveSetting "EventControl", "Inicio", _
      "chkMostrarIcono", chkMostrarIcono.Value
  SaveSetting "EventControl", "Inicio", _
      "chkverpantalla", chkVerPantalla.Value
  SaveSetting "EventControl", "Inicio", _
      "cboIdioma", cboIdioma.ListIndex
  SaveSetting "EventControl", "Inicio", _
      "gstrContraseña", _
      Contraseña_Encriptar(gstrContraseña)
  SaveSetting "EventControl", "Inicio", _
      "cboPrg", cboPrg.ListIndex
  SaveSetting "EventControl", "Inicio", _
      "gstrNumeroRegistro", gstrNumeroRegistro
  
  If chkAutoArranque.Value = vbUnchecked Then
    DeleteValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Run", _
    "EventControl"
  Else
    res = SetRegValue(HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Run", _
    "EventControl", _
    App.Path & "\EventControl.exe")
  End If
   
#End If

End Sub
Public Sub VerControlTiempo_Actualizar()
  Dim dDate As Date
  
  dDate = Format(MyDate, "DD/MM/YYYY")

  'Controla Windows
  If gtWin.Off = CDate("0") Then
    If gtWin.On <> CDate("0") Then
      If giWin = updWinIdx.Value And lblFecha = dDate Then
        lblWinSesion = Format(Time - gtWin.On, "h:mm:ss")
        lblWinSesion.ToolTipText = gtWin.On & " - " & Time
      End If
    End If
  Else
    If giWin = updWinIdx.Value And lblFecha = dDate Then
      lblWinSesion = Format(gtWin.Off - gtWin.On, "h:mm:ss")
      lblWinSesion.ToolTipText = gtWin.On & " - " & gtWin.Off
    End If
    Call Evento_Grabar(1, MyDate, gtWin.On, gtWin.Off)
    gaWin(giWin).On = gtWin.On
    gaWin(giWin).Off = gtWin.Off
    gtWin.On = CDate("0")
    gtWin.Off = CDate("0")
  End If
  
  'Controla el Salva Pantallas
  If gtScr.Off = CDate("0") Then
    If gtScr.On <> CDate("0") Then
      If giScr = updScrIdx.Value And lblFecha = dDate Then
        lblScrSesion = Format(Time - gtScr.On, "h:mm:ss")
        lblScrSesion.ToolTipText = gtScr.On & " - " & Time
      End If
    End If
  Else
    If giScr = updScrIdx.Value And lblFecha = dDate Then
      lblScrSesion = Format(gtScr.Off - gtScr.On, "h:mm:ss")
      lblScrSesion.ToolTipText = gtScr.On & " - " & gtScr.Off
    End If
    Call Evento_Grabar(2, MyDate, gtScr.On, gtScr.Off)
    gaScr(giScr).On = gtScr.On
    gaScr(giScr).Off = gtScr.Off
    gtScr.On = CDate("0")
    gtScr.Off = CDate("0")
  End If
  
  'Controla el Módem
  If gtMdm.Off = CDate("0") Then
    If gtMdm.On <> CDate("0") Then
      If giMdm = updMdmIdx.Value And lblFecha = dDate Then
        lblMdmSesion = Format(Time - gtMdm.On, "h:mm:ss")
        lblMdmSesion.ToolTipText = gtMdm.On & " - " & Time
      End If
    End If
  Else
    If giMdm = updMdmIdx.Value And lblFecha = dDate Then
      lblMdmSesion = Format(gtMdm.Off - gtMdm.On, "h:mm:ss")
      lblMdmSesion.ToolTipText = gtMdm.On & " - " & gtMdm.Off
    End If
    Call Evento_Grabar(3, MyDate, gtMdm.On, gtMdm.Off)
    gaMdm(giMdm).On = gtMdm.On
    gaMdm(giMdm).Off = gtMdm.Off
    gtMdm.On = CDate("0")
    gtMdm.Off = CDate("0")
  End If

  'Controla el Cronometro
  If gtTmr.Off = CDate("0") Then
    If gtTmr.On <> CDate("0") Then
      If giTmr = updTmrIdx.Value And lblFecha = dDate Then
        lblTmrSesion = Format(Time - gtTmr.On, "h:mm:ss")
        lblTmrSesion.ToolTipText = gtTmr.On & " - " & Time
      End If
    End If
  Else
    If giTmr = updTmrIdx.Value And lblFecha = dDate Then
      lblTmrSesion = Format(gtTmr.Off - gtTmr.On, "h:mm:ss")
      lblTmrSesion.ToolTipText = gtTmr.On & " - " & gtTmr.Off
    End If
    Call Evento_Grabar(4, MyDate, gtTmr.On, gtTmr.Off)
    gaTmr(giTmr).On = gtTmr.On
    gaTmr(giTmr).Off = gtTmr.Off
    gtTmr.On = CDate("0")
    gtTmr.Off = CDate("0")
  End If

  'Controla los Programas
  iPrograma = 0
  Do While gaPrograma(iPrograma, 1) <> ""
    iPrg = iPrograma + 1
    If gtPrg(iPrg).Off = CDate("0") Then
      If gtPrg(gnPrg).On <> CDate("0") Then
        If giPrg(gnPrg) = updPrgIdx.Value And lblFecha = dDate Then
          lblPrgSesion = Format(Time - gtPrg(gnPrg).On, "h:mm:ss")
          lblPrgSesion.ToolTipText = gtPrg(gnPrg).On & " - " & Time
        End If
      End If
    Else
      If giPrg(gnPrg) = updPrgIdx.Value And lblFecha = dDate Then
        lblPrgSesion = Format(gtPrg(gnPrg).Off - gtPrg(gnPrg).On, "h:mm:ss")
        lblPrgSesion.ToolTipText = gtPrg(gnPrg).On & " - " _
                                  & gtPrg(gnPrg).Off
      End If
      Call Evento_Grabar(iPrg + 4, MyDate, _
        gtPrg(iPrg).On, gtPrg(iPrg).Off)
      gaPrg(iPrg, giPrg(iPrg)).On = gtPrg(iPrg).On
      gaPrg(iPrg, giPrg(iPrg)).Off = gtPrg(iPrg).Off
      gtPrg(iPrg).On = CDate("0")
      gtPrg(iPrg).Off = CDate("0")
    End If
  iPrograma = iPrograma + 1
  Loop

End Sub

Public Sub gaHistorial_Leer()

  ReDim gaHistorial(0 To 999)

'  If grsHistorial.RecordCount > 0 Then
'    grsHistorial.MoveFirst
'  End If
'  Do While Not grsHistorial.EOF
'    fecha = grsHistorial.Fields("url_fecha")
'    If fecha = MyDate Then
'       gaHistorial(x) = grsHistorial.Fields("url_adress")
'       x = x + 1
'    End If
'    grsHistorial.MoveNext
'  Loop
 
End Sub


