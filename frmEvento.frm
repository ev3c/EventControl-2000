VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEvento 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpFechaDesde 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510467
      CurrentDate     =   36545
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox lstEvento 
      Height          =   1035
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpFechaHasta 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510467
      CurrentDate     =   36545
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   3855
      Begin VB.Label lblFechaHasta 
         Alignment       =   1  'Right Justify
         Caption         =   "=<  Hasta Fecha"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFechaDesde 
         Caption         =   "Desde Fecha  >="
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label lblInfo0 
      Caption         =   "Eventos de:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblInfo1 
      Caption         =   "Para borrar los Eventos (Fecha, HoraOn y HoraOff)"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   3855
   End
End
Attribute VB_Name = "frmEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBorrar_Click()
  Dim dteFecha As Date
  Dim gstrFormatoFecha As String
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsgPrg = "Borrar Eventos"
      strMsgBorrar = "Seguro que desea Borrar los Eventos de:  " & lstEvento.Text
      gstrFormatoFecha = "dd/MM/yyyy"
      strMsgFechaErronea = "Fecha Desde es mayor que Fecha Hasta"
    Case EC_ENGLISH
      strMsgPrg = "Delete Events"
      strMsgBorrar = "Sure you wants to Delete the Events of:  " & lstEvento.Text
      gstrFormatoFecha = "MM/dd/yyyy"
      strMsgFechaErronea = "From Date it's bigger than To Date"
    Case EC_CATALA
      strMsgPrg = "Borrar Events"
      strMsgBorrar = "Segur que desitja Borrar els Events de:  " & lstEvento.Text
      gstrFormatoFecha = "dd/MM/yyyy"
      strMsgFechaErronea = "Data Des de és major que Data Fins a"
  End Select
  
  If dtpFechaDesde > dtpFechaHasta Then
    MsgBox strMsgFechaErronea, vbInformation, strMsgPrg
    Exit Sub
  End If
  
  If frmEventControl.Contraseña_Entrar Then
    intBorrar = MsgBox(strMsgBorrar, vbYesNo + vbCritical, strMsgPrg)
    If intBorrar = vbYes Then
      If grsEvento.RecordCount > 0 Then
        grsEvento.MoveFirst
      End If
      Do While Not grsEvento.EOF
        dteFecha = Format(grsEvento.Fields("Fecha"), _
                   gstrFormatoFecha)
        If dteFecha >= dtpFechaDesde And _
           dteFecha <= dtpFechaHasta Then
          Select Case lstEvento.ListIndex
            Case 0
              grsEvento.Delete
              'grsEvento.Update
            Case Else
              If grsEvento.Fields("ProgramaID") = _
                 lstEvento.ListIndex Then
                grsEvento.Delete
                'grsEvento.Update
              End If
          End Select
        End If
        grsEvento.MoveNext
      Loop
      
    End If
  End If
End Sub

Private Sub cmdSalir_Click()
  Unload frmEvento
End Sub


Private Sub Form_Load()
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strEvento0 = "Todos los Eventos"
      cmdBorrar.Caption = "&Borrar"
      cmdSalir.Caption = "&Salir"
      frmEvento.Caption = "Borrar Eventos"
      lblInfo0.Caption = "Eventos de:"
      lblInfo1.Caption = "Para Borrar los Eventos (Fecha, HoraOn y HoraOff)"
      
      lblFechaDesde = "Desde Fecha  >="
      lblFechaHasta = "=<  Hasta Fecha"
      dtpFechaDesde.CustomFormat = gstrFormatoFecha
      dtpFechaHasta.CustomFormat = gstrFormatoFecha
      
    Case EC_ENGLISH
      strEvento0 = "All the Events"
      cmdBorrar.Caption = "&Delete"
      cmdSalir.Caption = "&Exit"
      frmEvento.Caption = "Delete Events"
      lblInfo0.Caption = "Events of:"
      lblInfo1.Caption = "To Delete the Events (Date, TimeOn and TimeOff)"
  
      lblFechaDesde = "From Date  >="
      lblFechaHasta = "=<  To Date"
      dtpFechaDesde.CustomFormat = gstrFormatoFecha
      dtpFechaHasta.CustomFormat = gstrFormatoFecha
  
    Case EC_CATALA
      strEvento0 = "Tots els Events"
      cmdBorrar.Caption = "&Borrar"
      cmdSalir.Caption = "&Sortir"
      frmEvento.Caption = "Borrar Events"
      lblInfo0.Caption = "Events de:"
      lblInfo1.Caption = "Per Borrar els Events (Data, HoraOn i HoraOff)"
      
      lblFechaDesde = "Des de Data  >="
      lblFechaHasta = "=<  Fins a Data"
      dtpFechaDesde.CustomFormat = gstrFormatoFecha
      dtpFechaHasta.CustomFormat = gstrFormatoFecha
      
  End Select
    
  dtpFechaDesde.Value = dtpFechaDesde.MinDate
  dtpFechaHasta.Value = dtpFechaHasta.MaxDate
    
  lstEvento.Clear
  lstEvento.AddItem strEvento0
  grsPrograma.MoveFirst
  Do While Not grsPrograma.EOF
    lstEvento.AddItem grsPrograma.Fields("Nombre")
  grsPrograma.MoveNext
  Loop
  lstEvento.ListIndex = frmPrograma.lblID
End Sub

