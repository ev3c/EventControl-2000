VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpFechaDesde 
      Height          =   375
      Left            =   600
      TabIndex        =   4
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
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      TabIndex        =   5
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
      TabIndex        =   6
      Top             =   1680
      Width           =   3855
      Begin VB.Label lblFechaHasta 
         Alignment       =   1  'Right Justify
         Caption         =   "=<  Hasta Fecha"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFechaDesde 
         Caption         =   "Desde Fecha  >="
         Height          =   255
         Left            =   240
         TabIndex        =   7
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
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mintPrnLinea As Integer   'Numero de línea impresa
Dim mintPrnLen As Integer     'Longitud de la página

Private Sub cmdImprimir_Click()
  Dim dteFecha As Date
  Dim intID As Integer    'Ultimo Evento Impreso
  Dim dteTiempo As Date   'Suma de los tiempos
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsgPrg = "Imprimir Eventos"
      strMsgFechaErronea = "Fecha Desde es mayor que Fecha Hasta"
      strMsgInfo = "Los Eventos activos en este momento no serán listados"
    Case EC_ENGLISH
      strMsgPrg = "Print Events"
      strMsgFechaErronea = "From Date it's bigger than To Date"
      strMsgInfo = "Active Events on this moment can't be listed"
    Case EC_CATALA
      strMsgPrg = "Imprimir Events"
      strMsgFechaErronea = "Data Des de és major que Data Fins a"
      strMsgInfo = "Els Events actius en aquest moment no seran llistats"
  End Select
  
  If dtpFechaDesde > dtpFechaHasta Then
    MsgBox strMsgFechaErronea, vbInformation, strMsgPrg
    Exit Sub
  End If
    
  Set grsEvento = gdb.OpenRecordset("SELECT * " & _
                  "FROM tblEvento " & _
                  "ORDER BY ProgramaID, Fecha, HoraOn")

  If grsEvento.RecordCount > 0 Then
    grsEvento.MoveFirst
  End If
  
  mintPrnLinea = 0      'Línea impresa numero 0
  
  Call Imprimir_Titulo
  Call Imprimir_SubTitulo
   
  Do While Not grsEvento.EOF
    dteFecha = Format(grsEvento.Fields("Fecha"), _
               gstrFormatoFecha)
    If dteFecha >= dtpFechaDesde And _
       dteFecha <= dtpFechaHasta Then
      
      If grsEvento.Fields("ProgramaID") = _
             lstEvento.ListIndex Or _
             lstEvento.ListIndex = 0 Then
         
        If intID <> grsEvento.Fields("ProgramaID") Then
          If intID <> 0 Then
            Call Imprimir_SubTotal(dteTiempo)
            dteTiempo = #12:00:00 AM#
            intID = grsEvento.Fields("ProgramaID")
          Else
            If dteTiempo <> #12:00:00 AM# Or intID = 0 Then
              intID = grsEvento.Fields("ProgramaID")
            End If
          End If
        End If
         
         dteTiempo = dteTiempo + _
                     CDate(grsEvento.Fields("HoraOff") - _
                     grsEvento.Fields("HoraOn"))
         
         Call Imprimir_Evento
       
      End If
    End If
    grsEvento.MoveNext
  Loop
  
  Call Imprimir_SubTotal(dteTiempo)
  
  Printer.EndDoc
  
  MsgBox strMsgInfo, vbInformation, strMsgPrg
  
  Call cmdSalir_Click
  
End Sub

Private Sub cmdSalir_Click()
  Unload frmImprimir
End Sub


Private Sub Form_Load()
  
  mintPrnLen = 65       'Longitud de la página A4
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strEvento0 = "Todos los Eventos"
      cmdImprimir.Caption = "&Imprimir"
      cmdSalir.Caption = "&Salir"
      frmImprimir.Caption = "Imprimir Eventos"
      lblInfo0.Caption = "Eventos de:"
            
      lblFechaDesde = "Desde Fecha  >="
      lblFechaHasta = "=<  Hasta Fecha"
      dtpFechaDesde.CustomFormat = gstrFormatoFecha
      dtpFechaHasta.CustomFormat = gstrFormatoFecha
      
    Case EC_ENGLISH
      strEvento0 = "All the Events"
      cmdImprimir.Caption = "&Print"
      cmdSalir.Caption = "&Exit"
      frmImprimir.Caption = "Print Events"
      lblInfo0.Caption = "Events of:"
   
      lblFechaDesde = "From Date  >="
      lblFechaHasta = "=<  To Date"
      dtpFechaDesde.CustomFormat = gstrFormatoFecha
      dtpFechaHasta.CustomFormat = gstrFormatoFecha
  
    Case EC_CATALA
      strEvento0 = "Tots els Events"
      cmdImprimir.Caption = "&Imprimir"
      cmdSalir.Caption = "&Sortir"
      frmImprimir.Caption = "Imprimir Events"
      lblInfo0.Caption = "Events de:"
            
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

Private Sub Imprimir_Titulo()

  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strTitulo = "Listado de Eventos"
      strDesde = "Desde : "
      strHasta = "Hasta : "
    Case EC_ENGLISH
      strTitulo = "Report of Events"
      strDesde = "From : "
      strHasta = "To   : "
    Case EC_CATALA
      strTitulo = "Llistat d'Events"
      strDesde = "Des de : "
      strHasta = "Fins a : "
  End Select
  
  Printer.Font.Name = "Courier New"
  Printer.Font.size = 30
  Printer.Font.Underline = True
  Printer.Font.Bold = True
  For x = 1 To 6
    Printer.Print
  Next
  
  Printer.Print Spc(7); strTitulo
  Printer.Font.Bold = False
  Printer.Font.Underline = False
  Printer.Print
  Printer.Print Spc(7); strDesde; Format(dtpFechaDesde, _
                                  gstrFormatoFecha)
  Printer.Print
  Printer.Print Spc(7); strHasta; Format(dtpFechaHasta, _
                                  gstrFormatoFecha)
  
  Printer.Font.size = 10
  Printer.Font.Underline = False
  Printer.Font.Bold = False
  Printer.NewPage
  
End Sub


Private Sub Imprimir_SubTitulo()
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strSubTitulo = "ID     Nombre Evento       Fecha          Hora On      Hora Off     Tiempo Uso"
    Case EC_ENGLISH
      strSubTitulo = "ID     Event Name          Date           Time On      Time Off     Time of Use"
    Case EC_CATALA
      strSubTitulo = "ID     Nom Event           Data           Hora On      Hora Off     Tiempo Uso"
  End Select

  If mintPrnLinea > mintPrnLen - 8 Then
    Printer.NewPage
    mintPrnLinea = 0
  End If
  
  Printer.Font.Name = "Courier New"
  Printer.Font.size = 10
  Printer.Font.Underline = True
  Printer.Font.Bold = True
  
  For x = 1 To 5
    Printer.Print
  Next x
  Printer.Print Spc(10); strSubTitulo
  Printer.Print
  
  Printer.Font.Bold = False
  Printer.Font.Underline = False
  
  mintPrnLinea = 7

End Sub

Private Sub Imprimir_Evento()
  Dim strID, strNombre, strFecha, strHoraOn, _
      strHoraOff, strTiempo, strEvento As String
  Static dteTiempo As Date
  
  strID = Format(grsEvento.Fields("ProgramaID"), "00")
  
  grsPrograma.MoveFirst
  grsPrograma.Move (Int(strID) - 1)
  strNombre = Left$(grsPrograma.Fields("Nombre") & "               ", 15)
  strFecha = Format(grsEvento.Fields("Fecha"), gstrFormatoFecha)
  strHoraOn = Format(grsEvento.Fields("HoraOn"), "Hh:Nn:Ss")
  strHoraOff = Format(grsEvento.Fields("HoraOff"), "Hh:Nn:Ss")
  strTiempo = Format(grsEvento.Fields("HoraOff") - _
                    grsEvento.Fields("HoraOn"), "Hh:Nn:Ss")
  strEvento = strID & _
              "     " & strNombre & _
              "     " & strFecha & _
              "     " & strHoraOn & _
              "     " & strHoraOff & _
              "     " & strTiempo

 
  If mintPrnLinea > mintPrnLen Then
    Call Imprimir_SubTitulo
  End If
  
  Printer.Print Spc(10); strEvento
  mintPrnLinea = mintPrnLinea + 1
  
End Sub

Private Sub Imprimir_SubTotal(dteTiempo As Date)
  
    If mintPrnLinea > mintPrnLen - 5 Then
      Printer.NewPage
      Call Imprimir_SubTitulo
    End If
    
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.Print
    Printer.Print Spc(58); "Total............... "; Hora_Suma(dteTiempo)
    Printer.Print
    Printer.Print
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    mintPrnLinea = mintPrnLinea + 4
  
End Sub
