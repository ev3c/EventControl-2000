VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInternetHistorial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Internet Historial"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBorrarHistorial 
      Caption         =   "Borrar Historial"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   1080
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
      Left            =   7560
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox lstHistorial 
      Height          =   3960
      Left            =   120
      MouseIcon       =   "frmInternetHistorial.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
   Begin MSComCtl2.DTPicker dtpFechaHistorial 
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24379395
      CurrentDate     =   36545
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha Historial:"
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmInternetHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
  Unload frmInternetHistorial
End Sub

Private Sub cmdBorrarHistorial_Click()
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsgPrg = "Borrar Historial"
      strMsgBorrar = "Seguro que desea Borrar el Historial de Internet"
    Case EC_ENGLISH
      strMsgPrg = "Delete History"
      strMsgBorrar = "Sure you wants to Delete Internet History"
    Case EC_CATALA
      strMsgPrg = "Borrar Historial"
      strMsgBorrar = "Segur que desitja Borrar l'Historial d'Internet"
  End Select
  
  If frmEventControl.Contraseña_Entrar Then
    intBorrar = MsgBox(strMsgBorrar, vbYesNo + vbCritical, strMsgPrg)
    If intBorrar = vbYes Then
      If grsHistorial.RecordCount > 0 Then
        grsHistorial.MoveFirst
      End If
      Do While Not grsHistorial.EOF
        grsHistorial.Delete
        grsHistorial.MoveNext
      Loop
      
      Call Historial_Compactar
      
      Call frmEventControl.gaHistorial_Leer
      Call cmdAceptar_Click
      
    End If
  End If

End Sub

Private Sub dtpFechaHistorial_Change()

  Call InternetHistorial_Actualizar(dtpFechaHistorial)

End Sub


Private Sub Form_Load()
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      frmInternetHistorial.Caption = "Internet Historial"
      cmdAceptar.Caption = "&Aceptar"
      cmdBorrarHistorial.Caption = "&Borrar Historial"
      lblFecha.Caption = "&Fecha Historial:"
      lstHistorial.ToolTipText = "Doble click para ir a la página web"
    
    Case EC_ENGLISH
      frmInternetHistorial.Caption = "Internet History"
      cmdAceptar.Caption = "&Accept"
      cmdBorrarHistorial.Caption = "&Delete History"
      lblFecha.Caption = "&History Date:"
      lstHistorial.ToolTipText = "Double click to go to web page"
  
    Case EC_CATALA
      frmInternetHistorial.Caption = "Internet Historial"
      cmdAceptar.Caption = "&Acceptar"
      cmdBorrarHistorial.Caption = "&Borrar Historial"
      lblFecha.Caption = "&Data Historial:"
      lstHistorial.ToolTipText = "Doble click per anar a la pàgina web"
  
  End Select
      
  dtpFechaHistorial.CustomFormat = gstrFormatoFecha
  dtpFechaHistorial.Value = MyDate
  
  Call InternetHistorial_Actualizar(MyDate)
  
End Sub

Private Sub InternetHistorial_Actualizar(fecha)
  
  Set grsHistorial = gdbHistorial.OpenRecordset( _
      "SELECT * " & _
      "FROM tblHistorial " & _
      "ORDER BY url_fecha, url_hora ")

  lstHistorial.Clear

  If grsHistorial.RecordCount > 0 Then
    grsHistorial.MoveFirst
  End If
  Do While Not grsHistorial.EOF
    If grsHistorial.Fields("url_fecha") = _
       dtpFechaHistorial.Value Then
      lstHistorial.AddItem Format( _
                   grsHistorial.Fields("url_fecha"), _
                   gstrFormatoFecha) & _
                   " - " & grsHistorial.Fields("url_hora") & _
                   " - " & grsHistorial.Fields("url_adress")
    End If
    grsHistorial.MoveNext
  Loop

End Sub

Private Sub lstHistorial_DblClick()
  Dim sURL As String
   sURL = lstHistorial
   sURL = Trim(Right(sURL, Len(sURL) - 24))
   
   Call ShellExecute(Me.hWnd, "open", sURL, "", "", 5)
   
   Call cmdAceptar_Click

End Sub

Private Sub Historial_Compactar()
  
  'Cierra Historial.mdb
  grsHistorial.Close
  Set grsHistorial = Nothing
  gdbHistorial.Close
  Set gdbHistorial = Nothing
 
  If gstrContraseña = "" Then
    DBEngine.CompactDatabase App.Path & "\Historial.mdb", _
      App.Path & "\Hist_Co.mdb"
  Else
    DBEngine.CompactDatabase App.Path & "\Historial.mdb", _
      App.Path & "\Hist_Co.mdb", , , ";PWD=" & gstrContraseña
  End If
  
  'Borra Historial
  If Dir(App.Path & "\Historial.mdb") <> "" Then
     Kill App.Path & "\Historial.mdb"
  End If
  
  'Renombra Hist_Co.mdb
  Name App.Path & "\Hist_Co.mdb" As _
       App.Path & "\Historial.mdb"
  
  
  'Reabre Historial
  Set gdbHistorial = DBEngine.Workspaces(0).OpenDatabase( _
       App.Path & "\Historial.mdb", True, False, _
       ";PWD=" & gstrContraseña)
  
  Set grsHistorial = gdbHistorial.OpenRecordset("SELECT * " & _
                  "FROM tblHistorial " & _
                  "ORDER BY url_fecha, url_hora")

End Sub
