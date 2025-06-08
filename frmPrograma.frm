VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrograma 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdMover 
      Height          =   375
      Index           =   3
      Left            =   1320
      Picture         =   "frmPrograma.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdMover 
      Height          =   375
      Index           =   2
      Left            =   960
      Picture         =   "frmPrograma.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdMover 
      Height          =   375
      Index           =   1
      Left            =   600
      Picture         =   "frmPrograma.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdMover 
      Height          =   375
      Index           =   0
      Left            =   240
      Picture         =   "frmPrograma.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
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
      Left            =   6120
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame fraPrograma 
      Caption         =   "Programas"
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      Begin VB.CheckBox chkContraseña 
         Caption         =   "Protejer con contraseña"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdBorrarEventos 
         Caption         =   "Borrar &Eventos"
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   17
         Top             =   720
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog dlgNombre 
         Left            =   6600
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar"
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Path:"
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
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblNombre 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre:"
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
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
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
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAñadir 
      Caption         =   "&Añadir"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBorrarEventos_Click()
  strBookMark = grsPrograma.Bookmark
  Load frmEvento
  frmEvento.Show vbModal
  grsPrograma.Bookmark = strBookMark
End Sub

Private Sub cmdAceptar_Click()
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsg0 = "Debe especificar el Nombre y el Path del Programa a Añadir"
      strMsg1 = "El Path del Programa a Añadir no existe"
      strMsg2 = "Añadir Programa"
    Case EC_ENGLISH
      strMsg0 = "You should specify the Name and the Path of the Program to Add"
      strMsg1 = "The Path of the Program to Add don't exist"
      strMsg2 = "Add Program"
    Case EC_CATALA
      strMsg0 = "Ha d'especificar el Nom i el Path del Programa a Afegir"
      strMsg1 = "El Path del Programa a Afegir no existeix"
      strMsg2 = "Afegir Programa"
  End Select
  If txtNombre = "" Or txtPath = "" Then
    MsgBox strMsg0, vbInformation, strMsg2
  Else
    If Dir(txtPath) = "" Then
      MsgBox strMsg1, vbInformation, strMsg2
    Else
      Call Campos_Grabar
      grsPrograma.Update
      Call Command_Mostrar
      Call cmdMover_Click(3)
      
    End If
  End If
End Sub

Private Sub cmdAñadir_Click()
  Call Command_Ocultar
  grsPrograma.MoveLast
  lblID = grsPrograma.Fields("ID") + 1
  grsPrograma.AddNew
  txtNombre = ""
  txtPath = ""
  txtNombre.SetFocus
End Sub

Private Sub cmdBorrar_Click()
  Dim intBorrar As Integer
  Dim intID, intBorradoID As Integer
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      strMsgPrg = "Borrar Programa"
      strMsgBorrar = "Seguro que desea Borrar el Programa " & grsPrograma.Fields("Nombre")
      strMsgNoBorrar = "No puede Borrar un Programa que está siendo Cronometrado"
      strMsgUltimo = "Debe haber cómo mínimo un Programa definido"
    Case EC_ENGLISH
      strMsgPrg = "Delete Program"
      strMsgBorrar = "Sure you wants to Delete the Program" & grsPrograma.Fields("Nombre")
      strMsgNoBorrar = "It cannot Erase a Program that it is being Cronometred"
      strMsgUltimo = "It should have how minimum one defined Program"
    Case EC_CATALA
      strMsgPrg = "Borrar Programa"
      strMsgBorrar = "Segur que desitja Borrar el Programa " & grsPrograma.Fields("Nombre")
      strMsgNoBorrar = "No pot Borrar un Programa que està essent Cronometrat"
      strMsgUltimo = "Hi ha d'haber com a mínim un Programa definit"
  End Select
  
  intBorradoID = grsPrograma.Fields("ID")
  
  If gtPrg(intBorradoID - 4).On <> CDate("0") Then
    MsgBox strMsgNoBorrar, vbInformation, strMsgPrg
  Else
    
    If grsPrograma.RecordCount <= 5 Then
      MsgBox strMsgUltimo, vbInformation, strMsgPrg
    Else
      
      If frmEventControl.Contraseña_Entrar Then
        intBorrar = MsgBox(strMsgBorrar, vbYesNo + vbCritical, strMsgPrg)
        If intBorrar = vbYes Then
          grsPrograma.Delete
          grsPrograma.MoveFirst
          Do While Not grsPrograma.EOF
            intID = grsPrograma.Fields("ID")
            If intID > intBorradoID Then
              grsPrograma.Edit
              grsPrograma.Fields("ID") = intID - 1
              grsPrograma.Update
            End If
            grsPrograma.MoveNext
          Loop
          
          If grsEvento.RecordCount > 0 Then
            grsEvento.MoveFirst
          End If
          Do While Not grsEvento.EOF
            intID = grsEvento.Fields("ProgramaID")
            If intID = intBorradoID Then
              grsEvento.Delete
            End If
            If intID > intBorradoID Then
              grsEvento.Edit
              grsEvento.Fields("ProgramaID") = intID - 1
              grsEvento.Update
            End If
            'grsEvento.Update
            grsEvento.MoveNext
          Loop
          
          Call cmdMover_Click(0)
          
        End If
      End If
    End If
  End If
End Sub

Private Sub cmdCancelar_Click()
  grsPrograma.CancelUpdate
  Call Command_Mostrar
  Call cmdMover_Click(3)
  'Call Campos_Ver
End Sub

Private Sub cmdExaminar_Click()
       
  On Error GoTo errHandler
  
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      dlgNombre.DialogTitle = "Examinar Programas"
      dlgNombre.Filter = "Archivos de Programa (*.exe) |*.exe"
    Case EC_ENGLISH
      dlgNombre.DialogTitle = "Examine Programs"
      dlgNombre.Filter = "Program Files (*.exe) |*.exe"
    Case EC_CATALA
      dlgNombre.DialogTitle = "Examinar Programes"
      dlgNombre.Filter = "Arxius de Programa (*.exe) |*.exe"
  End Select
  
  dlgNombre.CancelError = True
  dlgNombre.InitDir = "c:\"
  dlgNombre.ShowOpen
  
  txtPath = dlgNombre.FileName
      
  Exit Sub
    
errHandler:
txtFichero = Error
End Sub

Private Sub cmdModificar_Click()
  If frmEventControl.Contraseña_Entrar Then
    Call Command_Ocultar
    grsPrograma.Edit
    txtNombre.SetFocus
  End If
End Sub

Private Sub cmdSalir_Click()
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
  gaPrograma(iPrg, 0) = ""
  
  iPrg = 0
  iIndex = frmEventControl.cboPrg.ListIndex
  frmEventControl.cboPrg.Clear
  Do While gaPrograma(iPrg, 0) <> ""
    frmEventControl.cboPrg.AddItem gaPrograma(iPrg, 0)
    iPrg = iPrg + 1
  Loop
  If iIndex >= frmEventControl.cboPrg.ListCount Then
    frmEventControl.cboPrg.ListIndex = 0
  Else
    frmEventControl.cboPrg.ListIndex = iIndex
  End If
  Unload frmPrograma
End Sub

Private Sub Form_Load()
  Select Case frmEventControl.cboIdioma.ListIndex
    Case EC_ESPAÑOL
      cmdAñadir.Caption = "&Añadir"
      cmdBorrar.Caption = "&Borrar"
      cmdExaminar.Caption = "&Examinar HD"
      cmdModificar.Caption = "&Modificar"
      cmdSalir.Caption = "&Salir"
      cmdAceptar.Caption = "&Aceptar"
      cmdCancelar.Caption = "&Cancelar"
      cmdBorrarEventos.Caption = "Borrar &Eventos"
      fraPrograma.Caption = "Programas"
      frmPrograma.Caption = "Gestionar Programas y Borrar Eventos"
      lblNombre.Caption = "Nombre:"
      chkContraseña.Caption = "Protejer con Contraseña"
      chkContraseña.ToolTipText = "Solo se podrá ejecutar el programa si se conoce la contraseña"
    Case EC_ENGLISH
      cmdAñadir.Caption = "&Add"
      cmdBorrar.Caption = "&Delete"
      cmdExaminar.Caption = "E&xamine HD"
      cmdModificar.Caption = "&Modify"
      cmdSalir.Caption = "&Exit"
      cmdAceptar.Caption = "&Accept"
      cmdCancelar.Caption = "&Cancel"
      cmdBorrarEventos.Caption = "Delete &Events"
      fraPrograma.Caption = "Programs"
      frmPrograma.Caption = "Negotiate Programs and Delete Events"
      lblNombre.Caption = "Name:"
      chkContraseña.Caption = "Protect by Password"
      chkContraseña.ToolTipText = "Only can execute the program if know the Password"
    Case EC_CATALA
      cmdAñadir.Caption = "&Afegir"
      cmdBorrar.Caption = "&Borrar"
      cmdExaminar.Caption = "&Examinar HD"
      cmdModificar.Caption = "&Modificar"
      cmdSalir.Caption = "&Sortir"
      cmdAceptar.Caption = "&Acceptar"
      cmdCancelar.Caption = "&Cancelar"
      cmdBorrarEventos.Caption = "Borrar &Events"
      fraPrograma.Caption = "Programes"
      frmPrograma.Caption = "Gestionar Programes y Borrar Events"
      lblNombre.Caption = "Nom:"
      chkContraseña.Caption = "Protegir amb Contrasenya"
      chkContraseña.ToolTipText = "Només es podrà executar el programa si es coneix la contrasenya"
  End Select
  
  Call Command_Mostrar
  Call cmdMover_Click(0)

End Sub
Private Sub cmdMover_Click(Index As Integer)
  Select Case Index
    Case 0
      grsPrograma.MoveFirst
    Case 1
      grsPrograma.MovePrevious
    Case 2
      grsPrograma.MoveNext
    Case 3
      grsPrograma.MoveLast
  End Select

  If grsPrograma.BOF Then grsPrograma.MoveFirst
  If grsPrograma.EOF Then grsPrograma.MoveLast
  
  Call Campos_Ver
  
  If lblID < 5 Then
    cmdModificar.Enabled = False
    cmdBorrar.Enabled = False
  Else
    cmdModificar.Enabled = True
    cmdBorrar.Enabled = True
  End If
  
End Sub

Private Sub Campos_Ver()
  lblID = ""
  txtNombre = ""
  txtPath = ""
  chkContraseña = 0
  
  lblID = grsPrograma.Fields("ID")
  txtNombre = grsPrograma.Fields("Nombre")
  If grsPrograma.Fields("Path") <> Chr(0) Then
    txtPath = grsPrograma.Fields("Path")
  End If
  If grsPrograma.Fields("Password") <> Chr(0) Then
     chkContraseña = grsPrograma.Fields("Password")
  End If

End Sub

Private Sub Campos_Grabar()
  grsPrograma.Fields("ID") = lblID
  grsPrograma.Fields("Nombre") = txtNombre
  grsPrograma.Fields("Path") = txtPath
  grsPrograma.Fields("Password") = chkContraseña
End Sub

Private Sub Command_Ocultar()
  cmdAñadir.Visible = False
  cmdModificar.Visible = False
  cmdBorrar.Visible = False
  cmdSalir.Visible = False
  
  cmdMover.Item(0).Enabled = False
  cmdMover.Item(1).Enabled = False
  cmdMover.Item(2).Enabled = False
  cmdMover.Item(3).Enabled = False
  
  cmdAceptar.Visible = True
  cmdCancelar.Visible = True
  cmdExaminar.Enabled = True
  
  chkContraseña.Enabled = True
  
  cmdBorrarEventos.Enabled = False
  
  txtNombre.Enabled = True
  txtPath.Enabled = True
End Sub

Private Sub Command_Mostrar()
  cmdAñadir.Visible = True
  cmdModificar.Visible = True
  cmdBorrar.Visible = True
  cmdSalir.Visible = True
  
  cmdAceptar.Visible = False
  cmdCancelar.Visible = False
  cmdExaminar.Enabled = False
  
  cmdMover.Item(0).Enabled = True
  cmdMover.Item(1).Enabled = True
  cmdMover.Item(2).Enabled = True
  cmdMover.Item(3).Enabled = True
  
  chkContraseña.Enabled = False
  
  cmdBorrarEventos.Enabled = True
  
  txtNombre.Enabled = False
  txtPath.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If cmdSalir.Visible = False Then
    Cancel = True
  End If
End Sub
