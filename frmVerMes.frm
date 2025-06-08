VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerMes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendario"
   ClientHeight    =   3225
   ClientLeft      =   405
   ClientTop       =   645
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24510466
      CurrentDate     =   36395
   End
End
Attribute VB_Name = "frmVerMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
    Unload frmVerMes
End Sub

Private Sub Form_Load()
    Select Case frmEventControl.cboIdioma.ListIndex
      Case EC_ESPAÑOL
        frmVerMes.Caption = "Calendario"
        cmdCerrar.Caption = "&Cerrar"
      Case EC_ENGLISH
        frmVerMes.Caption = "Calendar"
        cmdCerrar.Caption = "&Close"
    End Select
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    frmEventControl.txtFecha.Text = Format(DateClicked, "DD/MM/YYYY")
End Sub
