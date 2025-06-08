VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form1.Hide
End Sub

Private Sub Form_Load()
    Dim lRet As Long
    lRet = RegisterHotKey(Me.hWnd, &HB000&, MOD_ALT Or MOD_CONTROL, vbKeyE)
    Call Subclasifica_Ventana(Form1.hWnd)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnregisterHotKey(Me.hWnd, &HB000&)
    Call Ventana_Normal(Form1.hWnd)
End Sub
 
