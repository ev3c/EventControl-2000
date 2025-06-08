VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Obtener Número Série HD"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Obtener Número"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Número del HD:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Numero de Validación"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim sUnidad, sNumeroHD As String
    Debug.Print App.Path
    sUnidad = Left$(App.Path, 3)
    sNumeroHD = UCase$(Text1)
    Text1 = sNumeroHD
    
    Label1 = sNumeroHD
    Text2 = Encriptar(sNumeroHD)
End Sub




