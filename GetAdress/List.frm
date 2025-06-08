VERSION 5.00
Begin VB.Form frmWindowList 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   1620
   ClientTop       =   1545
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7335
   Begin VB.ListBox List1 
      Height          =   3765
      ItemData        =   "List.frx":0000
      Left            =   120
      List            =   "List.frx":0002
      MouseIcon       =   "List.frx":0004
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   720
      Width           =   7095
   End
   Begin VB.CommandButton cmdFindAddress 
      Caption         =   "Find Address"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmWindowList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lParam As Long) As Long
' Start the enumeration.
Private Sub cmdFindAddress_Click()
        EnumWindows AddressOf EnumProc, 0
End Sub

