VERSION 5.00
Begin VB.Form frmColores 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Posición"
      Height          =   1695
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   2775
      Begin VB.OptionButton Option6 
         Caption         =   "Abajo"
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Arriba"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color"
      Height          =   2415
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2895
      Begin VB.OptionButton Option4 
         Caption         =   "Amarillo"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Verde"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Rojo"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Azul"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox txtCaja 
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
End
Attribute VB_Name = "frmColores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    txtCaja.Top = 0
End Sub

Private Sub Option1_Click()
    txtCaja.BackColor = vbBlue
End Sub

Private Sub Option2_Click()
    txtCaja.BackColor = vbRed
End Sub

Private Sub Option3_Click()
    txtCaja.BackColor = vbGreen
End Sub

Private Sub Option4_Click()
    txtCaja.BackColor = vbYellow
End Sub

Private Sub Option5_Click()
    'Arriba
    txtCaja.Top = 0
End Sub

Private Sub Option6_Click()
    'Abajo
    txtCaja.Top = frmColores.ScaleHeight - txtCaja.Height
    'Se puede poner txtCaja.Top=1 pero podría superar la escala del form
End Sub

Private Sub TextCaja_Change()

End Sub
