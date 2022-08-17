VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000015&
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cono"
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pentágono"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Trapecio"
      Height          =   735
      Left            =   4560
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Triángulo"
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Círculo"
      Height          =   735
      Left            =   360
      MaskColor       =   &H00404040&
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Calculadora de Áreas"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Show
    Form1.Hide
End Sub

Private Sub Command2_Click()
    Form3.Show
    Form1.Hide
End Sub

Private Sub Command3_Click()
    Form4.Show
    Form1.Hide
End Sub

Private Sub Command4_Click()
    Form5.Show
    Form1.Hide
End Sub

Private Sub Command5_Click()
    Form6.Show
    Form1.Hide
End Sub

Private Sub Command6_Click()
    End
End Sub
