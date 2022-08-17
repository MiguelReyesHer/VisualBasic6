VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000A&
   Caption         =   "Form2"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4245
   LinkTopic       =   "Form2"
   ScaleHeight     =   4140
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Radio:"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Circulo"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As Double
    a = (Val(Text2.Text) * Val(Text2.Text) * 3.1416)
    Label3.Caption = "El resultado es: " & r
End Sub

Private Sub Command2_Click()
    Text2.Text = ""
    Label3.Caption = ""
End Sub

Private Sub Command3_Click()
    Form1.Show
    Form2.Hide
End Sub

Private Sub Form_Load()
    Text2.Text = ""
    Label3.Caption = ""
End Sub

Private Sub Text1_Change()

End Sub

