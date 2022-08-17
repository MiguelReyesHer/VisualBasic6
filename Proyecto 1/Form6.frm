VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H8000000A&
   Caption         =   "Form6"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4155
   LinkTopic       =   "Form6"
   ScaleHeight     =   3945
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "Generatriz:"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Radio de la base:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Cono"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a, s As Double
    s = (Val(Text1.Text) + Val(Text2.Text))
    r = (Val(Text1.Text) * 3.1416)
    a = (r * s)
    Label4.Caption = "El resultado es: " & a

End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Label4.Caption = ""
End Sub

Private Sub Command3_Click()
    Form1.Show
    Form6.Hide
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Label4.Caption = ""
End Sub
