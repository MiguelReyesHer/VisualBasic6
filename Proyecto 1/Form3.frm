VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000A&
   Caption         =   "Form3"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   LinkTopic       =   "Form3"
   ScaleHeight     =   7260
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   3360
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Line Line5 
      X1              =   3480
      X2              =   2280
      Y1              =   3360
      Y2              =   1440
   End
   Begin VB.Line Line4 
      X1              =   1080
      X2              =   2280
      Y1              =   3360
      Y2              =   1440
   End
   Begin VB.Line Line3 
      X1              =   1080
      X2              =   3480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   3480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   1440
      Y2              =   3360
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "Altura:"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Base:"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Triángulo"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As Double
    a = ((Val(Text1.Text) * Val(Text2.Text)) / 2)
    Label4.Caption = "El resultado es: " & a
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Label4.Caption = ""
End Sub

Private Sub Command3_Click()
    Form1.Show
    Form3.Hide
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Label4.Caption = ""
End Sub

