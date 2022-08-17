VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000A&
   Caption         =   "Form4"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   LinkTopic       =   "Form4"
   ScaleHeight     =   6375
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "Altura:"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "Base mayor:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Base menor:"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Trapecio"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Line Line5 
      X1              =   3720
      X2              =   4560
      Y1              =   1680
      Y2              =   3240
   End
   Begin VB.Line Line4 
      X1              =   720
      X2              =   1680
      Y1              =   3240
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   1680
      X2              =   1680
      Y1              =   1680
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   4560
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   3720
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As Double
    a = (((Val(Text1.Text) + Val(Text2.Text)) * Val(Text3.Text)) / 2)
    Label5.Caption = "El resultado es: " & a
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Label5.Caption = ""
End Sub

Private Sub Command3_Click()
    Form1.Show
    Form4.Hide
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Label5.Caption = ""
End Sub
