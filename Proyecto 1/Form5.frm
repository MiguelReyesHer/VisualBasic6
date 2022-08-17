VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H8000000A&
   Caption         =   "Form5"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   LinkTopic       =   "Form5"
   ScaleHeight     =   5730
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "Apotema:"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Lado:"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Pentágono"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Line Line7 
      X1              =   1320
      X2              =   1800
      Y1              =   1920
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   2160
      X2              =   1320
      Y1              =   1920
      Y2              =   1080
   End
   Begin VB.Line Line5 
      X1              =   480
      X2              =   1320
      Y1              =   1920
      Y2              =   1080
   End
   Begin VB.Line Line4 
      X1              =   1800
      X2              =   2160
      Y1              =   2760
      Y2              =   1920
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   840
      Y1              =   1920
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   1080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   1800
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As Double
    a = (((5 * Val(Text1.Text)) * Val(Text2.Text)) / 2)
    Label4.Caption = "El resultado es: " & a
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Label4.Caption = ""
End Sub

Private Sub Command3_Click()
    Form1.Show
    Form2.Hide
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Label4.Caption = ""
End Sub

