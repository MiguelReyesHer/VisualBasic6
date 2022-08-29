VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   21210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   615
      Left            =   19320
      TabIndex        =   3
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reinicio"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inicio"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Left            =   1680
      Top             =   7320
   End
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   7320
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   7320
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   7320
      Width           =   3615
   End
   Begin VB.Image carro3 
      Height          =   975
      Left            =   120
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Image carro2 
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":2C932
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image carro1 
      Height          =   840
      Left            =   120
      Picture         =   "Form1.frx":59264
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   0
      Picture         =   "Form1.frx":85B96
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, x, y, z As Double
Private Sub Command1_Click()
    a = Rnd * 600
    b = Rnd * 600
    c = Rnd * 600
    Timer1.Interval = 100
    Timer2.Interval = 100
    Timer3.Interval = 100
End Sub

Private Sub Command2_Click()
    Label1.Caption = ""
    carro1.Left = 0
    carro2.Left = 0
    carro3.Left = 0
    Timer1.Interval = 0
    Timer2.Interval = 0
    Timer3.Interval = 0
    a = 0
    b = 0
    c = 0
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    Label1.Caption = ""
End Sub

Private Sub Image3_Click()

End Sub

Private Sub Timer1_Timer()
    x = carro1.Left + a
    carro1.Left = x
    
    If x > 14000 Then ''14000 es el tiempo que el auto toma en llegar hasta el punto de meta, este valor cambia de acuerdo a la pista
        Label1.Caption = "Ganó auto 1"
        Timer1.Interval = 0
        Timer2.Interval = 0
        Timer3.Interval = 0
    End If
End Sub

Private Sub Timer2_Timer()
    y = carro2.Left + b
    carro2.Left = y
    
    If y > 14000 Then
        Label1.Caption = "Ganó auto 2"
        Timer1.Interval = 0
        Timer2.Interval = 0
        Timer3.Interval = 0
    End If
End Sub

Private Sub Timer3_Timer()
    z = carro3.Left + c
    carro3.Left = z
    
    If z > 14000 Then
        Label1.Caption = "Ganó auto 3"
        Timer1.Interval = 0
        Timer2.Interval = 0
        Timer3.Interval = 0
    End If
End Sub
