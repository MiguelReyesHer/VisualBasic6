VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15630
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   15630
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer12 
      Left            =   15240
      Top             =   4680
   End
   Begin VB.Timer Timer11 
      Left            =   15240
      Top             =   4320
   End
   Begin VB.Timer Timer10 
      Left            =   15240
      Top             =   3960
   End
   Begin VB.Timer Timer9 
      Left            =   15240
      Top             =   3600
   End
   Begin VB.Timer Timer8 
      Left            =   15240
      Top             =   3000
   End
   Begin VB.Timer Timer7 
      Left            =   15240
      Top             =   2640
   End
   Begin VB.Timer Timer6 
      Left            =   15240
      Top             =   2280
   End
   Begin VB.Timer Timer5 
      Left            =   15240
      Top             =   1920
   End
   Begin VB.Timer Timer4 
      Left            =   15240
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Left            =   15240
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Left            =   15240
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Left            =   15240
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   735
      Left            =   14040
      TabIndex        =   2
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reiniciar carrera"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "Iniciar carrera"
      Height          =   495
      Left            =   4560
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Label1"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Image carro3 
      Height          =   735
      Left            =   1560
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image carro2 
      Height          =   735
      Left            =   1560
      Picture         =   "Form1.frx":FC5D
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image carro1 
      Height          =   615
      Left            =   1560
      Picture         =   "Form1.frx":1F8BA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   9495
      Left            =   120
      Picture         =   "Form1.frx":2F517
      Stretch         =   -1  'True
      Top             =   120
      Width           =   15015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, x1, x2, x3, x4, y1, y2, y3, y4, z1, z2, z3, z4 As Double


Private Sub Command1_Click()
    a = Rnd * 600
    b = Rnd * 600
    c = Rnd * 600
    Timer1.Interval = 100
    Timer5.Interval = 100
    Timer9.Interval = 100
End Sub

Private Sub Command2_Click()
    Label1.Caption = ""
    carro1.Left = 1500
    carro2.Left = 1500
    carro3.Left = 1500
    carro1.Top = 1000
    carro2.Top = 2000
    carro3.Top = 3000

    Timer1.Interval = 0
    Timer2.Interval = 0
    Timer3.Interval = 0
    Timer4.Interval = 0
    Timer5.Interval = 0
    Timer6.Interval = 0
    Timer7.Interval = 0
    Timer8.Interval = 0
    Timer9.Interval = 0
    Timer10.Interval = 0
    Timer11.Interval = 0
    Timer12.Interval = 0
    a = 0
    b = 0
    c = 0
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    Label1.Caption = ""
    carro1.Left = 1500
    carro2.Left = 1500
    carro3.Left = 1500
    carro1.Top = 1000
    carro2.Top = 2000
    carro3.Top = 3000
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Timer1_Timer()
    x1 = carro1.Left + a
    carro1.Left = x1
    
    If x1 > 12000 Then
        Timer1.Interval = 0
        Timer2.Interval = 100
    End If
End Sub

Private Sub Timer10_Timer()
    z2 = carro3.Top + c
    carro3.Top = z2
    
    If z2 > 7000 Then
        Timer10.Interval = 0
        Timer11.Interval = 100
    End If
End Sub

Private Sub Timer11_Timer()
    z3 = carro3.Left - c
    carro3.Left = z3
    
    If z3 < 1500 Then
        Timer11.Interval = 0
        Timer12.Interval = 100
    End If
End Sub

Private Sub Timer12_Timer()
    z4 = carro3.Top - c
    carro3.Top = z4
    
    If z4 < 1000 Then
        Label1.Caption = "¡¡¡Ganó el Auto 3!!!"
        Timer1.Interval = 0
        Timer2.Interval = 0
        Timer3.Interval = 0
        Timer4.Interval = 0
        Timer5.Interval = 0
        Timer6.Interval = 0
        Timer7.Interval = 0
        Timer8.Interval = 0
        Timer12.Interval = 0
    End If
End Sub

Private Sub Timer2_Timer()
    x2 = carro1.Top + a
    carro1.Top = x2
    
    If x2 > 7000 Then
        Timer2.Interval = 0
        Timer3.Interval = 100
    End If
End Sub

Private Sub Timer3_Timer()
    x3 = carro1.Left - a
    carro1.Left = x3
    
    If x3 < 1500 Then
        Timer3.Interval = 0
        Timer4.Interval = 100
    End If
End Sub

Private Sub Timer4_Timer()
    x4 = carro1.Top - a
    carro1.Top = x4
    
    If x4 < 1000 Then
        Label1.Caption = "¡¡¡Ganó el Auto 1!!!"
        Timer4.Interval = 0
        Timer5.Interval = 0
        Timer6.Interval = 0
        Timer7.Interval = 0
        Timer8.Interval = 0
        Timer9.Interval = 0
        Timer10.Interval = 0
        Timer11.Interval = 0
        Timer12.Interval = 0
    End If
End Sub

Private Sub Timer5_Timer()
    y1 = carro2.Left + b
    carro2.Left = y1
    
    If y1 > 12000 Then
        Timer5.Interval = 0
        Timer6.Interval = 100
    End If
End Sub

Private Sub Timer6_Timer()
    y2 = carro2.Top + b
    carro2.Top = y2
    
    If y2 > 7000 Then
        Timer6.Interval = 0
        Timer7.Interval = 100
    End If
End Sub

Private Sub Timer7_Timer()
    y3 = carro2.Left - b
    carro2.Left = y3
    
    If y3 < 1500 Then
        Timer7.Interval = 0
        Timer8.Interval = 100
    End If
End Sub

Private Sub Timer8_Timer()
    y4 = carro2.Top - b
    carro2.Top = y4
    
    If y4 < 1000 Then
        Label1.Caption = "¡¡¡Ganó el Auto 2!!!"
        Timer1.Interval = 0
        Timer2.Interval = 0
        Timer3.Interval = 0
        Timer4.Interval = 0
        Timer8.Interval = 0
        Timer9.Interval = 0
        Timer10.Interval = 0
        Timer11.Interval = 0
        Timer12.Interval = 0
    End If
End Sub

Private Sub Timer9_Timer()
    z1 = carro3.Left + c
    carro3.Left = z1
    
    If z1 > 12000 Then
        Timer9.Interval = 0
        Timer10.Interval = 100
    End If
End Sub
