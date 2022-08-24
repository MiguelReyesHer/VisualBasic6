VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tron"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Detener auto 3"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Detener auto 2"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Detener auto 1"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Left            =   1200
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Left            =   720
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   4920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ejecutar"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
   Begin VB.Image auto_2 
      Height          =   1080
      Left            =   4320
      Picture         =   "Form1.frx":17C59
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1560
   End
   Begin VB.Image auto_1 
      Height          =   855
      Left            =   240
      Picture         =   "Form1.frx":215F8
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Image auto_3 
      Height          =   855
      Left            =   2280
      Picture         =   "Form1.frx":31A99
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Var para todo el programa
Dim A, B, C As Single

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Timer1.Interval = 0
        Check2.Value = 0
        Check3.Value = 0
    Else
        Timer1.Interval = 50
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Timer2.Interval = 0
        Check1.Value = 0
        Check3.Value = 0
    Else
        Timer2.Interval = 80
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Timer3.Interval = 0
        Check1.Value = 0
        Check2.Value = 0
    Else
        Timer3.Interval = 20
    End If
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    Timer1.Interval = 100
    Timer2.Interval = 100
    Timer3.Interval = 100
End Sub

Private Sub Timer1_Timer()
    auto_1.Left = auto_1.Left + 50
End Sub

Private Sub Timer2_Timer()
    auto_2.Left = auto_2.Left - 50
End Sub

Private Sub Timer3_Timer()
    auto_3.Top = auto_3.Top + 50
End Sub
