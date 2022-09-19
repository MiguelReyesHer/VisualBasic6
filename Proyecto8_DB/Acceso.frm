VERSION 5.00
Begin VB.Form Acceso 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   ":Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1 = "" And Text2 = "" Then
        MsgBox "Es importante llenar los campos indicados", vbInformation, "Datos incorrectos"
        ElseIf Text1 = "" Then
            MsgBox "Es nesesario ingresar un usuario", vbInformation, "Usuario necesario"
        ElseIf Text2 = "" Then
            MsgBox "Debe ingresar una contraseña", vbInformation, "Contraseña necesaria"
        
        ElseIf Text1 = "Miguel" And Text2 = "12345" Then
            Form1.Hide
            Form2.Show
            Else
                MsgBox "Los datos ingreados son erroneos", vbInformation, "Uusario y/o contraseña incorrectos"
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
End Sub

