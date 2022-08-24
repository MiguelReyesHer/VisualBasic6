VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresar"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim usuario As String
    Dim contraseña As String 'las mejores contraseñas llevan números y letras
    
    usuario = "Miguel"
    contraseña = "12345"
    usuario2 = Trim(Text1.Text)
    contraseña2 = Trim(Text2.Text)
    
    If (usuario <> usuario2) Then
        MsgBox "Usuario incorrecto", vbExclamation + vbOKOnly, "Datos incorrectos"
        Else
            If (contraseña <> contraseña2) Then
                MsgBox "Contraseña incorrecta", vbExclamation + vbOKOnly, "Datos incorrectos"
        Else
            MsgBox "Bienvenido", vbExclamation + vbOKOnly, "Cuenta verificada"
            Me.Hide
            Form2.Show
            End If
    End If
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
End Sub
