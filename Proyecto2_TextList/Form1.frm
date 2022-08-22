VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000005&
         Caption         =   "Salir"
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   3480
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre;"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'El Text list está vacío y no podemos agregar nada
    If Text1 = "" Then
        MsgBox "Debe ingresar un nombre para poder agregar un elemento", vbQuestion + vbOKOnly, "Datos incompletos"
    End If
    
    'Agregar contenido del Textbox al textlist
    List1.AddItem Text1
End Sub

Private Sub Command2_Click()
    'Si la lista está vacía no podemos eliminar
    If List1.ListIndex <> -1 Then
    
    'Eliminamos el elemento que se encuentra seleccionado
    List1.RemoveItem List1.ListIndex
    
    End If
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    List1.Text = ""
    Text1.Text = ""
End Sub

