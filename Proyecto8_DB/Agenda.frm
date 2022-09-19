VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Agenda 
   Caption         =   "Form2"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   3120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\migue\Desktop\Proyecto8_DB\Agenda (mdb).mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\migue\Desktop\Proyecto8_DB\Agenda (mdb).mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Agenda"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "e-mail"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "Telefono de casa"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "Celular"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Correo"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Tel. Casa"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Celular"
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contacto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "Agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    Form1.Show
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub

