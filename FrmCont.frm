VERSION 5.00
Begin VB.Form FrmCont 
   Caption         =   "Jack 3D"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   Picture         =   "FrmCont.frx":0000
   ScaleHeight     =   4290
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox DesMaq 
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox PreLuz 
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox PreFil 
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdmodificar 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modificar Constantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5325
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Desgaste Maquina"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Luz"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Lhoras 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Filamento"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "FrmCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub cmdmodificar_Click()
With RsConstantes
    .Requery
    .AddNew
    !FecCam = Date
    !PreFil = CDbl(PreFil.Text)
    !PreLuz = CDbl(PreLuz.Text)
    !DesMaq = CDbl(DesMaq.Text)
    .Update
    MsgBox ("El cambio se Realizo Correctamente")
    Unload Me
End With
End Sub

Private Sub Form_Load()
AbrirTablaConstantes

End Sub


