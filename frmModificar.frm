VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmModificar 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   Picture         =   "frmModificar.frx":0000
   ScaleHeight     =   4965
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Grupo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox NumCli 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox NomCli 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero Cliente"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Cliente"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmModificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VariableId As Integer

Private Sub Command1_Click()
If NomCli.Text = "" Then MsgBox ("Ingrese un Nombre"), vbInformation, "Alerta": NomCli.SetFocus: Exit Sub
If NumCli.Text = "" Then MsgBox ("Ingrese un Numero"), vbInformation, "Alerta": NumCli.SetFocus: Exit Sub
If Grupo.Text = "" Then MsgBox ("Ingrese un Grupo"), vbInformation, "Alerta": Grupo.SetFocus: Exit Sub

If Not IsNumeric(NumCli.Text) Then MsgBox ("Ingrese numeros"), vbInformation, "Alerta": NumCli = "": NumCli.SetFocus: Exit Sub

With RsCliente
    .Requery
    .Find "NomCli='" & Trim(DataCombo1.Text) & "'"
        !NomCli = NomCli.Text
        !NumCli = NumCli.Text
        !Grupo = Grupo.Text
    .UpdateBatch
    MsgBox ("La modificacion se realizo correctamente")
    Unload Me
End With
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)
NomCli.Enabled = True
NumCli.Enabled = True
Grupo.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Form_Load()
AbrirTablaCliente
Set DataCombo1.RowSource = RsCliente
DataCombo1.ListField = "NomCli"
DataCombo1.DataField = "NomCli"
End Sub


