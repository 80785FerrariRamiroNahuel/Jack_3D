VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmElim 
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   Picture         =   "frmElim.frx":0000
   ScaleHeight     =   2595
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
End
Attribute VB_Name = "frmElim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Seguro que quiere eliminar a '" & DataCombo1.Text & "'", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    With RsCliente
        .Requery
        .Find "NomCli='" & Trim(DataCombo1.Text) & "'"
        .Delete
        .Requery
        
    End With
End If
MsgBox ("El cliente fue eliminado Correctamente"), vbInformation, "Alerta"
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
AbrirTablaCliente
Set DataCombo1.RowSource = RsCliente
DataCombo1.ListField = "NomCli"
DataCombo1.DataField = "NomCli"
End Sub
Private Sub DataCombo1_Change()
Command1.Enabled = True
With RsCliente
    .Requery
    .Find "NomCli='" & Trim(DataCombo1.Text) & "'"
    If .EOF Then MsgBox ("Error"), vbInformation, "Alerta": Exit Sub
    VariableId = !NroCli
End With
End Sub
