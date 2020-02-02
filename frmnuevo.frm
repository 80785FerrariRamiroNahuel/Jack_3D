VERSION 5.00
Begin VB.Form frmnuevo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jack 3D"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmnuevo.frx":0000
   ScaleHeight     =   2835
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtgrupo 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtnumero 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmnuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txtnombre.Text = "" Then MsgBox ("Ingrese un Nombre"), vbInformation, "Alerta": txtnombre.SetFocus: Exit Sub
If txtnumero.Text = "" Then MsgBox ("Ingrese un Numero"), vbInformation, "Alerta": txtnumero.SetFocus: Exit Sub
If txtgrupo.Text = "" Then MsgBox ("Ingrese un Grupo"), vbInformation, "Alerta": txtgrupo.SetFocus: Exit Sub

If Not IsNumeric(txtnumero.Text) Then MsgBox ("Ingrese numeros"), vbInformation, "Alerta": txtnumero = "": txtnumero.SetFocus: Exit Sub
With RsCliente
    .Requery
    .Find "NomCli='" & Trim(txtnombre.Text) & "'"
    If .EOF Then
        .AddNew
        !NomCli = txtnombre.Text
        !NumCli = txtnumero
        !Grupo = txtgrupo.Text
        .Update
        MsgBox ("El cliente fue agregado Satisfactoriamente"), vbOKOnly, "Alerta"
        Unload Me
    Else
        MsgBox ("El nombre ya existe")
        txtnombre.Text = ""
        txtnombre.SetFocus
        
    End If
    
    
    
    
  
    
    
    
    
    
    
    
    
End With
End Sub

Private Sub Command2_Click()
txtnombre.Text = ""
txtnumero.Text = ""
txtgrupo.Text = ""
Unload Me
End Sub

