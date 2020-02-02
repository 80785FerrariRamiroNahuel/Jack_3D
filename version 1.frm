VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jack 3D"
   ClientHeight    =   6615
   ClientLeft      =   2055
   ClientTop       =   2340
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "version 1.frx":0000
   ScaleHeight     =   6615
   ScaleWidth      =   5505
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAgr 
      Caption         =   "Agregar al Cliente"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "Calcular"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox PreTot 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Ganancia 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Peso 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox PreUni 
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox PreGan 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Horas 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Precion con Ganancia:"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Total:"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Lpeso 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Lganancia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ganancia:"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Unitario:"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Lhoras 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Horas:"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
   Begin VB.Menu Mmenu 
      Caption         =   "Menu"
      Begin VB.Menu Magrcli 
         Caption         =   "Agregar Cliente"
      End
      Begin VB.Menu Mmodcli 
         Caption         =   "Modificar Cliente"
      End
      Begin VB.Menu Melicli 
         Caption         =   "Eliminar Cliente"
      End
      Begin VB.Menu mcon 
         Caption         =   "Modificar Constantes"
      End
      Begin VB.Menu Mfab 
         Caption         =   "Ver Fabricaciones"
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VariableId As Integer

Private Sub cmdCal_Click()
p1 = DataGrid1.Columns(1)
p2 = DataGrid1.Columns(2)
p3 = DataGrid1.Columns(3)
PreUni.Text = CDbl(p1) + CDbl(p2) + CDbl(p3) + CDbl(Horas.Text) + CDbl(Peso.Text)
PreTot.Text = CDbl(p1) + CDbl(p2) + CDbl(p3) + CDbl(Horas.Text) * CDbl(Peso.Text)
PreGan.Text = CDbl(p1) + CDbl(p2) + CDbl(p3) + CDbl(Horas.Text) + CDbl(Peso.Text) + CDbl(Ganancia.Text)
Horas.Text = ""
Peso.Text = ""
Ganancia.Text = ""
End Sub
Private Sub cmdAgr_Click()
With RsFabricacion
    .Requery
    .AddNew
    !PreUniFab = CDbl(PreUni.Text)
    !PreTotFab = CDbl(PreTot.Text)
    !PreGanFab = CDbl(PreGan.Text)
    !FecFab = Date
    !NroCli = VariableId
    .Update
    MsgBox ("Se Agrego Correctamente"), vbOKOnly, "Jack 3D"
End With
Horas.Text = ""
Peso.Text = ""
Ganancia.Text = ""
End Sub

Private Sub DataCombo1_Change()
With RsCliente
    .Requery
    .Find "NomCli='" & Trim(DataCombo1.Text) & "'"
    If .EOF Then MsgBox ("Error"), vbInformation, "Alerta": Exit Sub
    VariableId = !NroCli
End With
End Sub

Private Sub DataCombo1_Click(Area As Integer)
DataCombo1.Locked = False
cmdAgr.Enabled = True
cmdCal.Enabled = True
Horas.Enabled = True
Peso.Enabled = True
Ganancia.Enabled = True

End Sub
Private Sub Form_Load()
AbrirTablaCliente
Set DataCombo1.RowSource = RsCliente
DataCombo1.ListField = "NomCli"
DataCombo1.DataField = "NomCli"
AbrirTablaFabricacion
AbrirTablaConstantes
Set DataGrid1.DataSource = RsConstantes

End Sub

Private Sub Magrcli_Click()
frmnuevo.Show

End Sub

Private Sub mcon_Click()
FrmCont.Show
End Sub

Private Sub Melicli_Click()
frmElim.Show
End Sub

Private Sub Mfab_Click()
frmFab.Show vbModal
End Sub

Private Sub Mmodcli_Click()
frmModificar.Show

End Sub
