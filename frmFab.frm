VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFab 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   Picture         =   "frmFab.frx":0000
   ScaleHeight     =   5460
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   2415
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmFab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Text1_Change()

End Sub
