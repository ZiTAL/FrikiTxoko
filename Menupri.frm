VERSION 5.00
Begin VB.Form Menupri 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Friki Txoko"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Egileak"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Historialan informiek"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Totalak"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Txokoko Kideak"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zorrak"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresuek/Gastuek"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "ARTiKE BiDEA 3 TXOKO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Menupri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    IngresoGasto.Show vbModal
End Sub

Private Sub Command2_Click()
    FormuZor.Show vbModal
End Sub

Private Sub Command3_Click()
    FormuKide.Show vbModal
End Sub

Private Sub Command4_Click()
    FormuTotal.Show vbModal
End Sub

Private Sub Command5_Click()
    FormuInforme.Show vbModal
End Sub

Private Sub Command6_Click()
    FormuEgileak.Show vbModal
End Sub

Private Sub Form_Load()
    Entorno.Cone.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\datubase.mdb;Persist Security Info=False"
    Entorno.Cone.Open
End Sub


