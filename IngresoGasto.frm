VERSION 5.00
Begin VB.Form IngresoGasto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresuek eta gastuek"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "=>"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Lokaleko Gastuek"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Gastuek:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingresuek:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "IngresoGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
    FormuLokal.Show vbModal
End Sub

Private Sub Command3_Click()
    LokalInGas.Caption = "Gastue"
    LokalInGas.Command1.Caption = "Atara"
    LokalInGas.Show vbModal
End Sub

Private Sub Command4_Click()
    If Combo1.ListIndex = 0 Then
        FormuLokal.Show vbModal
    End If
    If Combo1.ListIndex = 1 Then
        LokalInGas.Caption = "Ingresue"
        LokalInGas.Command1.Caption = "Sartun"
        LokalInGas.Show vbModal
    End If
End Sub

Private Sub Form_Load()
    Combo1.AddItem "Lokala paga"
    Combo1.AddItem "Beste ingreso batzuk"
    Combo1.ListIndex = 0
    
End Sub
