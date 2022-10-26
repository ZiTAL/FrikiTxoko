VERSION 5.00
Begin VB.Form FormuTotal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Totalak €-tan"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Danetara:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Total zor:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Total kuentan:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FormuTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Private Sub Form_Activate()
   Dim comando As String
    Dim RecHisto As Recordset
    Dim RecZor As Recordset
    Dim valor As Integer
    comando = "select sum(kantidadie) from historial"
    Set RecHisto = Bd.OpenRecordset(comando)
    RecHisto.MoveFirst
    valor = RecHisto.Fields(0).Value
    Text1.Text = valor
    
    '----------------
    comando = "select sum(kantidadie) from zorrak"
    Set RecZor = Bd.OpenRecordset(comando)
    RecZor.MoveFirst
    valor = RecZor.Fields(0).Value
    Text2.Text = valor
    '----------------
    Text3.Text = Val(Text1.Text) + Val(Text2.Text)
End Sub

Private Sub Form_Load()
    Set Wk = CreateWorkspace("", "admin", "", dbUseJet)
    Set Bd = Wk.OpenDatabase(App.Path & "\datubase.mdb")
    
End Sub
