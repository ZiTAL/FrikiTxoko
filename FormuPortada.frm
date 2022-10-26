VERSION 5.00
Begin VB.Form FormuPortada 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6720
      Top             =   1320
   End
   Begin VB.Image Image1 
      Height          =   1965
      Left            =   0
      Picture         =   "FormuPortada.frx":0000
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "FormuPortada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Unload Me
    Menupri.Show vbModal
    Timer1.Enabled = False
End Sub
