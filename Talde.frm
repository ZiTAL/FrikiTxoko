VERSION 5.00
Begin VB.Form Talde 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2685
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Talde.frx":0000
      Left            =   120
      List            =   "Talde.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Talde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Urten As Boolean

Private Sub Form_Load()
Urten = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Urten = True Then
        Rec.MoveFirst
        While Not Rec.EOF
            If Rec.Fields(0) = KideTalde Then
                DatuakIkusi
                Exit Sub
            End If
            Rec.MoveNext
        Wend
    End If
End Sub



Private Sub List1_DblClick()
    Dim Listie As Integer
    
    Urten = False
    Listie = Val(Left(List1.Text, 2))
    
    Rec.MoveFirst
    While Not Rec.EOF
        If Listie = Rec.Fields(0) Then
            Blokie
            DatuakIkusi
            
            Kontsulta
            Unload Me
            Exit Sub
        End If
        Rec.MoveNext
    Wend
End Sub
