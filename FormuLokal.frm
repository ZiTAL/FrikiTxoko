VERSION 5.00
Begin VB.Form FormuLokal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lokala paga dauienak eta eztauienak"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Sartu 
      Caption         =   "Sartu"
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ez"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Bai"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3960
      ItemData        =   "FormuLokal.frx":0000
      Left            =   3480
      List            =   "FormuLokal.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Kantidadie €-tan"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Paga dau honek gandulek?"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1980
   End
   Begin VB.Label Label3 
      Caption         =   "Abizena:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Izena:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Zenbakia:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FormuLokal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim cadena As String
    Option1(0).Value = True
    

    Set Wk = CreateWorkspace("", "admin", "", dbUseJet)
    Set Bd = Wk.OpenDatabase(App.Path & "\datubase.mdb")
    
    
    cadena = "select zenbakia,izena,abizena from sozioak"
    Set Rec = Bd.OpenRecordset(cadena)
    Rec.MoveFirst
    While Not Rec.EOF
        List1.AddItem Rec.Fields(0) & vbTab & Rec.Fields(1) & vbTab & Rec.Fields(2)
        Rec.MoveNext
    Wend
    Rec.MoveFirst
    DatuakikusiIngre
    
End Sub


Private Sub List1_DblClick()
Dim Listie As String
    Listie = Val(Left(List1.Text, 2))
    
    Rec.MoveFirst
    While Not Rec.EOF
        If Listie = Rec.Fields(0) Then
          DatuakikusiIngre
            Exit Sub
        End If
        Rec.MoveNext
    Wend
    
    

End Sub

Public Sub DatuakikusiIngre()
    Text1.Text = Rec.Fields(0).Value
    Text2.Text = Rec.Fields(1).Value
    Text3.Text = Rec.Fields(2).Value
    
End Sub

Private Sub Sartu_Click()
Dim comando As String
Dim vData As Date

If Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Datie eta kantidadie derrigorrezkuek diez", vbOKOnly, "Error"
    Exit Sub
End If
If Option1(0).Value = True Then
    vdate = (Text4.Text)
    comando = "insert into historial values(" & Val(Text1.Text) & ",#" & vdate & "#," & Val(Text5.Text) & ",'lokala paga','ingresue')"
    Bd.Execute comando
    MsgBox "Ingresua ondo sartu da historial taulan", vbOKOnly, "Ingresua"
    Text4.Text = " "
    Text5.Text = " "
    Text4.SetFocus
Else
    vdate = (Text4.Text)
    comando = "insert into zorrak values(" & Val(Text1.Text) & ",#" & vdate & "#," & Val(Text5.Text) & ",'lokala ez da pagatie')"
    Bd.Execute comando
    comando = "update sozioak set zorrak=1 " & "where zenbakia=" & Val(Text1.Text)
    Bd.Execute comando
    MsgBox "Bazkide honek ez du ordaindu, beraz zorrak taulan 22 € sartu jakoz ", vbOKOnly, "Zorra"
    Text4.Text = " "
    Text5.Text = " "
    Text4.SetFocus
End If

    

End Sub



Private Sub Text4_LostFocus()
    If Text4.Text = "" Then
        MsgBox "Datie sartunbizu derrigorrez", vbOKOnly, "Error"
        Text4.SetFocus
    End If
    
End Sub

Private Sub Text5_Change()
    If Text5.Text = "" Then
        MsgBox "Kantidadie derrigorrezkue da", vbOKOnly, "Error"
        Text5.SetFocus
    End If
End Sub
