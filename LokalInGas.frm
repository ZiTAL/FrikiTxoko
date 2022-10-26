VERSION 5.00
Begin VB.Form LokalInGas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Zenbakia:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Kantidadie:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Zergaitije:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Abizena:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Izena:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "LokalInGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Wk As Workspace
Dim Bd As Database
Dim Rec As Recordset


Private Sub Command1_Click()
    Dim RecFecha As Recordset
    Dim comando As String
    Dim vdate As Date
    comando = "select date() from sozioak"
    Set RecFecha = Bd.OpenRecordset(comando)
    RecFecha.MoveFirst
    vdate = RecFecha.Fields(0).Value
    RecFecha.Close
    
    Dim vzenbakia As Integer
    Dim vzergaitia As String
    
    Dim vkantidadie As Integer
    If Text1(3).Text = "" Or Text1(4).Text = "" Then
        MsgBox "Zergaitije eta Kantidadie bete", vbOKOnly, "Error"
        Text1(3).SetFocus
    Else
    vzenbakia = Val(Text1(0).Text)
    vzergaitia = Trim(Text1(3).Text)
    vkantidadie = Val(Text1(4).Text)
    vdate = Date
    
    
    If Command1.Caption = "Sartun" Then
        BaiEz.Caption = "Dirue Sartun lokaleko kuentan"
        BaiEz.Label1.Caption = "Seguru zauz dirue sartutiegaz?"
        BaiEz.Show vbModal
        If Erantzuna = True Then
            comando = "insert into historial values(" & vzenbakia & ",#" & vdate & "#," & vkantidadie & ",'" & vzergaitia & "','ingresue')"
            Bd.Execute comando
            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text1(3).Text = ""
            Text1(4).Text = ""
            Text1(0).SetFocus
            
        End If
    Else
        BaiEz.Caption = "Dirue Atara lokaleko kuentan"
        BaiEz.Label1.Caption = "Seguru zauz dirue ataratiegaz?"
        BaiEz.Show vbModal
        If Erantzuna = True Then
            comando = "insert into historial values(" & vzenbakia & ",#" & vdate & "#,-" & vkantidadie & ",'" & vzergaitia & "','gastue')"
            Bd.Execute comando
            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(2).Text = ""
            Text1(3).Text = ""
            Text1(4).Text = ""
            Text1(0).SetFocus
        End If
    End If
    End If
End Sub

Private Sub Form_Activate()
    Text1(0).SetFocus
End Sub

Private Sub Form_Load()
    Set Wk = CreateWorkspace("", "admin", "", dbUseJet)
    Set Bd = Wk.OpenDatabase(App.Path & "\datubase.mdb")
    Set Rec = Bd.OpenRecordset("sozioak", dbOpenTable)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim vmaximo As Integer
    Dim RecMax As Recordset
    Dim comando As String
    Dim topa As Boolean
    
    comando = "select max(zenbakia) from sozioak"
    Set RecMax = Bd.OpenRecordset(comando)
    RecMax.MoveFirst
    vmaximo = RecMax.Fields(0).Value
    
    topa = False
    Dim numero As Integer
    numero = Val(Text1(0).Text)
    Rec.MoveFirst
    While Not Rec.EOF
        If numero = Rec.Fields(0).Value Then
            topa = True
            Text1(1).Text = Rec.Fields(1).Value
            Text1(2).Text = Rec.Fields(2).Value
        End If
        Rec.MoveNext
        
    Wend
    If topa = False Then
        MsgBox "Kodiguek 1-etik " & vmaximo & "-ra dauz", vbOKOnly, "Error"
        Text1(0).SetFocus
        Text1(0).Text = ""
        Text1(1).Text = ""
        Text1(3).Text = ""
        
    End If
    
End Sub
