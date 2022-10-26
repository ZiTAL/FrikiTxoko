VERSION 5.00
Begin VB.Form FormuInforme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informiek"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Historial osue"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FormuInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    If Combo1.ListIndex = 3 Or Combo1.ListIndex = 4 Then
        Text1.Visible = False
        Text1.Text = ""
    Else
        Text1.Visible = True
    End If
End Sub

Private Sub Command1_Click()
    Informe01.Show vbModal
End Sub

Private Sub Command2_Click()
    If Text1.Text = "" And Text1.Visible = True Then
        MsgBox "Sartun zeuzer kajan", vbOKOnly, "Error"
        Text1.SetFocus
    Else
    If Combo1.ListIndex = 0 Then
        Variable01 = Val(Text1.Text)
        If Entorno.rsCommand2.State = adStateOpen Then
            Entorno.rsCommand2.Close
        End If
        Entorno.Command2 (Variable01)
        Informe02.Show vbModal
    End If
    If Combo1.ListIndex = 1 Then
        Variable02 = Trim(Text1.Text)
        If Entorno.rsCommand3.State = adStateOpen Then
            Entorno.rsCommand3.Close
        End If
        Entorno.Command3 (Variable02)
        Informe03.Show vbModal
    End If
    If Combo1.ListIndex = 2 Then
        variable03 = Text1.Text
        If Entorno.rsCommand4.State = adStateOpen Then
            Entorno.rsCommand4.Close
        End If
        Entorno.Command4 (variable03)
        Informe04.Show vbModal
    End If
    If Combo1.ListIndex = 3 Then
        variable04 = "ingresue"
        Informe05.Caption = "HiSTORiALA iNGRESUENA"
        If Entorno.rsCommand5.State = adStateOpen Then
            Entorno.rsCommand5.Close
        End If
        Entorno.Command5 (variable04)
        Informe05.Show vbModal
    End If
    If Combo1.ListIndex = 4 Then
        variable04 = "gastue"
        Informe05.Caption = "HiSTORiALA GASTUENA"
        If Entorno.rsCommand5.State = adStateOpen Then
            Entorno.rsCommand5.Close
        End If
        Entorno.Command5 (variable04)
        Informe05.Show vbModal
    End If
    End If
  
End Sub

Private Sub Form_Activate()
        'Kill "*.tmp"
End Sub

Private Sub Form_Load()

    Combo1.Clear
    Combo1.AddItem "Zenbakia"
    Combo1.AddItem "Izena"
    Combo1.AddItem "Data"
    Combo1.AddItem "Ingresuek"
    Combo1.AddItem "Gastuek"
    Combo1.ListIndex = 0
    
    
End Sub
