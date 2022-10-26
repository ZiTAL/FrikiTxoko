VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormuKide 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lokaleko kidiek"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Graba 
      Height          =   495
      Left            =   3360
      Picture         =   "FormuKide.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Graba"
      Top             =   240
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.CommandButton Barrije 
      Height          =   495
      Left            =   3360
      Picture         =   "FormuKide.frx":0534
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Registro barrije"
      Top             =   240
      Width           =   570
   End
   Begin VB.CommandButton Kandaue 
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Ezabatu"
      Top             =   840
      Width           =   570
   End
   Begin VB.CommandButton BotoiTalde 
      Height          =   495
      Left            =   3360
      Picture         =   "FormuKide.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Taldeko Kidiek"
      Top             =   2040
      Width           =   570
   End
   Begin VB.CommandButton Elimine 
      Height          =   495
      Left            =   3360
      Picture         =   "FormuKide.frx":0B63
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Ezabatu"
      Top             =   1440
      Width           =   570
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   26
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormuKide.frx":1037
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormuKide.frx":1169
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormuKide.frx":1296
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Rejilla 
      Height          =   1935
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Azkana 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      ToolTipText     =   "Azkaningokue"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Hurrengue 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      ToolTipText     =   "Hurrengue"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Arinaukue 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      ToolTipText     =   "Ariñaukue"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Lehenengue 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      ToolTipText     =   "Lehengue"
      Top             =   4200
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3765
      ItemData        =   "FormuKide.frx":177A
      Left            =   4200
      List            =   "FormuKide.frx":177C
      TabIndex        =   16
      Top             =   240
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   14
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Zorran datuek:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Zorrak:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Taldie:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Telefonoa:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Helbidea:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Jaiotze data:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Abizena:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Izena:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Zenbakia:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FormuKide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private A As Integer



Private Sub Arinaukue_Click()
   Rec.MovePrevious
   If Rec.BOF Then Rec.MoveFirst
   Blokie
   DatuakIkusi
   
   Kontsulta
   
End Sub

Private Sub Azkana_Click()
    Rec.MoveLast
    Blokie
    DatuakIkusi
    
    Kontsulta
    
End Sub

Private Sub Barrije_Click()
    Dim i As Integer
    Dim comando As String
    Dim RecMax As Recordset
    Dim vmax As Integer
    Rec.AddNew
    For i = 1 To 6
        Text1(i).Text = ""
    Next i
    DesBlokie
    Text1(0).Enabled = False
    comando = "select max(zenbakia) from sozioak"
    Set RecMax = Bd.OpenRecordset(comando)
    RecMax.MoveFirst
    vmax = RecMax.Fields(0).Value
    Text1(0).Text = vmax + 1
    Check1.Value = 0
    Check1.Enabled = False
    Text1(1).SetFocus
    FormuKide.Height = 4440
    FormuKide.Width = 4080
    
    Barrije.Visible = False
    Graba.Visible = True
    Elimine.Visible = False
    BotoiTalde.Visible = False
    Kandaue.Visible = False
    
    Lehenengue.Enabled = False
    Arinaukue.Enabled = False
    Hurrengue.Enabled = False
    Azkana.Enabled = False
    List1.Enabled = False
    Rejilla.Visible = False
End Sub

Private Sub BotoiTalde_Click()
    Dim TaldeZenbaki As Integer
    KideTalde = Rec.Fields(0)
    
    TaldeZenbaki = Rec.Fields("taldie")
    Talde.Caption = TaldeZenbaki & ". taldie"
    
    Rec.MoveFirst
    While Not Rec.EOF
        If TaldeZenbaki = Rec.Fields("taldie") Then
            Talde.List1.AddItem Rec.Fields(0) & vbTab & Rec.Fields(1) & vbTab & Rec.Fields(2)
        End If
        Rec.MoveNext
    Wend
    Talde.Show vbModal
End Sub


Private Sub Elimine_Click()
        BaiEzCaption = "Kidea Ezabatu"
        BaiEzLabel = Rec.Fields(0) & " " & Rec.Fields(1) & " " & Rec.Fields(2) & " guzu borra?"
        BaiEz.Show vbModal
        If Erantzuna = True Then
            Rec.Delete
            Rec.MoveFirst
            List1.Clear
            While Not Rec.EOF
                List1.AddItem Rec.Fields(0) & vbTab & Rec.Fields(1) & vbTab & Rec.Fields(2)
            Rec.MoveNext
            Wend
            Rec.MoveFirst
            DatuakIkusi
            Kontsulta
        End If
            
End Sub

Private Sub Form_Load()
    A = 1
    
    Kandaue.Picture = ImageList1.ListImages(1).Picture
    Kandaue.ToolTipText = "Ezin da modifikaziorik egin"
    Kandaue.DownPicture = ImageList1.ListImages(2).Picture
    
    'Base de datos-a eta "sozioak" tablie igiri
    Set Wk = CreateWorkspace("", "admin", "", dbUseJet)
    Set Bd = Wk.OpenDatabase(App.Path & "\datubase.mdb")
    Set Rec = Bd.OpenRecordset("sozioak", dbOpenTable)
    
    Rec.MoveFirst
    DatuakIkusi
    
    While Not Rec.EOF
        List1.AddItem Rec.Fields(0) & vbTab & Rec.Fields(1) & vbTab & Rec.Fields(2)
        Rec.MoveNext
    Wend
    Rec.MoveFirst
    Kontsulta
    Blokie
End Sub

Private Sub Graba_Click()
    If Text1(1).Text = "" Or Text1(2).Text = "" Or Text1(3).Text = "" Or Text1(4).Text = "" Or Text1(5).Text = "" Or Text1(6).Text = "" Then
        MsgBox "Kaja danak bete bidiez", vbOKOnly, "Error"
        Text1(1).SetFocus
    Else
    BaiEz.Caption = "Graba"
    BaiEz.Label1.Caption = Trim(FormuKide.Text1(1).Text) & " " & Trim(FormuKide.Text1(2).Text) & " graba sozijuen taulen?"
    BaiEz.Show vbModal
    If Erantzuna = True Then
        DatuakSartu
        Rec.Update
        FormuKide.List1.Clear
        Rec.MoveFirst
        While Not Rec.EOF
            List1.AddItem Rec.Fields(0) & vbTab & Rec.Fields(1) & vbTab & Rec.Fields(2)
            Rec.MoveNext
        Wend
        Rec.MoveLast
    Else
        Rec.MoveFirst
        DatuakIkusi
        
    End If
    Text1(0).Enabled = True
    Blokie
    Kontsulta
    If Rejilla.Visible = True Then
            FormuKide.Height = 7650
    Else
            FormuKide.Height = 5025
    End If
    FormuKide.Width = 6795
     Barrije.Visible = True
    Graba.Visible = False
    Elimine.Visible = True
    BotoiTalde.Visible = True
    Kandaue.Visible = True
    
    Lehenengue.Enabled = True
    Arinaukue.Enabled = True
    Hurrengue.Enabled = True
    Azkana.Enabled = True
    List1.Enabled = True
    End If
    
    
    
End Sub

Private Sub Hurrengue_Click()
    Rec.MoveNext
    If Rec.EOF Then Rec.MoveLast
    Blokie
    DatuakIkusi
    
    Kontsulta
    
End Sub

Private Sub Kandaue_Click()
        
    If A = 1 Then
        A = 2
        Kandaue.DownPicture = ImageList1.ListImages(1).Picture
        Text1(0).Enabled = False
        Kandaue.ToolTipText = "Aldatu daiteke"
        DesBlokie
        Check1.Enabled = False
        Rec.Edit
        FormuKide.Height = 4440
        FormuKide.Width = 4080
        
        Lehenengue.Enabled = False
        Arinaukue.Enabled = False
        Hurrengue.Enabled = False
        Azkana.Enabled = False
        Elimine.Visible = False
        BotoiTalde.Visible = False
        Barrije.Visible = False
        
    Else
        A = 1
        Kandaue.DownPicture = ImageList1.ListImages(2).Picture
        BaiEzCaption = "Datuak aldatu"
        BaiEzLabel = Rec.Fields(0) & " " & Rec.Fields(1) & " " & Rec.Fields(2) & " -ren datuak aldatu?"
        BaiEz.Show vbModal
        
        If Erantzuna = True Then
            DatuakSartu
            Rec.Update
            FormuKide.List1.Clear
            Rec.MoveFirst
            While Not Rec.EOF
                List1.AddItem Rec.Fields(0) & vbTab & Rec.Fields(1) & vbTab & Rec.Fields(2)
                Rec.MoveNext
            Wend
            Rec.MoveFirst
            While Rec.Fields(0) <> Val(Text1(0).Text)
                Rec.MoveNext
            Wend
            Blokie
        Else
            Rec.Update
            DatuakIkusi
            Blokie
            Check1.Enabled = False
        End If
        Text1(0).Enabled = True
    
        Kandaue.ToolTipText = "Ezin da modifikaziorik egin"
        FormuKide.Width = 6795
        If Rejilla.Visible = True Then
            FormuKide.Height = 7650
        Else
            FormuKide.Height = 5025
        End If
        Lehenengue.Enabled = True
        Arinaukue.Enabled = True
        Hurrengue.Enabled = True
        Azkana.Enabled = True
        Elimine.Visible = True
        BotoiTalde.Visible = True
        Barrije.Visible = True
    End If
    
    Kandaue.Picture = ImageList1.ListImages(A).Picture
    
       
    
End Sub

Private Sub Lehenengue_Click()
    
        Rec.MoveFirst
        Blokie
        DatuakIkusi
        Kontsulta
        
    
    
End Sub

Private Sub List1_DblClick()
    Dim Listie As String
    Listie = Val(Left(List1.Text, 2))
    
    Rec.MoveFirst
    While Not Rec.EOF
        If Listie = Rec.Fields(0) Then
            Blokie
            DatuakIkusi
            
            Kontsulta
            
            Exit Sub
        End If
        Rec.MoveNext
    Wend
    
    

End Sub



