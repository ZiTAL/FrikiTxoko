VERSION 5.00
Begin VB.Form FormuZor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOrosuek"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
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
      Height          =   495
      Left            =   3360
      TabIndex        =   23
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command8 
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
      Height          =   495
      Left            =   2520
      TabIndex        =   22
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command7 
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
      Height          =   495
      Left            =   1680
      TabIndex        =   21
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command6 
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
      Height          =   495
      Left            =   840
      TabIndex        =   20
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox ZorAbi 
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox ZorIzen 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox ZorZen 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox TotZor 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5535
      Begin VB.CommandButton Command5 
         Caption         =   "Zorra Kendu"
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">>"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox ZorKantidadie 
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox ZorZergaitia 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox ZorData 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Kantidadie:"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Zergaitia:"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Total zor:"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Izena:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Abizena:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Zenbakia:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FormuZor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ZorRec As Recordset

Private Sub Command1_Click()
    ZorRec.MoveFirst
    DatuakIkusiZor2
End Sub

Private Sub Command2_Click()
    ZorRec.MovePrevious
    If ZorRec.BOF Then
        ZorRec.MoveFirst
    End If
    DatuakIkusiZor2
End Sub

Private Sub Command3_Click()
    ZorRec.MoveNext
    If ZorRec.EOF Then
        ZorRec.MoveLast
    End If
    DatuakIkusiZor2
End Sub

Private Sub Command4_Click()
    ZorRec.MoveLast
    DatuakIkusiZor2
End Sub

Private Sub Command5_Click()
    Dim HistorialRec As Recordset
    Dim Historial As String
    
    Dim Zordun As String
    
    Dim vZenbakia As Integer
    Dim vData As Date
    Dim vKantidadie As Integer
    Dim vZergaitia As String
    Dim vIngas As String
    
    BaiEzCaption = "Zorra kendu"
    BaiEzLabel = FormuZor.ZorZen & " " & FormuZor.ZorIzen & " " & FormuZor.ZorAbi & " -ri zorra kendu?"
    BaiEz.Show vbModal
    If Erantzuna = True Then
            
    
        '-----------------------------------
        vZenbakia = Val(FormuZor.ZorZen.Text)
        vData = Trim(FormuZor.ZorData.Text)
        vKantidadie = Val(FormuZor.ZorKantidadie.Text)
        vZergaitia = Trim(FormuZor.ZorZergaitia.Text)
        vIngas = "ingresue"
    
        
        Historial = "insert into historial (zenbakia,data,kantidadie,zergaitia,ingas) values (" & vZenbakia & ",#" & vData & "#," & vKantidadie & ",'" & vZergaitia & "','" & vIngas & "')"
        Bd.Execute Historial
        '--------------------------
        ZorRec.Delete
        
        If ZorRec.RecordCount <> 0 Then
            ZorRec.MoveFirst
            DatuakIkusiZor
            KontsultaZor
        Else
            Dim Up As String
            Dim UpRec As Recordset
            
            Up = " update sozioak set zorrak='Falso' where zenbakia=" & Val(FormuZor.ZorZen.Text)
            Bd.Execute Up
            Zordun = "select distinct sozioak.Zenbakia,Izena,Abizena from sozioak inner join zorrak on zorrak.zenbakia=sozioak.zenbakia"
            Set Rec = Bd.OpenRecordset(Zordun)
            
            If Rec.RecordCount <> 0 Then
                Rec.MoveFirst
                DatuakIkusiZor
                KontsultaZor
            Else
                Unload Me
                MsgBox "Iñok eztako zorrik :)", vbOKOnly, "Zorge"
            End If
        End If
    End If
    
    
End Sub

Private Sub Command6_Click()
    Rec.MoveFirst
    DatuakIkusiZor
    KontsultaZor
End Sub

Private Sub Command7_Click()
    Rec.MovePrevious
    If Rec.BOF Then
        Rec.MoveFirst
    End If
    DatuakIkusiZor
    KontsultaZor
    
End Sub

Private Sub Command8_Click()
    Rec.MoveNext
    If Rec.EOF Then
        Rec.MoveLast
    End If
    DatuakIkusiZor
    KontsultaZor
End Sub

Private Sub Command9_Click()
    Rec.MoveLast
    DatuakIkusiZor
    KontsultaZor
End Sub

Private Sub Form_Activate()

    Dim Zordun As String
    Dim cadena As String

    
    
    Set Wk = CreateWorkspace("", "admin", "", dbUseJet)
    Set Bd = Wk.OpenDatabase(App.Path & "\datubase.mdb")
    
    cadena = "select data,kantidadie,zergaitia from zorrak where zenbakia=" & Val(FormuZor.ZorZen.Text)
    Set ZorRec = Bd.OpenRecordset(cadena)
    
    
    Zordun = "select distinct sozioak.Zenbakia,Izena,Abizena from sozioak inner join zorrak on zorrak.zenbakia=sozioak.zenbakia"
    'ZorDun = "SELECT sozioak.Zenbakia, sozioak.Izena, sozioak.Abizena FROM sozioak INNER JOIN zorrak ON sozioak.Zenbakia = zorrak.zenbakia GROUP BY sozioak.Zenbakia, sozioak.Izena, sozioak.Abizena"
    Set Rec = Bd.OpenRecordset(Zordun)
    '-------------------------------------
    If ZorRec.RecordCount <> 0 Then
            ZorRec.MoveFirst
            DatuakIkusiZor
            KontsultaZor
        Else
            Zordun = "select distinct sozioak.Zenbakia,Izena,Abizena from sozioak inner join zorrak on zorrak.zenbakia=sozioak.zenbakia"
            Set Rec = Bd.OpenRecordset(Zordun)
            
            If Rec.RecordCount <> 0 Then
                Rec.MoveFirst
                DatuakIkusiZor
                KontsultaZor
            Else
                Unload Me
                MsgBox "Iñok eztako zorrik :)", vbOKOnly, "ZorGe"
                
                Exit Sub
               
            End If
        End If
End Sub

Private Sub Form_Load()
    Dim Zordun As String
    Dim cadena As String

    
    
    Set Wk = CreateWorkspace("", "admin", "", dbUseJet)
    Set Bd = Wk.OpenDatabase(App.Path & "\datubase.mdb")
    
    cadena = "select data,kantidadie,zergaitia from zorrak where zenbakia=" & Val(FormuZor.ZorZen.Text)
    Set ZorRec = Bd.OpenRecordset(cadena)
    
    
    Zordun = "select distinct sozioak.Zenbakia,Izena,Abizena from sozioak inner join zorrak on zorrak.zenbakia=sozioak.zenbakia"
    'ZorDun = "SELECT sozioak.Zenbakia, sozioak.Izena, sozioak.Abizena FROM sozioak INNER JOIN zorrak ON sozioak.Zenbakia = zorrak.zenbakia GROUP BY sozioak.Zenbakia, sozioak.Izena, sozioak.Abizena"
    Set Rec = Bd.OpenRecordset(Zordun)
    If ZorRec.RecordCount <> 0 Then
            ZorRec.MoveFirst
            DatuakIkusiZor
            KontsultaZor
        Else
            Zordun = "select distinct sozioak.Zenbakia,Izena,Abizena from sozioak inner join zorrak on zorrak.zenbakia=sozioak.zenbakia"
            Set Rec = Bd.OpenRecordset(Zordun)
            
            If Rec.RecordCount <> 0 Then
                Rec.MoveFirst
                DatuakIkusiZor
                KontsultaZor
            Else
                FormuZor.Hide
                
                Exit Sub
               
            End If
        End If
    
End Sub

Private Sub KontsultaZor()
    
    
    Dim cadena As String
    
    Dim TotalRec As Recordset
    Dim Total As String
    
    cadena = "select data,kantidadie,zergaitia from zorrak where zenbakia=" & Val(FormuZor.ZorZen.Text)
    Set ZorRec = Bd.OpenRecordset(cadena)
    
    ZorRec.MoveFirst
    FormuZor.ZorData.Text = ZorRec.Fields(0).Value
    FormuZor.ZorZergaitia.Text = ZorRec.Fields(2).Value
    FormuZor.ZorKantidadie.Text = ZorRec.Fields(1).Value
    
    Total = "select sum(kantidadie) from zorrak where zenbakia=" & Val(FormuZor.ZorZen.Text)
    Set TotalRec = Bd.OpenRecordset(Total)
    FormuZor.TotZor = TotalRec.Fields(0).Value
    
End Sub
Private Sub DatuakIkusiZor()

    FormuZor.ZorZen.Text = Rec.Fields("Zenbakia")
    FormuZor.ZorIzen.Text = Rec.Fields("Izena")
    FormuZor.ZorAbi.Text = Rec.Fields("Abizena")
End Sub
Private Sub DatuakIkusiZor2()

    FormuZor.ZorData.Text = ZorRec.Fields(0)
    FormuZor.ZorKantidadie.Text = ZorRec.Fields(1)
    FormuZor.ZorZergaitia.Text = ZorRec.Fields(2)
End Sub
