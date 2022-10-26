VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormuEgileak 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Egileak:"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   55
      ImageHeight     =   68
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormuEgileak.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormuEgileak.frx":2CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormuEgileak.frx":5744
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   120
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   120
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "www.zital.tk"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "zitalman@hotmail.com"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Iban Bilbao Barturen"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "motxene@msn.com"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Iñigo Allika Barrena"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "gabantxo@hotmail.com"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Roberto Gabantxo Martin"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FormuEgileak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = vbBlack
    Label2.FontBold = False
    Label4.ForeColor = vbBlack
    Label4.FontBold = False
    Label6.ForeColor = vbBlack
    Label6.FontBold = False
    Label7.ForeColor = vbBlack
    Label7.FontBold = False
    '-----------
    Image1.Picture = ImageList1.ListImages(1).Picture
    Image2.Picture = ImageList1.ListImages(2).Picture
    Image3.Picture = ImageList1.ListImages(3).Picture
    
End Sub

Private Sub Label2_Click()
    Dim sLink As String
       
           'e-mail
           sLink = "gabantxo@hotmail.com"
           ShellExecute 0, vbNullString, "mailto:" & sLink, vbNullString, _
           vbNullString, vbNormalFocus

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.FontBold = True
    Label2.ForeColor = vbBlue
    
End Sub

Private Sub Label4_Click()
    Dim sLink As String
       
           'e-mail
           sLink = "motxene@msn.com"
           ShellExecute 0, vbNullString, "mailto:" & sLink, vbNullString, _
           vbNullString, vbNormalFocus
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Label4.FontBold = True
    Label4.ForeColor = vbBlue
End Sub

Private Sub Label6_Click()
    Dim sLink As String
       
           'e-mail
           sLink = "zitalman@hotmail.com"
           ShellExecute 0, vbNullString, "mailto:" & sLink, vbNullString, _
           vbNullString, vbNormalFocus
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Label6.FontBold = True
    Label6.ForeColor = vbBlue
End Sub

Private Sub Label7_Click()
    Dim sLink As String
    sLink = "http://www.zital.tk"
           ShellExecute 0, vbNullString, sLink, vbNullString, _
           vbNullString, vbNormalFocus

End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label7.FontBold = True
    Label7.ForeColor = vbBlue
End Sub
