Attribute VB_Name = "Module1"
Option Explicit
Public Variable01 As Integer
Public Variable02 As String
Public variable03 As Date
Public variable04 As String
Public variable05 As Date
Public variable06 As Date

Public Wk As Workspace
Public Bd As Database
Public Rec As Recordset

Public Erantzuna As Boolean

Public BaiEzCaption As String
Public BaiEzLabel As String

Public KideTalde As Integer


Public Sub DatuakIkusi()
    
    FormuKide.Text1(0).Text = Rec.Fields("Zenbakia")
    FormuKide.Text1(1).Text = Rec.Fields("Izena")
    FormuKide.Text1(2).Text = Rec.Fields("Abizena")
    FormuKide.Text1(3).Text = Rec.Fields("Jaiotze_data")
    FormuKide.Text1(4).Text = Rec.Fields("Helbidea")
    FormuKide.Text1(5).Text = Rec.Fields("Telefonoa")
    FormuKide.Text1(6).Text = Rec.Fields("Taldie")
    If Rec.Fields("Zorrak") = "Verdadero" Then
        FormuKide.Check1.Value = 1
    Else
        FormuKide.Check1.Value = 0
    End If
End Sub

Public Sub DatuakSartu()
    Rec.Fields(0) = FormuKide.Text1(0).Text
    Rec.Fields(1) = FormuKide.Text1(1).Text
    Rec.Fields(2) = FormuKide.Text1(2).Text
    Rec.Fields(3) = FormuKide.Text1(3).Text
    Rec.Fields(4) = FormuKide.Text1(4).Text
    Rec.Fields(5) = FormuKide.Text1(5).Text
    Rec.Fields(6) = FormuKide.Text1(6).Text
End Sub

Public Sub Blokie()

    Dim i As Integer
    For i = 0 To FormuKide.Controls.Count - 1
        If TypeOf FormuKide.Controls(i) Is TextBox Then
            FormuKide.Controls(i).Locked = True
        End If
    Next i
    FormuKide.Check1.Enabled = False
End Sub
Public Sub DesBlokie()

    Dim i As Integer
    For i = 0 To FormuKide.Controls.Count - 1
        If TypeOf FormuKide.Controls(i) Is TextBox Then
            FormuKide.Controls(i).Locked = False
        End If
    Next i
    FormuKide.Check1.Enabled = True
End Sub

Public Sub Kontsulta()
    Dim i As Integer
    Dim Rec2 As Recordset
    Dim cadena As String
    Dim consulta As Recordset
    Dim Lerro As Integer
    
    cadena = "select data,kantidadie,zergaitia from zorrak where zenbakia=" & Val(FormuKide.Text1(0).Text)
    Set Rec2 = Bd.OpenRecordset(cadena)

    Lerro = Rec2.RecordCount
    If Lerro = 0 Then
        FormuKide.Rejilla.Visible = False
        FormuKide.Height = 5025
        FormuKide.Elimine.Enabled = True
        Exit Sub
    Else
        Rec2.MoveLast
        Lerro = Rec2.RecordCount
        FormuKide.Elimine.Enabled = False
        FormuKide.Rejilla.Height = (Lerro + 1) * 275
    End If
    FormuKide.Height = 7875

    With FormuKide.Rejilla
        .ColWidth(0) = 0
        .Cols = 4
        .Rows = Rec2.RecordCount + 1
        .Row = 0
        '--------
        .Col = 1
        .Text = "Egune"
        .ColWidth(1) = 1600
        '--------
        .Col = 2
        .Text = "kantidadie"
        .ColWidth(2) = 1000
        '--------
        .Col = 3
        .Text = "Zergaitia"
        .ColWidth(3) = 2000
        '--------
        .Visible = True
        Rec2.MoveFirst
        i = 1
        Do While Not Rec2.EOF
            .Row = i
            '-------
            .Col = 1
            .Text = Rec2.Fields(0)
            '-------
            .Col = 2
            .Text = Rec2.Fields(1)
            '-------
            .Col = 3
            .Text = Rec2.Fields(2)
            '-------
            i = i + 1
            Rec2.MoveNext
        Loop
        
        End With
        Rec2.Close
        

End Sub

