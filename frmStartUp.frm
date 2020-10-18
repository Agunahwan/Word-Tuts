VERSION 5.00
Begin VB.Form frmStartUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WordTuts"
   ClientHeight    =   1365
   ClientLeft      =   1620
   ClientTop       =   720
   ClientWidth     =   2055
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblContinue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Continue Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image imgNewBiru 
      Height          =   495
      Left            =   120
      Picture         =   "frmStartUp.frx":17D2A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image imgContinueBiru 
      Height          =   495
      Left            =   120
      Picture         =   "frmStartUp.frx":1BB89
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image imgNewHijau 
      Height          =   495
      Left            =   120
      Picture         =   "frmStartUp.frx":1F9E8
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image imgContinueHijau 
      Height          =   495
      Left            =   120
      Picture         =   "frmStartUp.frx":23733
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image imgBack 
      Height          =   1575
      Left            =   0
      Picture         =   "frmStartUp.frx":2747E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error Resume Next
    Win = False
    If App.PrevInstance = True Then End
    With imgBack
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = Me.Height
    End With
    If PeriksaFile = False Then
        Unload Me
        Level = 1
        frmUtama.Show
        Exit Sub
    End If
End Sub

Sub Biru()
On Error Resume Next
    imgContinueBiru.Visible = True
    imgNewBiru.Visible = True
    imgContinueHijau.Visible = False
    imgNewHijau.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Biru
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Biru
End Sub

Private Sub imgContinueBiru_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Biru
    imgContinueBiru.Visible = False
    imgContinueHijau.Visible = True
End Sub

Private Sub imgContinueHijau_Click()
    BukaLevel
    Unload Me
    frmUtama.Show
End Sub

Private Sub imgContinueHijau_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgContinueHijau.BorderStyle = 1
End Sub

Private Sub imgContinueHijau_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgContinueHijau.BorderStyle = 0
End Sub

Private Sub imgNewBiru_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Biru
    imgNewBiru.Visible = False
    imgNewHijau.Visible = True
End Sub

Private Sub imgNewHijau_Click()
On Error Resume Next
Dim Alamat As String
    Alamat = App.Path
    If Right(Alamat, 1) <> "\" Then
        Alamat = Alamat & "\"
    End If
    Kill Alamat & "data.wt"
    Unload Me
    Level = 1
    frmUtama.Show
End Sub

Private Sub imgNewHijau_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNewHijau.BorderStyle = 1
End Sub

Private Sub imgNewHijau_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNewHijau.BorderStyle = 0
End Sub

Private Sub lblContinue_Click()
    BukaLevel
    Unload Me
    frmUtama.Show
End Sub

Private Sub lblContinue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgContinueHijau.BorderStyle = 1
End Sub

Private Sub lblContinue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Biru
    imgContinueBiru.Visible = False
    imgContinueHijau.Visible = True
End Sub

Private Sub lblContinue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgContinueHijau.BorderStyle = 0
End Sub

Private Sub lblNew_Click()
    Call imgNewHijau_Click
End Sub

Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNewHijau.BorderStyle = 1
End Sub

Private Sub lblNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Biru
    imgNewBiru.Visible = False
    imgNewHijau.Visible = True
End Sub

Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgNewHijau.BorderStyle = 0
End Sub

Sub BukaLevel()
On Error GoTo Salah
Dim Alamat As String
Dim Kode As String
Dim LevelSave As Integer
    Alamat = App.Path
    If Right(Alamat, 1) <> "\" Then
        Alamat = Alamat & "\"
    End If
    Open Alamat & "data.wt" For Input As #1
    Do While Not EOF(1)
        Input #1, Kode
    Loop
    Close #1
    LevelSave = CariKode(Kode)
    Level = LevelSave + 1
    If Level > 8 Then
        frmUtama.tmrTime.Interval = 1000 - ((8 * 50) + ((Level - 8) * 30))
    Else
        frmUtama.tmrTime.Interval = 1000 - ((Level - 1) * 50)
    End If
Salah:
End Sub

Function CariKode(Nilai As String) As Integer
On Error Resume Next
    Select Case Nilai
    Case "FLC5BT"
        CariKode = 1
    Case "HJ3DF6"
        CariKode = 2
    Case "SI6Y7U"
        CariKode = 3
    Case "A3ERT3"
        CariKode = 4
    Case "8LJK6T"
        CariKode = 5
    Case "TY8GHF"
        CariKode = 6
    Case "IOPN87"
        CariKode = 7
    Case "QW5KL6"
        CariKode = 8
    Case "NF5DS1"
        CariKode = 9
    Case "Z9B3C1"
        CariKode = 10
    Case "FGH45S"
        CariKode = 11
    Case "5AS31H"
        CariKode = 12
    End Select
End Function

Function PeriksaFile() As Boolean
PeriksaFile = True
On Error Resume Next
Dim Object
Dim Alamat As String
    Set Object = CreateObject("scripting.filesystemobject")
    Alamat = App.Path
    If Right(Alamat, 1) <> "\" Then
        Alamat = Alamat & "\"
    End If
    If Not Object.fileexists(Alamat & "data.wt") Then
        PeriksaFile = False
    End If
End Function
