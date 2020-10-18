VERSION 5.00
Begin VB.Form frmUtama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Word Tuts"
   ClientHeight    =   6165
   ClientLeft      =   510
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrJawab 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   5040
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   5400
   End
   Begin VB.TextBox txtJawab 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Image imgListScore 
      Height          =   375
      Left            =   4200
      Picture         =   "frmUtama.frx":17D2A
      Stretch         =   -1  'True
      ToolTipText     =   "List Score"
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblJmlKata 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH KATA :"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4680
      Y1              =   2270
      Y2              =   2270
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   4680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4680
      Y1              =   1550
      Y2              =   1550
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   4680
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "L E V E L :"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblJawab 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   69.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   720
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   4200
      Picture         =   "frmUtama.frx":1C7BA
      Stretch         =   -1  'True
      ToolTipText     =   "Help"
      Top             =   0
      Width           =   495
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4680
      Y1              =   950
      Y2              =   950
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   4680
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2009"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "By Agunahwan Absin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image imgLogo 
      Height          =   735
      Left            =   120
      Picture         =   "frmUtama.frx":217EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "T I M E :"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   4305
      X2              =   4305
      Y1              =   3120
      Y2              =   4800
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4320
      X2              =   4320
      Y1              =   3120
      Y2              =   4800
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   345
      X2              =   345
      Y1              =   3120
      Y2              =   4800
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   360
      X2              =   360
      Y1              =   3120
      Y2              =   4800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   4320
      Y1              =   4785
      Y2              =   4785
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   360
      X2              =   4320
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   4320
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   360
      X2              =   4320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblKata 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE:"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "frmUtama.frx":27B51
      Stretch         =   -1  'True
      Top             =   960
      Width           =   4695
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################################################
'###########    Nama       : Word Tuts          ##########################
'###########    Version    : 1.1                ##########################
'###########    Programmer : Agunahwan Absin    ##########################
'###########    Copyright (c) 2009              ##########################
'#########################################################################

Option Explicit
Dim Waktu, JmlKata
Dim iJawab As Integer
Dim ReKata As String

'Variabel untuk memanggil Database
Dim Nomor(10000) As Long
Dim Word(10000) As String
'Batasan menang=40 kata

Private Sub Form_Load()
'Dim i As Integer
    'Ascii "0"=48 sampai "9"=57
    'Ascii "A"=65 sampai "Z"=90
    'Ascii "a"=97 sampai "z"=122
    'For i = Asc("0") To Asc("9")
    '    lblKata.Caption = lblKata.Caption & " " & CStr(i)
    'Next

Dim i, Cari As Integer
Dim Baris As String
    Open App.Path & "\kata.db" For Input As #1
    Jumlah = LOF(1)
    i = 1
    Do While Not EOF(1)
        Line Input #1, Baris
        Baris = Trim(Baris)
        Cari = InStr(1, Baris, ":")
        Nomor(i) = CLng(Left(Baris, Cari - 1))
        Word(i) = CStr(Right(Baris, Len(Baris) - Cari))
        i = i + 1
    Loop
    Jumlah = i - 2
    'Me.Caption = CStr(Jumlah)
    Close #1
    
    Inisial
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With Me
        .Height = 6615
        .Width = 4815
    End With
End Sub

Private Sub imgHelp_Click()
    tmrTime.Enabled = False
    Bantuan
    tmrTime.Enabled = True
    'frmScore.Show
End Sub

Private Sub imgListScore_Click()
    tmrTime.Enabled = False
    frmListScore.Show vbModal
End Sub

Private Sub tmrJawab_Timer()
    iJawab = iJawab + 1
    If iJawab = 1 Then
        lblJawab.Visible = False
        tmrJawab.Enabled = False
        lblKata.Caption = ReKata
    End If
End Sub

Private Sub tmrTime_Timer()
    If Waktu - 1 >= 0 Then
        Waktu = Waktu - 1
        lblTime.Caption = CStr(Waktu)
    Else
        If JmlKata < 10 Then
            MsgBox "Maaf, anda kalah pada level " & CStr(Level) & " dengan score=" & lblScore.Caption & " dan jumlah kata=" & CStr(JmlKata) & ".", vbOKOnly + vbExclamation, "Kalah"
            frmScore.Show vbModal
            End
        Else
            Simpan CInt(Level)
            MsgBox "Anda menang dengan score=" & lblScore.Caption & " dan jumlah kata=" & CStr(JmlKata) & ".", vbOKOnly + vbInformation, "Selamat"
        End If
        frmBonus.Show vbModal
        tmrTime.Enabled = False
        Level = Level + 1
        If Level > 8 Then
            tmrTime.Interval = tmrTime.Interval - 30
        Else
            tmrTime.Interval = tmrTime.Interval - 50
        End If
        If Level > 12 Then
            Win = True
            frmWin.Show
            Me.Visible = False
        Else
            Inisial
        End If
    End If
End Sub

Private Sub txtJawab_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtJawab.Text) <> 0 Then
        If KeyCode = 13 Then
            txtJawab.Text = LCase(txtJawab.Text)
            Cocok
        Else
            KeyCode = 0
        End If
    End If
End Sub

'================================= Prosedur-Prosedur ==================================================================

Sub Cocok()
    If LCase(txtJawab.Text) = LCase(lblKata.Caption) Then
        lblScore.Caption = CStr(Val(lblScore.Caption) + 10)
        txtJawab.Text = ""
        txtJawab.SetFocus
        Kata
        JmlKata = JmlKata + 1
        lblJmlKata.Caption = CStr(JmlKata)
        If JmlKata < 10 Then
            lblJmlKata.ForeColor = &HFF&
        Else
            lblJmlKata.ForeColor = &HFF0000
        End If
    Else
        Beep
        lblScore.Caption = CStr(Val(lblScore.Caption) - 5)
        txtJawab.Text = ""
        txtJawab.SetFocus
        Kata
        ReKata = lblKata.Caption
        lblKata.Caption = ""
        iJawab = 0
        lblJawab.Visible = True
        tmrJawab.Enabled = True
    End If
End Sub

Sub Kata()
Dim Min, Max, Kode As Integer
'Dim Sesuai As Boolean
    'Sesuai = False
    Max = Jumlah
    Min = 1
    'Do
        Randomize Timer
        Kode = Int((Max - Min + 1) * Rnd) + Min
    '    If (Kode >= 48 And Kode <= 57) Or (Kode >= 65 And Kode <= 90) Or (Kode >= 97 And Kode <= 122) Then
    '        Sesuai = True
    '    End If
    'Loop Until Sesuai = True
    lblKata.Caption = LCase(Word(Kode)) 'Chr(Kode)
End Sub

Sub Bantuan()
    MsgBox "Untuk memainkan Game, lihat kata yang muncul, ketikkan kata pada kotak." & vbNewLine & _
           "Jika kata yang diketikkan telah sesuai, tekan Enter." & vbNewLine & _
           "Jika kata yang Anda masukkan sesuai, score Anda akan bertambah 10." & vbNewLine & _
           "Jika kata yang Anda masukkan tidak sesuai, score Anda akan berkurang 5." & vbNewLine & _
           "Anda diharuskan menjawab benar 10 kata dengan waktu yang ditentukan dan berkurang." & vbNewLine & _
           "Jika Anda telah mencapai 10 atau lebih kata yang benar, Anda akan memasuki level selanjutnya." & vbNewLine & _
           "Semakin tinggi level yang Anda mainkan, waktu yang diberikan berjalan akan semakin cepat." & vbNewLine & _
           "Game ini berakhir hingga level 12.", vbOKOnly + vbInformation, "Cara memainkan"
End Sub

Sub Inisial()
Dim Pesan
    Pesan = MsgBox("Anda memasuki level=" & CStr(Level) & "." & vbNewLine & "Jumlah kata yang harus dicapai untuk memenangkan game adalah 10 kata.", vbOKOnly + vbInformation, "Informasi")
    lblLevel.Caption = CStr(Level)
    If Pesan = vbOK Then
        tmrTime.Enabled = True
    End If
    Kata
    JmlKata = 0
    Waktu = 120
    lblTime.Caption = CStr(Waktu)
    
    lblJmlKata.Caption = CStr(JmlKata)
    If JmlKata < 10 Then
        lblJmlKata.ForeColor = &HFF&
    Else
        lblJmlKata.ForeColor = &HFF0000
    End If
End Sub

Sub Simpan(Data As Integer)
On Error Resume Next
Dim Alamat As String
    Alamat = App.Path
    If Right(Alamat, 1) <> "\" Then
        Alamat = Alamat & "\"
    End If
    If Data < 12 Then
        Open Alamat & "data.wt" For Append As #1
        Write #1, Kode(Data)
        Close #1
    End If
End Sub

Function Kode(Nomor As Integer) As String
    Select Case Nomor
    Case 1
        Kode = "FLC5BT"
    Case 2
        Kode = "HJ3DF6"
    Case 3
        Kode = "SI6Y7U"
    Case 4
        Kode = "A3ERT3"
    Case 5
        Kode = "8LJK6T"
    Case 6
        Kode = "TY8GHF"
    Case 7
        Kode = "IOPN87"
    Case 8
        Kode = "QW5KL6"
    Case 9
        Kode = "NF5DS1"
    Case 10
        Kode = "Z9B3C1"
    Case 11
        Kode = "FGH45S"
    Case 12
        Kode = "5AS31H"
    End Select
End Function
