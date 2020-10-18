VERSION 5.00
Begin VB.Form frmListScore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Score"
   ClientHeight    =   4815
   ClientLeft      =   3165
   ClientTop       =   1785
   ClientWidth     =   6150
   Icon            =   "frmListScore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstScore 
      Height          =   3765
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   5
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   4
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6000
      Y1              =   4070
      Y2              =   4070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   6000
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   0
      Picture         =   "frmListScore.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmListScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub DaftarData()
On Error GoTo Salah
Dim DataScore(10000), LevelScore(10000) As Integer
Dim NamaScore(10000) As String
Dim Titik1, Titik2 As Integer
Dim i As Integer
Dim Baris As String
    Open App.Path & "\data.sro" For Input As #1
    
    'Memulai Data
    lstScore.AddItem Space(10) & "Nama" & Space(30) & "Level" & Space(30) & "Score"
    
    'Memasukkan data pada database ke variabel-variabel
    i = 1
    Do While Not EOF(1)
        Line Input #1, Baris
        Baris = Trim(Baris)
        Titik1 = InStr(1, Baris, ":")
        Titik2 = InStr(Titik1 + 1, Baris, ":")
        DataScore(i) = CInt(Mid(Baris, Titik1 + 1, Titik2 - Titik1 - 1))
        LevelScore(i) = CInt(Left(Baris, Titik1 - 1))
        NamaScore(i) = CStr(Right(Baris, Len(Baris) - Titik2))
        i = i + 1
    Loop
    Close #1

    'Proses pengurutan data
    Dim j, k As Integer
    Dim Temp As Integer
    For j = 1 To i - 1
        k = i
        While (k > j)
            If DataScore(k) < DataScore(k - 1) Then
                Temp = DataScore(k)
                DataScore(k) = DataScore(k - 1)
                DataScore(k - 1) = Temp
            End If
            k = k - 1
        Wend
    Next
    
    'Memasukkan data ke dalam List
    Dim l As Integer
    For l = 1 To i - 1
        lstScore.AddItem Space(10) & NamaScore(l) & Space(35) & LevelScore(l) & Space(35) & DataScore(l)
    Next
Salah:
End Sub

Private Sub cmdOK_Click()
    Unload Me
    frmUtama.tmrTime.Enabled = True
End Sub

Private Sub Form_Load()
    DaftarData
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmUtama.tmrTime.Enabled = True
End Sub
