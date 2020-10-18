VERSION 5.00
Begin VB.Form frmScore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Score"
   ClientHeight    =   2055
   ClientLeft      =   2895
   ClientTop       =   1740
   ClientWidth     =   6015
   ControlBox      =   0   'False
   Icon            =   "frmScore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNama 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   4575
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
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblLevel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5880
      Y1              =   1310
      Y2              =   1310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   5880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Level            :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Score           :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Nama  :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image imgBackGround 
      Height          =   2055
      Left            =   0
      Picture         =   "frmScore.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Inisial()
On Error Resume Next
    lblScore.Caption = frmUtama.lblScore.Caption
    If Win = False Then
        lblLevel.Caption = Level
    Else
        lblLevel.Caption = "12"
    End If
    txtNama.SetFocus
End Sub

Private Sub cmdOK_Click()
    If txtNama.Text <> "" Then
        SimpanScore
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Inisial
End Sub

Sub ProsesSimpan()
On Error GoTo Salah
Dim DataScore(10000), LevelScore(10000) As Integer
Dim NamaScore(10000) As String
Dim Titik1, Titik2 As Integer
Dim i, j As Integer
Dim Awal, Akhir As Integer
Dim Baris As String
    Open App.Path & "\data.sro" For Input As #1
    
    'Memasukkan data pada database ke variabel-variabel
    i = 1
    Do While Not EOF(1)
        Line Input #1, Baris
        Baris = Trim(Baris)
        Titik1 = InStr(1, Baris, ":")
        Titik2 = InStr(Titik1 + 1, Baris, ":")
        DataScore(i) = Val(Mid(Baris, Titik1 + 1, Titik2 - Titik1 - 1))
        LevelScore(i) = Val(Left(Baris, Titik1 - 1))
        NamaScore(i) = CStr(Right(Baris, Len(Baris) - Titik2))
        i = i + 1
    Loop
    Close #1
    
    Akhir = 0
    Awal = 0
    If i > 1 Then
        For j = 1 To i - 1
            If DataScore(j) <= Val(lblScore.Caption) Then
                Awal = j
                If j - 1 > 0 Then
                    Akhir = j - 1
                End If
                Exit For
            End If
        Next
        
        'Proses penyimpanan ke dalam file
        Dim Satu, Dua As Integer
        On Error Resume Next
            With Me
                Baris = .lblLevel.Caption & ":" & .lblScore.Caption & ":" & .txtNama.Text
            End With
            Open App.Path & "\data.sro" For Append As #1
                
                For Satu = 1 To Akhir
                    Print #1, LevelScore(Satu) & ":" & DataScore(Satu) & ":" & NamaScore(Satu) & vbNewLine
                Next
                
                Print #1, Baris & vbNewLine
                
                For Dua = Awal To i - 1
                    Print #1, LevelScore(Dua) & ":" & DataScore(Dua) & ":" & NamaScore(Dua) & vbNewLine
                Next
            
            Close #1
    Else
        SimpanScore
    End If
Salah:
End Sub
