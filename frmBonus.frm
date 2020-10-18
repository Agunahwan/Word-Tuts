VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBonus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bonus Gambar"
   ClientHeight    =   5625
   ClientLeft      =   2670
   ClientTop       =   1605
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "frmBonus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgSimpan 
      Left            =   4560
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   2490
      TabIndex        =   0
      Top             =   5055
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selamat Anda memenangkan level ini, Anda dapat menyimpan Gambar di bawah ini sebagai bonus."
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   6375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6360
      Y1              =   4905
      Y2              =   4905
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   6360
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Image imgBonus 
      BorderStyle     =   1  'Fixed Single
      Height          =   4095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   720
      Width           =   6375
   End
   Begin VB.Image imgBackGround 
      Height          =   5775
      Left            =   0
      Picture         =   "frmBonus.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    SimpanGambar
    Unload Me
End Sub

Sub SimpanGambar()
On Error Resume Next
Dim Objek, Pesan
Set Objek = CreateObject("scripting.filesystemobject")
    With Me.cdgSimpan
        .DialogTitle = "Simpan Gambar"
        .CancelError = True
        .InitDir = App.Path
        .Filter = "File JPEG(*.jpeg;*.jpg)|*.jpeg;*.jpg|"
        .ShowSave
        If Len(.FileName) <> 0 Then
            If Objek.fileexists(CStr(.FileName)) Then
                Pesan = MsgBox(" File " & " telah ada. Apa ingin ditimpa ?", vbQuestion + vbYesNo, "Timpa")
                If Pesan = vbNo Then
                    Exit Sub
                End If
            End If
            SavePicture imgBonus.Picture, .FileName
        End If
    End With
End Sub

Sub PanggilGambar()
On Error GoTo Salah
'Dim Min, Max, Kode As Integer
    'Max = 25
    'Min = 0
    'Randomize Timer
    'Kode = Int((Max - Min + 1) * Rnd) + Min
    imgBonus.Picture = frmGambar.imgGambar(Level).Picture
Salah:
End Sub

Private Sub Form_Load()
    PanggilGambar
End Sub
