VERSION 5.00
Begin VB.Form frmWin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selamat"
   ClientHeight    =   3150
   ClientLeft      =   2805
   ClientTop       =   1650
   ClientWidth     =   5895
   Icon            =   "frmWin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5760
      Y1              =   2390
      Y2              =   2390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   5760
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblUcapan 
      BackStyle       =   0  'Transparent
      Caption         =   "Anda memang orang yang hebat karena telah menyelesaikan Game ini hingga selesai. Terima Kasih Anda telah memainkan Game ini."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S E L A M A T"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   240
      Picture         =   "frmWin.frx":17D2A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   3135
      Left            =   0
      Picture         =   "frmWin.frx":19A2A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    frmScore.Show vbModal
    End
End Sub

Private Sub Form_Load()
    lblUcapan.Caption = "Anda memang orang yang hebat karena telah menyelesaikan Game ini hingga selesai dengan score " & frmUtama.lblScore.Caption & ". Terima Kasih Anda telah memainkan Game ini."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
