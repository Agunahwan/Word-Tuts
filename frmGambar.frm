VERSION 5.00
Begin VB.Form frmGambar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4065
   ClientLeft      =   2685
   ClientTop       =   1620
   ClientWidth     =   6120
   Icon            =   "frmGambar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   12
      Left            =   3720
      Picture         =   "frmGambar.frx":000C
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   11
      Left            =   1320
      Picture         =   "frmGambar.frx":2E2C7
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   10
      Left            =   4920
      Picture         =   "frmGambar.frx":6CF35
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   9
      Left            =   3720
      Picture         =   "frmGambar.frx":7BEF6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   8
      Left            =   2520
      Picture         =   "frmGambar.frx":107A23
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   7
      Left            =   1320
      Picture         =   "frmGambar.frx":1346CA
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   6
      Left            =   120
      Picture         =   "frmGambar.frx":13D64F
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   5
      Left            =   4920
      Picture         =   "frmGambar.frx":151E64
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   4
      Left            =   3720
      Picture         =   "frmGambar.frx":164A30
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   3
      Left            =   2520
      Picture         =   "frmGambar.frx":17AA1A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   2
      Left            =   1320
      Picture         =   "frmGambar.frx":18CE6C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgGambar 
      Height          =   1215
      Index           =   1
      Left            =   120
      Picture         =   "frmGambar.frx":1BA5D9
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmGambar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

