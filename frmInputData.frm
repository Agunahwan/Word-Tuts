VERSION 5.00
Begin VB.Form frmInputData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data WordTuts"
   ClientHeight    =   6000
   ClientLeft      =   345
   ClientTop       =   675
   ClientWidth     =   4590
   Icon            =   "frmInputData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstData 
      Height          =   2985
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtKata 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   1
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar Kata :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4560
      Y1              =   5150
      Y2              =   5150
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   4560
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   110
      X2              =   110
      Y1              =   960
      Y2              =   4560
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4440
      Y1              =   4550
      Y2              =   4550
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   4430
      X2              =   4430
      Y1              =   960
      Y2              =   4560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4440
      Y1              =   950
      Y2              =   950
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4560
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   4560
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image imgLogo 
      Height          =   735
      Left            =   120
      Picture         =   "frmInputData.frx":17D2A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
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
      TabIndex        =   5
      Top             =   120
      Width           =   2295
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
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kata :"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   855
   End
   Begin VB.Image imgBackGround 
      Height          =   5295
      Left            =   0
      Picture         =   "frmInputData.frx":1E091
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4695
   End
End
Attribute VB_Name = "frmInputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub DaftarKata()
On Error Resume Next
Dim Baris As String
Dim Cari
    Open App.Path & "\kata.db" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Baris
        Baris = Trim(Baris)
        Cari = InStr(1, Baris, ":")
        lstData.AddItem LCase(CStr(Right(Baris, Len(Baris) - Cari)))
    Loop
    Close #1
End Sub

Sub SimpanData()
Dim i As Integer
    Open App.Path & "\kata.db" For Output As #1
    For i = 0 To Me.lstData.ListCount - 1
        Print #1, i & ":" & lstData.List(i)
    Next
    Close #1
End Sub

Function Sama(Kata As String) As Boolean
Dim i As Integer
    Sama = False
    For i = 0 To Me.lstData.ListCount - 1
        If LCase(Kata) = LCase(lstData.List(i)) Then
            Sama = True
        End If
    Next
End Function

Private Sub cmdOK_Click()
    If txtKata.Text <> "" And Sama(txtKata.Text) = False Then
        lstData.AddItem LCase(txtKata.Text)
        SimpanData
    Else
        If Sama(txtKata.Text) = True Then
            MsgBox "Kata yang Anda masukkan telah ada dalam database, silahkan masukkan kata yang lain.", vbOKOnly + vbCritical, "Sama"
        Else
            MsgBox "Kotak masih dalam keadaan kosong, isi dahulu kata yang akan dimasukkan.", vbOKOnly + vbCritical, "Kosong"
        End If
    End If
    txtKata.Text = ""
    txtKata.SetFocus
End Sub

Private Sub Form_Load()
    DaftarKata
End Sub

Private Sub lstData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyDelete Then
        HapusData
    End If
End Sub

Sub HapusData()
On Error Resume Next
Dim a As Long
    For a = 0 To lstData.ListCount - 1
        If lstData.Selected(a) Then
            lstData.RemoveItem a
        End If
    Next
    SimpanData
End Sub

Private Sub txtKata_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtKata.Text) <> 0 Then
        If KeyCode = 13 Then
            cmdOK_Click
        Else
            KeyCode = 0
        End If
    End If
End Sub
