Attribute VB_Name = "mdlUtama"
Option Explicit

Public Level, Jumlah As Integer
Public Win As Boolean

Sub Main()
    Periksa
    
    Level = 1
    If App.PrevInstance Then End
    frmUtama.Bantuan
    frmUtama.Show
End Sub

Public Sub Periksa()
Dim Objek
Set Objek = CreateObject("scripting.filesystemobject")
    
    'Memeriksa apakah database ada
    If Not Objek.fileexists(App.Path & "\kata.db") Then
        MsgBox "Database tidak ditemukan", vbOKOnly + vbExclamation, "Error"
        End 'Jika database tidak ada, maka program keluar
    End If
End Sub

Public Sub SimpanScore()
On Error GoTo Salah
Dim Baris As String
On Error Resume Next
    With frmScore
        Baris = .lblLevel.Caption & ":" & .lblScore.Caption & ":" & .txtNama.Text
    End With
    Open App.Path & "\data.sro" For Append As #1
        Print #1, Baris
    Close #1
Salah:
End Sub

'Public Type DATABASE
'    Nomor As Long
'    Word As String
'End Type

