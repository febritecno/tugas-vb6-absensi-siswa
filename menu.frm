VERSION 5.00
Begin VB.Form menu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MENU UTAMA ABSENSI"
   ClientHeight    =   7425
   ClientLeft      =   2925
   ClientTop       =   1830
   ClientWidth     =   11970
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menu.frx":0000
   ScaleHeight     =   30.938
   ScaleMode       =   0  'User
   ScaleWidth      =   99.75
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu menu 
      Caption         =   "File"
      Begin VB.Menu login 
         Caption         =   "Login"
         Checked         =   -1  'True
      End
      Begin VB.Menu keluar 
         Caption         =   "Admin Keluar"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu admin 
      Caption         =   "Admin"
      Visible         =   0   'False
      Begin VB.Menu mulai 
         Caption         =   "Mulai Absen Siswa"
      End
      Begin VB.Menu dtsis 
         Caption         =   "Data Siswa"
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu lapor 
         Caption         =   "Laporan Absen Siswa"
      End
      Begin VB.Menu lapor2 
         Caption         =   "Laporan Data Siswa"
      End
   End
   Begin VB.Menu lihat 
      Caption         =   "SHOW ABSENSI"
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dtsis_Click()
Form3.Show
End Sub

Private Sub exit_Click()
pesan = MsgBox("Anda Yakin Ingin Keluar Dari Program ini?", vbQuestion + vbYesNo, "Keluar")
If pesan = vbYes Then
A
Form1.Hide
End
Unload Me
Else
End If
End Sub


Private Sub keluar_Click()
admin.Visible = False
login.Visible = True
lihat.Visible = True
keluar.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
A
End Sub
Public Sub A()
Do
    Me.Top = Me.Top + 40
    Me.Move Me.Left, Me.Top
    DoEvents
    Loop Until Me.Top > Screen.Height - 500
End Sub


Private Sub lapor_Click()
koneksi
RsAbsen.Open "select * from absen", ConN
If Not RsAbsen.EOF Then
Set DataReport1.DataSource = RsAbsen
DataReport1.Show
Else
MsgBox "Tidak ada Data"
End If
End Sub
Private Sub lapor2_Click()
koneksi
RsAbsen.Open "select * from siswa", ConN
If Not RsAbsen.EOF Then
Set DataReport2.DataSource = RsAbsen
DataReport2.Show
Else
MsgBox "Tidak ada Data"
End If
End Sub

Private Sub lihat_Click()
Form4.Show
End Sub

Private Sub login_Click()
Form2.Show
End Sub

Private Sub mulai_Click()
Form1.Show
End Sub
