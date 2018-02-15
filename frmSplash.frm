VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4260
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   10500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   3  'Dash-Dot
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4260
   ScaleWidth      =   10500
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar Bar 
      Height          =   210
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   -600
      Top             =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "by Febrian Dwi Putra / FTI-B"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "APLIKASI ABSENSI SEDERHANA"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   6975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8640
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   0
      Top             =   2160
      Width           =   12015
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Bar.Value = Bar.Value + 2
Screen.MousePointer = vbHourglass
Label4.Caption = Bar.Value & " %"
If Bar.Value < 20 Then
ElseIf Bar.Value < 40 Then
ElseIf Bar.Value < 60 Then
ElseIf Bar.Value < 80 Then
ElseIf Bar.Value < 100 Then
End If
If Bar.Value = 100 Then
If Timer1.Interval >= 1 Then
Unload frmSplash
menu.Show
Screen.MousePointer = vbDefault
End If
End If
Exit Sub
End Sub

