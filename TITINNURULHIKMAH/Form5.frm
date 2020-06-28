VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFC0FF&
   Caption         =   "FORM UTAMA"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   FillColor       =   &H00FF80FF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form5"
   ScaleHeight     =   3030
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   11280
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PENILAIAN"
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BOBOT KRITERIA"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DATA LOKASI"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   13335
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF80FF&
         Caption         =   "DATA PEDAGANG"
         Height          =   375
         Left            =   240
         MaskColor       =   &H00FF80FF&
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Caption         =   "SISTEM PENUNJANG KEPUTUSAN MENENTUKAN LOKASI PASAR UNTUK  PEDAGANG PADA KANTOR PELAYANAN PASAR KOTA DUMAI MENGUNAKAN METODE SAW"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   9405
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
