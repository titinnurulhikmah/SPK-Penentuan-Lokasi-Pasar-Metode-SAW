VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form3 (BOBOT KRITERIA)"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13350
   LinkTopic       =   "Form3"
   ScaleHeight     =   6945
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1560
      Top             =   6120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=konekdblokasipasir"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "konekdblokasipasir"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbl_bobotkriteria"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   2175
      Left            =   1560
      TabIndex        =   9
      Top             =   3840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "keluar"
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "simpan"
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total Bobot"
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Dagangan"
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Jarak Tempuh"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ukuran Kios"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Kd Bobot"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BOBOT KRITERIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   2550
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If Text1.Text = "" Then
    MsgBox "Kd Bobot kosong", vbExclamation, "pesan"
    Text1.SetFocus
    Exit Sub
    End If
    If Text2.Text = "" Then
    MsgBox "Ukuran Kios kosong", vbExclamation, "pesan"
    Text2.SetFocus
    Exit Sub
    End If
    If Text3.Text = "" Then
    MsgBox "Jarak Tempuh  kosong", vbExclamation, "pesan"
    Text3.SetFocus
    Exit Sub
    End If
    If Text4.Text = "" Then
    MsgBox "Jenis dagangan kosong", vbExclamation, "pesan"
    Text2.SetFocus
    Exit Sub
    End If
Set lahanpasar = New ADODB.Recordset
lahanpasar.Open "select * from tbl_bobotkriteria where kd_bobot='" & Text1.Text & "'", konekdblokasipasir
If Not lahanpasar.EOF Then
MsgBox "kode Bobot sudah digunakan", vbCritical, "pesan"
Text1.Text = ""
Text1.SetFocus
Exit Sub
Else
konekdblokasipasir.Execute "insert into tbl_bobotkriteria(kd_bobot,c1,c2,c3) value ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "')"
MsgBox "DATA TERSIMPAN"
Call segar
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

