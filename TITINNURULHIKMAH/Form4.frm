VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form4 (PENILAIAN)"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   LinkTopic       =   "Form4"
   ScaleHeight     =   7695
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   2880
      TabIndex        =   17
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "keluar"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "hapus"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "edit"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "simpan"
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   600
      Top             =   7200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   873
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
      RecordSource    =   "tbl_penilaian"
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
      Bindings        =   "Form4.frx":0000
      Height          =   2415
      Left            =   600
      TabIndex        =   12
      Top             =   4680
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4260
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
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   2880
      TabIndex        =   9
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "kode bobot"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   4080
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "kode lokasi"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "kode pedagang"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "status"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "tanggal permohonan"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nomor permohonan"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PENILAIAN"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   360
      Width           =   1605
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub
