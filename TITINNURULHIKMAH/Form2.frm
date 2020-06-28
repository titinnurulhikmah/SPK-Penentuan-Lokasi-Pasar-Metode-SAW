VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form2 (LOKASI)"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   LinkTopic       =   "Form2"
   ScaleHeight     =   7110
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "keluar"
      Height          =   495
      Left            =   8760
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2880
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   600
      Top             =   6360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1085
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
      RecordSource    =   "tbl_lokasi"
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
   Begin VB.CommandButton Command4 
      Caption         =   "edit"
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "batal"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton command2 
      Caption         =   "hapus"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "simpan"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   2535
      Left            =   480
      TabIndex        =   7
      Top             =   3600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   14
      Top             =   3000
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "alamat pasar"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nama pasar"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Kode lokasi"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATA LOKASI"
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
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   2010
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lahanpasar As New ADODB.Recordset

Private Sub Command3_Click()
Call kosong
End Sub

Private Sub Command4_Click()
konekdblokasipasir.Execute "update tbl_lokasi set nm_pasar ='" & Text2.Text & "',alamat='" & Text3.Text & "' where kd_lokasi='" & Text1.Text & "'"
MsgBox "data berhasil di edit", vbInformation, "pesan"
Call segar
Call kosong
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Dim hapus As String
hapus = MsgBox("yakin akan menghapus data ini", vbYesNo, "pesan")
If hapus = vbYes Then
konekdblokasipasir.Execute "delete from tbl_lokasi where kd_lokasi='" & Text1.Text & "'"
Call segar
Call kosong
Text1.SetFocus
MsgBox "data telah dihapus", vbExclamation, "pesan"
End If
End Sub

Private Sub Command1_Click()
    If Text1.Text = "" Then
    MsgBox "kode lokasi kosong", vbExclamation, "pesan"
    Text1.SetFocus
    Exit Sub
    End If
    If Text2.Text = "" Then
    MsgBox "nama pasar kosong", vbExclamation, "pesan"
    Text2.SetFocus
    Exit Sub
    End If
    If Text3.Text = "" Then
    MsgBox "alamat  kosong", vbExclamation, "pesan"
    Text3.SetFocus
    Exit Sub
    End If
Set lahanpasar = New ADODB.Recordset
lahanpasar.Open "select * from tbl_lokasi where kd_lokasi='" & Text1.Text & "'", konekdblokasipasir
If Not lahanpasar.EOF Then
MsgBox "kode lokasi sudah digunakan", vbCritical, "pesan"
Text1.Text = ""
Text1.SetFocus
Exit Sub
Else
konekdblokasipasir.Execute "insert into tbl_lokasi(kd_lokasi,nm_pasar,alamat) value ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "')"
MsgBox "DATA TERSIMPAN"
Call segar
Text1.SetFocus
End If

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
Text1.Text = lahanpasar!kd_lokasi
Text2.Text = lahanpasar!nm_pasar
Text3.Text = lahanpasar!alamat
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = lahanpasar
With lahanpasar
End With
Call edit_grid
End Sub
Sub tampil_data()
Set lahanpasar = New ADODB.Recordset
lahanpasar.ActiveConnection = konekdblokasipasir
lahanpasar.CursorLocation = adUseClient
lahanpasar.LockType = adLockOptimistic
lahanpasar.Source = "select * from tbl_lokasi"
lahanpasar.Open
End Sub
Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "Kode lokasi"
    .Columns(1).Caption = "Nama pasar"
    .Columns(2).Caption = "alamat pasar"
    .Columns(0).Width = 2000
    .Columns(1).Width = 2000
    .Columns(2).Width = 3000
End With
End Sub
Sub segar()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = lahanpasar
With DataGrid1
Call edit_grid
End With
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub text4_Change()
Set lahanpasar = New ADODB.Recordset
lahanpasar.Open "select * from tbl_lokasi where kd_lokasi like '%" & Text4.Text & "%'", konekdblokasipasir
If Not lahanpasar.EOF Then
Set DataGrid1.DataSource = lahanpasar
Call edit_grid
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lahanpasar.State = adStateOpen Then lahanpasar.Close
lahanpasar.Open "select * from tbl_lokasi where kd_lokasi like '%" & Text4.Text & "%'", konekdblokasipasir
If Not lahanpasar.EOF Then
Text1.Text = lahanpasar!kd_lokasi
Text2.Text = lahanpasar!nm_pasar
Text3.Text = lahanpasar!alamat

Call segar
End If
End If
End Sub

