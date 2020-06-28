VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form1 (PEDAGANG)"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17205
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   17205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "keluar"
      Height          =   375
      Left            =   9240
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   2160
      TabIndex        =   19
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   3120
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   480
      Top             =   7680
      Width           =   14535
      _ExtentX        =   25638
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
      RecordSource    =   "tbl_pedagang"
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
      Left            =   7080
      TabIndex        =   14
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "batal"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hapus"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "simpan"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2175
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   14775
      _ExtentX        =   26061
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2160
      TabIndex        =   7
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "cari"
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
      Left            =   360
      TabIndex        =   18
      Top             =   4800
      Width           =   330
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "DATA PEDAGANG"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   2640
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "no. hp"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ukuran kios"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "jenis dagangan"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "nama pedagang"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "tanggal registrasi"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   1185
   End
   Begin VB.Label i 
      AutoSize        =   -1  'True
      Caption         =   "id pedagang"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lahanpasar As New ADODB.Recordset

Private Sub Command3_Click()
Call kosong
End Sub

Private Sub Command4_Click()
konekdblokasipasir.Execute "update tbl_pedagang set nm_pedagang ='" & Text2.Text & "',tgl_registrasi='" & Text3.Text & "',jns_dagangan='" & Text4.Text & "',ukuran_kios='" & Text5.Text & "',no_hp='" & Text6.Text & "' where id_pedagang='" & Text1.Text & "'"
MsgBox "data berhasil di edit", vbInformation, "pesan"
Call segar
Call kosong
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Dim hapus As String
hapus = MsgBox("yakin akan menghapus data ini", vbYesNo, "pesan")
If hapus = vbYes Then
konekdblokasipasir.Execute "delete from tbl_pedagang where id_pedagang='" & Text1.Text & "'"
Call segar
Call kosong
Text1.SetFocus
MsgBox "data telah dihapus", vbExclamation, "pesan"
End If
End Sub

Private Sub Command1_Click()
    If Text1.Text = "" Then
    MsgBox "Id pedagang kosong", vbExclamation, "pesan"
    Text1.SetFocus
    Exit Sub
    End If
    If Text2.Text = "" Then
    MsgBox "nama pedagang kosong", vbExclamation, "pesan"
    Text2.SetFocus
    Exit Sub
    End If
    If Text3.Text = "" Then
    MsgBox "tanggal registrasi  kosong", vbExclamation, "pesan"
    Text3.SetFocus
    Exit Sub
    End If
    If Text4.Text = "" Then
    MsgBox "jenis dagangan  kosong", vbExclamation, "pesan"
    Text4.SetFocus
    Exit Sub
    End If
    If Text5.Text = "" Then
    MsgBox "ukuran kios  kosong", vbExclamation, "pesan"
    Text5.SetFocus
    Exit Sub
    End If
    If Text6.Text = "" Then
    MsgBox "no hp  kosong", vbExclamation, "pesan"
    Text6.SetFocus
    Exit Sub
    End If
Set lahanpasar = New ADODB.Recordset
lahanpasar.Open "select * from tbl_pedagang where id_pedagang='" & Text1.Text & "'", konekdblokasipasir
If Not lahanpasar.EOF Then
MsgBox "id pedagang sudah digunakan", vbCritical, "pesan"
Text1.Text = ""
Text1.SetFocus
Exit Sub
Else
konekdblokasipasir.Execute "insert into tbl_pedagang(id_pedagang,nm_pedagang,tgl_registrasi,jns_dagangan,ukuran_kios,no_hp) value ('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "')"
MsgBox "DATA TERSIMPAN"
Call segar
Text1.SetFocus
End If

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
Text1.Text = lahanpasar!id_pedagang
Text2.Text = lahanpasar!nm_pedagang
Text3.Text = lahanpasar!tgl_registrasi
Text4.Text = lahanpasar!jns_dagangan
Text5.Text = lahanpasar!ukuran_kios
Text6.Text = lahanpasar!no_hp
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
lahanpasar.Source = "select * from tbl_pedagang"
lahanpasar.Open
End Sub
Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "id pedagang"
    .Columns(1).Caption = "nama pedagang"
    .Columns(2).Caption = "tanggal registrasi"
    .Columns(3).Caption = "jenis dagangan"
    .Columns(4).Caption = "ukuran kios"
    .Columns(5).Caption = "no. hp"
    .Columns(0).Width = 2000
    .Columns(1).Width = 2000
    .Columns(2).Width = 3000
    .Columns(3).Width = 3000
    .Columns(4).Width = 3000
    .Columns(5).Width = 3000
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
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub
Private Sub text7_Change()
Set lahanpasar = New ADODB.Recordset
lahanpasar.Open "select * from tbl_pedagang where id_pedagang like '%" & Text7.Text & "%'", konekdblokasipasir
If Not lahanpasar.EOF Then
Set DataGrid1.DataSource = lahanpasar
Call edit_grid
End If
End Sub
Private Sub text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lahanpasar.State = adStateOpen Then lahanpasar.Close
lahanpasar.Open "select * from tbl_pedagang where id_pedagang like '%" & Text7.Text & "%'", konekdblokasipasir
If Not lahanpasar.EOF Then
Text1.Text = lahanpasar!id_pedagang
Text2.Text = lahanpasar!nm_pedagang
Text3.Text = lahanpasar!tgl_registrasi
Text4.Text = lahanpasar!jns_dagangan
Text5.Text = lahanpasar!ukuran_kios
Text6.Text = lahanpasar!no_hp

Call segar
End If
End If
End Sub
