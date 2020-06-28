Attribute VB_Name = "Module1"
Option Explicit
Public konekdblokasipasir As New ADODB.Connection
Sub bukadb()
    Set konekdblokasipasir = New ADODB.Connection
    konekdblokasipasir.CursorLocation = adUseClient
    konekdblokasipasir.ConnectionString = "driver={mysql odbc 3.51 driver};server=localhost;database=lahanpasar;uid=root;option"
    On Error GoTo pesan
    If konekdblokasipasir.State = adStateClosed Then konekdblokasipasir.Open
Exit Sub
pesan:
 MsgBox "Maaf ! Tidak Bisa Terkoneksi KeDatabase", vbInformation, "Pesan"
    End
End Sub
