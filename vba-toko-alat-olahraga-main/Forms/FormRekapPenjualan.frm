VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormRekapPenjualan 
   Caption         =   "Form Rekap Penjualan"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   OleObjectBlob   =   "FormRekapPenjualan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormRekapPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range
Dim cariByNamaBarang As Range

Private Sub UserForm_Initialize()
    ComboBoxNamaBarang.List = wsPenjualanBarang.Range("D2:D" & getBarisPenjualanBarang).Value
    
    Call bersihForm
    TextBoxIdRekapPenjualan.Text = buatIdRekapPenjualan
End Sub

Private Sub bersihForm()
    TextBoxIdRekapPenjualan.Text = vbNullString
    ComboBoxNamaBarang.Text = vbNullString
    TextBoxBulan.Text = vbNullString
    TextBoxTahun.Text = vbNullString
    TextBoxJumlahPenjualan.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariRekapPenjualan("A", TextBoxIdRekapPenjualan.Value)
    Set cariByNamaBarang = cariPenjualanBarang("D", ComboBoxNamaBarang.Value)
    Dim baris As Long
    Dim idBarang As String
    
    idBarang = cariByNamaBarang.Offset(0, -1).Value
    
    If cariById Is Nothing Then
        baris = getBarisRekapPenjualan + 1
    Else
        baris = cariById.Row
    End If
    
    Dim isiData As Variant
    isiData = Array(TextBoxIdRekapPenjualan.Value, idBarang, _
                    ComboBoxNamaBarang.Value, TextBoxBulan.Value, _
                    TextBoxTahun.Value, TextBoxJumlahPenjualan.Value)
                    
    wsRekapPenjualan.Range("A" & baris).Resize(1, 6).Value = isiData
    
    MsgBox "Data Berhasil Disimpan!", vbInformation
    Call bersihForm
    TextBoxIdRekapPenjualan.Text = buatIdRekapPenjualan
End Sub

Private Sub CmdBtnLoad_Click()
    Set cariById = cariRekapPenjualan("A", TextBoxIdRekapPenjualan.Value)
    
    If TextBoxIdRekapPenjualan.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "ID Barang Tidak Ditemukan!", vbInformation
        Call bersihForm
        TextBoxIdRekapPenjualan.Text = buatIdRekapPenjualan
    Else
        TextBoxIdRekapPenjualan.Text = cariById.Value
        ComboBoxNamaBarang.Text = cariById.Offset(0, 2).Value
        TextBoxBulan.Text = cariById.Offset(0, 3).Value
        TextBoxTahun.Text = cariById.Offset(0, 4).Value
        TextBoxJumlahPenjualan.Text = cariById.Offset(0, 5).Value
    End If
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdRekapPenjualan.Text = buatIdRekapPenjualan
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariRekapPenjualan("A", TextBoxIdRekapPenjualan.Value)
    
    If TextBoxIdRekapPenjualan.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "ID Barang Tidak Ditemukan", vbInformation
    Else
        cariById.EntireRow.Delete
        MsgBox "Data Berhasil Di Hapus!", vbInformation
    End If
    
    Call bersihForm
    TextBoxIdRekapPenjualan.Text = buatIdRekapPenjualan
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub
