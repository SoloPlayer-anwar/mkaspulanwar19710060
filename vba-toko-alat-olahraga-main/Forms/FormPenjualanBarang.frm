VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPenjualanBarang 
   Caption         =   "Form Penjualan Barang"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7545
   OleObjectBlob   =   "FormPenjualanBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPenjualanBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range
Dim cariByNamaBarang As Range
Dim tanggalTerjual As Date

Private Sub UserForm_Initialize()
    ComboBoxNamaBarang.List = wsMasterBarang.Range("B2:B" & getBarisMasterBarang).Value
    
    Call bersihForm
    TextBoxIdPenjualanBarang.Text = buatIdPenjualanBarang
End Sub

Private Sub bersihForm()
    TextBoxIdPenjualanBarang.Text = vbNullString
    TextBoxTanggalTerjual.Text = vbNullString
    ComboBoxNamaBarang.Text = vbNullString
    TextBoxJumlahPenjualan.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariPenjualanBarang("A", TextBoxIdPenjualanBarang.Value)
    Set cariByNamaBarang = cariMasterBarang("B", ComboBoxNamaBarang.Value)
    Dim baris As Long
    
    Dim idBarang As String
    
    idBarang = cariByNamaBarang.Offset(0, -1).Value
    tanggalTerjual = TextBoxTanggalTerjual.Value
    
    If cariById Is Nothing Then
        baris = getBarisPenjualanBarang + 1
    Else
        baris = cariById.Row
    End If
        
    ' Rekap Penjualan
    
    Dim isiData As Variant
    isiData = Array(TextBoxIdPenjualanBarang.Value, tanggalTerjual, _
                    idBarang, ComboBoxNamaBarang.Value, _
                    TextBoxJumlahPenjualan.Value)
                                   
    wsPenjualanBarang.Range("A" & baris).Resize(1, 5).Value = isiData
    
    MsgBox "Data Berhasil Disimpan!", vbInformation
    Call bersihForm
    TextBoxIdPenjualanBarang.Text = buatIdPenjualanBarang
    UpdateRekapPenjualan
End Sub

Private Sub CmdBtnLoad_Click()
    Set cariById = cariPenjualanBarang("A", TextBoxIdPenjualanBarang.Value)
    
    If TextBoxIdPenjualanBarang.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "ID Barang Tidak Ditemukan!", vbInformation
        Call bersihForm
        TextBoxIdPenjualanBarang.Text = buatIdPenjualanBarang
    Else
        tanggalTerjual = cariById.Offset(0, 1).Value
        TextBoxIdPenjualanBarang.Text = cariById.Value
        TextBoxTanggalTerjual.Text = tanggalTerjual
        ComboBoxNamaBarang.Text = cariById.Offset(0, 3).Value
        TextBoxJumlahPenjualan.Text = cariById.Offset(0, 4).Value
    End If
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdPenjualanBarang.Text = buatIdPenjualanBarang
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariPenjualanBarang("A", TextBoxIdPenjualanBarang.Value)
    
    If TextBoxIdPenjualanBarang.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "Data ID Barang Tidak Ditemukan!", vbInformation
    Else
        cariById.EntireRow.Delete
        MsgBox "Data Berhasil Di Hapus!", vbInformation
    End If
    
    Call bersihForm
    TextBoxIdPenjualanBarang.Text = buatIdPenjualanBarang
    UpdateRekapPenjualan
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub

