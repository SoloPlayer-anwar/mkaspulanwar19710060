VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormBarangMasuk 
   Caption         =   "Form Barang Masuk"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "FormBarangMasuk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormBarangMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range
Dim cariByIdMasterBarang As Range
Dim tanggalMasuk As Date

Private Sub UserForm_Initialize()
   ComboBoxNamaBarang.List = wsMasterBarang.Range("B2:B" & getBarisMasterBarang).Value

   Call bersihForm
   TextBoxIdBarangMasuk.Text = buatIdBarangMasuk
End Sub

Private Sub bersihForm()
   TextBoxIdBarangMasuk.Text = vbNullString
   TextBoxTanggalMasuk.Text = vbNullString
   ComboBoxNamaBarang.Text = vbNullString
   TextBoxJumlahMasuk.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariBarangMasuk("A", TextBoxIdBarangMasuk.Value)
    Set cariByIdMasterBarang = cariMasterBarang("B", ComboBoxNamaBarang.Value)
    
    Dim idBarang As String
    
    idBarang = cariByIdMasterBarang.Offset(0, -1).Value
    tanggalMasuk = TextBoxTanggalMasuk.Value
    
    Dim baris As Long
    
    If cariById Is Nothing Then
        baris = getBarisBarangMasuk + 1
    Else
        baris = cariById.Row
    End If
    
    Dim isiData As Variant
    isiData = Array(TextBoxIdBarangMasuk.Value, tanggalMasuk, _
                    idBarang, ComboBoxNamaBarang.Value, _
                    TextBoxJumlahMasuk.Value)
                                
    wsBarangMasuk.Range("A" & baris).Resize(1, 5).Value = isiData
    MsgBox "Data Berhasil Disimpan!", vbInformation
    Call bersihForm
    TextBoxIdBarangMasuk.Text = buatIdBarangMasuk
End Sub

Private Sub CmdBtnLoad_Click()
    Set cariById = cariBarangMasuk("A", TextBoxIdBarangMasuk.Value)
    
    If TextBoxIdBarangMasuk.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbInformation
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "ID Barang Tidak Ditemukan!", vbInformation
        Call bersihForm
        TextBoxIdBarangMasuk.Text = buatIdBarangMasuk
    Else
        tanggalMasuk = cariById.Offset(0, 1).Value
        TextBoxIdBarangMasuk.Text = cariById.Value
        TextBoxTanggalMasuk.Text = tanggalMasuk
        ComboBoxNamaBarang.Text = cariById.Offset(0, 3).Value
        TextBoxJumlahMasuk.Text = cariById.Offset(0, 4).Value
    End If
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdBarangMasuk.Text = buatIdBarangMasuk
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariBarangMasuk("A", TextBoxIdBarangMasuk.Value)
    
    If TextBoxIdBarangMasuk.Text = "" Then
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
    TextBoxIdBarangMasuk.Text = buatIdBarangMasuk
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub

