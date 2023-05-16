VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormKategoriBarang 
   Caption         =   "Form Kategori Barang"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   OleObjectBlob   =   "FormKategoriBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormKategoriBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range

Private Sub UserForm_Initialize()
    TextBoxIdKategoriBarang.Text = buatIdKategoriBarang
End Sub

Private Sub bersihForm()
    TextBoxIdKategoriBarang.Text = vbNullString
    TextBoxKategoriBarang.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariKategoriBarang("A", TextBoxIdKategoriBarang.Value)
    
    Dim baris As Long
    
    If cariById Is Nothing Then
        baris = getBarisKategoriBarang + 1
    Else
        baris = cariById.Row
    End If
    
    Dim isiData As Variant
    isiData = Array(TextBoxIdKategoriBarang.Value, TextBoxKategoriBarang.Value)
    
    wsKategoriBarang.Range("A" & baris).Resize(1, 2).Value = isiData
    MsgBox "Data Berhasil Disimpan!", vbInformation
    Call bersihForm
    TextBoxIdKategoriBarang.Text = buatIdKategoriBarang
End Sub

Private Sub CmdBtnLoad_Click()
    Set cariById = cariKategoriBarang("A", TextBoxIdKategoriBarang.Value)
    
    If TextBoxIdKategoriBarang.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "ID Barang Tidak Ditemukan!", vbInformation
        Call bersihForm
        TextBoxIdKategoriBarang.Text = buatIdKategoriBarang
    Else
        TextBoxIdKategoriBarang.Text = cariById.Value
        TextBoxKategoriBarang.Text = cariById.Offset(0, 1).Value
    End If
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdKategoriBarang.Text = buatIdKategoriBarang
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariKategoriBarang("A", TextBoxIdKategoriBarang.Value)
    
    If TextBoxIdKategoriBarang.Text = "" Then
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
    TextBoxIdKategoriBarang.Text = buatIdKategoriBarang
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub
