VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMasterBarang 
   Caption         =   "Form Master Barang"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "FormMasterBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMasterBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range
Dim cariByIdMerekBarang As Range
Dim cariByIdKategoriBarang As Range

Private Sub UserForm_Initialize()
    ComboBoxMerekBarang.List = wsMerekBarang.Range("B2:B" & getBarisMerekBarang).Value
    ComboBoxKategoriBarang.List = wsKategoriBarang.Range("B2:B" & getBarisKategoriBarang).Value
    
    Call bersihForm
    TextBoxIdBarang.Text = buatIdMasterBarang
End Sub

Private Sub bersihForm()
    TextBoxIdBarang.Text = vbNullString
    TextBoxNamaBarang.Text = vbNullString
    ComboBoxMerekBarang.Text = vbNullString
    ComboBoxKategoriBarang.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariMasterBarang("A", TextBoxIdBarang.Value)
    Set cariByIdMerekBarang = cariMerekBarang("B", ComboBoxMerekBarang.Value)
    Set cariByIdKategoriBarang = cariKategoriBarang("B", ComboBoxKategoriBarang.Value)
    
    Dim idMerekBarang As String
    Dim idKategoriBarang As String
    
    idMerekBarang = cariByIdMerekBarang.Offset(0, -1).Value
    idKategoriBarang = cariByIdKategoriBarang.Offset(0, -1).Value
    
    Dim baris As Long
    
    If cariById Is Nothing Then
        baris = getBarisMasterBarang + 1
    Else
        baris = cariById.Row
    End If
    
    Dim isiData As Variant
    isiData = Array(TextBoxIdBarang.Value, TextBoxNamaBarang.Value, _
                    idMerekBarang, ComboBoxMerekBarang.Value, _
                    idKategoriBarang, ComboBoxKategoriBarang.Value)
                    
    wsMasterBarang.Range("A" & baris).Resize(1, 6).Value = isiData
    MsgBox "Data Berhasil Disimpan!", vbInformation
    Call bersihForm
    TextBoxIdBarang.Text = buatIdMasterBarang
End Sub

Private Sub CmdBtnLoad_Click()
    Set cariById = cariMasterBarang("A", TextBoxIdBarang.Value)
    
    If TextBoxIdBarang.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "ID Barang Tidak Ditemukan!", vbInformation
        Call bersihForm
        TextBoxIdBarang.Text = buatIdMasterBarang
    Else
        TextBoxIdBarang.Text = cariById.Value
        TextBoxNamaBarang.Text = cariById.Offset(0, 1).Value
        ComboBoxMerekBarang.Text = cariById.Offset(0, 3).Value
        ComboBoxKategoriBarang.Text = cariById.Offset(0, 5).Value
    End If
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdBarang.Text = buatIdMasterBarang
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariMasterBarang("A", TextBoxIdBarang.Value)
    
    If TextBoxIdBarang.Text = "" Then
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
    TextBoxIdBarang.Text = buatIdMasterBarang
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub

