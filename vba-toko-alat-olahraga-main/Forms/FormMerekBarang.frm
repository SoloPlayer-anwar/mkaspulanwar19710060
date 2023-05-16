VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMerekBarang 
   Caption         =   "Form Merek Barang"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "FormMerekBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMerekBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cariById As Range

Private Sub UserForm_Initialize()
    TextBoxIdMerekBarang.Text = buatIdMerekBarang
End Sub

Private Sub bersihForm()
    TextBoxIdMerekBarang.Text = vbNullString
    TextBoxMerekBarang.Text = vbNullString
End Sub

Private Sub CmdBtnSimpan_Click()
    Set cariById = cariMerekBarang("A", TextBoxIdMerekBarang.Value)
    Dim baris As Long
    
    If cariById Is Nothing Then
        baris = getBarisMerekBarang + 1
    Else
        baris = cariById.Row
    End If
    
    Dim isiData As Variant
    isiData = Array(TextBoxIdMerekBarang.Value, TextBoxMerekBarang.Value)
    
    wsMerekBarang.Range("A" & baris).Resize(1, 2).Value = isiData
    MsgBox "Data berhasil disimpan!", vbInformation
    Call bersihForm
    TextBoxIdMerekBarang.Text = buatIdMerekBarang
End Sub

Private Sub CmdBtnLoad_Click()
    Set cariById = cariMerekBarang("A", TextBoxIdMerekBarang.Value)

    If TextBoxIdMerekBarang.Text = "" Then
        MsgBox "Silahkan Isi ID Barang!", vbCritical
        Exit Sub
    End If
    
    If cariById Is Nothing Then
        MsgBox "ID Barang Tidak Ditemukan!", vbInformation
        Call bersihForm
        TextBoxIdMerekBarang.Text = buatIdMerekBarang
    Else
        TextBoxIdMerekBarang.Text = cariById.Value
        TextBoxMerekBarang.Text = cariById.Offset(0, 1).Value
    End If
End Sub

Private Sub CmdBtnBatal_Click()
    Call bersihForm
    TextBoxIdMerekBarang.Text = buatIdMerekBarang
End Sub

Private Sub CmdBtnHapus_Click()
    Set cariById = cariMerekBarang("A", TextBoxIdMerekBarang.Value)

    If TextBoxIdMerekBarang.Text = "" Then
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
    TextBoxIdMerekBarang.Text = buatIdMerekBarang
End Sub

Private Sub CmdBtnKeluar_Click()
    Unload Me
End Sub


