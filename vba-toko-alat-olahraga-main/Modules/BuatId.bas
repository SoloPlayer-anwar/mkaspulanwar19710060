Attribute VB_Name = "BuatId"
Option Explicit

Public Function buatIdMerekBarang() As String
    Dim baris As Long
    Dim idLama As String

    baris = getBarisMerekBarang

    If Not baris = 1 Then
        idLama = wsMerekBarang.Range("A" & baris)
        baris = CLng(Right(idLama, 4)) + 1
    End If

    buatIdMerekBarang = "IDMB" & Format(baris, "0000")
End Function

Public Function buatIdKategoriBarang() As String
    Dim baris As Long
    Dim idLama As String
    
    baris = getBarisKategoriBarang
    
    If Not baris = 1 Then
        idLama = wsKategoriBarang.Range("A" & baris)
        baris = CLng(Right(idLama, 4)) + 1
    End If
    
    buatIdKategoriBarang = "IDKB" & Format(baris, "0000")
End Function

Public Function buatIdMasterBarang() As String
    Dim baris As Long
    Dim idLama As String
    
    baris = getBarisMasterBarang
    
    If Not baris = 1 Then
        idLama = wsMasterBarang.Range("A" & baris)
        baris = CLng(Right(idLama, 4)) + 1
    End If
    
    buatIdMasterBarang = "IDB" & Format(baris, "0000")
End Function

Public Function buatIdBarangMasuk() As String
    Dim baris As Long
    Dim idLama As String

    baris = getBarisBarangMasuk

    If Not baris = 1 Then
        idLama = wsBarangMasuk.Range("A" & baris)
        baris = CLng(Right(idLama, 4)) + 1
    End If

    buatIdBarangMasuk = "IDBM" & Format(baris, "0000")
End Function

Public Function buatIdPenjualanBarang() As String
    Dim baris As Long
    Dim idLama As String

    baris = getBarisPenjualanBarang

    If Not baris = 1 Then
        idLama = wsPenjualanBarang.Range("A" & baris)
        baris = CLng(Right(idLama, 4)) + 1
    End If

    buatIdPenjualanBarang = "IDPB" & Format(baris, "0000")
End Function

Public Function buatIdRekapPenjualan() As String
    Dim baris As Long
    Dim idLama As String

    baris = getBarisRekapPenjualan

    If Not baris = 1 Then
        idLama = wsRekapPenjualan.Range("A" & baris)
        baris = CLng(Right(idLama, 4)) + 1
    End If

    buatIdRekapPenjualan = "IDRP" & Format(baris, "0000")
End Function
