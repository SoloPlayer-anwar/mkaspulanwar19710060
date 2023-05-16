Attribute VB_Name = "UpdateRekap"
Sub ShapeUpdateRekapPenjualan_Click()
    UpdateRekapPenjualan
    MsgBox "Data berhasil diupdate!", vbInformation, "Update Rekap Penjualan"
End Sub

Sub UpdateRekapPenjualan()
    Call SetWorksheets
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim idRekap As String
    Dim idBarang As String
    Dim namaBarang As String
    Dim bulan As String
    Dim tahun As String
    Dim jumlahPenjualan As Integer
    Dim totalPenjualan As Integer
    Dim found As Boolean
    
    ' Clear data in Rekap Penjualan sheet
    'wsRekapPenjualan.Range("A2:F" & wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row).ClearContents
    
    ' Check if Rekap Penjualan sheet has data
    If wsRekapPenjualan.Cells(2, 1).Value <> "" Then
        ' Delete data starting from row 2
        wsRekapPenjualan.Range("A2:F" & wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row).ClearContents
    End If
    
    ' Find last row in Penjualan Barang sheet
    lastRow = wsPenjualanBarang.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row in Penjualan Barang sheet
    For i = 2 To lastRow
        ' Delete all data
        'wsRekapPenjualan.Cells(i, 1).Value = vbNullString
        'wsRekapPenjualan.Cells(i, 2).Value = vbNullString
        'wsRekapPenjualan.Cells(i, 3).Value = vbNullString
        'wsRekapPenjualan.Cells(i, 4).Value = vbNullString
        'wsRekapPenjualan.Cells(i, 5).Value = vbNullString
        'wsRekapPenjualan.Cells(i, 6).Value = vbNullString
    
        ' Get values from current row
        idBarang = wsPenjualanBarang.Cells(i, 3).Value
        namaBarang = wsPenjualanBarang.Cells(i, 4).Value
        jumlahPenjualan = wsPenjualanBarang.Cells(i, 5).Value
        bulan = Month(wsPenjualanBarang.Cells(i, 2).Value)
        tahun = Year(wsPenjualanBarang.Cells(i, 2).Value)
        
        ' Check if data already exists in Rekap Penjualan sheet
        found = False
        For j = 2 To wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row
            If wsRekapPenjualan.Cells(j, 2).Value = idBarang _
            And wsRekapPenjualan.Cells(j, 3).Value = namaBarang _
            And wsRekapPenjualan.Cells(j, 4).Value = bulan _
            And wsRekapPenjualan.Cells(j, 5).Value = tahun _
            Then
                ' Update existing data
                found = True
                totalPenjualan = wsRekapPenjualan.Cells(j, 6).Value + jumlahPenjualan
                wsRekapPenjualan.Cells(j, 6).Value = totalPenjualan
                Exit For
            End If
        Next j
        
        ' Add new data if not found in Rekap Penjualan sheet
        If Not found Then
            wsRekapPenjualan.Cells(wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row + 1, 1).Value = buatIdRekapPenjualan
            wsRekapPenjualan.Cells(wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row, 2).Value = idBarang
            wsRekapPenjualan.Cells(wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row, 3).Value = namaBarang
            wsRekapPenjualan.Cells(wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row, 4).Value = bulan
            wsRekapPenjualan.Cells(wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row, 5).Value = tahun
            wsRekapPenjualan.Cells(wsRekapPenjualan.Cells(Rows.Count, 1).End(xlUp).Row, 6).Value = jumlahPenjualan
        End If
    Next i
End Sub
