<h1 align="center">VBA - Aplikasi Toko Alat Olahraga</h1>
<p align="center">Repository ini merupakan sebuah aplikasi Toko Alat Olahraga yang memiliki fitur CRUD pada Excel dan dibuat menggunakan bahasa pemgrograman Visual Basic Application</p>

<img src="https://user-images.githubusercontent.com/64394320/227556953-848079c0-2898-40e3-8f71-23dc169710f7.png" alt="VBA - Toko Alat Olahraga">

# Daftar Isi

- [Layout Excel](#layout-excel)
  - [Sheet `Merek Barang`](#sheet-merek-barang)
  - [Sheet `Kategori Barang`](#sheet-kategori-barang)
  - [Sheet `Master Barang`](#sheet-master-barang)
  - [Sheet `Barang Masuk`](#sheet-barang-masuk)
  - [Sheet `Penjualan Barang`](#sheet-penjualan-barang)
  - [Sheet `Rekap Penjualan`](#sheet-rekap-penjualan)
- [VBA Design](#vba-design)
  - [Form `Merek Barang`](#form-merek-barang)
  - [Form `Kategori Barang`](#form-kategori-barang)
  - [Form `Master Barang`](#form-master-barang)
  - [Form `Barang Masuk`](#form-barang-masuk)
  - [Form `Penjualan Barang`](#form-penjualan-barang)
  - [Form `Rekap Penjualan`](#form-rekap-penjualan)
- [Aturan Penamaan](#aturan-penamaan)
  - [Penamaan Name Untuk Object dan isinya](#penamaan-name-untuk-object-dan-isinya)
  - [Penamaan Variabel, Procedure atau Function](#penamaan-variabel-procedure-atau-function)
  - [Penamaan Shape](#penamaan-shape)
- [ERD VBA](#erd-vba)
  - [Penjelasan ERD](#penjelasan-erd)
    - [Tabel `Merek Barang` dan `Master Barang`](#tabel-merek-barang-dan-master-barang)
    - [Tabel `Kategori Barang` dan `Master Barang`](#tabel-kategori-barang-dan-master-barang)
    - [Tabel `Master Barang` dan `Barang Masuk`](#tabel-master-barang-dan-barang-masuk)
    - [Tabel `Master Barang` dan `Penjualan Barang`](#tabel-master-barang-dan-penjualan-barang)
    - [Tabel `Penjualan Barang` dan `Rekap Penjualan`](#tabel-penjualan-barang-dan-rekap-penjualan)

## Layout Excel

Pada aplikasi ini memiliki layout excel atau design tabel nya sebagai berikut:

### Sheet `Merek Barang`

> **Catatan**:
>
> - `IDMB` = ID Merek Barang
> - Memiliki button trigger untuk memunculkan pop up form input untuk CRUD data

| ID Merek Barang | Merek Barang |
| --------------- | ------------ |
| IDMB0001        | Domyos       |
| IDMB0002        | Nabaiji      |
| IDMB0003        | Artengo      |
| ...             | ...          |

### Sheet `Kategori Barang`

> **Catatan**:
>
> - `IDKB` = ID Kategori Barang
> - Memiliki button trigger untuk memunculkan pop up form input untuk CRUD data

| ID Kategori Barang | Kategori Barang |
| ------------------ | --------------- |
| IDKB0001           | Kebugaran       |
| IDKB0002           | Renang          |
| IDKB0003           | Tenis           |
| ...                | ...             |

### Sheet `Master Barang`

> **Catatan**:
>
> - `IDB` = ID Barang
> - Memiliki button trigger untuk memunculkan pop up form input untuk CRUD data

| ID Barang | Nama Barang                            | ID Merek Barang | Merek Barang | ID Kategori Barang | Kategori Barang |
| --------- | -------------------------------------- | --------------- | ------------ | ------------------ | --------------- |
| IDB0001   | Adidas Samba Classic Soccer White - 42 | IDMB0008        | Adidas       | IDKB0008           | Lari            |
| IDB0002   | Carrier Eiger Sunature 65L             | IDMB0007        | Eiger        | IDKB0007           | Treking         |
| IDB0003   | Air Jordan 1 High Zoom Pink Oxford     | IDMB0012        | Nike         | IDKB0008           | Lari            |
| ...       | ...                                    | ...             | ...          | ...                | ...             |

### Sheet `Barang Masuk`

> **Catatan**:
>
> - `IDBM` = ID Barang Masuk
> - Format tanggal = `dd/mm/yy`
> - Memiliki button trigger untuk memunculkan pop up form input untuk CRUD data

| ID Barang Masuk | Tanggal Masuk | ID Barang | Nama Barang                            | Jumlah Masuk |
| --------------- | ------------- | --------- | -------------------------------------- | ------------ |
| IDBM0001        | 25/02/2023    | IDB0001   | Adidas Samba Classic Soccer White - 42 | 50           |
| IDBM0002        | 23/02/2023    | IDB0002   | Carrier Eiger Sunature 65L             | 24           |
| IDBM0003        | 26/02/2023    | IDB0003   | Air Jordan 1 High Zoom Pink Oxford     | 21           |
| ...             | ...           | ...       | ...                                    | ...          |

### Sheet `Penjualan Barang`

> **Catatan**:
>
> - `IDPB` = ID Penjualan Barang
> - Format tanggal = `dd/mm/yy`
> - Memiliki button trigger untuk memunculkan pop up form input untuk CRUD data

| ID Penjualan Barang | Tanggal Terjual | ID Barang | Nama Barang                                    | Jumlah Penjualan |
| ------------------- | --------------- | --------- | ---------------------------------------------- | ---------------- |
| IDPB0001            | 25/02/2023      | IDB0005   | Celana Renang Anak Yoke 900 - Hitam Hijau Biru | 2                |
| IDPB0002            | 21/02/2023      | IDB0001   | Adidas Samba Classic Soccer White - 42         | 4                |
| IDPB0003            | 02/02/2023      | IDB0002   | Carrier Eiger Sunature 65L                     | 5                |
| IDPB0004            | 24/02/2023      | IDB0001   | Adidas Samba Classic Soccer White - 42         | 2                |
| ...                 | ...             | ...       | ...                                            | ...              |

### Sheet `Rekap Penjualan`

> **Catatan**:
>
> - `IDRP` = ID Rekap Penjualan
> - Kriteria:
>   - Jika `Nama Barang` sama, `Bulan` dan `Tahun` sama = Diakumulasikan
>   - Jika `Nama Barang` sama, `Bulan` atau `Tahun` berbeda = Membuat data baru
> - Memiliki button trigger untuk memunculkan pop up form input untuk CRUD data

| ID Rekap Penjualan | ID Barang | Nama Barang                                    | Bulan | Tahun | Jumlah Penjualan |
| ------------------ | --------- | ---------------------------------------------- | ----- | ----- | ---------------- |
| IDRP0001           | IDB0005   | Celana Renang Anak Yoke 900 - Hitam Hijau Biru | 2     | 2023  | 2                |
| IDRP0002           | IDB0001   | Adidas Samba Classic Soccer White - 42         | 2     | 2023  | 6                |
| IDRP0003           | IDB0002   | Carrier Eiger Sunature 65L                     | 2     | 2023  | 5                |
| ...                | ...       | ...                                            | ...   | ...   | ...              |

## VBA Design

Sebelum aplikasi ini dibuat, saya terlebih dahulu membuat design aplikasu untuk masing-masing form input di setiap sheet nya menggunakan aplikasi Figma, berikut design nya:

### Form `Merek Barang`

![Sheet Merek Barang - Form Merek Barang](https://user-images.githubusercontent.com/64394320/227563871-91e3a00f-657b-4f88-bba2-ad4cdccf7b4b.png)

### Form `Kategori Barang`

![Sheet Kategori Barang - Form Kategori Barang](https://user-images.githubusercontent.com/64394320/227564619-c33e43ef-d064-4f36-948f-1988775c34ad.png)

### Form `Master Barang`

![Sheet Master Barang - Form Master Barang](https://user-images.githubusercontent.com/64394320/227564785-ac8e358c-0d0a-4fcf-a2e4-32b8ef971dc6.png)

### Form `Barang Masuk`

![Sheet Barang Masuk - Form Barang Masuk](https://user-images.githubusercontent.com/64394320/227565026-edd6bffd-3e57-4efd-96f5-70c62dabec97.png)

### Form `Penjualan Barang`

![Sheet Penjualan Barang - Form Penjualan Barang](https://user-images.githubusercontent.com/64394320/227565194-deabe32c-3102-490a-9403-c7691bf00ff2.png)

### Form `Rekap Penjualan`

![Sheet Rekap Penjualan - Form Rekap Penjualan](https://user-images.githubusercontent.com/64394320/227565382-efdc5995-379b-48e3-aa02-92f6d8df41f5.png)

## Aturan Penamaan

Sebelum saya membuat program atau coding VBA nya, saya terlebih dahulu membuat aturan penamaan agar mudah membedakan mana variable, procedure atau function, dll

### Penamaan Name Untuk Object dan isinya

- Form
- Label
- Text Box
- Combo Box
- Command Button

Untuk penamaan name untuk Object dan isinya menggunakan `Pascal Case`, contoh:

```vba
FormBarangMasuk, LabelIdBarangMasuk, TextBoxIdBarangMasuk, ComboBoxMerekBarang, CmdBtnSimpan
```

### Penamaan Variabel, Procedure atau Function

Untuk penamaan variabel, procedure atau function menggunakan `Camel Case`, contoh:

```vba
cariById(), bersihForm(), totalStok, getBarisMerekBarang()
```

### Penamaan Shape

Untuk penamaan shape atau button trigger untuk memunculkan pop up form input untuk CRUD data menggunakan `Pascal Case`, contoh:

```vba
ShapeFormMerekBarang()
```

## ERD VBA

Sebelum saya membuat program atau coding VBA nya, saya terlebih dahulu membuat ERD (Entity Relationship Diagram) Design untuk memudahkan atau membuat gambaran bagaimana relasi antar tabel pada aplikasi ini, berikut adalah ERD dan penjelasan nya:

![ERD - VBA Mini Project 1 Final](https://user-images.githubusercontent.com/64394320/227568883-dd8ce12b-e85c-474b-8aae-60f2d2b40299.png)

### Penjelasan ERD

Pada ERD aplikasi ini memiliki beberapa relasi antar tabel, berikut daftar relasi antar tabel nya:

#### Tabel `Merek Barang` dan `Master Barang`

- Tabel `merek_barang` memiliki relasi ke tabel `master_barang` melalui field `id_merek_barang`. Relasinya adalah `one-to-many`, dimana satu merek barang dapat memiliki banyak barang pada tabel `master_barang`.

#### Tabel `Kategori Barang` dan `Master Barang`

- Tabel `kategori_barang` memiliki relasi ke tabel `master_barang` melalui field `id_kategori_barang`. Relasinya adalah `one-to-many`, dimana satu kategori barang dapat memiliki banyak barang pada tabel `master_barang`.

#### Tabel `Master Barang` dan `Barang Masuk`

- Tabel `master_barang` memiliki relasi ke tabel `barang_masuk` melalui field `id_barang`. Relasinya adalah `one-to-many`, dimana satu barang dapat memiliki banyak transaksi barang masuk pada tabel `barang_masuk`.

#### Tabel `Master Barang` dan `Penjualan Barang`

- Tabel `master_barang` memiliki relasi ke tabel `penjualan_barang` melalui field `id_barang`. Relasinya adalah `one-to-many`, dimana satu barang dapat memiliki banyak transaksi penjualan barang pada tabel `penjualan_barang`.

#### Tabel `Penjualan Barang` dan `Rekap Penjualan`

- Tabel `penjualan_barang` memiliki relasi ke tabel `rekap_penjualan` melalui field `id_barang`. Relasinya adalah `many-to-one`, dimana banyak transaksi penjualan barang dapat memiliki satu rekap penjualan pada tabel `rekap_penjualan`.
