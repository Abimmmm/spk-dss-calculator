# DSS Calculator — Entropy, GRA, MEREC, MARCOS

Kalkulator sistem pendukung keputusan (SPK) multi-kriteria untuk menghitung peringkat alternatif menggunakan empat metode:

- **Entropy** — pembobotan kriteria berbasis informasi
- **GRA** (Grey Relational Analysis) — kedekatan alternatif terhadap solusi ideal
- **MEREC** (Method based on the Removal Effects of Criteria) — pembobotan kriteria
- **MARCOS** (Measurement of Alternatives and Ranking according to Compromise Solution) — utility dan ranking akhir

![Tampilan DSS Calculator](preview.jpg)

## Fitur

- Matriks keputusan dinamis: jumlah alternatif dan kriteria bebas
- Tipe kriteria **benefit** dan **cost**
- Perhitungan normalisasi, bobot, skor, dan ranking
- Perbandingan hasil GRA dan MARCOS dalam satu rekapitulasi
- Impor data dari Excel (`.xlsx`, `.xls`, `.csv`)
- Unduh template Excel
- Ekspor seluruh hasil perhitungan ke Excel
- Data contoh untuk mencoba aplikasi dengan cepat
- Berjalan tanpa database; perhitungan dilakukan di browser

## Teknologi

- PHP (entry point)
- JavaScript
- Tailwind CSS
- SheetJS / XLSX
- Metode MCDM: Entropy, GRA, MEREC, MARCOS

## Menjalankan

```bash
php -S localhost:8000
```

Buka `http://localhost:8000` lalu klik **Contoh Data** dan **Hitung Semua** untuk melihat alur lengkap.

## Format impor Excel

| Baris | Isi |
|---|---|
| 1 | Nama kriteria (kolom A kosong) |
| 2 | Tipe kriteria: `benefit` atau `cost` |
| 3+ | Nama alternatif di kolom A, nilai kriteria di kolom B dan seterusnya |

Gunakan tombol **Unduh Template Excel** untuk mendapatkan format yang benar.

## Catatan

Aplikasi ini adalah alat perhitungan, bukan sistem multi-pengguna. Data tidak disimpan di server; memuat ulang halaman akan mengosongkan input. Untuk penggunaan produksi, validasi rumus dengan contoh perhitungan manual/Excel dan pertimbangkan penyimpanan hasil di database.

Dibuat oleh **Abim Mustawa**, 2026.
