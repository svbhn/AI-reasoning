CARA MENJALANKAN PROGRAM:

1. Pastikan sudah menginstall Python 3 keatas
2. Simpan file berikut dalam folder yang sama:
   - restoran.py (file program utama)
   - restoran.xlsx (file data input)
3. Install library yang diperlukan:
   pip install openpyxl
4. Jalankan program dengan perintah:
   python restoran.py
5. Hasil akan:
   - Ditampilkan di terminal (top 5 restoran)
   - Disimpan dalam file peringkat.xlsx

STRUKTUR FILE INPUT:
- File Excel dengan 3 kolom:
  - Kolom A: ID Restoran
  - Kolom B: Kualitas Layanan (1-100)
  - Kolom C: Harga (Rp 25.000-55.000)

PENJELASAN OUTPUT:
- Skor rekomendasi (0-100%) dihitung berdasarkan:
  - Kualitas layanan (Buruk, Sedang, Baik, Sangat Baik)
  - Harga (Murah, Sedang, Mahal)
  - 12 aturan inferensi fuzzy
  - Metode defuzzifikasi Center of Gravity