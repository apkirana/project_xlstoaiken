# XLStoAIKEN Converter

XLStoAIKEN Converter adalah proyek Python yang menyajikan alat untuk mengonversi data dari file Excel (XLSX) ke format AIKEN yang digunakan untuk membuat dan mengimpor kuis ke dalam berbagai platform Learning Management System (LMS).

## Deskripsi

Proyek ini bertujuan untuk membuat konversi dari format Excel yang diberikan menjadi format AIKEN secara otomatis. Format Excel yang dimaksud memiliki struktur khusus yang terdiri dari kolom pertanyaan, pilihan jawaban, dan kolom jawaban benar. Alat ini membaca data dari file Excel, memproses data, dan menghasilkan file keluaran dalam format AIKEN.

## Struktur Kolom Input Excel

- Kolom 0: Pertanyaan
- Kolom 1-5: Pilihan jawaban (A, B, C, D, E)
- Kolom 6: Jawaban benar (A, B, C, D, atau E)

Berikut adalah contoh template Excel yang digunakan:

| Pertanyaan                     | A    | B     | C     | D     | E      | Jawaban Benar |
|--------------------------------|------|-------|-------|-------|--------|---------------|
| Apa ibukota dari Prancis?      | Bern | Berlin| Paris | Rome  | Madrid | C             |
| Planet terdekat ke matahari?   | Venus| Bumi  | Mars  | Merkuri | Jupiter | D          |
| Siapa penemu bola lampu?       | Edison | Newton | Galileo | Tesla | Bell | A        |

## Instalasi

1. Clone repositori ini:
    ```sh
    git clone https://github.com/apkirana/project_xlstoaiken.git
    ```
2. Pindah ke direktori proyek:
    ```sh
    cd project_xlstoaiken
    ```
3. Buat lingkungan virtual dan aktifkan:
    ```sh
    python -m venv env
    source env/bin/activate  # pada Windows gunakan `.\env\Scripts\activate`
    ```
4. Install dependensi:
    ```sh
    pip install -r requirements.txt
    ```

## Penggunaan

1. Letakkan file Excel yang ingin Anda konversi di dalam direktori proyek.
2. Edit `file_path` di dalam skrip Python dengan path ke file Excel Anda.
3. Jalankan skrip Python untuk melakukan konversi:
    ```sh
    python convert_to_aiken.py
    ```
4. File output dalam format AIKEN akan berada di direktori proyek dengan nama `output_aiken.txt`.

## Contoh Kode

Berikut contoh skrip Python (`convert_to_aiken.py`) yang digunakan untuk konversi:

```python
import pandas as pd

# Path ke file Excel Anda
file_path = '/path/to/your/template.xlsx'

# Membaca file Excel
df = pd.read_excel(file_path, header=None)

# Membuat peta konversi dari huruf ke indeks
answer_map = {
    'A': 0,
    'B': 1,
    'C': 2,
    'D': 3,
    'E': 4
}

output_path = 'output_aiken.txt'

# Membuka file output untuk penulisan
with open(output_path, 'w') as f:
    for index, row in df.iterrows():
        try:
            # Memeriksa apakah jawaban benar adalah huruf yang valid dan melakukan konversi
            correct_answer_index = answer_map[row[6].strip().upper()]
        except KeyError as e:
            print(f"Data error at row {index + 1}: {row[6]} is not a valid answer (A-E)")
            continue

        # Menulis pertanyaan tanpa nomor
        f.write(f"{row[0]}\n")
        
        # Menulis pilihan jawaban
        choices = ['A', 'B', 'C', 'D', 'E']
        for i in range(5):
            f.write(f"    {choices[i]}. {row[i + 1]}\n")
        
        # Menulis jawaban benar
        correct_answer = choices[correct_answer_index]
        f.write(f"ANSWER: {correct_answer}\n\n")

# Menampilkan pesan bahwa konversi selesai
print(f"Konversi selesai! File AIKEN dapat ditemukan di '{output_path}'")
```

## Kontribusi
Kontribusi sangat diterima! Jika Anda memiliki saran perbaikan atau menemukan bug, silakan buka Issues atau buat Pull Request.

## Lisensi
Proyek ini berlisensi di bawah MIT License - lihat LICENSE file untuk detail.

## Contact
- **Project Link**: [Air Quality Analysis on GitHub](https://github.com/apkirana/air-quality-analysis)
- **LinkedIn:** [Annisa Puspa Kirana](https://id.linkedin.com/in/annisapuspakirana/en)
- **Social Links:** [linktr.ee/puspakirana](http://linktr.ee/puspakirana)


