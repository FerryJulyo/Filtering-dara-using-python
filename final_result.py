import pandas as pd

# Mapping kategori berdasarkan awalan SongId
kategori_map = {
    '1': 'INDONESIA',
    '2': 'ENGLISH',
    '3': 'MANDARIN',
    '4': 'JEPANG',
    '5': 'KOREA',
    '6': 'INDIA',
    '7': 'FILIPINA',
    '8': 'THAILAND',
    '9': 'REMIX'
}

# Baca file hasil filter
print("🔄 Membaca FilteredResult.xlsx...")
df = pd.read_excel('Combined/FilteredResult.xlsx')

# Pastikan kolom SongId ada
if 'SongId' not in df.columns:
    raise KeyError("Kolom 'SongId' tidak ditemukan di FilteredResult.xlsx.")

# Buat writer untuk simpan ke beberapa sheet
output_file = 'Combined/FilteredByLanguage.xlsx'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# Proses per kategori
for kode, kategori in kategori_map.items():
    print(f"📂 Memproses kategori: {kategori}")
    filtered = df[df['SongId'].astype(str).str.startswith(kode)]
    if not filtered.empty:
        filtered.to_excel(writer, sheet_name=kategori, index=False)

# Simpan file
writer.close()
print(f"✅ File berhasil disimpan sebagai '{output_file}'")
