import pandas as pd
from time import sleep
from tqdm import tqdm

# Tampilkan log awal
print("🔄 Membaca file MasterCombined.xlsx dan TopHitCombined.xlsx...")

# Baca file Excel
master_df = pd.read_excel('Combined/MasterCombined.xlsx')
tophit_df = pd.read_excel('Combined/TopHitCombined.xlsx')

print("✅ File berhasil dibaca.")
print("\n📋 Kolom di MasterCombined:")
for i, col in enumerate(master_df.columns):
    print(f"{i}: {col}")

print("\n📋 Kolom di TopHitCombined:")
for i, col in enumerate(tophit_df.columns):
    print(f"{i}: {col}")

# Ganti sesuai nama kolom yang cocok dengan posisi J, L, DB, DD, DF
# Contoh di bawah ini hanya asumsi, ubah jika tidak sesuai
kolom_gabungan = ['Sing1', 'Sing2', 'Sing3', 'Sing4', 'Sing5']  # Sesuaikan jika perlu

# Cek apakah kolom-kolom gabungan ada
for kolom in kolom_gabungan:
    if kolom not in master_df.columns:
        raise KeyError(f"Kolom '{kolom}' tidak ditemukan di MasterCombined.xlsx. Periksa nama kolomnya.")

# Fungsi gabung kolom: mengabaikan cell kosong
def gabung_kolom(row):
    values = [row.get(k) for k in kolom_gabungan if pd.notna(row.get(k)) and str(row.get(k)).strip() != '']
    return ' - '.join(str(v).strip() for v in values)

# Tambahkan kolom gabungan
print("\n🔧 Membuat kolom gabungan dari MasterCombined...")
master_df['gabungan'] = master_df.apply(gabung_kolom, axis=1)

# Pastikan kolom yang digunakan untuk pencocokan ada
if 'Song' not in master_df.columns:
    raise KeyError("Kolom 'Song' tidak ditemukan di MasterCombined.")
if 'Judul Lagu' not in tophit_df.columns:
    raise KeyError("Kolom 'Judul Lagu' tidak ditemukan di TopHitCombined.")
if 'Penyanyi' not in tophit_df.columns:
    raise KeyError("Kolom 'Penyanyi' tidak ditemukan di TopHitCombined.")

# Filter data
print("🔍 Memproses filter data, harap tunggu...")

# Tambahkan animasi loading pakai tqdm
filtered_list = []
for i in tqdm(range(len(master_df)), desc="🔎 Memfilter"):
    row = master_df.iloc[i]
    matched = tophit_df[
        (tophit_df['Judul Lagu'] == row['Song']) &
        (tophit_df['Penyanyi'] == row['gabungan'])
    ]
    if not matched.empty:
        filtered_list.append(row)


filtered_df = pd.DataFrame(filtered_list)

# Simpan hasil
print("\n💾 Menyimpan hasil filter ke Combined/FilteredResult.xlsx...")
filtered_df.to_excel('Combined/FilteredResult.xlsx', index=False)

print("✅ Proses selesai. File tersimpan sebagai 'FilteredResult.xlsx'.")
