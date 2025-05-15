import pandas as pd
from openpyxl import load_workbook

# File path
master_file = 'Master.xlsx'
filtered_file = 'Combined/FilteredByLanguage.xlsx'
output_file = 'Combined/Master_TopHits_Flagged.xlsx'

print("🔄 Membaca semua sheet dari Master.xlsx dan FilteredByLanguage.xlsx...")
master_sheets = pd.read_excel(master_file, sheet_name=None)
filtered_sheets = pd.read_excel(filtered_file, sheet_name=None)
print(f"✅ Master.xlsx terdiri dari {len(master_sheets)} sheet.")
print(f"✅ FilteredByLanguage.xlsx terdiri dari {len(filtered_sheets)} sheet.\n")

# Siapkan dictionary untuk hasil akhir
result_sheets = {}

# Iterasi setiap sheet di Master.xlsx
for sheet_name, master_df in master_sheets.items():
    print(f"📄 Sheet: {sheet_name}")
    total_rows = len(master_df)

    if total_rows == 0:
        print(f"⚠️  Sheet ini kosong, dilewati.\n")
        master_df['Top Hits'] = ''
        result_sheets[sheet_name] = master_df
        continue

    # Cek apakah sheet juga ada di FilteredByLanguage
    if sheet_name in filtered_sheets:
        filtered_df = filtered_sheets[sheet_name]
        filtered_count = len(filtered_df)

        print(f"🔍 Sheet cocok ditemukan di FilteredByLanguage.xlsx ({filtered_count} baris).")

        # Pastikan kolom SongId ada di kedua file
        if 'SongId' not in master_df.columns:
            raise KeyError(f"❌ Kolom 'SongId' tidak ditemukan di sheet '{sheet_name}' Master.xlsx.")
        if 'SongId' not in filtered_df.columns:
            raise KeyError(f"❌ Kolom 'SongId' tidak ditemukan di sheet '{sheet_name}' FilteredByLanguage.xlsx.")

        # Ambil daftar SongId dari FilteredByLanguage
        filtered_ids = set(filtered_df['SongId'].dropna().astype(str))

        # Tambahkan kolom "Top Hits"
        master_df['Top Hits'] = master_df['SongId'].astype(str).apply(
            lambda x: '✅' if x in filtered_ids else ''
        )

        # Siapkan mapping Jumlah Pengguna berdasarkan SongId
        if 'Jumlah Pengguna' not in filtered_df.columns:
            raise KeyError(f"❌ Kolom 'Jumlah Pengguna' tidak ditemukan di sheet '{sheet_name}' FilteredByLanguage.xlsx.")

        pengguna_map = dict(zip(filtered_df['SongId'].astype(str), filtered_df['Jumlah Pengguna']))

        # Ambil nilai 'Jumlah Pengguna' dari mapping
        jumlah_pengguna_series = master_df['SongId'].astype(str).map(pengguna_map)

        # Sisipkan kolom "Jumlah Pengguna" setelah "Top Hits"
        top_hits_index = master_df.columns.get_loc('Top Hits')
        cols = list(master_df.columns)
        cols.insert(top_hits_index + 1, 'Jumlah Pengguna')
        master_df['Jumlah Pengguna'] = jumlah_pengguna_series
        master_df = master_df[cols]

        hits_count = (master_df['Top Hits'] == '✅').sum()
        print(f"✅ Ditandai {hits_count} lagu sebagai 'Top Hits' dari total {total_rows} lagu.\n")

    else:
        print(f"⚠️  Sheet ini tidak ditemukan di FilteredByLanguage.xlsx, kolom 'Top Hits' dikosongkan.\n")
        master_df['Top Hits'] = ''

    # Simpan hasil
    result_sheets[sheet_name] = master_df

# Simpan ke Excel baru
print("💾 Menyimpan hasil ke file output...")
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for name, df in result_sheets.items():
        df.to_excel(writer, sheet_name=name, index=False)

print(f"\n✅ Semua sheet selesai diproses! Hasil disimpan sebagai: {output_file}")
