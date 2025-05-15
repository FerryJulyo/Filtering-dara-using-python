# import pandas as pd

# # File paths
# top_hits_file = 'Combined/Master_TopHits_Flagged.xlsx'
# vod_file = 'Master VOD.xlsx'
# output_file = 'Combined/Master_TopHits_Flagged_With_VodType.xlsx'

# print("🔄 Membaca semua sheet dari Master_TopHits_Flagged.xlsx...")
# top_hits_sheets = pd.read_excel(top_hits_file, sheet_name=None)
# print(f"✅ Ditemukan {len(top_hits_sheets)} sheet pada Master_TopHits_Flagged.xlsx\n")

# print("🔄 Membaca sheet 'Song' dari Master VOD.xlsx...")
# vod_song_df = pd.read_excel(vod_file, sheet_name='Song')
# print(f"✅ Sheet 'Song' dari Master VOD.xlsx terdiri dari {len(vod_song_df)} baris.\n")

# # Siapkan dictionary hasil akhir
# result_sheets = {}

# print("🔍 Memulai proses pencocokan dan penambahan kolom 'VodType Temp'...")
# for sheet_name, df in top_hits_sheets.items():
#     print(f"📄 Memproses sheet: {sheet_name} ({len(df)} baris)...")

#     # Cek kolom yang diperlukan
#     # required_columns = ['Song', 'Sing1', 'Sing2', 'Sing3', 'Sing4', 'Sing5']
#     required_columns = ['Song']
#     missing_cols = [col for col in required_columns if col not in df.columns]
#     if missing_cols:
#         print(f"⚠️  Sheet '{sheet_name}' tidak memiliki kolom {missing_cols}, kolom 'VodType Temp' akan dikosongkan.")
#         df['VodType Temp'] = ''
#         result_sheets[sheet_name] = df
#         continue

#     # Tambah kolom VodType Temp, default kosong
#     df['VodType Temp'] = ''

#     # Ubah semua kolom yang digunakan untuk matching ke string agar perbandingan konsisten
#     for col in required_columns:
#         df[col] = df[col].astype(str)
#         vod_song_df[col] = vod_song_df[col].astype(str)

#     # Buat key gabungan untuk pencocokan yang lebih cepat
#     df['match_key'] = df[required_columns].agg('|'.join, axis=1)
#     vod_song_df['match_key'] = vod_song_df[required_columns].agg('|'.join, axis=1)

#     # Buat dictionary mapping match_key ke VodType Temporary
#     vod_map = dict(zip(vod_song_df['match_key'], vod_song_df['VodType Temporary']))

#     # Isi kolom VodType Temp berdasarkan mapping
#     df['VodType Temp'] = df['match_key'].map(vod_map).fillna('')

#     # Hapus kolom bantu
#     df.drop(columns=['match_key'], inplace=True)

#     hit_count = df['VodType Temp'].astype(bool).sum()
#     print(f"✅ Ditemukan {hit_count} baris dengan nilai 'VodType Temp' di sheet {sheet_name}.\n")

#     # Simpan hasil sheet
#     result_sheets[sheet_name] = df

# # Simpan semua sheet ke file output
# print("💾 Menyimpan hasil ke file output...")
# with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#     for name, df in result_sheets.items():
#         df.to_excel(writer, sheet_name=name, index=False)

# print(f"\n✅ Semua sheet selesai diproses dan hasil disimpan di: {output_file}")


import pandas as pd

# File paths
top_hits_file = 'Combined/Master_TopHits_Flagged.xlsx'
vod_file = 'Master VOD.xlsx'
output_file = 'Combined/Master_TopHits_Flagged_With_VodType.xlsx'

print("🔄 Membaca semua sheet dari Master_TopHits_Flagged.xlsx...")
top_hits_sheets = pd.read_excel(top_hits_file, sheet_name=None)
print(f"✅ Ditemukan {len(top_hits_sheets)} sheet pada Master_TopHits_Flagged.xlsx\n")

print("🔄 Membaca sheet 'Song' dari Master VOD.xlsx...")
vod_song_df = pd.read_excel(vod_file, sheet_name='Song')
print(f"✅ Sheet 'Song' dari Master VOD.xlsx terdiri dari {len(vod_song_df)} baris.\n")

# Siapkan dictionary hasil akhir
result_sheets = {}

print("🔍 Memulai proses pencocokan dan penambahan kolom 'VodType Temp'...")
for sheet_name, df in top_hits_sheets.items():
    print(f"📄 Memproses sheet: {sheet_name} ({len(df)} baris)...")

    required_columns = ['Song', 'Sing1', 'Sing2', 'Sing3', 'Sing4', 'Sing5']
    # required_columns = ['Song']
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        print(f"⚠️  Sheet '{sheet_name}' tidak memiliki kolom {missing_cols}, kolom terkait akan dikosongkan.")
        for col in ['VodType Temp', 'SongId_VOD', 'Sing1_VOD', 'Sing2_VOD', 'Sing3_VOD', 'Sing4_VOD', 'Sing5_VOD']:
            df[col] = ''
        result_sheets[sheet_name] = df
        continue

    # Ubah ke string
    for col in required_columns:
        df[col] = df[col].astype(str)
        vod_song_df[col] = vod_song_df[col].astype(str)

    # Key gabungan untuk pencocokan
    df['match_key'] = df[required_columns].agg('|'.join, axis=1)
    vod_song_df['match_key'] = vod_song_df[required_columns].agg('|'.join, axis=1)

    # Buat mapping per kolom
    vodtype_map = dict(zip(vod_song_df['match_key'], vod_song_df.get('VodType Temporary', '')))
    songid_map = dict(zip(vod_song_df['match_key'], vod_song_df.get('SongId', '')))
    sing1_map = dict(zip(vod_song_df['match_key'], vod_song_df.get('Sing1', '')))
    sing2_map = dict(zip(vod_song_df['match_key'], vod_song_df.get('Sing2', '')))
    sing3_map = dict(zip(vod_song_df['match_key'], vod_song_df.get('Sing3', '')))
    sing4_map = dict(zip(vod_song_df['match_key'], vod_song_df.get('Sing4', '')))
    sing5_map = dict(zip(vod_song_df['match_key'], vod_song_df.get('Sing5', '')))

    # Isi kolom berdasarkan mapping
    df['VodType Temp'] = df['match_key'].map(vodtype_map).fillna('')
    df['SongId_VOD'] = df['match_key'].map(songid_map).fillna('')
    df['Sing1_VOD'] = df['match_key'].map(sing1_map).fillna('')
    df['Sing2_VOD'] = df['match_key'].map(sing2_map).fillna('')
    df['Sing3_VOD'] = df['match_key'].map(sing3_map).fillna('')
    df['Sing4_VOD'] = df['match_key'].map(sing4_map).fillna('')
    df['Sing5_VOD'] = df['match_key'].map(sing5_map).fillna('')

    # Hapus kolom bantu
    df.drop(columns=['match_key'], inplace=True)

    hit_count = df['VodType Temp'].astype(bool).sum()
    print(f"✅ Ditemukan {hit_count} baris dengan nilai 'VodType Temp' di sheet {sheet_name}.\n")

    result_sheets[sheet_name] = df

# Simpan ke file output
print("💾 Menyimpan hasil ke file output...")
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for name, df in result_sheets.items():
        df.to_excel(writer, sheet_name=name, index=False)

print(f"\n✅ Semua sheet selesai diproses dan hasil disimpan di: {output_file}")

# 

# 

# Menulis hasil ke file Excel
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for name, df in result_sheets.items():
        # Gantikan string 'nan' menjadi string kosong ('') pada kolom yang diinginkan
        columns_to_replace = ['Sing1', 'Sing2', 'Sing3', 'Sing4', 'Sing5', 
                              'Sing1_VOD', 'Sing2_VOD', 'Sing3_VOD', 'Sing4_VOD', 'Sing5_VOD']
        
        # Ganti string 'nan' menjadi ''
        df[columns_to_replace] = df[columns_to_replace].replace('nan', '')

        # Simpan ke file output
        df.to_excel(writer, sheet_name=name, index=False)

print(f"\n✅ Semua sheet selesai diproses dan hasil disimpan di: {output_file}")

