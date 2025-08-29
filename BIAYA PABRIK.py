import pandas as pd

# Baca file Excel
file_path = "cari.xlsx"
data = pd.read_excel(file_path)

# Rapikan nama kolom
data.columns = data.columns.str.strip()

# Bersihkan kolom angka
for col in ["AmountD", "AmountC"]:
    data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

# Kolom total
data["Total"] = data["AmountD"] - data["AmountC"]

# Copy agar tidak dobel
data_available = data.copy()

# Buat daftar kode akun yang mau di-exclude
exclude_accounts = [4220412001, 4220411001, 4220412003]

# Buang baris yang punya AccCode tersebut
data_available = data_available[~data_available["AccCode"].isin(exclude_accounts)]

results = []

# --- 0. Efluent ---
mask_eflu = (data_available["AccCode"] == 4220100015) & (data_available["Scode"] != "CJ")
total_eflu = data_available.loc[mask_eflu, "Total"].sum()
results.append({"Jenis Biaya": "Efluent", "Total": total_eflu})
data_available = data_available.loc[~mask_eflu]


# --- 1. Bahan & alat analisa ---
mask_analisa = ((data_available["AccCode"] == 4220100016) &
                (data_available["Scode"] != "CJ"))
total_analisa = data_available.loc[mask_analisa, "Total"].sum()
results.append({"Jenis Biaya": "Bahan & Alat Analisa", "Total": total_analisa})
data_available = data_available.loc[~mask_analisa]

# --- 2. Penerangan & Air ---
mask_air = (((data_available["AccCode"] == 4220100010) | (data_available["AccCode"] == 4220100014)) &
            (data_available["Scode"] != "CJ"))
total_air = data_available.loc[mask_air, "Total"].sum()
results.append({"Jenis Biaya": "Penerangan & Air", "Total": total_air})
data_available = data_available.loc[~mask_air]

# --- 3. Pengangkutan dalam pabrik ---
mask_pengangkutan = (data_available["AccCode"] == 4220100013) & (data_available["Scode"] != "CJ") | (data_available["Scode"] == "VC")
total_pengangkutan = data_available.loc[mask_pengangkutan, "Total"].sum()   
results.append({"Jenis Biaya": "Pengangkutan Dalam Pabrik", "Total": total_pengangkutan})
data_available = data_available.loc[~mask_pengangkutan]

# --- 4. Upkeep Factory Building ---
mask_upkeep = (data_available["AccCode"] == 4220100011) & (data_available["Scode"] != "CJ")
total_upkeep = data_available.loc[mask_upkeep, "Total"].sum()
results.append({"Jenis Biaya": "Upkeep Factory Building", "Total": total_upkeep})
data_available = data_available.loc[~mask_upkeep]

# --- 5. Operasional Despatch ---
mask_dispatch = (data_available["AccCode"] == 4220100017) & (data_available["Scode"] == "CB")
total_dispatch = data_available.loc[mask_dispatch, "Total"].sum()
results.append({"Jenis Biaya": "Operasional Despatch", "Total": total_dispatch})
data_available = data_available.loc[~mask_dispatch]

# --- 6. Umum ---
mask_umum = data_available["Description"].str.contains("cek|rubah|servis|penambahan|perbaikan|SBA|BSA|pabrikasi|ganti|pengelasan|pengecoran|pembersihan", case=False, na=False)
total_umum = data_available.loc[mask_umum, "Total"].sum()
results.append({"Jenis Biaya": "Umum", "Total": total_umum})
data_available = data_available.loc[~mask_umum]

# --- 7. Bahan kimia ---
mask_kimia = data_available["Description"].str.contains(
    "oksigen isi|soda ash|nalco|garam kasar|soda flake|calcium carbonat",
    case=False, na=False
)
total_kimia = data_available.loc[mask_kimia, "Total"].sum()
results.append({"Jenis Biaya": "Bahan Kimia", "Total": total_kimia})
data_available = data_available.loc[~mask_kimia]

# --- 8. Bahan bakar & pelumas ---
mask_bbm = data_available["Description"].str.contains(
    "oli deltalube|grease|gear oil|oli himatsu|solar|oli rimula|oli gardan spirax|pertalite|oli omala",
    case=False, na=False
)
total_bbm = data_available.loc[mask_bbm, "Total"].sum()
results.append({"Jenis Biaya": "Bahan Bakar & Pelumas", "Total": total_bbm})
data_available = data_available.loc[~mask_bbm]

# --- 9. Marketing (Tonindo - Tomy E Fee Penjualan) ---
mask_marketing = (data_available["AccCode"] == 4220100017) & \
                 (data_available["Description"].str.contains("Tomy E - Fee penjualan | Fee penjualan", case=False, na=False))
total_marketing = data_available.loc[mask_marketing, "Total"].sum()
results.append({"Jenis Biaya": "Marketing (Tonindo)", "Total": total_marketing})

# buang dari data_available
data_available = data_available.loc[~mask_marketing]

# --- 10. Perkakas kecil ---
mask_perkakas = (data_available["Total"] < 1_000_000) & ~(
    data_available["Description"].str.contains("oksigen isi|garam kasar|soda ash|soda flake|calcium carbonat|nalco",
                                              case=False, na=False)
)
total_perkakas = data_available.loc[mask_perkakas, "Total"].sum()
results.append({"Jenis Biaya": "Perkakas Kecil", "Total": total_perkakas})
data_available = data_available.loc[~mask_perkakas]

# --- 11. Overhead ---
mask_overhead = (
    (data_available["Total"] > 1_000_000)
    & ~(data_available["Scode"] == "CJ")
    & ~data_available["AccCode"].isin(exclude_accounts)   # pastikan exclude juga dicek
)

total_overhead = data_available.loc[mask_overhead, "Total"].sum()
results.append({"Jenis Biaya": "Overhead", "Total": total_overhead})
data_available = data_available.loc[~mask_overhead]


# Buat DataFrame hasil
df_results = pd.DataFrame(results)

# Simpan ke sheet baru
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_results.to_excel(writer, sheet_name="hasil_kelompok", index=False)

print(df_results)

# Atau cek total sisa
print("Total sisa:", data_available["Total"].sum())

# Kalau mau export ke Excel biar lebih jelas
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    data_available.to_excel(writer, sheet_name="sisa_data", index=False)
