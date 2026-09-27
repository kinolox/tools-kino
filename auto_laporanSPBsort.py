import pandas as pd
import requests

# 1. Load data
try:
    df = pd.read_excel('data_spb.xlsx')
except FileNotFoundError:
    print("Error: File 'data_spb.xlsx' tidak ditemukan. Pastikan nama file benar.")
    exit()

# --- INPUT TAHUN & RENTANG BULAN ---
tahun_input = input("Masukkan tahun (contoh: 2026): ")
rentang_bulan = input("Masukkan rentang bulan (contoh: 1-6 atau 1): ")

# Parsing rentang bulan
if '-' in rentang_bulan:
    b_awal, b_akhir = map(int, rentang_bulan.split('-'))
    daftar_bulan = range(b_awal, b_akhir + 1)
else:
    daftar_bulan = [int(rentang_bulan)]

# 2. Konfigurasi
session = requests.Session()
# Pastikan ci_session masih valid (cek di browser)
session.cookies.set('ci_session', '8d8f2d4acb2419ed6ae6e3d4a31d4e2897a53bb3')
url = "https://inventory.pms.web.id/LaporanSPB/getData"
headers = {
    'X-Requested-With': 'XMLHttpRequest', 
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
}

all_records = []
jumlah_item_dinamis = len(df)
total_langkah = jumlah_item_dinamis * len(daftar_bulan)

print(f"\n--- INFORMASI PROSES ---")
print(f"Total item di Excel: {jumlah_item_dinamis}")
print(f"Total langkah pencarian: {total_langkah} ({jumlah_item_dinamis} item x {len(daftar_bulan)} bulan)")
print(f"------------------------\n")

# 3. Looping
for index, row in df.iterrows():
    item_id = str(row['item_id'])
    
    for bulan in daftar_bulan:
        print(f"[{index + 1}/{jumlah_item_dinamis}] Mencari: {item_id} | Bulan: {bulan}")
        
        payload = {
            'cari': item_id,
            'doc_type': 'SPB',
            'tahun': tahun_input,
            'bulan': str(bulan),
            'jtStartIndex': 0,
            'jtPageSize': 100,
            'jtSorting': 'no_doc DESC'
        }
        
        try:
            response = session.post(url, data=payload, headers=headers)
            if response.status_code == 200:
                records = response.json().get('Records', [])
                for rec in records:
                    rec.update({'item_asal': item_id, 'bulan_laporan': bulan})
                    all_records.append(rec)
            else:
                print(f"  -> Gagal mengambil data untuk {item_id}")
        except Exception as e:
            print(f"  -> Error: {e}")

# 4. Simpan hasil
nama_file = f'laporan_spb_{tahun_input}_bulan_{rentang_bulan.replace("-", "sd")}.xlsx'
pd.DataFrame(all_records).to_excel(nama_file, index=False)
print(f"\nSelesai! Data berhasil disimpan di file: '{nama_file}'.")