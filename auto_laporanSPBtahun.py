import pandas as pd
import requests

# 1. Load data
df = pd.read_excel('data_spb.xlsx')

# --- INPUT TAHUN FLEKSIBEL ---
tahun_input = input("Masukkan tahun (contoh: 2026, atau range 2025-2026): ")

if '-' in tahun_input:
    y_awal, y_akhir = tahun_input.split('-')
    tahun_list = [str(y) for y in range(int(y_awal), int(y_akhir) + 1)]
else:
    tahun_list = [tahun_input]

# 2. Konfigurasi
session = requests.Session()
# PENTING: Pastikan cookie ci_session masih valid (cek di browser)
session.cookies.set('ci_session', 'be97dc31b6c7d77da162d7ed994e884ac0ce5170')
url = "https://inventory.pms.web.id/LaporanSPB/getData"
headers = {
    'X-Requested-With': 'XMLHttpRequest', 
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36'
}

all_records = []
print(f"\nMemulai pengambilan laporan untuk: {', '.join(tahun_list)}...\n")

# 3. Looping
for index, row in df.iterrows():
    item_id = str(row['item_id'])
    
    for tahun in tahun_list:
        payload = {
            'cari': item_id,
            'doc_type': 'SPB',
            'tahun': tahun,
            'jtStartIndex': 0,
            'jtPageSize': 100,
            'jtSorting': 'tanggal_spb ASC'
        }
        
        try:
            response = session.post(url, data=payload, headers=headers)
            if response.status_code == 200:
                records = response.json().get('Records', [])
                print(f"Item: {item_id} | Tahun: {tahun} | Ditemukan: {len(records)} data")
                
                for rec in records:
                    rec.update({'item_asal': item_id, 'tahun_laporan': tahun})
                    all_records.append(rec)
            else:
                print(f"Item: {item_id} | Tahun: {tahun} | Status: Gagal")
        except Exception as e:
            print(f"Error pada {item_id} tahun {tahun}: {e}")

# 4. Simpan hasil
nama_file = f'laporan_spb_{tahun_input.replace("-", "_sd_")}.xlsx'
pd.DataFrame(all_records).to_excel(nama_file, index=False)
print(f"\nSelesai! Data disimpan di '{nama_file}'.")