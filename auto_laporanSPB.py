import pandas as pd
import requests

# 1. Load daftar item dari Excel
# Pastikan file 'data_spb.xlsx' punya kolom 'item_id'
df = pd.read_excel('data_spb.xlsx')
df['Status_Laporan'] = ''
df['Total_Record'] = 0

# 2. Konfigurasi
url = "https://inventory.pms.web.id/LaporanSPB/getData"
session = requests.Session()
# PENTING: Update cookie ci_session terbaru Anda dari browser
session.cookies.set('ci_session', '0a710185098ccd43ad91e0c61703e07f4ce7af84')

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
}

all_records = [] 

print(f"Memulai pengambilan laporan untuk {len(df)} item...\n")

# 3. Looping untuk setiap item
for index, row in df.iterrows():
    item_id = str(row['item_id'])
    
    payload = {
        'cari': item_id,
        'doc_type': 'SPB',
        'tahun': '2026',
        'jtStartIndex': 0,
        'jtPageSize': 100,
        'jtSorting': 'tanggal_spb ASC'
    }
    
    try:
        response = session.post(url, data=payload, headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            records = data.get('Records', [])
            
            df.at[index, 'Status_Laporan'] = 'Sukses'
            df.at[index, 'Total_Record'] = len(records)
            print(f"Item: {item_id} | Ditemukan: {len(records)} data")
            
            # Simpan detail data
            for rec in records:
                all_records.append({
                    'item_asal': item_id,
                    'tanggal_spb': rec.get('tanggal_spb'),
                    'no_doc': rec.get('no_doc'),
                    'qty': rec.get('qty')
                })
        else:
            df.at[index, 'Status_Laporan'] = f'Gagal ({response.status_code})'
            print(f"Item: {item_id} | Status: Gagal")
            
    except Exception as e:
        df.at[index, 'Status_Laporan'] = 'Error'
        print(f"Item: {item_id} | Error: {str(e)}")

# 4. Simpan ke Excel
df.to_excel('log_laporan_spb.xlsx', index=False)
pd.DataFrame(all_records).to_excel('detail_data_spb.xlsx', index=False)

print("\nProses selesai!")
print("File 'log_laporan_spb.xlsx' dan 'detail_data_spb.xlsx' telah dibuat.")