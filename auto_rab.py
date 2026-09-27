import pandas as pd
import requests
import os

# 1. Load data dari Excel
# Pastikan file 'data_rab.xlsx' berada di folder yang sama dengan skrip ini
file_input = 'data_rab.xlsx'
df = pd.read_excel(file_input)

# Inisialisasi kolom log jika belum ada
df['Status'] = ''
df['Log_Pesan'] = ''

# 2. Konfigurasi Sesi
session = requests.Session()
# Ganti dengan ci_session terbaru dari browser Anda
session.cookies.set('ci_session', '15218e873175d4310031256aa66142f1544c0136')

url = "https://inventory.pms.web.id/inputRAB/save"
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    'Referer': 'https://inventory.pms.web.id/InputRAB',
    'X-Requested-With': 'XMLHttpRequest',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
}

print(f"Memulai proses input {len(df)} item...\n")

# 3. Looping untuk setiap baris
for index, row in df.iterrows():
    payload = {
        'item_id': str(row['item_id']),
        'qty': str(row['qty']),
        'divisi_id': str(row['divisi_id']),
        'tanggal_rab': str(row['tanggal_rab']),
        'konversi': '0',
        'keterangan': ''
    }
    
    try:
        response = session.post(url, data=payload, headers=headers)
        
        # Logika status berdasarkan respon server
        if response.text.strip() == 'true':
            status = 'Sukses'
            msg = 'Berhasil diinput'
        else:
            status = 'Gagal'
            msg = response.text.strip()
            
    except Exception as e:
        status = 'Error'
        msg = str(e)
    
    # Menampilkan status langsung di terminal (layar biru)
    print(f"Item: {row['item_id']} | Status: {status} | Pesan: {msg}")
    
    # Menyimpan status ke dataframe
    df.at[index, 'Status'] = status
    df.at[index, 'Log_Pesan'] = msg

# 4. Simpan hasil akhir ke file Excel baru
file_output = 'data_rab_hasil.xlsx'
df.to_excel(file_output, index=False)

print(f"\nProses selesai! Log lengkap telah disimpan di '{file_output}'.")