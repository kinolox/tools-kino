import pandas as pd
import requests

# 1. Load data dari file Excel
file_input = 'data_item.xlsx'  # Ganti dengan nama file Excel Anda
try:
    df = pd.read_excel(file_input)
except FileNotFoundError:
    print(f"Error: File '{file_input}' tidak ditemukan.")
    exit()

# Inisialisasi kolom hasil
df['Status'] = ''
df['Nama_Barang'] = ''
df['No_Part'] = ''

# 2. Konfigurasi Sesi & Endpoint berdasarkan Network Tab Anda
session = requests.Session()
session.cookies.set('ci_session', '19dbf514fa164e1d90017b37d05c3beba13fd935') # Pastikan cookie terbaru

url = "https://inventory.pms.web.id/Item/getDataAll?jtStartIndex=0&jtPageSize=100&jtSorting=nama_barang%20ASC"

headers = {
    'X-Requested-With': 'XMLHttpRequest',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
}

total_item = len(df)
print(f"Memeriksa {total_item} item...\n")

# 3. Looping pengecekan
for index, row in df.iterrows():
    item_id = str(row['item_id']).strip()
    
    payload = {'cari': item_id}
    
    try:
        response = session.post(url, data=payload, headers=headers)
        if response.status_code == 200:
            records = response.json().get('Records', [])
            
            # Cek ketersediaan data
            if len(records) > 0:
                status = 'Ada'
                nama_barang = records[0].get('nama_barang', '-')
                no_part = records[0].get('nomor_part', '-')
            else:
                status = 'Tidak Ada'
                nama_barang = '-'
                no_part = '-'
        else:
            status = 'Gagal'
            nama_barang = '-'
            no_part = '-'
            
    except Exception as e:
        status = 'Error'
        nama_barang = str(e)
        no_part = '-'
    
    # Tampilkan hasil di terminal secara real-time
    print(f"[{index + 1}/{total_item}] ID: {item_id} | Status: {status} | Nama: {nama_barang} | No. Part: {no_part}")
    
    # Simpan ke DataFrame
    df.at[index, 'Status'] = status
    df.at[index, 'Nama_Barang'] = nama_barang
    df.at[index, 'No_Part'] = no_part

# 4. Simpan hasil akhir ke file Excel baru
file_output = 'hasil_cek_barang.xlsx'
df.to_excel(file_output, index=False)

print(f"\nSelesai! Hasil telah disimpan dalam file '{file_output}'.")