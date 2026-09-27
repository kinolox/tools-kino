import pandas as pd
import time
from playwright.sync_api import sync_playwright

# =====================================================================
# CONFIGURATION / PENGATURAN UTAMA
# =====================================================================
URL_LOGIN = "https://inventory.pms.web.id/Login"

FILE_INPUT = r"input_cek_stok.xlsx"  # File Excel berisi daftar 'item_id'
FILE_OUTPUT = r"hasil_cek_stok.xlsx" # File Excel hasil rekap stok lengkap dengan satuan
# =====================================================================

# 1. BACA FILE EXCEL INPUT
try:
    df_excel = pd.read_excel(FILE_INPUT)
    print(f"[+] Berhasil memuat {len(df_excel)} baris data untuk dicek.")
except Exception as e:
    print(f"[-] Gagal membaca file Excel: {e}")
    exit()

if 'item_id' not in df_excel.columns:
    print("[-] Error: File Excel Anda harus memiliki kolom dengan nama header 'item_id'")
    exit()

hasil_cek = []

# 2. PROSES OTOMATISASI BROWSER DENGAN PLAYWRIGHT
with sync_playwright() as p:
    print("\n[*] Membuka browser Chrome...")
    browser = p.chromium.launch(headless=False, channel="chrome") 
    context = browser.new_context()
    page = context.new_page()
    
    # ATUR TIMEOUT MENJADI TIDAK TERBATAS AGAR TIDAK ERROR TIMEOUT
    page.set_default_timeout(0)
    
    print("[*] Menuju halaman login website...")
    page.goto(URL_LOGIN)
    
    print("\n" + "!"*60)
    print("[!] SILAKAN LAKUKAN LOGIN MANUAL DI MONITOR ANDA SEKARANG.")
    print("!"*60 + "\n")
    
    input("[!] Jika sudah berhasil login dan masuk ke dashboard, tekan ENTER di sini untuk melanjutkan...")
    
    print("[+] Robot mulai memeriksa stok...")

    # --- LOOP UTAMA ---
    for index, row in df_excel.iterrows():
        item_id = str(row.get('item_id', '')).strip()
        
        print(f"==================================================")
        print(f"[*] MEMERIKSA BARIS KE-{index+1} | Item ID: {item_id}")
        print(f"==================================================")

        try:
            # Tembak langsung API internal website untuk menarik data item
            response = page.request.post(
                "https://inventory.pms.web.id/Item/getDataAll?jtStartIndex=0&jtPageSize=10&jtSorting=nama_barang%20ASC",
                form={"cari": item_id}
            )
            
            stok_ditemukan = "0"
            nama_barang = "-"
            satuan_barang = "-"
            
            if response.status == 200:
                data = response.json()
                records = data.get("Records", [])
                if len(records) > 0:
                    item_data = records[0]
                    stok_ditemukan = str(item_data.get("stok", "0"))
                    nama_barang = str(item_data.get("nama_barang", "-"))
                    satuan_barang = str(item_data.get("satuan", "-"))
                    print(f"    [OK] Ditemukan -> Barang: {nama_barang} | Stok: {stok_ditemukan} {satuan_barang}")
                else:
                    stok_ditemukan = "TIDAK KETEMU"
                    print(f"    [-] Item ID '{item_id}' tidak ditemukan di sistem.")
            
            hasil_cek.append({
                "Item ID": item_id,
                "Nama Barang": nama_barang,
                "Satuan": satuan_barang,
                "Stok Tersedia": stok_ditemukan,
                "Status": "Sukses"
            })

        except Exception as e:
            print(f"    [-] Gagal memeriksa Item ID {item_id} -> {e}")
            hasil_cek.append({
                "Item ID": item_id,
                "Nama Barang": "-",
                "Satuan": "-",
                "Stok Tersedia": "ERROR",
                "Status": "Gagal",
                "Pesan": str(e)
            })

    browser.close()

# Simpan hasil pengecekan ke file Excel baru
df_hasil = pd.DataFrame(hasil_cek)
df_hasil.to_excel(FILE_OUTPUT, index=False)
print(f"\n[+] Pengecekan selesai! Rekap stok lengkap dengan satuan disimpan di: '{FILE_OUTPUT}'")