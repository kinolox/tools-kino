import pandas as pd
from playwright.sync_api import sync_playwright

# =====================================================================
# CONFIGURATION
# =====================================================================
URL_LOGIN = "https://inventory.pms.web.id/Login"
URL_API_CARI = "https://inventory.pms.web.id/Item/getDataAll?jtStartIndex=0&jtPageSize=10&jtSorting=nama_barang%20ASC"

FILE_INPUT = r"data_itemid.xlsx"  # Berisi kolom 'nama_barang' saja
FILE_OUTPUT = r"hasil_tarik_itemid.xlsx"  # File hasil lengkap
# =====================================================================

def tarik_data_item():
    # 1. Baca file Excel input
    try:
        df = pd.read_excel(FILE_INPUT)
        print(f"[+] Berhasil membaca {len(df)} baris data dari {FILE_INPUT}")
    except Exception as e:
        print(f"[-] Gagal membaca file Excel: {e}")
        return

    if 'nama_barang' not in df.columns:
        print("[-] Error: File Excel Anda harus memiliki kolom dengan nama header 'nama_barang'")
        return

    # List untuk menampung hasil data baru
    hasil_item_id = []
    hasil_nama_barang = []
    hasil_nomor_part = []
    hasil_satuan = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome")
        context = browser.new_context()
        page = context.new_page()
        
        # Login manual untuk mendapatkan sesi cookie yang valid
        page.goto(URL_LOGIN)
        print("\n" + "!"*50)
        print("[!] SILAKAN LOGIN MANUAL DI BROWSER SEKARANG.")
        print("!"*50 + "\n")
        
        input("[!] Jika sudah berhasil login dan masuk ke dashboard, tekan ENTER di sini untuk mulai menarik data...")
        
        print("\n[*] Memulai proses penarikan data item...")
        
        # Loop setiap baris nama barang di input
        for index, row in df.iterrows():
            keyword = str(row.get('nama_barang', '')).strip()
            
            if not keyword or keyword.lower() == 'nan':
                hasil_item_id.append("-")
                hasil_nama_barang.append("-")
                hasil_nomor_part.append("-")
                hasil_satuan.append("-")
                continue
                
            try:
                # Tembak API pencarian
                response = page.request.post(URL_API_CARI, form={"cari": keyword})
                
                if response.status == 200:
                    data = response.json()
                    records = data.get("Records", [])
                    if len(records) > 0:
                        # Ambil data teratas yang paling sesuai
                        item = records[0]
                        hasil_item_id.append(item.get("item_id", "-"))
                        hasil_nama_barang.append(item.get("nama_barang", "-"))
                        hasil_nomor_part.append(item.get("nomor_part", "-"))
                        hasil_satuan.append(item.get("satuan", "-"))
                        print(f"    [OK] '{keyword}' --> Ditemukan!")
                    else:
                        hasil_item_id.append("TIDAK KETEMU")
                        hasil_nama_barang.append(keyword)
                        hasil_nomor_part.append("-")
                        hasil_satuan.append("-")
                        print(f"    [-] '{keyword}' --> Tidak ditemukan di sistem")
                else:
                    hasil_item_id.append("ERROR API")
                    hasil_nama_barang.append(keyword)
                    hasil_nomor_part.append("-")
                    hasil_satuan.append("-")
                    print(f"    [-] '{keyword}' --> Gagal akses API")
            except Exception as e:
                hasil_item_id.append("ERROR")
                hasil_nama_barang.append(keyword)
                hasil_nomor_part.append("-")
                hasil_satuan.append("-")
                print(f"    [-] Error untuk '{keyword}': {e}")

        browser.close()

    # Masukkan hasil ke DataFrame baru
    df_output = pd.DataFrame({
        'pencarian_awal': df['nama_barang'],
        'item_id': hasil_item_id,
        'nama_barang': hasil_nama_barang,
        'nomor_part': hasil_nomor_part,
        'satuan': hasil_satuan
    })
    
    # Simpan ke file Excel output
    df_output.to_excel(FILE_OUTPUT, index=False)
    print(f"\n[+] Selesai! Data lengkap telah disimpan ke file: '{FILE_OUTPUT}'")

if __name__ == "__main__":
    tarik_data_item()