import pandas as pd
from playwright.sync_api import sync_playwright

# Konfigurasi
FILE_OPNAME = r"input_opname.xlsx" 
FILE_LOG = r"log_opname_auto.xlsx"
URL_SAVE = "https://inventory.pms.web.id/StokOpnameItem/save"

def proses_opname():
    df = pd.read_excel(FILE_OPNAME)
    log_hasil = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        page.goto("https://inventory.pms.web.id/Login")
        print("\n[!] Silakan login manual...")
        input("[!] Tekan ENTER setelah berhasil login...")
        
        for index, row in df.iterrows():
            # Payload disederhanakan: hanya item_id, tanggal, applicant_id, dan jumlah_opname
            # Pastikan nama kolom di Excel sesuai dengan yang ada di bawah ini
            payload = {
                "tanggal_opname": row['tanggal'],
                "applicant_id": str(row['applicant_id']),
                "item_id": row['item_id'],
                "jumlah_opname": row['jumlah_opname']
            }
            
            try:
                response = page.request.post(URL_SAVE, form=payload)
                if response.status == 200:
                    print(f"[+] Berhasil: {row['item_id']}")
                    log_hasil.append({**payload, "Status": "Sukses"})
                else:
                    print(f"[-] Gagal: {row['item_id']} (Status: {response.status})")
                    log_hasil.append({**payload, "Status": f"Gagal {response.status}"})
            except Exception as e:
                print(f"[-] Error: {e}")
                log_hasil.append({**payload, "Status": "Error", "Pesan": str(e)})

        browser.close()

    pd.DataFrame(log_hasil).to_excel(FILE_LOG, index=False)
    print(f"\n[+] Selesai. Log disimpan di: {FILE_LOG}")

if __name__ == "__main__":
    proses_opname()