import pandas as pd
import time
import re
from playwright.sync_api import sync_playwright

# =====================================================================
# CONFIGURATION / PENGATURAN UTAMA
# =====================================================================
URL_LOGIN = "https://inventory.pms.web.id/Login"

FILE_INPUT = r"input_pengeluaran.xlsx"
FILE_LOG = r"log_input_pms_auto.xlsx"
# =====================================================================

def konversi_tanggal_indo(teks_tanggal):
    teks_str = str(teks_tanggal).strip().lower()
    bulan_map = {
        'januari': '01', 'jan': '01', 'februari': '02', 'feb': '02',
        'maret': '03', 'mar': '03', 'april': '04', 'apr': '04', 'mei': '05',
        'juni': '06', 'jun': '06', 'juli': '07', 'jul': '07',
        'agustus': '08', 'ags': '08', 'agu': '08', 'september': '09', 'sep': '09',
        'oktober': '10', 'okt': '10', 'november': '11', 'nov': '11', 'desember': '12', 'des': '12'
    }
    try:
        return pd.to_datetime(teks_tanggal).strftime('%Y-%m-%d')
    except:
        pass
    partisi = teks_str.split()
    if len(partisi) == 3:
        tgl = partisi[0].zfill(2)
        bln = bulan_map.get(partisi[1], '01')
        thn = partisi[2]
        return f"{thn}-{bln}-{tgl}"
    return teks_str

# 1. BACA FILE EXCEL INPUT
try:
    df_excel = pd.read_excel(FILE_INPUT)
    print(f"[+] Berhasil memuat {len(df_excel)} baris data dari Excel.")
except Exception as e:
    print(f"[-] Gagal membaca file Excel: {e}")
    exit()

df_excel['tanggal_proses'] = df_excel['tanggal_keluar'].apply(konversi_tanggal_indo)
log_hasil = []

# 2. PROSES OTOMATISASI BROWSER DENGAN PLAYWRIGHT
with sync_playwright() as p:
    print("\n[*] Membuka browser Chrome...")
    browser = p.chromium.launch(headless=False, channel="chrome") 
    context = browser.new_context()
    page = context.new_page()
    
    # PENTING UNTUK 200 DATA: Menghapus batasan timeout total
    page.set_default_timeout(0)
    
    print("[*] Menuju halaman login website...")
    page.goto(URL_LOGIN)
    
    print("\n" + "!"*60)
    print("[!] SILAKAN LAKUKAN LOGIN MANUAL DI MONITOR ANDA SEKARANG.")
    print("!"*60 + "\n")
    
    halaman_siap = False
    for i in range(120): 
        try:
            if page.locator("input[name='tanggal_keluar']").is_visible():
                print("[+] Halaman Form Input Pengeluaran Terdeteksi!")
                halaman_siap = True
                break
        except:
            pass
        time.sleep(1)
        
    if not halaman_siap:
        print("[-] Waktu tunggu login manual habis. Program dihentikan.")
        browser.close()
        exit()
        
    page.wait_for_load_state("load")
    page.wait_for_timeout(2000)
    print("[+] Sesi login dikonfirmasi. Robot mulai memproses data...\n")

    # --- LOOP UTAMA ---
    for index, row in df_excel.iterrows():
        tanggal_sekarang = row['tanggal_proses']
        item_id_excel = str(row.get('item_id', '')).strip().split('.')[0]
        qty_excel = str(row.get('qty', '1')).strip().split('.')[0]
        unit_lambung_excel = str(row.get('unit_lambung', '')).strip()
        lokasi_excel = str(row.get('lokasi', '')).strip()
        
        divisi_id_val = str(row.get('divisi_id', '')).split('.')[0]
        applicant_id_val = str(row.get('applicant_id', '')).split('.')[0]
        
        sn_in = str(row.get('sn_in', '')).strip()
        sn_out = str(row.get('sn_out', '')).strip()
        posisi_raw = str(row.get('posisi', '')).strip()
        angka_posisi = re.sub(r'(?i)no\.?\s*', '', posisi_raw).strip()
        
        if not sn_out or sn_out.lower() == 'nan' or sn_out == '':
            keterangan_otomatis = f"{sn_in}(INN) NO.{angka_posisi}"
        else:
            keterangan_otomatis = f"{sn_in}(INN), {sn_out}(OUT) NO.{angka_posisi}"

        print(f"==================================================")
        print(f"[*] MEMPROSES DATA BARIS KE-{index+1} / {len(df_excel)}")
        print(f"==================================================")
        print(f"    -> Mengisi Lengkap: Tgl: {tanggal_sekarang} | Item: {item_id_excel} | Unit: {unit_lambung_excel}")

        try:
            no_sbk_otomatis = page.locator("input[name='no_doc']").first.input_value()
            
            page.locator("input[name='tanggal_keluar']").fill(tanggal_sekarang)
            page.locator("input[name='tanggal_keluar']").press("Tab")
            page.wait_for_timeout(500)

            if divisi_id_val and divisi_id_val != 'nan':
                page.locator("select[name='divisi_id']").select_option(value=divisi_id_val)
                page.wait_for_timeout(300)
                
            if applicant_id_val and applicant_id_val != 'nan':
                page.locator("select[name='applicant_id']").select_option(value=applicant_id_val)
                page.wait_for_timeout(500)

            input_barang = page.locator("input#nama_barang-flexdatalist, input[placeholder*='karakter']").first
            input_barang.click()
            page.wait_for_timeout(200)
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            page.wait_for_timeout(200)
            page.keyboard.type(item_id_excel, delay=150)
            
            target_preview_barang = page.locator("table, ul.flexdatalist-results").locator(f"tr:has-text('{item_id_excel}'), li:has-text('{item_id_excel}')").first
            target_preview_barang.wait_for(state="visible", timeout=10000)
            target_preview_barang.click()
            print("      [+] Sukses memilih Nama Barang!")
            page.wait_for_timeout(500)
            
            if qty_excel != "1" and qty_excel != "nan":
                page.locator("input[name='qty']").first.fill(qty_excel)
                page.wait_for_timeout(200)
            
            if unit_lambung_excel and unit_lambung_excel != 'nan':
                input_unit = page.locator("input#unit_id-flexdatalist, input[name='flexdatalist-unit_id']").first
                input_unit.click()
                page.wait_for_timeout(200)
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                page.wait_for_timeout(200)
                page.keyboard.type(unit_lambung_excel, delay=150)
                target_preview_unit = page.locator("ul.flexdatalist-results, table").locator(f"li:has-text('{unit_lambung_excel}'), tr:has-text('{unit_lambung_excel}')").first
                target_preview_unit.wait_for(state="visible", timeout=5000)
                target_preview_unit.click()
                print("      [+] Sukses memilih Unit Pemakai!")
                page.wait_for_timeout(500)
            
            if lokasi_excel and lokasi_excel != 'nan':
                input_lokasi = page.locator("input#lokasi_id-flexdatalist, input[name='flexdatalist-lokasi_id']").first
                input_lokasi.click()
                page.wait_for_timeout(200)
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                page.wait_for_timeout(200)
                page.keyboard.type(lokasi_excel, delay=150)
                target_preview_lokasi = page.locator("ul.flexdatalist-results:visible, .flexdatalist-results:visible").locator(f"li:has-text('{lokasi_excel}'), a:has-text('{lokasi_excel}')").first
                target_preview_lokasi.wait_for(state="visible", timeout=5000)
                target_preview_lokasi.click()
                print("      [+] SUKSES KLIK opsi Lokasi Pemakaian!")
                page.wait_for_timeout(500)
            
            page.locator("textarea[name='keterangan'], input[name='keterangan']").first.fill(keterangan_otomatis)
            page.wait_for_timeout(500)
            page.locator("button:has-text('ADD'), #btn-add").first.click()
            page.wait_for_timeout(2500) 
            print("      [+] Data Berhasil di-ADD!")
            
            # --- LOG SUKSES (TAMBAHAN UNIT) ---
            log_hasil.append({
                "Tanggal": tanggal_sekarang, "No SBK": no_sbk_otomatis, "Item ID": item_id_excel, 
                "SN IN": sn_in, "SN OUT": sn_out, "Posisi": posisi_raw, "Unit": unit_lambung_excel, "Status": "Sukses"
            })

        except Exception as e:
            print(f"      [-] Gagal memproses baris ini -> {e}")
            # --- LOG GAGAL (TAMBAHAN UNIT) ---
            log_hasil.append({
                "Tanggal": tanggal_sekarang, "No SBK": "-", "Item ID": item_id_excel, 
                "SN IN": sn_in, "SN OUT": sn_out, "Posisi": posisi_raw, "Unit": unit_lambung_excel, "Status": "Gagal", "Pesan": str(e)
            })

        print("[*] 1 Data Selesai. Langsung membersihkan halaman untuk data berikutnya...")
        try:
            page.locator("button:has-text('CLEAR & NEW'), [id*='clear']").first.click()
            page.wait_for_timeout(2500)
        except:
            page.reload()
            page.wait_for_load_state("load")
            page.wait_for_timeout(3000)

    browser.close()

df_log = pd.DataFrame(log_hasil)
df_log.to_excel(FILE_LOG, index=False)
print(f"[+] Selesai total! Rekap pekerjaan sukses disimpan di: '{FILE_LOG}'")