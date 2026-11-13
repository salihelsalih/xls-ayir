import os
import sys
import time
import xlrd
from xlutils.copy import copy

def dosya_sec_dialog():
    """Eğer parametresiz çalıştırılırsa kullanıcıya dosya seçme penceresi açar."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        dosya = filedialog.askopenfilename(
            title="Ayrılacak XLS Dosyasını Seçin",
            filetypes=[("Excel 97-2003 Çalışma Kitabı (*.xls)", "*.xls"), ("Tüm Dosyalar", "*.*")]
        )
        root.destroy()
        return dosya
    except Exception:
        return None

def xls_sayfalarini_stilleriyle_ayir(dosya_yolu, cikis_klasoru=None):
    """
    Büyük bir Excel (.xls) dosyasındaki tüm sayfaları (sheet'leri),
    orijinal hücre stilleri, renkleri ve biçimlendirmeleriyle birlikte
    ayrı ayrı .xls dosyalarına böler.
    """
    dosya_yolu = os.path.abspath(dosya_yolu.strip('\"\''))

    if not os.path.exists(dosya_yolu):
        print(f"\n[HATA] Dosya bulunamadı: '{dosya_yolu}'")
        return False

    # Varsayılan çıkış klasörü: Dosyanın bulunduğu dizinde 'Ayrılmış' klasörü
    if not cikis_klasoru:
        dosya_klasoru = os.path.dirname(dosya_yolu)
        cikis_klasoru = os.path.join(dosya_klasoru, "Ayrılmış")

    os.makedirs(cikis_klasoru, exist_ok=True)

    baslangic = time.time()
    dosya_adi = os.path.basename(dosya_yolu)
    print("=" * 60)
    print(f" XLS AYIRICI - HÜCRE VE STİL KORUMALI SAYFA AYIRMA")
    print("=" * 60)
    print(f"Hedef Dosya    : {dosya_adi}")
    print(f"Dosya Konumu   : {dosya_yolu}")
    print(f"Kayıt Klasörü  : {cikis_klasoru}")
    print("-" * 60)
    print("Excel dosyası okunuyor (renkler, fontlar ve kenarlıklar yükleniyor)...")

    try:
        # formatting_info=True: tüm renk, kenarlık ve font stillerini korur
        rb = xlrd.open_workbook(dosya_yolu, formatting_info=True)
    except Exception as e:
        print(f"\n[HATA] Excel dosyası açılırken bir sorun oluştu:")
        print(f"       {e}")
        print("\nNot: Bu araç .xls (Excel 97-2003) formatındaki dosyaları tam stil desteğiyle ayırır.")
        return False

    sayfa_sayisi = rb.nsheets
    sayfa_adlari = rb.sheet_names()

    print(f"Toplam {sayfa_sayisi} sayfa bulundu. Ayırma işlemi başlıyor...\n")

    for index, sayfa_adi in enumerate(sayfa_adlari):
        t_sayfa = time.time()

        # 1. Tüm çalışma kitabını stilleriyle birlikte kopyala
        wb_kopya = copy(rb)

        # 2. Seçili sayfa dışındaki tüm sayfaları sondan başa doğru sil
        for i in range(sayfa_sayisi - 1, -1, -1):
            if i != index:
                wb_kopya._Workbook__worksheets.pop(i)

        # 3. xlwt aktif sayfa indeksini sıfırla (IndexError hatasını önler)
        wb_kopya._Workbook__active_sheet = 0

        # 4. Dosya adındaki Windows yasaklı karakterlerini temizle (\ / : * ? " < > |)
        gecersiz_karakterler = r'\/:*?"<>|'
        temiz_ad = "".join("_" if c in gecersiz_karakterler else c for c in sayfa_adi).strip()

        # Dosyaların sıralı durması için: 01_SayfaAdi.xls
        yeni_dosya_adi = f"{index + 1:02d}_{temiz_ad}.xls"
        yeni_dosya_yolu = os.path.join(cikis_klasoru, yeni_dosya_adi)

        # 5. Dosyayı kaydet
        wb_kopya.save(yeni_dosya_yolu)

        # İstatistik bilgileri
        boyut_kb = os.path.getsize(yeni_dosya_yolu) / 1024
        sheet_obj = rb.sheet_by_index(index)
        gecen = time.time() - t_sayfa

        print(f"[{index + 1:02d}/{sayfa_sayisi:02d}] [OK] {yeni_dosya_adi:<30} "
              f"({sheet_obj.nrows} satır, {sheet_obj.ncols} sütun | {boyut_kb:.1f} KB | {gecen:.2f}s)")

    toplam_sure = time.time() - baslangic
    print("-" * 60)
    print(f"BAŞARILI: Toplam {sayfa_sayisi} sayfa {toplam_sure:.2f} saniyede ayrıldı!")
    print(f"Kaydedilen Klasör: {os.path.abspath(cikis_klasoru)}")
    print("=" * 60)

    # Klasörü Windows Gezgini'nde otomatik aç
    try:
        os.startfile(cikis_klasoru)
    except Exception:
        pass

    return True

def main():
    # 1. Komut satırından dosya yolu parametresi geldi mi? (Sağ tık seçeneği %1 gönderir)
    if len(sys.argv) > 1:
        dosya_yolu = sys.argv[1]
    else:
        # Parametre gelmediyse dosya seçme penceresi aç
        print("Lütfen açılan pencereden ayrılacak .xls dosyasını seçin...")
        secilen = dosya_sec_dialog()
        if secilen:
            dosya_yolu = secilen
        else:
            # Kullanıcı pencereyi kapattıysa konsoldan sor
            dosya_yolu = input("\nAyrılacak .xls dosyasının tam yolunu girin: ").strip('\"\'')

    if not dosya_yolu:
        print("\nHiçbir dosya belirtilmedi. Çıkılıyor.")
        time.sleep(2)
        return

    try:
        basari = xls_sayfalarini_stilleriyle_ayir(dosya_yolu)
    except Exception as e:
        print(f"\n[BEKLENMEYEN HATA]: {e}")
        basari = False

    # Pencerenin hemen kapanmaması için kullanıcıya bildir
    print("\nİşlem tamamlandı. Çıkmak için Enter tuşuna basın...")
    try:
        input()
    except Exception:
        time.sleep(3)

if __name__ == "__main__":
    main()
