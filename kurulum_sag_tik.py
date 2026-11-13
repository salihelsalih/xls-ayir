import os
import sys
import winreg

def sag_tik_menusu_ekle():
    script_dizini = os.path.dirname(os.path.abspath(__file__))
    ayir_py_yolu = os.path.join(script_dizini, "ayir.py")
    python_exe = sys.executable

    # Komut: python "ayir.py" "%1"
    komut = f'"{python_exe}" "{ayir_py_yolu}" "%1"'
    menu_metni = "XLS Sayfalarını Ayır"
    ikon = f'"{python_exe}",0'

    # Kayıt defteri yolları
    kayit_noktalari = [
        r"Software\Classes\SystemFileAssociations\.xls\shell\XlsAyir",
        r"Software\Classes\Excel.Sheet.8\shell\XlsAyir",
        r"Software\Classes\.xls\shell\XlsAyir"
    ]

    print("=" * 60)
    print(" XLS AYIRICI - SAĞ TIK MENÜSÜ ENTEGRASYONU")
    print("=" * 60)
    print(f"Python Yolu  : {python_exe}")
    print(f"Script Yolu  : {ayir_py_yolu}")
    print(f"Komut        : {komut}")
    print("-" * 60)

    for anahtar_yolu in kayit_noktalari:
        try:
            # Anahtar oluştur veya aç
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, anahtar_yolu) as anahtar:
                winreg.SetValueEx(anahtar, "", 0, winreg.REG_SZ, menu_metni)
                winreg.SetValueEx(anahtar, "Icon", 0, winreg.REG_SZ, ikon)

            # Alt komut anahtarı
            komut_anahtar_yolu = anahtar_yolu + r"\command"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, komut_anahtar_yolu) as komut_anahtar:
                winreg.SetValueEx(komut_anahtar, "", 0, winreg.REG_SZ, komut)

            print(f"[OK] Kayıt eklendi: HKCU\\{anahtar_yolu}")
        except Exception as e:
            print(f"[HATA] {anahtar_yolu}: {e}")

    print("=" * 60)
    print("Tebrikler! Sağ tık menüsü başarıyla eklendi.")
    print("Artık herhangi bir .xls dosyasına sağ tıklayıp:")
    print("-> 'XLS Sayfalarını Ayır' seçeneğini seçebilirsiniz.")
    print("=" * 60)

if __name__ == "__main__":
    sag_tik_menusu_ekle()
    if len(sys.argv) == 1:
        input("\nKapatmak için Enter'a basın...")
