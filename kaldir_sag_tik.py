import os
import sys
import winreg

def sag_tik_menusu_kaldir():
    kayit_noktalari = [
        r"Software\Classes\SystemFileAssociations\.xls\shell\XlsAyir",
        r"Software\Classes\Excel.Sheet.8\shell\XlsAyir",
        r"Software\Classes\.xls\shell\XlsAyir"
    ]

    print("=" * 60)
    print(" XLS AYIRICI - SAĞ TIK MENÜSÜ TEMİZLEME")
    print("=" * 60)

    for anahtar_yolu in kayit_noktalari:
        try:
            # Önce alt anahtar olan 'command' silinir
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, anahtar_yolu + r"\command")
            except FileNotFoundError:
                pass

            # Sonra anahtar silinir
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, anahtar_yolu)
            print(f"[SİLİNDİ] HKCU\\{anahtar_yolu}")
        except FileNotFoundError:
            print(f"[ZATEN YOK] HKCU\\{anahtar_yolu}")
        except Exception as e:
            print(f"[HATA] {anahtar_yolu}: {e}")

    print("=" * 60)
    print("Sağ tık menü kayıtları başarıyla kaldırıldı.")
    print("=" * 60)

if __name__ == "__main__":
    sag_tik_menusu_kaldir()
    if len(sys.argv) == 1:
        input("\nKapatmak için Enter'a basın...")
