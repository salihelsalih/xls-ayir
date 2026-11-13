import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

klasor = os.path.dirname(os.path.abspath(__file__))
girdi_dosya = os.path.join(klasor, "sadece_ihtiyac_olan.xlsx")
html_cikti = os.path.join(klasor, "kusursuz_tablo.html")

# İstediğiniz alan
baslangic_satir, bitis_satir = 1, 85
baslangic_sutun, bitis_sutun = 1, 81  # 81 = CC Sütunu

print("Excel okunuyor ve görsel biçimlendirme aktarılıyor...")
wb = load_workbook(girdi_dosya, data_only=True)
ws = wb.active

html_content = """
<html>
<head>
<meta charset="utf-8">
<style>
    table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px; }
    td, th { border: 1px solid #ccc; padding: 4px 6px; white-space: nowrap; }
</style>
</head>
<body>
<table>
"""

for r in range(baslangic_satir, bitis_satir + 1):
    html_content += "<tr>"
    for c in range(baslangic_sutun, bitis_sutun + 1):
        cell = ws.cell(row=r, column=c)
        val = "" if cell.value is None else str(cell.value)
        
        # Hücre Arka Plan Rengi Alma
        bg_color = ""
        if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
            rgb = str(cell.fill.start_color.rgb)
            if len(rgb) == 8: # ARGB formatı
                rgb = rgb[2:]
            if len(rgb) == 6:
                bg_color = f"background-color: #{rgb};"

        # Yazı Tipi Kalınlık / Renk
        font_style = ""
        if cell.font:
            if cell.font.bold:
                font_style += "font-weight: bold;"
            if cell.font.color and cell.font.color.rgb:
                fc = str(cell.font.color.rgb)
                if len(fc) == 8: fc = fc[2:]
                if len(fc) == 6: font_style += f"color: #{fc};"

        style = f"style='{bg_color} {font_style}'"
        html_content += f"<td {style}>{val}</td>"
    html_content += "</tr>"

html_content += """
</table>
</body>
</html>
"""

with open(html_cikti, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"\nİşlem 2 saniyede tamamlandı! '{html_cikti}' dosyası oluşturuldu.")