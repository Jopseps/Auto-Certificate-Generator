import fitz  # PyMuPDF
import pandas as pd
import os

# 1. Klasör kontrolü
if not os.path.exists('sertifikalar'):
    os.makedirs('sertifikalar')

# Fontu yükle
custom_font = fitz.Font(fontfile="LibreBaskerville.ttf")

# Verileri oku
df = pd.read_excel('ROTALIST.xlsx')

for index, row in df.iterrows():
    doc = fitz.open("template.pdf")
    page = doc[0]  
    sayfa_genisligi = page.rect.width
    
    isim = str(row['Column 2']).upper().strip()
    font_boyutu = 35.2
    page.insert_font(fontname="f1", fontfile="LibreBaskerville.ttf")
    
    # --- SATIR BÖLME MANTIĞI ---
    satirlar = []
    if len(isim) >= 15 and " " in isim:
        # İsmi boşluklardan böl
        kelimeler = isim.split(" ")
        orta_nokta = len(kelimeler) // 2
        # Kelimeleri iki gruba ayır
        satirlar.append(" ".join(kelimeler[:orta_nokta + (1 if len(kelimeler) % 2 != 0 else 0)]))
        satirlar.append(" ".join(kelimeler[orta_nokta + (1 if len(kelimeler) % 2 != 0 else 0):]))
    else:
        # İsim kısa ise tek satır
        satirlar.append(isim)

    # --- YAZDIRMA VE HİZALAMA ---
    hedef_y = 307
    # İki satır varsa, ilk satırı biraz yukarı, ikinciyi biraz aşağı almalıyız
    satir_araligi = 40 # İki satır arasındaki dikey mesafe (point)
    
    for i, satir in enumerate(satirlar):
        # Her satırın genişliğini ayrı hesapla (Ortalama için şart)
        metin_genisligi = custom_font.text_length(satir, fontsize=font_boyutu)
        
        # Tam merkezi bul (Kendi kodundaki -100 offsetini korudum ama gerekmiyorsa kaldırabilirsin)
        merkez_x = ((sayfa_genisligi - metin_genisligi) / 2) - 100
        
        # Eğer iki satır varsa, y koordinatını ayarla
        if len(satirlar) > 1:
            # İlk satırı hedef_y'den biraz yukarı, ikinciyi biraz aşağı basar
            guncel_y = hedef_y - (satir_araligi / 2) + (i * satir_araligi)
        else:
            guncel_y = hedef_y

        page.insert_text(
            (merkez_x, guncel_y),
            satir,
            fontname="f1",
            fontsize=font_boyutu,
            color=(0.003, 0.105, 0.329)
        )

    # Kaydet
    temiz_isim = "".join(c for c in isim if c.isalnum() or c in (' ', '_')).rstrip()
    cikti_yolu = f"sertifikalar/{temiz_isim.replace(' ', '_')}.pdf"
    
    doc.save(cikti_yolu)
    doc.close()
    print(f"Oluşturuldu: {cikti_yolu}")

print("Tüm işlemler bitti!")
