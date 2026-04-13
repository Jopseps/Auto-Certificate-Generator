import fitz  # PyMuPDF
import pandas as pd
import os

# 1. Klasör kontrolü
if not os.path.exists('sertifikalar'):
    os.makedirs('sertifikalar')

# Fontu yükle (Dizindeki dosya adıyla aynı olmalı)
custom_font = fitz.Font(fontfile="LibreBaskerville.ttf")

# Verileri oku
df = pd.read_excel('ROTALIST.xlsx')

for index, row in df.iterrows():
    doc = fitz.open("template.pdf")
    page = doc[0]  
    sayfa_genisligi = page.rect.width

    # İsmi temizle ve büyük harf yap
    arinmisIsim = str(row['Column 2']).replace('i', 'İ').replace('ç', 'Ç').replace('ö', 'Ö').replace('ü', 'Ü').replace('ş', 'Ş') 
    isim = arinmisIsim.upper().strip()

    font_boyutu = 35.2
    page.insert_font(fontname="f1", fontfile="LibreBaskerville.ttf")

    # --- GELİŞMİŞ SATIR BÖLME MANTIĞI ---
    satirlar = []
    kelimeler = isim.split()

    if len(isim) >= 19 and len(kelimeler) >= 2:
        # Eğer 3 veya daha fazla kelime varsa (örn: İbrahim Fatmanur Arkadaş)
        if len(kelimeler) >= 3:
            # İlk iki kelimeyi yukarı, son kelimeyi (soyismi) aşağı alalım
            satirlar.append(" ".join(kelimeler[:-1])) # "IBRAHIM FATMANUR"
            satirlar.append(kelimeler[-1])           # "ARKADAS"
        else:
            # 2 kelimeyse (örn: Dilara Mahmudoğlu) tam ortadan böl
            satirlar.append(kelimeler[0])
            satirlar.append(kelimeler[1])
        print(f"Bölündü: {satirlar}")
    else:
        satirlar = [isim]

    # --- YAZDIRMA AYARLARI ---
    hedef_y = 307
    satir_araligi = 45 # İki satır arasındaki boşluk
    
    for i, satir in enumerate(satirlar):
        # Satır genişliğini her satır için ayrı ölç
        metin_genisligi = custom_font.text_length(satir, fontsize=font_boyutu)
        
        # TAM MERKEZLEME (Eğer sola/sağa kayık dersen buradaki 0 ile oyna)
        # Önceki kodundaki -100'ü kaldırdım, çünkü o ismi sola itiyordu.
        merkez_x = ((sayfa_genisligi - metin_genisligi) / 2) - 100
        
        if len(satirlar) > 1:
            # İlk satırı 22.5 yukarı, ikinciyi 22.5 aşağı çeker
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

    # Kaydetme ve Dosya Çakışması Kontrolü
    temiz_isim = "".join(c for c in isim if c.isalnum() or c in (' ', '_')).rstrip()
    #cikti_yolu = f"sertifikalar/{temiz_isim.replace(' ', '_')}.pdf"
    cikti_yolu = f"sertifikalar/{temiz_isim}.pdf"
    try:
        # Eğer dosya zaten varsa önce siliyoruz (güncelleme garantisi)
        if os.path.exists(cikti_yolu):
            os.remove(cikti_yolu)
        doc.save(cikti_yolu)
        doc.close()
        print(f"Başarılı: {cikti_yolu}")
    except Exception as e:
        print(f"!!! HATA: {isim} dosyası kaydedilemedi. PDF açık olabilir! -> {e}")

print("\nBütün sertifikalar hazır!")
