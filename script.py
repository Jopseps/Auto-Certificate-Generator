import fitz  # PyMuPDF
import pandas as pd
import os

# Klasör kontrolü
if not os.path.exists('sertifikalar'):
    os.makedirs('sertifikalar')

# Verileri oku
df = pd.read_excel('deneme.xlsx')

for index, row in df.iterrows():
    # 1. PDF taslağını aç
    doc = fitz.open("template.pdf")
    page = doc[0]  # İlk sayfa
    
    isim = str(row['Column 2']).upper()

    # 2. Fontu kaydet (Dışarıdan font kullanmak için)
    # font_yolu daha önceki LibreBaskerville.ttf dosyan olsun
    # 'fontname' kısmına istediğin bir isim verebilirsin (örn: "f1")
    page.insert_font(fontname="f1", fontfile="LibreBaskerville.ttf")

    # 2. İSİM GENİŞLİĞİNİ HESAPLA (Kaymayı önleyen kritik adım)
    # Bu fonksiyon, yazdığın ismin o font ve boyutta kaç "point" yer kapladığını söyler.
    metin_genisligi = fitz.get_text_length(isim, fontname="f1", fontsize=35.2)

    # 3. X KOORDİNATINI HESAPLA
    # (Sayfa Genişliği - Metin Genişliği) / 2 yaparak metni tam yatay merkeze alırız.
    merkez_x = (sayfa_genisligi - metin_genisligi) / 2
    
    # Senin belirlediğin yükseklik (Y değeri)
    # Not: Resimdeki 307 değeri PDF'de biraz farklı bir yere düşebilir, 
    # gerekirse bu rakamı yukarı-aşağı (örn: 280 veya 350) kaydırarak dene.
    hedef_y = 307
    page.insert_text(
        (merkez_x, hedef_y),           # Koordinatlar (x, y)
        isim, 
        fontname="f1", 
        fontsize=35.2, 
        color=(0.003, 0.105, 0.329)
    )

    # 4. Kaydet
    temiz_isim = "".join(c for c in isim if c.isalnum() or c in (' ', '_')).rstrip()
    cikti_yolu = f"sertifikalar/{temiz_isim.replace(' ', '_')}.pdf"
    
    doc.save(cikti_yolu)
    doc.close()
    print(f"Oluşturuldu: {cikti_yolu}")

print("Tüm işlemler bitti!")
