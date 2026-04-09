import fitz  # PyMuPDF
import pandas as pd
import os

# 1. Klasör kontrolü
if not os.path.exists('sertifikalar'):
    os.makedirs('sertifikalar')

# 2. Fontu döngü dışında bir kez yükleyelim (Hesaplama yapabilmek için)
# Bu satır fontun ölçülerini kütüphaneye tanıtır
custom_font = fitz.Font(fontfile="LibreBaskerville.ttf")

# Verileri oku
df = pd.read_excel('deneme.xlsx')

for index, row in df.iterrows():
    # PDF taslağını aç
    doc = fitz.open("template.pdf")
    page = doc[0]  
    
    # SAYFA GENİŞLİĞİNİ AL (Bu satırı ekledik)
    sayfa_genisligi = page.rect.width
    
    isim = str(row['Column 2']).upper()
    font_boyutu = 35.2

    # PDF'e fontu göm
    page.insert_font(fontname="f1", fontfile="LibreBaskerville.ttf")

    # DOĞRU HESAPLAMA YÖNTEMİ:
    # get_text_length yerine yüklediğimiz font objesinin kendi metodunu kullanıyoruz
    metin_genisligi = custom_font.text_length(isim, fontsize=font_boyutu)

    # X KOORDİNATINI HESAPLA
    merkez_x = (sayfa_genisligi - metin_genisligi) / 2
    
    # Yüksekliği ayarla (Kırmızı artının olduğu yer)
    hedef_y = 307 

    # Yazıyı yerleştir
    page.insert_text(
        (merkez_x - 100, hedef_y),
        isim, 
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
