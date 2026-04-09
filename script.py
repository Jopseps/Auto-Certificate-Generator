from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os

# 1. Klasörü kontrol et, yoksa oluştur
if not os.path.exists('sertifikalar'):
    os.makedirs('sertifikalar')

# 2. Verileri oku (Dosya adının doğru olduğundan emin ol)
# Not: openpyxl yüklü olmalı (pip install openpyxl)
df = pd.read_excel('deneme.xlsx') 

# 3. Font ayarları
# Font dosyasının kodun olduğu klasörde olduğundan emin ol
font_yolu = "LibreBaskerville.ttf" 
font_boyutu = 60
font = ImageFont.truetype(font_yolu, font_boyutu)

for index, row in df.iterrows():
    # Taslağı her seferinde yeniden açıyoruz
    img = Image.open("template.png")
    draw = ImageDraw.Draw(img)
    
    isim = str(row['Column 2']).upper() # İsmi büyük harfe çevir
    
    # --- İSİMLERİ ORTALAMAK İÇİN ---
    # Resmin genişliğini al
    W, H = img.size
    # Yazının kapladığı alanı hesapla (ortalamak için)
    bbox = draw.textbbox((0, 0), isim, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    
    # Koordinatlar: (Genişlik/2 - YazıGenişliği/2) yaparak tam yatay merkeze koyarız
    # Y koordinatını (400) kendi görseline göre yukarı-aşağı ayarlayabilirsin
    x_koordinati = ((W - w) / 2) - 200
    y_koordinati = 620 
    
    # Yazıyı yazdır
    draw.text((x_koordinati, y_koordinati), isim, fill="#011b54", font=font)
    
    # PDF olarak kaydet
    rgb_img = img.convert('RGB')
    
    # Dosya ismindeki boşlukları alt tire yapalım ve hatalı karakterleri temizleyelim
    temiz_isim = "".join(c for c in isim if c.isalnum() or c in (' ', '_')).rstrip()
    dosya_adi = f"sertifikalar/{temiz_isim.replace(' ', '_')}.pdf"
    
    try:
        rgb_img.save(dosya_adi, "PDF", resolution=100.0)
        print(f"Başarıyla oluşturuldu: {dosya_adi}")
    except Exception as e:
        print(f"Hata oluştu ({isim}): {e}")

print("İşlem tamam!") 
