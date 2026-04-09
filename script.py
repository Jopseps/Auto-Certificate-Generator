import fitz  # PyMuPDF
import pandas as pd
import os

# 1. Klasör kontrolü
if not os.path.exists('sertifikalar'):
    os.makedirs('sertifikalar')

# Fontu yükle (Hesaplama için)
custom_font = fitz.Font(fontfile="LibreBaskerville.ttf")

# Verileri oku
df = pd.read_excel('ROTALIST.xlsx')

for index, row in df.iterrows():
    doc = fitz.open("template.pdf")
    page = doc[0]  
    sayfa_genisligi = page.rect.width
    
    # İsmi temizle ve büyük harf yap
    isim = str(row['Column 2']).upper().strip()
    
    # HATA AYIKLAMA: Python bu ismi kaç karakter sayıyor terminalde görelim
    print(f"İşleniyor: {isim} (Uzunluk: {len(isim)})")

    font_boyutu = 35.2
    page.insert_font(fontname="f1", fontfile="LibreBaskerville.ttf")
    
    # --- YENİLENMİŞ SATIR BÖLME MANTIĞI ---
    satirlar = []
    # 15 karakter ve üzeri ise ve içinde boşluk varsa böl
    if len(isim) >= 15 and " " in isim:
        kelimeler = isim.split() # split() çift boşlukları da halleder
        orta = len(kelimeler) // 2
        
        # Eğer 2 kelimeyse (örn: DİLARA MAHMUDOĞLU) orta 1 olur.
        # [:1] -> DİLARA, [1:] -> MAHMUDOĞLU olur.
        s1 = " ".join(kelimeler[:orta])
        s2 = " ".join(kelimeler[orta:])
        
        satirlar = [s1, s2]
        print(f"  -> BÖLÜNDÜ: '{s1}' ve '{s2}'")
    else:
        satirlar = [isim]

    # --- YAZDIRMA VE HİZALAMA ---
    hedef_y = 307
    satir_araligi = 45 # Satır arasını biraz daha açalım ki net görünsün
    
    for i, satir in enumerate(satirlar):
        # Satır genişliğini ölç
        metin_genisligi = custom_font.text_length(satir, fontsize=font_boyutu)
        
        # TAM MERKEZLEME: -100'ü kaldırdım, istersen küçük bir değer (-10 gibi) ekle
        merkez_x = ((sayfa_genisligi - metin_genisligi) / 2) - 100
        
        # Y koordinatı hesabı
        if len(satirlar) > 1:
            # İlk satırı 22 birim yukarı, ikinciyi 23 birim aşağı koyar (toplam 45)
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
    
    try:
        doc.save(cikti_yolu)
        doc.close()
    except Exception as e:
        print(f"  !! HATA: {cikti_yolu} kaydedilemedi. Dosya açık olabilir mi? {e}")

print("\nTüm işlemler bitti! Sertifikalar klasörünü kontrol edebilirsin.")
