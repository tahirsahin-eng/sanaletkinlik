# Saffet Çam Ortaokulu — Uzay Çadırı ve Sanal Gerçeklik (VR) Etkinliği

T.C. Yalova Valiliği • Merkez / Saffet Çam Ortaokulu Müdürlüğü (Aralık 2026)

## Mimari Yapı (GitHub Pages + Google Drive / Google E-Tablolar)

Bu proje, veli başvuru ekranı ile sayman yönetim panelini birbirinden ayırarak güvenli ve bulut tabanlı çalışacak şekilde yapılandırılmıştır:

1. **`index.html` (Veli Başvuru Formu):** Velilerin sınıf ve öğrenci seçimi yapıp 13 içerik arasından 5 adet VR içeriğini seçerek başvuru yaptığı herkese açık sayfadır. Yönetici listesi veya ödeme onay butonları bu sayfada yer almaz.
2. **`sayman.html` (Sayman Yönetim Paneli):** Okul Aile Birliği Saymanı **Fatma Erdem** için ayrılmış yönetim ve tahsilat ekranıdır. Şifre doğrulaması ile açılır; Google E-Tablolar'daki kayıtları listeler, 150 TL ödeme onaylarını günceller ve Google Drive'a CSV yedeği alır.
3. **`data.js` (Ortak Veri Modülü):** 13 VR içeriği ve 23 şubelik (676 öğrenci) güncel okul listesini barındırır.
4. **`Code.gs` (Google Apps Script):** Google Drive üzerindeki E-Tabloya veri yazma (`doPost`), sayman doğrulama (`doGet`) ve Google Drive klasörüne otomatik CSV yedekleme (`backup_drive`) işlemlerini yürütür.

## Kurulum Adımları

1. **Google Drive & Apps Script:**
   - Google Drive'da yeni bir **Google E-Tablo** oluşturun.
   - **Uzantılar > Apps Script** menüsünü açıp `Code.gs` içeriğini yapıştırın.
   - **Dağıt > Yeni Dağıtım > Web Uygulaması** (*Çalıştıran: Ben*, *Erişimi olanlar: Herkes*) olarak yayınlayın.
2. **GitHub Pages Yayını:**
   - `index.html`, `sayman.html`, `data.js` ve `README.md` dosyalarını GitHub deponuza yükleyin.
   - **Settings > Pages** bölümünden `main` dalını seçerek yayına alın.
