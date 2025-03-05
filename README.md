# Autodesk Inventor Automation - Windows Forms Uygulaması

Bu proje, **Autodesk Inventor API** kullanarak Inventor'ı başlatan, yeni bir parça oluşturan, 2D Sketch ekleyerek bir şekil çizen ve ardından projeyi kaydedip kapatan bir **Windows Forms (WinForms) uygulamasıdır**.

---

## 🚀 **Kurulum ve Çalıştırma**

### **1️⃣ Gereksinimler**
- **Windows 10/11** işletim sistemi
- **Visual Studio 2022** (veya daha yeni)
- **.NET 8.0 SDK**
- **Autodesk Inventor** (2022 ve sonrası tavsiye edilir)

---

### **2️⃣ Projeyi Çalıştırmadan Önce Yapılması Gerekenler**

#### **📌 Autodesk Inventor COM Referanslarını Ekleyin**
Autodesk Inventor API’si bir **COM bileşeni** olarak çalışır. **NuGet paketi yoktur**, eğer proje içinde ekli gözükmüyorsa manuel olarak eklemeniz gerekir:

1. **Solution Explorer (Çözüm Gezgini)** → Projeye sağ tıklayın → **"Add" → "Reference..."** seçeneğini seçin.
2. Açılan pencerede **"COM" sekmesine** gidin.
3. **"Autodesk Inventor Object Library"** seçeneğini bulun ve işaretleyin.
4. **OK** butonuna basarak projeye ekleyin.

---

### **3️⃣ Uygulamayı Çalıştırma**

#### **📌 Projeyi Derleme ve Başlatma**
1. **Visual Studio'da projeyi açın.**
2. **Projeyi derleyin (Ctrl + Shift + B).**
3. **Uygulamayı başlatın (Ctrl + F5).**
4. **Arayüzde bulunan butonları sırasıyla kullanarak Autodesk Inventor üzerinde işlem yapabilirsiniz.**

#### **📌 Butonların İşlevleri**
- **"Autodesk Inventor Başlat"** → Autodesk Inventor'ı başlatır.
- **"Yeni Parça Oluştur"** → Inventor'da yeni bir parça (`.ipt` dosyası) oluşturur.
- **"Dikdörtgen Çiz"** → Yeni bir 2D Sketch ekler ve oluşturulan bu Sketch içine dikdörtgen çizer.
- **"Üçgen Çiz"** → Yeni bir 2D Sketch ekler ve oluşturulan bu Sketch içine üçgen çizer.
- **"Daire Çiz"** → Yeni bir 2D Sketch ekler ve oluşturulan bu Sketch içine daire çizer.
- **"Beşgen Çiz"** → Yeni bir 2D Sketch ekler ve oluşturulan bu Sketch içine beşgen çizer.
- **"Altıgen Çiz"** → Yeni bir 2D Sketch ekler ve oluşturulan bu Sketch içine altıgen çizer.
- **"Kaydet ve Kapat"** → Dosyayı `C:\Temp\InventorPart_YYYYMMDD_HHMMSS.ipt` olarak kaydeder ve kapatır.


