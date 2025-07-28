# ServiceDesk & ServiceDesk MSP Importer 🚀

## English

### 🎉 Description

A super-friendly Python 3 GUI tool to import tickets from an Excel sheet into ManageEngine ServiceDesk MSP via the V3 API using your Technician Key 🔑. Watch the magic happen with a progress window, live logs, and detailed console/file (`app.log`) logging!

**Key Features:**

* 📥 **Excel → JSON**: Reads your rows and converts to JSON payloads.

  ### 📝 Excel Sheet Structure

Prepare your Excel file with the following columns (headers are case-sensitive):

| Column Name          | Type      | API Field       | Notes                                             | Mandatory                       |
| -------------------- | --------- | --------------- | ------------------------------------------------- | ------------------------------- |
| Requester            | Text      | requester       | Name or email of the requester                    | Yes                             |
| Site                 | Text      | site            | Site name                                         | No                              |
| Account              | Text      | account         | Account identifier (mandatory in ServiceDesk MSP) | Yes (SDP-MSP)                   |
| Subject              | Text      | subject         | Short summary                                     | Yes                             |
| Item                 | Text      | item            | Item name                                         | No                              |
| Request Type         | Text      | request\_type   | Request type                                      | No                              |
| Level                | Text      | level           | Level                                             | No                              |
| Urgency              | Text      | urgency         | Urgency                                           | No                              |
| Impact               | Text      | impact          | Impact                                            | No                              |
| Technician           | Text      | technician      | Technician name                                   | No                              |
| Category             | Text      | category        | Category                                          | No                              |
| Subcategory          | Text      | subcategory     | Subcategory                                       | No                              |
| Priority             | Text      | priority        | Priority                                          | No                              |
| Group                | Text      | group           | Group name                                        | No                              |
| Status               | Text      | status          | Status                                            | No                              |
| Template             | Text      | template        | Template name                                     | No                              |
| Description          | Text      | description     | Detailed description                              | No                              |
| Resolution           | Text      | resolution      | Resolution content                                | No (Yes if Status is Closed)    |
| Created Time         | Date Text | created\_time   | Format `DD/MM/YYYY HH:MM AM/PM`                   | No (defaults to today if empty) |
| Completed Time       | Date Text | udf\_sline\_301 | Format `DD/MM/YYYY HH:MM AM/PM` (optional)        | No                              |
| Resolved Time        | Date Text | udf\_sline\_302 | Format `DD/MM/YYYY HH:MM AM/PM` (optional)        | No                              |
| udf\_sline\_\* fields| Text      | udf\_sline\_\*  | Custom string sline fields                        | No                              |
| udf\_date_\_\* fields| timestamp | udf\_date\_\*   | Custom timestamp value fields                     | No                              |
| udf\_pick\_\* fields | Text      | udf\_pick\_\*   | Custom picklist fields                            | No                              |

> **Note:** Ensure your Excel has no empty header rows, dates match the format exactly, and if you fill the `Template` column, all fields required by that template must be populated. This section serves as a template to build your sheet before running the importer.

* ❌ **Skip 'Not Assigned'**: Automatically removes any fields with `Not Assigned`.

* ⏰ **Date Parsing**: Converts `DD/MM/YYYY HH:MM AM/PM` dates to Unix ms timestamps.

* 🔄 **Field Mapping**:

  * **Standard**: requester, site, account, subject, description, resolution…

  * **UDFs**: If you want to preserve the original Completed Time and Resolved Time values from another ServiceDesk product, you can assign them to an additional field (we used `udf_sline_301` and `udf_sline_302`). You can view the transfer of these values at the database level in the `Update_time_query.sql` file.

  * **Picklist UDFs** (`udf_pick_*`): For certain ServiceDesk products, values should be entered with the `{'name': ...}` wrapper; for others, you can assign directly as `udf_pick_*: "..."`. If your ServiceDesk structure doesn’t accept the `{'name'}` format, you can comment out lines 151–154.

* 🚑 **Error Handling**: On HTTP 400 group errors, retries without the group field.

* 🔐 **SSL Off**: Disables SSL verification & suppresses warnings (for dev/self-signed).

* ⚙️ **Configurable Logs**: Set `LOG_FILE`, `LOG_LEVEL`, `ERROR_LOG_FILE` via env vars.

### 📋 Prerequisites

* Python 3.7+
* `requests`, `pandas`, `openpyxl`, `tkinter`

```bash
pip install requests pandas openpyxl
```

### ⚙️ Configuration

1. **ServiceDesk Domain** (e.g. `example.manageengine.com:8080`)
2. **Technician Key** (your API key)
3. Optional env vars:

   * `LOG_FILE` (default: `app.log`)
   * `LOG_LEVEL` (default: `INFO`)
   * `ERROR_LOG_FILE` (default: `error.log`)

### 🚀 Usage

```bash
./importer.py
```

1. Enter domain & key in the GUI.
2. Choose your Excel file.
3. Enjoy the progress bar & live logs.
4. See a summary of successes & failures at the end 🎉

---

## Türkçe

### 🎉 Açıklama

Python 3 ile yazılmış samimi bir GUI aracı! Excel tablonuzdaki kayıtları ManageEngine ServiceDesk MSP'ye V3 API ile aktarır. Technician Key 🔑 ile giriş yapın, ilerleme çubuğu ve canlı log'lar eşliğinde transferin tadını çıkarın; detaylar konsol ve `app.log` dosyasında saklansın.

**Özellikler:**

* 📥 **Excel → JSON**: Satırları okur, JSON'a dönüştürür.

  ### 📝 Excel Sayfa Yapısı

Excel dosyanızı aşağıdaki sütunlarla hazırlayın (başlıklar büyük/küçük harfe duyarlıdır):

| Sütun Adı              | Tür         | API Alanı       | Açıklama                                                   | Zorunluluk             |
| ---------------------- | ----------- | --------------- | --------------------------------------------               | ---------------------- |
| Requester              | Metin       | requester       | İstek sahibinin adı veya e-posta adresi                    | Evet         |
| Site                   | Metin       | site            | Site adı                                                   | Hayır        |
| Account                | Metin       | account         | Hesap kimliği (ServiceDesk MSP Ortamı için zorunludur.)    | Evet  (SDP-MSP) |
| Subject                | Metin       | subject         | Kısa özet                                                  | Evet         |
| Item                   | Metin       | item            | Öğe adı                                                    | Hayır        |
| Request Type           | Metin       | request\_type   | İstek türü                                                 | Hayır        |
| Level                  | Metin       | level           | Seviye                                                     | Hayır        |
| Urgency                | Metin       | urgency         | Aciliyet                                                   | Hayır        |
| Impact                 | Metin       | impact          | Etki                                                       | Hayır        |
| Technician             | Metin       | technician      | Tekniker adı                                               | Hayır        |
| Category               | Metin       | category        | Kategori                                                   | Hayır        |
| Subcategory            | Metin       | subcategory     | Alt kategori                                               | Hayır        |
| Priority               | Metin       | priority        | Öncelik                                                    | Hayır        |
| Group                  | Metin       | group           | Grup adı                                                   | Hayır        |
| Status                 | Metin       | status          | Durum                                                      | Hayır        |
| Template               | Metin       | template        | Şablon adı                                                 | Hayır        |
| Description            | Metin       | description     | Detaylı açıklama                                           | Hayır        |
| Resolution             | Metin       | resolution      | Çözüm içeriği                                              | Hayır (Status eğer Closed ise Evet)       |
| Created Time           | Tarih Metin | created\_time   | `DD/MM/YYYY SS:DD AM/PM` formatı                           | Hayır (Girilmezse o gün oluşturulur) |
| Completed Time         | Tarih Metin | udf\_sline\_301 | `DD/MM/YYYY SS:DD AM/PM` formatı (opsiyonel)               | Hayır        |
| Resolved Time          | Tarih Metin | udf\_sline\_302 | `DD/MM/YYYY SS:DD AM/PM` formatı (opsiyonel)               | Hayır        |
| udf\_sline\_\* alanları| Metin       | udf\_sline\_\*  | Özel string metin alanları                                 | Hayır        |
| udf\_date\_\* alanları | timestamp   | udf\_date\_\*   | Özel timestamp veri alanları                               | Hayır        |
| udf\_pick\_\* alanları | Metin       | udf\_pick\_\*   | Özel picklist alanları                                     | Hayır        |

> **Not:** Excel dosyanızda boş başlık satırı olmadığından emin olun, tarih formatının tam eşleştiğini kontrol edin ve `Template` sütununa veri girerseniz o şablon için zorunlu olan tüm alanların dolu olduğundan emin olun. Bu bölüm, importer’ı çalıştırmadan önce sayfanızı oluşturmanız için bir şablon görevi görür.

* ❌ **Not Assigned?**: `Not Assigned` değerli alanlar atlanır.

* ⏰ **Tarih Dönüşümü**: `GG/AA/YYYY SS:DD AM/PM` formatı Unix ms'e çevrilir.

* 🔄 **Alan Eşleme**:

  * **Standart**: requester, site, account, subject, description, resolution…

  * **UDF**: Eğer diğer ServiceDesk ürününüzdeki Completed Time ve Resolved Time değerlerini korumak isterseniz, onları ek bir alana atayabilirsiniz (biz `udf_sline_301` ve `udf_sline_302` kullandık). Bu değerlerin aktarımını veritabanı seviyesinde `Update_time_query.sql` dosyasından görüntüleyebilirsiniz.

  * **Picklist UDF'ler** (`udf_pick_*`): Bazı ServiceDesk ürünleri için değerler `{'name': ...}` formatında girilmeli; bazıları için doğrudan `udf_pick_*: "..."` olarak atayabilirsiniz. ServiceDesk yapısı `{'name'}` formatını kabul etmiyorsa 151–154. satırları yorum satırı olarak işaretleyebilirsiniz.

* 🚑 **Hata Yönetimi**: HTTP 400 grup hatasında grup alanı çıkarılarak tekrar dener.

* 🔐 **SSL Kapatıldı**: SSL doğrulama yok, uyarılar bastırıldı (geliştirme için).

* ⚙️ **Log Ayarları**: `LOG_FILE`, `LOG_LEVEL`, `ERROR_LOG_FILE` env ile kontrol edin.

### 📋 Gereksinimler

* Python 3.7+
* `requests`, `pandas`, `openpyxl`, `tkinter`

```bash
pip install requests pandas openpyxl
```

### ⚙️ Yapılandırma

1. **Domain** (örn. `example.manageengine.com:8080`)
2. **Technician Key** (API anahtarınız)
3. Opsiyonel env:

   * `LOG_FILE` (varsayılan: `app.log`)
   * `LOG_LEVEL` (varsayılan: `INFO`)
   * `ERROR_LOG_FILE` (varsayılan: `error.log`)

### 🚀 Kullanım

```bash
./importer.py
```

1. Domain & key girin.
2. Excel dosyanızı seçin.
3. İlerleme ve log'ların keyfini çıkarın.
4. Sonunda başarı ve başarısız satır sayısını görün 🎉

---

## 🐍 Virtual Environment & Standalone Execution

**Create and Activate Virtualenv**

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

**Build Standalone Executable**
Use PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile importer.py
```

The executable will be in `dist/importer`. No Python needed on target.

**Deploy to Server Without Python**
Copy the generated executable and run it on the server. If on Linux, build on a similar distro or use Docker.

---

## 🐍 Sanal Ortam ve Bağımsız Çalıştırma

**Sanal Ortam Oluşturma ve Aktifleştirme**

```bash
python3 -m venv venv
source venv/bin/activate  # Windows için: venv\Scripts\activate
pip install -r requirements.txt
```

**Bağımsız Executable Oluşturma**
PyInstaller kullanın:

```bash
pip install pyinstaller
pyinstaller --onefile importer.py
```

Çıktı `dist/importer` klasöründe. Hedefte Python gerekmez.

**Sunucuya Python Olmadan Dağıtım**
Oluşturulan executable'ı kopyalayın ve sunucuda çalıştırın. Linux için benzer bir distro üzerinde derleyin veya bir Docker imajı kullanın.
