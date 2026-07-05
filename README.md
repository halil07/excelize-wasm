# @halil07/excelize

[![npm version](https://img.shields.io/npm/v/@halil07/excelize.svg)](https://www.npmjs.com/package/@halil07/excelize)
[![License](https://img.shields.io/badge/license-BSD--3--Clause-blue.svg)](LICENSE)

Tarayıcıda ve Node.js ortamında, WebAssembly sayesinde Microsoft Excel (`XLSX` / `XLSM` / `XLTM` / `XLTX`) dosyalarını okuyup yazmanızı sağlayan bir npm paketidir.

Bu paket, Go dilinde yazılmış güçlü [excelize](https://github.com/xuri/excelize) kütüphanesinin WebAssembly bağlamasıdır. Excelize kaynağı doğrudan GitHub’dan (`github.com/xuri/excelize/v2`) alınır; bu repo yalnızca **WASM bağlama katmanını** (`wasm/excelize_wasm.go`) ve kullanımı kolaylaştıran **JavaScript yükleyicisini** (`wasm/excelize.js`) içerir.

---

## WebAssembly Mimarisi

`@halil07/excelize`, Go kaynak kodunu `GOOS=js GOARCH=wasm` hedefiyle derleyerek `dist/excelize.wasm` dosyasını üretir. JavaScript tarafında ise `wasm/excelize.js` yükleyicisi, bu `.wasm` modülünü çalışma zamanına yükler ve Go tarafında dışa aktarılan ham API’yi (`ExcelizeWasm`) kullanıcı dostu, hata fırlatan, sözde senkron bir JavaScript API’sine dönüştürür.

### Nasıl Çalışır?

1. **Go tarafı (`wasm/excelize_wasm.go`)**
   - `syscall/js` paketi kullanılarak Excelize fonksiyonları JavaScript dünyasına açılır.
   - Oluşturulan her `*excelize.File` nesnesi, sayısal bir tutamaç (handle) ile kaydedilir; `close(id)` çağrısıyla bellekten silinir.
   - Her fonksiyon başarı durumunda `{ ok: true, data: ... }`, hata durumunda `{ ok: false, err: "..." }` şeklinde sonuç döner.
   - Tarayıcıda geçici dosya oluşturulamayacağı için disk tabanlı `OpenFile` / `SaveAs` API’leri bağlanmamıştır; bunun yerine bellek içi `openFile(Uint8Array)` ve `save()` kullanılır.
   - Büyük dosyaların geçici dosya yoluna düşmesini önlemek amacıyla `UnzipSizeLimit` ve `UnzipXMLSizeLimit` değerleri ~1 TB olarak sabitlenir.

2. **JavaScript tarafı (`wasm/excelize.js`)**
   - Node.js’te `dist/excelize.wasm` ve `dist/wasm_exec.js` dosyalarını otomatik olarak yükler.
   - Tarayıcıda sayfanın `wasm_exec.js`’i yükleyip `WebAssembly.instantiateStreaming` ile modülü başlatmasını bekler.
   - `Excelize.excelizeInit()` çağrısı tamamlandıktan sonra `Excelize.newFile()`, `Excelize.openFile(bytes)` ve diğer yöntemler kullanılabilir.
   - Ham `{ ok, data }` / `{ ok: false, err }` yapısı `Excelize.raw` üzerinden hâlâ erişilebilir; üst düzey API ise hataları `throw` eder.

3. **Derleme çıktıları (`dist/`)**

   | Dosya | Açıklama |
   |-------|----------|
   | `excelize.wasm` | Go kaynağının WASM derlemesi |
   | `wasm_exec.js` | Go çalışma zamanı desteği (Go toolchain’inden kopyalanır) |
   | `excelize.js` | UMD yükleyici (tarayıcı + eski CJS) |
   | `excelize.cjs` | CommonJS yükleyici |
   | `excelize.mjs` | ES Modül sarmalayıcı |
   | `excelize.d.ts` / `.d.mts` / `.d.cts` | TypeScript tanımları |

### WebAssembly Özellikleri

- **Çapraz platform:** Aynı `.wasm` dosyası hem tarayıcıda hem Node.js’te çalışır.
- **Güvenli bellek yönetimi:** Dosya tutamaçları açıkça `close()` ile kapatılır; Go tarafındaki `sync.Mutex` ile eşzamanlı erişim korunur.
- **JSON köprüsü:** Karmaşık seçenek yapıları (`Style`, `Table`, `DataValidation`, `Chart`, …) düz JS nesneleri olarak gönderilir, Go tarafında JSON olarak çözümlenir.
- **Uint8Array köprüsü:** Excel dosyaları ve resimler `Uint8Array` / `ArrayBuffer` olarak aktarılır; kopyalama `js.CopyBytesToGo` / `js.CopyBytesToJS` ile yapılır.
- **Tarayıcı uyumlu:** Geçici dosya kullanımı tamamen devre dışı bırakılmıştır; tüm işlemler bellekte gerçekleşir.

---

## Özellikler

- Bellekte yeni çalışma kitapları oluşturma ve `Uint8Array` olarak dışa aktarma.
- Mevcut `XLSX` dosyalarını `Uint8Array` / `ArrayBuffer` ile açma.
- Hücre okuma/yazma, formül, satır ve sütun işlemleri.
- Stil, koşullu biçimlendirme, birleştirme, tablo, veri doğrulama.
- Resim, grafik, köprü ekleme.
- Sayfa düzeni, sayfa görünümü, belge özellikleri, formül hesaplama.

---

## Kurulum

```bash
npm install @halil07/excelize
# veya
yarn add @halil07/excelize
# veya
pnpm add @halil07/excelize
```

---

## Hızlı Başlangıç

### Node.js — ES Modülleri

```js
import { excelizeInit } from "@halil07/excelize";
import fs from "fs";

const Excelize = await excelizeInit(); // wasm çalışma zamanını otomatik yükler

const f = Excelize.newFile();
f.setCellValue("Sheet1", "A1", "Merhaba Node.js!");
f.setCellValue("Sheet1", "A2", 41);
f.setCellFormula("Sheet1", "A3", "=A2+1");
console.log(f.getRows("Sheet1"));

const bytes = f.save(); // Uint8Array
f.close();
fs.writeFileSync("cikti.xlsx", bytes);
```

### Node.js — CommonJS

```js
const fs = require("fs");
const { excelizeInit } = require("@halil07/excelize");

(async () => {
  const Excelize = await excelizeInit(); // wasm çalışma zamanını otomatik yükler

  const f = Excelize.newFile();
  f.setCellValue("Sheet1", "A1", "Merhaba Node.js!");
  f.setCellValue("Sheet1", "A2", 41);
  f.setCellFormula("Sheet1", "A3", "=A2+1");
  console.log(f.getRows("Sheet1"));

  const bytes = f.save(); // Uint8Array
  f.close();
  fs.writeFileSync("cikti.xlsx", bytes);
})();
```

### TypeScript

Paket içinde TypeScript tanımları (`dist/excelize.d.ts`) gelir. `excelizeInit`’i `await` ettikten sonra tam olarak tiplenmiş `Excelize` API’sine erişirsiniz:

```ts
import { excelizeInit, File } from "@halil07/excelize";

const Excelize = await excelizeInit();
const f: File = Excelize.newFile();
f.setCellValue("Sheet1", "A1", "Merhaba TypeScript!");
const bytes: Uint8Array = f.save();
f.close();
```

### Tarayıcı

`dist/wasm_exec.js` ve yükleyiciyi yükleyip `.wasm` modülünü `new Go()` ile başlatın. `Excelize.ready()` çağrısından sonra API kullanılabilir. Çalışan bir örnek için [`wasm/example.html`](wasm/example.html) dosyasına bakabilirsiniz.

```html
<script src="node_modules/@halil07/excelize/dist/wasm_exec.js"></script>
<script src="node_modules/@halil07/excelize/dist/excelize.js"></script>
<script>
  const go = new Go();
  WebAssembly.instantiateStreaming(
    fetch("node_modules/@halil07/excelize/dist/excelize.wasm"),
    go.importObject
  ).then((res) => {
    go.run(res.instance); // globalThis.ExcelizeWasm’i kaydeder
  });
</script>
```

Tarayıcıda oluşturulan dosyayı indirmek için:

```js
await Excelize.ready();
const f = Excelize.newFile();
f.setCellValue("Sheet1", "A1", "Merhaba WebAssembly!");
const bytes = f.save();
const blob = new Blob([bytes], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
});
const url = URL.createObjectURL(blob);
const a = document.createElement("a");
a.href = url;
a.download = "cikti.xlsx";
a.click();
URL.revokeObjectURL(url);
f.close();
```

---

## WebAssembly ile Derleme

Projeyi kaynaktan derlemek için Go ≥ 1.25 ve Node.js gerekir.

```bash
npm install
npm run build      # scripts/build.js çapraz platform destekler
```

Bu komut şu çıktıları üretir:

- `dist/excelize.wasm`
- `dist/wasm_exec.js`
- `dist/excelize.js`
- `dist/excelize.cjs`
- `dist/excelize.mjs`
- `dist/excelize.d.ts`, `dist/excelize.d.mts`, `dist/excelize.d.cts`

Windows / PowerShell alternatifi:

```powershell
./wasm/build.ps1
```

Manuel Go derlemesi (ileri düzey):

```bash
GOOS=js GOARCH=wasm go build -trimpath -ldflags="-s -w" -o dist/excelize.wasm ./wasm
```

---

## API Yüzeyi

`excelizeInit`’i içe aktarıp `await` ettikten sonra başlatılmış `Excelize` API’sine ulaşırsınız. Dönen nesne `newFile()` / `openFile()` ve `File` sınıfını içerir. Her yöntem başarılı olursa hesaplanan değeri (`Uint8Array`, `string`, `number`, `Array`, …), başarısız olursa bir JS `Error` fırlatır. Alt düzey `{ ok, data }` / `{ ok:false, err }` bağlamaları başlatma sonrası `Excelize.raw` üzerinden erişilebilir.

| Alan | Yöntemler |
|------|-----------|
| Çalışma Kitabı | `newFile`, `openFile`, `save`, `close`, `getSheetList`, `getSheetCount`, `defaultSheetName` |
| Sayfalar | `newSheet`, `getSheetIndex`, `getSheetName`, `setSheetName`, `deleteSheet`, `setSheetVisible`, `getSheetVisible`, `getSheetMap`, `setActiveSheet`, `getActiveSheetIndex`, `setSheetBackgroundFromBytes`, `setSheetDimension`, `getSheetDimension` |
| Hücreler | `setCellValue`, `getCellValue`, `setCellFormula`, `getCellFormula`, `setSheetRow`, `setSheetCol`, `getRows`, `getCols` |
| Satırlar / Sütunlar | `setColWidth`, `getColWidth`, `setColVisible`, `getColVisible`, `setRowHeight`, `getRowHeight`, `setRowVisible`, `getRowVisible` |
| Anahat | `setColOutlineLevel`, `getColOutlineLevel`, `getRowOutlineLevel` |
| Stiller | `newStyle`, `getCellStyle`, `setCellStyle`, `setConditionalFormat`, `getConditionalFormats` |
| Birleştirmeler | `mergeCell`, `unmergeCell`, `getMergeCells` |
| Tablolar | `addTable`, `getTables`, `deleteTable` |
| Doğrulama | `addDataValidation` |
| Resimler | `addPictureFromBytes`, `getPictures` |
| Grafikler | `addChart`, `addChartSheet` |
| Köprüler | `setCellHyperLink`, `getCellHyperLink` |
| Düzen / Görünüm / Özellikler | `setPageLayout`, `getPageLayout`, `setSheetView`, `getSheetView`, `setSheetProps`, `getSheetProps` |
| Belge Özellikleri | `setAppProps`, `getAppProps`, `setDocProps`, `getDocProps` |
| Hesaplama | `calcCellValue`, `setCalcProps`, `getCalcProps` |

Karmaşık seçenek yapıları (`Style`, `Table`, `DataValidation`, `Chart`, …) düz JS nesneleri olarak iletilir ve ilgili Go struct’ına JSON ile çözümlenir.

---

## Sınırlamalar

- Tarayıcı / Node.js WASM ortamında kullanılabilir bir yazılabilir dosya sistemi olmadığından, disk tabanlı `OpenFile` / `SaveAs` API’leri bilerek bağlanmamıştır. Bunun yerine bellek içi `openFile(Uint8Array)` / `save()` kullanın.
- Akış (`StreamWriter`) API’si ve çok büyük çalışma sayfaları için geçici dosya yedek yolu devre dışı bırakılmıştır (tarayıcı wasm geçici dosya oluşturamaz). Çok büyük çalışma kitapları tamamen bellekte tutulur.

---

## Yayınlama

Yayınlanan npm paketi yalnızca önceden derlenmiş `dist/` yapıtlarını, README ve LICENSE dosyalarını içerir (`package.json` içindeki `files` alanına bakın). `pnpm publish`, `prepublishOnly` betiği aracılığıyla derlemeyi otomatik olarak çalıştırır.

### GitHub Actions ile Otomatik Yayınlama

Repo’ya `v*.*.*` formatında bir git etiketi push edildiğinde, GitHub Actions otomatik olarak derler, test eder ve npm’e yayınlar. Workflow’lar bağımlılık yönetimi için **pnpm** kullanır.

Gerekli kurulum:

1. npm’de oturum açın ve bir **Access Token** (Automation veya Publish token) oluşturun.
2. GitHub repo ayarlarından **Settings → Secrets and variables → Actions → New repository secret** yolunu izleyin.
3. Adı `NPM_TOKEN`, değeri npm token’ınız olacak şekilde bir secret ekleyin.
4. Yeni sürüm yayınlamak için:

   ```bash
   # package.json içindeki version alanını güncelleyin
   pnpm version patch   # veya minor / major
   git push origin main --tags
   ```

   veya manuel olarak:

   ```bash
   git tag v0.2.0
   git push origin v0.2.0
   ```

Etiket push edildiğinde `.github/workflows/release.yml` çalışır; başarılı testlerin ardından `pnpm publish --access public --no-git-checks` ile paket yayınlanır.

---

## Lisans

BSD-3-Clause — aynı lisans upstream excelize kütüphanesi tarafından da kullanılır.
