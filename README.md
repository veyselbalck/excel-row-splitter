# Excel Row Splitter

A VBA macro that splits the active Excel worksheet into new worksheets containing up to 5,000 rows each.

## Requirements

- Microsoft Excel with macros enabled
- A workbook containing the worksheet to split

## Usage

1. Open the workbook and select the worksheet you want to split.
2. Press `Alt+F11` to open the Visual Basic Editor.
3. Select **File > Import File...** and import `SplitExcelRows.bas`.
4. Return to Excel, press `Alt+F8`, select `SplitExcelRows`, and run it.

The macro determines the last used row by checking column A. It copies rows in groups of up to 5,000, starting at row 1, to newly created worksheets named `Sheet N`. Existing worksheet names that conflict with the generated names may prevent the macro from completing. Save a backup before running it on important workbooks.

## E-Commerce Exports

You can run the macro on a compatible Excel report exported from Ticimax or another e-commerce platform. This is a file-based utility; it does not connect to or exchange data directly with any platform.

## File

- `SplitExcelRows.bas` — VBA macro source file.

---

# Excel Satır Bölücü

Aktif Excel çalışma sayfasını, her biri en fazla 5.000 satır içeren yeni sayfalara bölen bir VBA makrosu.

## Gereksinimler

- Makroların etkin olduğu Microsoft Excel
- Bölünecek sayfayı içeren bir çalışma kitabı

## Kullanım

1. Çalışma kitabını açın ve bölmek istediğiniz sayfayı seçin.
2. Visual Basic Düzenleyicisi'ni açmak için `Alt+F11` tuşlarına basın.
3. **File > Import File...** seçeneğini kullanarak `SplitExcelRows.bas` dosyasını içe aktarın.
4. Excel'e dönün, `Alt+F8` tuşlarına basın, `SplitExcelRows` makrosunu seçip çalıştırın.

Makro, son kullanılan satırı A sütununa göre belirler. 1. satırdan başlayarak satırları en fazla 5.000 satırlık gruplara ayırır ve `Sheet N` adlı yeni çalışma sayfalarına kopyalar. Oluşturulacak sayfa adları mevcut sayfa adlarıyla çakışırsa makro tamamlanamayabilir. Önemli çalışma kitaplarında kullanmadan önce yedek alın.

## E-Ticaret Çıktıları

Makroyu Ticimax veya başka bir e-ticaret platformundan dışa aktarılan uyumlu bir Excel raporunda kullanabilirsiniz. Bu araç dosya üzerinden çalışır; herhangi bir platforma doğrudan bağlanmaz veya veri aktarımı yapmaz.

## Dosya

- `SplitExcelRows.bas` — VBA makro kaynak dosyası.
