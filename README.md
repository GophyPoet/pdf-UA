# pdf/UA — локальная подготовка документов

Приложение, которое принимает `.doc / .docx / .xls / .xlsx / .odt / .ods`
и выдаёт **редактируемый `.odt`**, **`PDF/UA` (PDF 1.7, tagged)** и
человекочитаемый отчёт о выполненных исправлениях и остаточных рисках.

Обработка автоматическая: помимо конвертации, применяются
нормализация текста, восстановление семантики (Title, заголовки,
списки), починка таблиц, локальный OCR изображений с генерацией
alt text и внедрением текстового эквивалента, проверка правил
доступности и контролируемый экспорт PDF/UA с фиксированными
параметрами.

Работает **полностью локально**. Никаких внешних API.

---

## Возможности

| Этап | Что делает |
|---|---|
| intake | Определяет тип (`text` / `spreadsheet`), снимает stale `.~lock` |
| title-detection | Находит заглавие: существующий `Heading 1`/`Title` → первый содержательный абзац → очищенное имя файла. Пишет в `DocumentProperties.Title` |
| normalize-text | Свёртывает множественные пробелы, удаляет декоративные табуляции, заменяет цепочки `___` на `[поле для заполнения]` или удаляет их, удаляет управляющие символы и лишние пустые абзацы (не уничтожая абзацы с инлайн-картинками) |
| restore-headings | Эвристически распознаёт псевдозаголовки (жирный + короткий + крупный / нумерация `1.`/`1.1`) и назначает `Heading 1..3` |
| repair-tables (Writer) | Разбивает объединённые ячейки через `XTextTableCursor.splitRange`, заполняет пустые ячейки `—`, убирает пустые строки/столбцы, помечает первую строку как заголовочную; при остаточных объединениях DOCX аккуратно сдаётся и пишет риск |
| repair-spreadsheet (Calc) | Размерживает все merged regions, пропускает пустые строки/столбцы, заполняет пустые ячейки `—` |
| spreadsheet-to-ODT | Создаёт новый Writer-документ: заголовок документа, для каждого листа `Heading 2` + полноценная Writer-таблица с header row |
| process-images | Обходит `GraphicObjects` + shapes на `DrawPage`; экспортирует графику через `GraphicProvider`; локально OCR'ит Tesseract'ом (rus+eng) с возвратом уровня уверенности; присваивает `Description`/`Title`; при наличии значимого текста добавляет параграф «Текстовое содержимое изображения: …» сразу после картинки |
| alt-text | Эвристика: tiny/rule → decorative; печать/штамп; скриншот/интерфейс; обычный текст — «Изображение с текстом: …»; низкая confidence — alt помечается как предположительный, текст в тело не вставляется |
| accessibility-rule-engine | 9 собственных правил (Title, Heading, multi-spaces, tabs, underscores, merged cells, empty cells/rows, images without alt), ранжируются по severity |
| pdf-export | `writer_pdf_Export` с `FilterData`: PDF 1.7, `PDFUACompliance=True`, `UseTaggedPDF=True`, JPEG 90%, ReduceImageResolution=150 DPI |
| reporting | `*.report.json` и `*.report.txt` (на русском) со всеми числами и остаточными рисками |

---

## Зависимости

Системные:
* LibreOffice 7+ (`soffice` в `$PATH`, поддержка `python3-uno`)
* Tesseract 5 с языковыми пакетами `rus` и `eng`
* Python 3.10+

Установка на Ubuntu 24.04 / Debian:

```bash
sudo apt install -y libreoffice libreoffice-script-provider-python python3-uno \
                    tesseract-ocr tesseract-ocr-rus tesseract-ocr-eng
pip install -r requirements.txt
```

---

## Запуск

### Веб-интерфейс (основной способ)

```bash
python -m pdfua.cli serve --host 127.0.0.1 --port 8000
```

Открыть в браузере `http://127.0.0.1:8000/`. Страница позволяет
перетащить файл, увидеть текущий этап, логи, финальный отчёт по
правилам и скачать ODT / PDF / JSON-отчёт / TXT-отчёт.

### CLI

```bash
python -m pdfua.cli convert input.docx output_dir/
python -m pdfua.cli convert tests/fixtures/sample_sheet.xlsx out/
python -m pdfua.cli convert input.docx out -v   # подробный лог
```

---

## Структура проекта

```
pdfua/
  __init__.py
  uno_bridge.py       # soffice subprocess + UNO-коннект по сокету
  pipeline.py         # оркестратор всех этапов
  intake              # (встроено в pipeline.detect_type)
  normalizer.py       # очистка/нормализация текста
  title_headings.py   # Title + promotion Heading 1..3
  images.py           # картинки, OCR, alt, текстовый эквивалент
  ocr.py              # Tesseract с confidence
  alt_text.py         # правила для alt text
  tables.py           # починка Writer-таблиц
  spreadsheet.py      # Calc → чистый ODT с Writer-таблицами
  rules.py            # собственный rule engine
  pdf_export.py       # PDF/UA с FilterData
  report.py           # JSON + TXT отчёт
  cli.py              # CLI: convert / serve
  server.py           # Flask web-UI
  templates/index.html
tests/
  make_fixtures.py    # генерация тестовых DOCX/XLSX через UNO
  test_unit.py        # юнит-тесты (не требуют soffice)
  test_e2e.py         # полный прогон пайплайна
```

---

## Что тестируется

* **`tests/test_unit.py`** — OCR на синтетическом изображении,
  решатель alt text на 4 классах (tiny/rule/readable/low-confidence),
  regex нормализации.
* **`tests/test_e2e.py`** — полный прогон на:
  * DOCX с множественными пробелами, табуляциями, псевдозаголовком,
    цепочкой подчёркиваний, таблицей с объединёнными и пустыми
    ячейками, изображением со встроенным текстом
  * XLSX с двумя листами, объединённой шапкой и пустым столбцом
  
  Проверяется: `PDF/UA` marker в выходе, `%PDF-1.7`, что OCR нашёл
  текст, что таблицы получили header row, что нет error-level правил.

Запуск:

```bash
python tests/test_unit.py
python tests/test_e2e.py
```

---

## Как именно получается PDF/UA

`pdfua.pdf_export.export_pdfua` строит вектор
`com.sun.star.beans.PropertyValue` и кладёт его в `FilterData`:

```
SelectPdfVersion      = 17       # PDF 1.7
PDFUACompliance       = True     # идентификатор PDF/UA
UseTaggedPDF          = True     # структурированный (tagged) PDF
ExportBookmarks       = True
UseLosslessCompression= False
Quality               = 90       # JPEG 90%
ReduceImageResolution = True
MaxImageResolution    = 150      # 150 DPI
```

«PageRange» / «Selection» не задаются → экспортируются все страницы.
`Title` в `DocumentProperties` ставится ДО экспорта, чтобы PDF унёс
его в метаданные.

Проверка выходного PDF:

```
$ pdfinfo "out.pdf" | grep -E 'PDF version|Tagged|Title'
Title:        ИНСТРУКЦИЯ ПО ОФОРМЛЕНИЮ ДОКУМЕНТА
Tagged:       yes
PDF version:  1.7
$ grep -c pdfuaid:part out.pdf
1
```

---

## OCR и alt text

**OCR.** `pytesseract.image_to_data` даёт per-word confidence.
Изображение предварительно увеличивается, если короткая сторона < 1000px.
Если языковые данные `rus` отсутствуют — фолбэк на `eng`. Возвращается
`OcrResult(text, confidence, lang, word_count, usable)` где
`usable = word_count > 0 and confidence >= 0.55`.

**alt text.** Логика в `pdfua/alt_text.py` делает честный выбор:

1. изображение < 80px или явная декоративная линия → `decorative=True`,
   alt пустой, текст в документ не вставляется;
2. уверенный OCR + ключевые слова (`печать/stamp/подпись`) →
   «Печать документа. Текст: …»;
3. уверенный OCR + длинный текст / UI-лексика → «Скриншот документа
   или интерфейса. Первая строка: …» + текст вставляется в документ
   отдельным абзацем;
4. уверенный OCR иначе → «Изображение с текстом: …»;
5. OCR нашёл что-то, но confidence низкая → «Изображение,
   предположительно содержит текст: …»; текст НЕ вставляется (честно
   помечаем недостоверное);
6. OCR пустой → размер-ориентированное нейтральное описание.

Для каждого изображения в отчёте фиксируется `ocr_confidence`,
`ocr_usable`, `alt_confidence`, `reasoning`.

---

## Остаточные ограничения

1. **DOCX с нестандартными таблицами (irregular rowspan/colspan).**
   `XTextTableCursor.splitRange` не всегда способен аккуратно
   размержить ячейки после DOCX-импорта. При опасности ухудшения мы
   аккуратно сдаёмся, фиксируем `residual_merge_tables` и помечаем
   пайплайн-статус как `fixed-with-warnings`. PDF/UA-экспорт при этом
   всё равно проходит; ручная доработка может потребоваться для
   идеального tagging'а таблицы.

2. **OCR сложных сканов / рукописного текста.** Tesseract даёт низкую
   confidence — это не скрывается, в отчёте записывается
   `ocr_usable=False` и alt text помечается как предположительный.

3. **Заглавие документа.** Если в документе ни `Heading 1`, ни
   содержательный первый абзац — используется очищенное имя файла.
   Это задокументировано в отчёте в поле `title_source`.

4. **Spreadsheet → ODT.** Табличные файлы линеаризуются: каждый лист
   становится Heading 2 + отдельной Writer-таблицей. Формулы,
   формирование цветов и условное форматирование теряются намеренно —
   они не имеют смысла в tagged PDF.

5. **Старые `.doc` / `.xls`.** Открываются через тот же UNO-конвейер,
   но импорт может потерять нетривиальные свойства документа
   (вложенные объекты OLE, макросы) — это поведение LibreOffice.

---

## Пример вход → выход

Вход: `tests/fixtures/sample_doc.docx` (генерируется
`python tests/make_fixtures.py`) — псевдозаголовок без стиля,
двойные пробелы, табуляции между словами, цепочка `______`, таблица
3×3 с объединённой ячейкой и пустой ячейкой, PNG-изображение со
словами «СКРИНШОТ ДОКУМЕНТА …».

Выход:
* `ИНСТРУКЦИЯ ПО ОФОРМЛЕНИЮ ДОКУМЕНТА.odt` — с `Heading 1`,
  нормализованным текстом, таблицей с header row и альт-текстом.
* `ИНСТРУКЦИЯ ПО ОФОРМЛЕНИЮ ДОКУМЕНТА.pdf` — PDF 1.7, tagged, с
  PDF/UA marker, `Title` в метаданных.
* `*.report.json` и `*.report.txt` — полный отчёт.

Фрагмент отчёта:

```
Заглавие: ИНСТРУКЦИЯ ПО ОФОРМЛЕНИЮ ДОКУМЕНТА  (источник: first-paragraph)
Статус: fixed-with-warnings

== Нормализация текста =====================
 • multi_spaces_collapsed: 6
 • tabs_removed: 2
 • underscore_lines_handled: 2

== Изображения / OCR / alt text ============
 • total: 1
 • with_usable_ocr: 1
   ○ #0 Image1 (600x200): alt='Изображение с текстом: СКРИНШОТ ДОКУМЕНТА';
     ocr_conf=0.96; ocr_usable=да; text_equivalent=да

== Проверка доступности ====================
PDF/UA ready: да
Всего правил: 9, прошло: 9, ошибки: 0
```

---

## Когда использовать только CLI, а когда веб-UI

* **Веб-UI** — основной рабочий путь, подходит для
  индивидуальной пакетной обработки с визуальным контролем отчёта
  и скачиванием артефактов.
* **CLI** — для скриптов, cron, CI-пайплайнов и быстрой проверки.
