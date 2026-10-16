# Stage 1 — Discovery Report

Date: 2026-10-16
Status: **Discovery only. No application code yet.** Critical inputs are still missing (see §11).

> Privacy note: the screenshots contain real employee data. This report doesn't reproduce any
> names, IDs, dates, phone numbers or other personal values. All examples are synthetic.

---

## 0. Artifacts found

| Artifact | Location | Status |
|---|---|---|
| Project folder `~/Desktop/Personnel` | — | Empty |
| Source workbook (`.xlsx`) | — | **Not provided.** Only photos of the screen |
| `InputExcel_1.png`, `InputExcel_2.png` | `~/Downloads` | Photos of the source workbook, columns A–L and L–AA |
| `FinalExcel_1/2/3.jpeg` | `~/Downloads` | Photos of the report workbook «کارنامه پرسنلی2 (9)» |
| Colleague's formula screenshot | — | **Not provided.** None of the photos show a formula in the formula bar |
| Report template (`.xlsx`) | — | **Not provided** |

Because I only have photos, I can't verify any of these: cell data types, number formats, formulas,
merged ranges, hidden sheets or columns, the print setup, or the exact cell addresses.

---

## 1. Source workbook schema (as observed in the photos)

- Worksheet tab visible: `Sheet1`. At least one other tab is present. **Sheet selection must not be assumed.**
- Row 1 is empty or a title row. **The header is row 2. Data starts at row 3.** The header row is styled in blue with wrapped text.
- There are 27 columns, A–AA. They are in exactly the order given in the brief (A=ردیف … AA=تلفن همراه).
  The header text in the photo appears to match the brief, including «محل خدمت ـ اداره کل».
  The exact dash character (ـ U+0640 vs – / -) must be confirmed against the real file.
- The formula bar for one date cell shows `1386/03/16`. That means **dates are stored as Jalali text, `yyyy/mm/dd`**, not as Excel date serials.
  The cell displays Persian digits, which is a font or locale effect.
- National ID and mobile values display leading zeros, so they're probably stored as text.
- Service duration is already split into days, months and years: سال/ماه/روز خدمت در دیوان and سابقه خدمت دولتی.
  **These are authoritative and must not be recalculated.**
- Values observed in some columns:
  - وضعیت اشتغال: e.g. شاغل, and a leave-type value
  - نوع استخدام: رسمی / رسمی آزمایشی / پیمانی
  - وضعیت تاهل: متاهل / مجرد / **معیل**
- **Data-quality issues observed:**
  - `-` is used in the child-count columns, apparently to mean "none".
  - At least one row's تاریخ تولد cell holds a value shaped like an identifier rather than a date.
  - Some organizational titles are long and wrap across lines.
  - Some personnel numbers have very different lengths.

## 2. Worksheet and header detection strategy

1. Scan every worksheet, hidden ones included, but flag hidden sheets. In each sheet, scan the first N rows (for example 10) for a row that contains **all 27 exact headers**. Before comparing, apply only whitespace trimming and Unicode NFC normalization. There is no fuzzy matching.
2. If exactly one sheet/row qualifies, use it. If none qualifies, report the missing headers in Persian. If more than one qualifies, ask the user to choose.
3. Map columns **by header text**, never by letter. Ignore extra columns.
4. Read with `openpyxl` using `data_only=True`, which gives the cached values of formula cells. Warn if a formula cell has no cached value.

## 3. Employee lookup

- Normalize the input: Persian ۰–۹ and Arabic ٠–٩ become ASCII; strip spaces, ZWNJ and RTL marks. Keep the result as a string and preserve leading zeros.
- Source national IDs are normalized the same way. If a cell is numeric (an int or float), restore the zeros by left-padding to 10 digits **and flag a warning**, because Excel will already have dropped the zeros.
  (Open question: should padding be allowed at all?)
- Exact match only. With 0 hits, show «کارمندی با این شماره ملی یافت نشد». With 2 or more hits, block and list the row numbers. Duplicates are detected at load time and summarized.

## 4. Report template structure (from the photos)

Workbook «کارنامه پرسنلی2 (9)». Its current practice is **one worksheet per employee, with the employee's name as the tab name**.
Font: **B Nazanin 12**. Columns roughly B–X, sheet RTL.

| # | Section (exact label) | Approx. rows | Content |
|---|---|---|---|
| 1 | کارنامه پرسنلی | 1–2 | Title |
| 2 | اطلاعات فردی | 4–19, right side (B–J) | نام ونام خانوادگی، کدملی، کدپرسنلی، تاریخ تولد، مدرک تحصیلی، رشته تحصیلی، شماره تماس، وضعیت تاهل (checkboxes متأهل / مجرد)، تعدادفرزندان split into **دختر / پسر** |
| 3 | اطلاعات پرسنلی | 4–19, left side (K–S) | نوع استخدام، پست سازمانی، محل خدمت، تاریخ شروع به کار، سوابق خدمتی/تجربی، سوابق آموزشی، سوابق ارزیابی (تنبیهات / تشویق ها) |
| 3b | Yellow box (no header) | 6–19 | «سمت های قبلی:» with list «۱-»، «توقف پست:» |
| 4 | عملکردوتوانمندی کارکنان | 20–21 | Section banner |
| 5 | ارزیابی توانمندی وشایستگی | ~23–42 | ردیف / شاخص / نمره / شرح. 7 criteria: تخصص وتسلط بروظایف، دقت وکیفیت انجام کار، مسئولیت پذیری، توانائی حل مسئله ومدیریت امور، همکاری وکارتیمی، تعامل وارتباط مؤثر، ابتکاروارائه پیشنهادموثر. Then **جمع امتیاز** |
| 6 | کارنامه نظم وانضباط اداری | ~23–44 | ردیف / شاخص / نمره / توضیحات/مستندات. 7 rows: ورود وخروج به موقع، میزان تأخیر/غیبت غیرموجه، رعایت شئونات اداری، پایبندی به اصول اخلاق حرفه ای، آراستگی وپوشش، شرکت درمراسم مذهبی_فرهنگی، تعهدات وتذکرات (with two sub-cells **تذکر** / **تعهد**) |
| 7 | جمع بندی وپیشنهادات مدیریتی | ~23–40 | 6 options with icons: شایسته تقدیروتشویق، مناسب استمراردرمسئولیت فعلی، شایسته ارتقای شغلی، مستعدپذیرش مسئولیت بالاتر، نیازمندبرنامه بهبودعملکرد، نیازمندبررسی مجددمسئولیت. Then «توضیحات تکمیلی:» |
| 8 | سوابق ارزیابی ۵ سال گذشته | ~46–61 | سال / نمره کل / رتبه / روندعملکرد / نقاط قوت شاخص / مواردنیازمند بهبود. Years ۱۴۰۰…۱۴۰۴. Then **میانگین ۵ ساله** |
| 9 | رضایت سنجی | ~49–58 | ردیف / شاخص / امتیاز / شرح: ۱ رضایت مدیرمستقیم، ۲ رضایت ارباب رجوع. Static note «منبع ارزیابی: سامانه سامع/ دفترارتباطات مردمی/مراجعات حضوری/تقدیرهاوشکایات» |
| 10 | جمع بندی نهائی ارزیابی عملکرد | ~48–62+ | امتیازعملکردواحد/شعبه، امتیازعملکردفردی، امتیازتوانمندی حرفه ای، امتیازنظم وانضباط اداری، امتیازرضایت مدیر، امتیازتکریم ارباب رجوع, … **cut off. More rows exist below.** |

**Not visible in any photo:**
- rows below about 62 (rest of §10, any total or grade, and **signature/approval areas**)
- the right-hand part of rows 38–40
- the print area and page setup

Labels above are transcribed as written in the template, including its spacing.

## 5. Source → report field mapping (draft — **needs confirmation**)

| Report field | Source header | Confidence |
|---|---|---|
| نام ونام خانوادگی | نام + « » + نام خانوادگی | High |
| کدملی | شماره ملی | High |
| کدپرسنلی | شماره پرسنلی | High |
| تاریخ تولد | تاریخ تولد | High |
| مدرک تحصیلی | مدرک تحصیلی | High |
| رشته تحصیلی | رشته تحصیلی | High |
| شماره تماس | تلفن همراه | High |
| وضعیت تاهل (متأهل/مجرد) | وضعیت تاهل | **Unclear.** How is «معیل» shown? |
| تعدادفرزندان دختر / پسر | تعداد اولاد واقعی? در ابلاغ ریالی? | **Blocked.** The source has no split by gender |
| نوع استخدام | نوع استخدام | High |
| پست سازمانی | عنوان پست سازمانی در ابلاغ | Medium |
| محل خدمت | محل خدمت ـ اداره کل and/or ـ اداره/بخش, or واحد سازمانی در ابلاغ? | **Unclear** |
| تاریخ شروع به کار | تاریخ ورود به دیوان or تاریخ استخدام در خدمات دولتی? | **Unclear** |
| سوابق خدمتی/تجربی | the سال/ماه/روز service columns? Manual? | **Unclear** |
| سوابق آموزشی، سوابق ارزیابی (تنبیهات/تشویق‌ها)، سمت‌های قبلی، توقف پست | — | Presumed manual |
| Sections 5–10 | — | Presumed manual, except totals and the average |

Unused source columns: ردیف، وضعیت اشتغال، تاریخ آخرین انتصاب، and the service-duration columns
(unless they map to سوابق خدمتی).

## 6. Field classification

- **Automatic:** §2 fields and §3 نوع استخدام/پست/محل خدمت/تاریخ شروع (pending confirmation)
- **Manual:** all scores (نمره/امتیاز), شرح/توضیحات, تذکر/تعهد, the management recommendation choice, توضیحات تکمیلی, the 5-year history table, سوابق آموزشی/ارزیابی, سمت‌های قبلی, توقف پست
- **Static:** every label, the criterion names, the years ۱۴۰۰–۱۴۰۴ (to be confirmed), and the رضایت‌سنجی source note
- **Calculated (rule needed):** جمع امتیاز (×2 sections), میانگین ۵ ساله, and possibly the final summary scores. **No formula has been seen, so none will be invented.**

## 7. Validation rules

Known so far: exact-header presence, a unique national ID, and required auto fields flagged when empty.
**Unknown:** score ranges per criterion, whether scores are integers, which manual fields are required,
and whether the recommendation is single-choice or multiple-choice.

## 8. Excel output

- Preferred approach: populate a **blank copy of the real template** with openpyxl, keeping merges, borders, fills and print setup. Output one employee per file.
- **"Embedded fonts" isn't possible in `.xlsx`.** Excel (unlike Word/PowerPoint) has no font embedding. The cell font is set by name, and the font must be installed on the viewing PC. Mitigation: ship the font files with the installer and show a font-check message in the app.
- Icons and shapes in §7 and the checkboxes: openpyxl **drops shapes, images and form controls** when it re-saves a workbook. I'll need to see the real template to choose an approach, for example ☑/☐ characters, or keeping the images.

## 9. Markdown and HTML

Both are built from one report model: the model is rendered once to Markdown and once to HTML through Jinja2.
- HTML uses `dir="rtl"` and `lang="fa"`, with `unicode-bidi: isolate` / `<bdi>` around IDs, dates and phone numbers, and `@media print` set to A4.
- Markdown is UTF-8, with GFM tables per section. Merged cells and icons can't be represented and will be documented.

## 10. Fonts (this development Mac)

| Font | Installed here? |
|---|---|
| B Nazanin | Yes (Regular, Bold) |
| B Mitra | Yes (Regular, Bold) |
| Vazirmatn FD | **No.** Only `Vazirmatn` (variable) is installed. The FD (Farsi-digit) build is OFL-licensed and can be bundled |
| Times New Roman | Yes (system) |

The B-series fonts are freeware, but their redistribution terms should be confirmed before bundling them.
The browser view will load bundled `@font-face` files locally, so it doesn't depend on fonts installed in Windows.

## 11. Unresolved questions (blocking first)

1. **Real files.** Please provide the following:
   - (a) the **blank report template** `.xlsx`, or one sheet with the personal data cleared
   - (b) a **synthetic or anonymized copy** of the source workbook with the same structure, sheets and formats. Even 5 fake rows is enough.
   - (c) the **colleague's formula screenshot**, which wasn't in the files
2. The mapping rows marked Unclear/Blocked in §5: محل خدمت, تاریخ شروع به کار, سوابق خدمتی/تجربی, دختر/پسر, معیل.
3. Score scales (min/max) for each table. How جمع امتیاز, میانگین ۵ ساله and the final summary (§10) are computed, or whether they're entered manually.
4. The hidden rest of the report: rows below 62, and any signature/approval block.
5. Is the year list ۱۴۰۰–۱۴۰۴ fixed, or should it roll with the current year?
6. Recommendation (§7): single choice or multiple choice?
7. Output: one employee per downloaded file (my assumption), or append sheets to a shared workbook as the colleague does today?
8. **Windows build and testing.** I'm on macOS, so I can't build or test a Windows package here. Is there a Windows PC (or a Windows VM or CI runner) available for testing the package?
9. Font list: the brief mentions B Mitra/B Nazanin in some places and adds Vazirmatn FD + Times New Roman elsewhere. Should all four be offered?

## 12. Proposed stack

**Python 3.12 + Flask + openpyxl + Jinja2**, packaged as a **PyInstaller one-folder app** with an Inno Setup installer. It installs per user, needs no admin rights, and puts a desktop shortcut that opens the browser.

| Option | Verdict |
|---|---|
| Flask | Small; full control over RTL HTML and print CSS; openpyxl is the best library for filling templates; packages well with PyInstaller. **Chosen.** |
| FastAPI | Async/OpenAPI features aren't needed; it's slightly heavier |
| Streamlit | Poor control of RTL, print layout and form layout; packaging is awkward |
| ASP.NET Core | Single-file publish is excellent, but Excel template filling (ClosedXML/EPPlus) is fine, not better, and the build is heavier |
| Node/Express | Excel libraries are weaker at preserving templates; packaging is less mature |

The server binds to `127.0.0.1` on a free port. Data is held only in process memory. Downloads are streamed from memory, with no temp files.

---

# Update — 2026-10-16 (round 2)

## A. Decisions confirmed by the user

| Topic | Decision |
|---|---|
| محل خدمت | ← source column «محل خدمت ـ اداره/بخش» |
| تاریخ شروع به کار | ← source column «تاریخ ورود به دیوان» |
| دختر / پسر | Entered manually by the user |
| وضعیت تاهل | Two boxes side by side, متأهل and مجرد, with a tick (☑) under whichever applies |
| Scoring | **The user enters every score.** The app calculates nothing: جمع امتیاز, میانگین ۵ ساله and all final scores are manual |
| Signature / final score | A block is added: «نمره نهایی ارزیابی عملکرد», «نام، سمت و امضای ارزیاب», «تاریخ» |
| 5-year years | Fixed at ۱۴۰۰–۱۴۰۴ for now |
| Output | **Added to a shared workbook**, one sheet per employee, as the colleague does today |
| Fonts | Bundle all four with the app |

Still open:
- How should «معیل» (seen in the real data) be marked, given the form only has متأهل / مجرد?
- Should the recommendation be a single choice or multiple choice? The template allows ticking any number.
- What are the remaining rows of «جمع بندی نهائی ارزیابی عملکرد»? Only 6 were visible in the photo.

## B. Colleague's formula (Formula.HEIC)

The formula is in the personnel-code cell of a sheet titled «فرم مشخصات همکار»:

```
=IF($D$6<>"اطلاعات نمایش داده شد","",
    IF(VLOOKUP($D$5,'اطلاعات همکاران'!$A$5:$N$504,2,FALSE)="","",
       VLOOKUP($D$5,'اطلاعات همکاران'!$A$5:$N$504,2,FALSE)))
```

- **Lookup key:** `$D$5`, the national ID the user types into the yellow box. It's absolute.
- **Source:** sheet «اطلاعات همکاران», range `A5:N504` (absolute). The **national ID must be in column A**, data starts at row 5, and there are at most 500 rows.
- **Column returned:** the 2nd column, by position. In that workbook it's کد پرسنلی. Other fields (نام, نام خانوادگی, نام پدر …) presumably use the indexes 3, 4, …
- **Gate:** `$D$6` is a status cell. Nothing is shown unless it reads «اطلاعات نمایش داده شد».
- **Error handling:** an empty source cell returns "" instead of 0. A missing ID still raises **#N/A**; there's no IFERROR.
- `FALSE` means exact match, which is good.

**What the app keeps and what it fixes:**
- It keeps the exact-match lookup on the national ID.
- It does not keep the fragile parts, which are:
  - the position-based column indexes
  - the key having to be the first column (in the real file the national ID is column G)
  - the fixed 500-row range
  - the silent #N/A
  - no duplicate detection
  - text vs. number ID mismatch, which drops leading zeros

The app looks columns up **by header name**, scans all rows, reports duplicates, and normalizes digits.
This workbook (with «نام پدر» and a sheet named «اطلاعات همکاران») is different from the 27-column source; the app follows the 27-column source.

## C. Fake personnel workbook (`Template/fake personnel workbook.xlsx`)

- There is one sheet, `Sheet1`, set to RTL. There are no hidden sheets or columns, no merges and no formulas. **The header is in row 1** (in the real file it's row 2), and there are 5 synthetic rows.
- Every value, including the numbers, is **text written with Persian digits** (۰–۹), for example IDs `۰۱۵…`, dates `۱۳۶۱/۰۱/۱۹`, and counts `۰`. Normalizing digits is therefore essential.
- Day/month/year service columns are empty.
- **Header text differs from the brief in 4 places.** The app must not silently "fix" these:

| Brief | Fake workbook |
|---|---|
| محل خدمت ـ اداره کل | محل خدمت - اداره کل (hyphen-minus) |
| محل خدمت ـ اداره/بخش | محل خدمت - اداره/بخش (hyphen-minus) |
| روزسابقه خدمت دولتی | روز سابقه خدمت دولتی (extra space) |
| ماه/سال سابقه خدمت دولتی | ماه/سال سابقه **خدمتی** دولتی |

  Proposal: the mapping config lists an **explicit set of accepted header spellings** for each field. These are exact strings, not fuzzy matching. It will include both spellings above, and the dash variants ـ – -.
  Header detection scans the first rows, so it handles both row 1 and row 2.

## D. Blank report template — built

- `Template/report_template.xlsx` has one blank sheet, «قالب». It's RTL, uses B Nazanin, prints on A4 landscape at 1 page wide, and is split into 2 pages:
  - page 1: personal info
  - page 2: evaluation tables, the final score and the signature block
- `Template/report_template_map.json` maps each field id to its cell. It has 96 fillable cells, and the app will use only this map.
- Generator: `tools/build_template.py`. Previews: `Template/preview/page1.png` and `page2.png`, rendered by Microsoft Excel.
- Deviations from the photos:
  - Drawn shapes (the star, cloud and flag icons) became coloured Unicode symbols, because openpyxl can't keep shapes.
  - A ☐ column was added to «جمع بندی وپیشنهادات مدیریتی» so the chosen option can be ticked.
  - Cell positions are approximations; they are not copied from the original.
