---
name: persian-docx
description: >-
  Use ONLY when the user explicitly requests generating, creating, styling, or exporting
  a Microsoft Word (.docx) document file in Persian. Do NOT activate this skill for general Persian
  conversations, answering questions, knowledge summaries, debriefs, or standard chat responses.
---

# Persian DOCX Styler & Generator Skill

> [!IMPORTANT]
> این اسکیل **فقط و فقط** زمانی فعال و اجرا می‌شود که کاربر صراحتاً درخواست ساخت، ویرایش، قالب‌بندی یا خروجی گرفتن فایل **Microsoft Word (.docx)** داشته باشد. برای پاسخ‌های متنی عادی، پرسش و پاسخ، گزارش‌های چت یا خلاصه‌سازی‌ها نباید از این اسکیل یا تولید فایل استفاده شود.

این اسکیل راهنمای جامع و ابزار استاندارد برای ساخت اسناد رسمی، فنی، اداری و آموزشی **Microsoft Word (.docx)** با پشتیبانی کامل از تنظیمات پنجره **Paragraph (Indents and Spacing / Direction: Left-to-Right & Right-to-Left)**، فونت استاندارد **Vazirmatn** و پالت‌های رنگی حرفه‌ای است.

---

## ۱. تنظیمات استاندارد پاراگراف (Paragraph: Indents and Spacing & Direction)

تنظیمات پاراگراف در مایکروسافت ورد در بخش `Indents and Spacing` شامل پارامترهای زیر است که در این اسکیل به صورت کامل و اختصاصی در لایه OpenXML پشتیبانی می‌شوند:

### الف) جهت پاراگراف (Direction)
* **`direction: right-to-left` (RTL):** فعال‌سازی تگ `<w:bidi w:val="1"/>` در `<w:pPr>` و تگ `<w:rtl w:val="1"/>` در `<w:rPr>`. مناسب متون فارسی، تیترها و شماره‌گذاری‌های فارسی.
* **`direction: left-to-right` (LTR):** تنظیم `<w:bidi w:val="0"/>` در `<w:pPr>` و `<w:rtl w:val="0"/>` در `<w:rPr>`. مناسب بخش‌های انگلیسی، نقل‌قول‌های لاتین، کدهای فنی، فرمول‌ها و مراجع بین‌المللی.

### ب) فواصل و تورفتگی‌ها (Indents)
* **Left Indent (`w:ind w:left`):** تورفتگی از سمت چپ (مناسب ایجاد تو رفتگی در پاراگراف‌های انگلیسی یا زیرمجموعه‌ها).
* **Right Indent (`w:ind w:right`):** تورفتگی از سمت راست (مناسب برای پاراگراف‌های فارسی، لیست‌های گلوله‌ای و کادرهای نقل‌قول).
* **First Line Indent (`w:ind w:firstLine`):** تورفتگی خط اول پاراگراف (معمولاً ۰.۵ تا ۱ سانتی‌متر برای کتاب‌ها و مقالات).

### ج) فاصله‌گذاری خطوط و پاراگراف‌ها (Spacing)
* **Space Before (`w:spacing w:before`):** فاصله قبل از پاراگراف (برای تیتر ۱: ۱۸pt، تیتر ۲: ۱۴pt، تیتر ۳: ۱۰pt).
* **Space After (`w:spacing w:after`):** فاصله بعد از پاراگراف (برای متون بدنه: ۶pt).
* **Line Spacing (`w:spacing w:line`):**
  * برای متون فارسی (Vazirmatn): **۱.۳ تا ۱.۳۵ برابر** (`line_spacing=1.3`) جهت خوانایی ایده‌آل و عدم تداخل اعراب و حروف سرکش‌دار.
  * برای متون انگلیسی LTR: **۱.۱۵ برابر** (`line_spacing=1.15`).

---

## ۲. پالت‌های رنگی استاندارد (Color Palettes)

این اسکیل دارای **۶ پالت رنگی استاندارد و کاربردی** است:

| شناسه پالت | نام فارسی | رنگ اصلی (Primary) | رنگ ثانویه (Secondary) | رنگ تأکیدی (Accent) | رنگ متن تیره (Dark Text) | پس‌زمینه لایت (Light Tint) | کاربرد |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `navy` | **سرمه‌ای کلاسیک (سازمانی)** | `#1E3A8A` | `#3B82F6` | `#D97706` | `#1E293B` | `#F1F5F9` | گزارش‌های سازمانی، قراردادها، اداری |
| `emerald` | **سبز زمردی (مدرن / زیست‌محیطی)** | `#065F46` | `#10B981` | `#F59E0B` | `#132E27` | `#ECFDF5` | گزارش‌های پایداری، سلامت، مالی، بازرگانی |
| `ruby` | **یاقوتی لوکس (مدیریتی)** | `#881337` | `#E11D48` | `#CA8A04` | `#27181C` | `#FFF1F2` | گزارش‌های ویژه هیئت مدیره، تحلیل استراتژیک |
| `ocean` | **اقیانوسی / تیل مدرن** | `#0F766E` | `#14B8A6` | `#F97316` | `#1E293B` | `#F0FDFA` | اسناد فنی، استارتاپ‌ها، گزارش‌های توسعه |
| `violet` | **بنفش نوآوری (خلاقانه)** | `#4C1D95` | `#8B5CF6` | `#EC4899` | `#1E1B4B` | `#F5F3FF` | پروپوزال‌های نوآورانه، بازاریابی، طراحی |
| `slate` | **طوسی زغالی (تکنیکال / مینیمال)** | `#334155` | `#64748B` | `#2563EB` | `#0F172A` | `#F8FAFC` | اسناد مهندسی، مستندات نرم‌افزاری، استانداردها |

---

## ۳. نحوه استفاده سریع (Quick Start)

### الف) استفاده از اسکریپت آماده پایتون با تعیین Direction

```bash
# تولید سند با جهت پیش‌فرض راست‌به‌چپ (RTL)
python scripts/persian_docx_builder.py --palette navy --direction right-to-left --output report_fa.docx

# تولید سند با جهت چپ‌به‌راست (LTR)
python scripts/persian_docx_builder.py --palette slate --direction left-to-right --output report_en.docx
```

### ب) نمونه کد پایتون با کنترل کامل Indents, Spacing و Direction

```python
from scripts.persian_docx_builder import PersianDocxBuilder

# ایجاد سند با پالت سرمه‌ای سازمانی
doc = PersianDocxBuilder(palette="navy", font_name="Vazirmatn")

# افزودن عنوان
doc.add_title("گزارش جامع وضعیت پروژه", subtitle="نسخه ترکیبی فارسی و انگلیسی")

# ۱. پاراگراف فارسی (RTL با فواصل استاندارد)
doc.add_heading_1("۱. بخش فارسی (Right-to-Left)")
doc.add_paragraph(
    "این پاراگراف به صورت راست‌به‌چپ (RTL) با فاصله خطوط ۱.۳ و تراز دوطرفه درج شده است.",
    direction="right-to-left",
    space_after_pt=6.0,
    line_spacing=1.3
)

# ۲. پاراگراف انگلیسی (Left-to-Right با تنظیم Indents و Spacing)
doc.add_heading_2("2. English Section (Left-to-Right)")
doc.add_ltr_paragraph(
    "This paragraph is explicitly configured with direction: left-to-right, left alignment, "
    "and a left indentation of 0.5 cm.",
    left_indent_cm=0.5,
    space_after_pt=6.0,
    line_spacing=1.15
)

# ۳. کادر اعلان و جدول
doc.add_callout(
    text="تمام تنظیمات پنجره Paragraph در نرم‌افزار Word به درستی اعمال شده است.",
    title="تأیید تنظیمات",
    callout_type="success"
)

# ذخیره خروجی
doc.save("output_document.docx")
```

---

## ۴. اجزای موجود در این اسکیل

* **[اسکریپت سازنده سند (`scripts/persian_docx_builder.py`)](./scripts/persian_docx_builder.py):** ماژول کامل پایتون با توابع اختصاصی برای تنظیم `direction`, `indents` و `spacing`.
* **[راهنمای پالت‌های رنگی (`references/palettes.md`)](./references/palettes.md):** مشخصات رنگ‌های ۶ پالت.
* **[راهنمای تایپوگرافی و OpenXML (`references/typography_guide.md`)](./references/typography_guide.md):** راهنمای جزئیات فنی تگ‌های XML برای `w:ind`، `w:spacing`، `w:bidi` و `w:rtl`.
