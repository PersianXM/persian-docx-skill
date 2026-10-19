# راهنمای فنی تایپوگرافی و ساختار OpenXML برای اسناد فارسی (RTL)

تولید فایل‌های Word (.docx) استاندارد برای زبان‌های راست‌به‌چپ (مانند فارسی و عربی) چالش‌هایی دارد که بدون تنظیم مستقیم تگ‌های OpenXML، ظاهر سند در Microsoft Word با مشکلاتی نظیر عدم اتصال صحیح حروف، جابجایی علائم نگارشی و ترتیب نادرست ستون‌های جدول مواجه می‌شود.

---

## ۱. جدول مقادیر استاندارد تایپوگرافی فارسی

| عنصر (Element) | اندازه (Pt) | ضخامت (Weight) | رنگ پیشنهادی | فاصله قبل (Before) | فاصله بعد (After) | فاصله خطوط (Line Spacing) | چینش (Alignment) |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **عنوان اصلی سند (Title)** | ۲۴ - ۲۶ | Bold | Primary Palette | ۱۸ pt | ۸ pt | ۱.۱۵ | راست / وسط |
| **زیرعنوان سند (Subtitle)** | ۱۴ - ۱۶ | Regular/Medium | Secondary Palette | ۰ pt | ۱۲ pt | ۱.۲ | راست / وسط |
| **تیتر سطح ۱ (Heading 1)** | ۱۸ - ۲۰ | Bold | Primary Palette | ۱۶ - ۱۸ pt | ۶ pt | ۱.۲ | راست |
| **تیتر سطح ۲ (Heading 2)** | ۱۵ - ۱۶ | Bold | Secondary Palette | ۱۲ - ۱۴ pt | ۴ pt | ۱.۲ | راست |
| **تیتر سطح ۳ (Heading 3)** | ۱۳ - ۱۴ | SemiBold | Dark Text | ۸ - ۱۰ pt | ۳ pt | ۱.۲۵ | راست |
| **تیتر سطح ۴ (Heading 4)** | ۱۱.۵ - ۱۲ | SemiBold | Muted Text | ۶ - ۸ pt | ۲ pt | ۱.۲۵ | راست |
| **متن بدنه (Body Paragraph)**| ۱۱ - ۱۲ | Regular | Dark Text (`#1E293B`) | ۰ pt | ۶ pt | **۱.۳ - ۱.۳۵** | Justified (دوطرفه) |
| **آیتم‌های لیست (Bullet/Num)** | ۱۱ | Regular | Dark Text | ۲ pt | ۳ pt | ۱.۲۵ | راست (با تورفتگی 0.6cm) |
| **سربرگ جدول (Table Header)** | ۱۰ - ۱۱ | Bold | سفید (`#FFFFFF`) | ۰ pt | ۰ pt | ۱.۱۵ | وسط / راست |
| **سلول‌های داده (Data Cells)** | ۹.۵ - ۱۰.۵ | Regular | Dark Text | ۰ pt | ۰ pt | ۱.۱۵ | وسط (اعداد) / راست (متن) |
| **کادر نکته (Callout)** | ۱۰.۵ - ۱۱ | Regular | Dark Text | ۰ pt | ۲ pt | ۱.۳ | راست |
| **پاورقی و پی‌نوشت (Footnote)** | ۹ - ۹.۵ | Regular | Muted Text | ۰ pt | ۲ pt | ۱.۱۵ | راست |

> **نکته کلیدی فاصله خطوط:** فونت‌های فارسی (به‌ویژه Vazirmatn) به دلیل وجود اعراب، کاف سرکش‌دار و حرکات عمودی، نیازمند فاصله خطوط حداقل **۱.۳ برابر** هستند. تنظیم فاصله خط روی ۱.۰ یا ۱.۱۵ باعث فشردگی و خستگی چشم خواننده می‌گردد.

---

## ۲. مشخصات فونت استاندارد Vazirmatn و مپینگ قلم‌ها

در OpenXML سند Word، هر Run دارای تنظیمات فونت چندگانه است:
* `w:ascii` و `w:hAnsi`: برای حروف انگلیسی و لاتین.
* `w:cs` (Complex Script) و `w:eastAsia`: برای حروف فارسی و عربی.

### نمونه XML استاندارد برای یک Run فارسی:
```xml
<w:rPr>
  <w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Vazirmatn" w:eastAsia="Vazirmatn"/>
  <w:rtl w:val="1"/>
  <w:sz w:val="23"/> <!-- 11.5 pt (23 half-points) -->
  <w:szCs w:val="23"/>
  <w:color w:val="1E293B"/>
</w:rPr>
```

---

## ۳. تنظیمات حیاتی RTL در لایه‌های مختلف OpenXML

### الف) سطح بخش (Section Level - `w:sectPr`)
برای این‌که Word بفهمد صفحه سند راست‌به‌چپ است و شماره‌گذاری و جهت صفحات هماهنگ شود:
```xml
<w:sectPr>
  <w:bidi w:val="1"/>
  <!-- Margins: 2.5cm = 1417 dxa -->
  <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/>
</w:sectPr>
```

### ب) سطح پاراگراف (Paragraph Level - `w:pPr`)
```xml
<w:pPr>
  <w:bidi w:val="1"/>
  <w:jc w:val="both"/> <!-- Justify alignment -->
  <w:spacing w:before="0" w:after="120" w:line="312" w:lineRule="auto"/> <!-- 1.3x spacing -->
</w:pPr>
```

### ج) سطح جدول (Table Level - `w:tblPr`)
در Word اگر تگ `w:tblBidi` وجود نداشته باشد، ستون اول جدول در سمت چپ قرار می‌گیرد (چپ‌به‌راست). با اضافه کردن این تگ، ستون اول در سمت راست قرار می‌گیرد:
```xml
<w:tblPr>
  <w:tblBidi w:val="1"/>
  <w:jc w:val="center"/>
</w:tblPr>
```
