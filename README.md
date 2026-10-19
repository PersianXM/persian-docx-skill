# persian-docx skill

<div dir="rtl">

اسکیلی برای ساخت اسناد **Microsoft Word (.docx)** فارسی با پشتیبانی کامل راست‌به‌چپ (RTL) و چپ‌به‌راست (LTR)، فونت **Vazirmatn**، تنظیمات دقیق Indents و Spacing و شش پالت رنگی حرفه‌ای.

این اسکیل با **Claude Code** و **Google Antigravity** کار می‌کند.

## نصب

### Claude Code

```bash
git clone https://github.com/PersianXM/persian-docx-skill.git ~/.claude/skills/persian-docx
```

### Google Antigravity

```bash
git clone https://github.com/PersianXM/persian-docx-skill.git ~/.gemini/config/skills/persian-docx
```

### پیش‌نیازها

```bash
pip install python-docx
```

فونت [Vazirmatn](https://github.com/rastikerdar/vazirmatn) باید روی سیستم نصب باشد تا سند درست نمایش داده شود.

## استفاده

بعد از نصب کافی است از دستیار صراحتاً یک فایل Word فارسی بخواهید؛ مثلاً:

> یک گزارش docx فارسی درباره وضعیت پروژه با پالت سرمه‌ای بساز.

یا اسکریپت را مستقیم اجرا کنید:

```bash
python scripts/persian_docx_builder.py --palette navy --direction right-to-left --output report_fa.docx
```

## پالت‌ها

`navy` · `emerald` · `ruby` · `ocean` · `violet` · `slate`. جزئیات در [references/palettes.md](references/palettes.md).

## ساختار

| مسیر | توضیح |
| --- | --- |
| `SKILL.md` | دستورالعمل اصلی اسکیل |
| `scripts/persian_docx_builder.py` | ماژول پایتون سازنده سند |
| `references/` | راهنمای پالت‌ها و تایپوگرافی OpenXML |
| `examples/sample_doc.py` | نمونه کد |
| `output/` | دو سند نمونه خروجی |

</div>
