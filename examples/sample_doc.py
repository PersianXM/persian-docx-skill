"""
Example script demonstrating Persian DOCX generation with various palettes and elements.
"""

import os
import sys

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Ensure scripts directory is in python path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, os.path.join(parent_dir, "scripts"))

from persian_docx_builder import PersianDocxBuilder, PALETTES


def generate_all_samples():
    output_dir = os.path.join(parent_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    print("🚀 در حال تولید اسناد نمونه با پالت‌های مختلف...")

    # Sample 1: Navy Corporate Proposal
    doc_navy = PersianDocxBuilder(palette="navy", font_name="Vazirmatn")
    doc_navy.set_header_footer("طرح توجیهی ارتقای زیرساخت سازمانی", "صفحه ۱ | شرکت فناوران داده")
    doc_navy.add_title("طرح پیشنهادی مدرن‌سازی شبکه و امنیت اطلاعات", subtitle="نسخه ۱.۴ - اسفند ۱۴۰۴", author="مهندس پارسا رضایی", date="۱۴۰۴/۱۲/۰۱")
    
    doc_navy.add_heading_1("۱. خلاصه مدیریتی و اهداف کلان")
    doc_navy.add_paragraph(
        "هدف از تدوین این سند، ارائه نقشه راه جامع جهت به‌روزرسانی زیرساخت‌های فناوری اطلاعات سازمان، افزایش پایداری سرویس‌ها "
        "و پیاده‌سازی سازوکارهای نوین امنیتی بر مبنای استاندارد ISO/IEC 27001 می‌باشد. اجرای این پروژه سبب کاهش ۷۰ درصدی زمان توقف ناخواسته سیستم‌ها خواهد شد."
    )
    doc_navy.add_callout(
        text="تمامی اقدامات اجرایی در ساعات خارج از اداری و تعطیلات پایان هفته انجام خواهد شد تا خللی در ارائه خدمات روزمره به مشتریان ایجاد نشود.",
        title="تعهد عدم ایجاد وقفه در عملیات",
        callout_type="success"
    )

    doc_navy.add_heading_2("۲. مشخصات فنی فازهای پیاده‌سازی")
    headers = ["فاز", "عنوان فاز", "مدت زمان", "پیش‌نیاز", "هزینه تخمینی (میلیون تومان)"]
    rows = [
        ["فاز ۱", "ارزیابی امنیتی و اسکن آسیب‌پذیری‌ها", "۲ هفته", "دسترسی به تجهیزات", "۴۵"],
        ["فاز ۲", "ارتقای هسته سوییچینگ و فایروال", "۳ هفته", "فاز ۱", "۱۲۰"],
        ["فاز ۳", "راه‌اندازی پشتیبان‌گیری خودکار ابری", "۲ هفته", "فاز ۲", "۶۰"],
        ["فاز ۴", "آزمون نفوذ و آموزش پرسنل", "۱ هفته", "فاز ۳", "۳۵"],
    ]
    doc_navy.add_table(headers=headers, rows=rows, col_widths=[1.5, 6.0, 2.5, 3.0, 3.5])
    
    doc_navy.add_heading_2("۳. نیازمندی‌ها و الزامات کلیدی")
    doc_navy.add_bullet_item("تأمین پهنای باند اختصاصی حداقل ۱۰۰ مگابیت بر ثانیه", bold_title="ارتباطات شبکه")
    doc_navy.add_bullet_item("تخصیص فضای دیتاسنتر با سرمایش استاندارد و برق اضطراری (UPS)", bold_title="زیرساخت فیزیکی")
    doc_navy.add_bullet_item("معرفی رابط فنی از سوی کارفرما جهت هماهنگی‌های دسترسی", bold_title="تیم کارفرما")

    doc_navy.save(os.path.join(output_dir, "sample_corporate_navy.docx"))

    # Sample 2: Emerald Modern Biotech / Financial
    doc_emerald = PersianDocxBuilder(palette="emerald", font_name="Vazirmatn")
    doc_emerald.set_header_footer("گزارش عملکرد مالی و پایداری زیست‌محیطی", "سال مالی ۱۴۰۴")
    doc_emerald.add_title("گزارش سالانه پایداری و مسئولیت اجتماعی", subtitle="بررسی شاخص‌های ESG و بازدهی سرمایه‌گذاری سبز", author="دپارتمان سرمایه‌گذاری سبز", date="۱۴۰۴/۱۲/۰۱")
    
    doc_emerald.add_heading_1("۱. بیانیه مسئولیت زیست‌محیطی")
    doc_emerald.add_paragraph("در سال مالی گذشته، شرکت توانسته است با بهینه‌سازی زنجیره تأمین و استفاده از انرژی‌های تجدیدپذیر، ردپای کربن خود را به میزان ۲۸ درصد کاهش دهد.")
    
    doc_emerald.add_callout(
        text="سرمایه‌گذاری در فناوری‌های پاک منجر به صرفه‌جویی سالانه بالغ بر ۳ میلیارد تومان در هزینه‌های انرژی گردیده است.",
        title="دستاورد برتر سال",
        callout_type="info"
    )

    doc_emerald.save(os.path.join(output_dir, "sample_sustainability_emerald.docx"))

    print(f"✨ تمامی اسناد نمونه با موفقیت در پوشه '{output_dir}' ایجاد شدند.")


if __name__ == "__main__":
    generate_all_samples()
