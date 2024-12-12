from telegram import Update, Document
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import openpyxl
import os
from dotenv import load_dotenv

# ذخیره مسیر فایل اکسل
EXCEL_FILE_PATH = "uploaded_data.xlsx"

# دستور /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "سلام! لطفاً فایل اکسل خود را ارسال کنید تا آن را پردازش کنم. سپس کلمه‌ای برای جستجو بفرستید."
    )

# دریافت فایل اکسل از کاربر
async def receive_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document: Document = update.message.document

    if document.file_name.endswith(".xlsx"):
        # دانلود فایل
        file = await document.get_file()
        await file.download_to_drive(EXCEL_FILE_PATH)

        await update.message.reply_text("فایل اکسل با موفقیت ذخیره شد! حالا کلمه‌ای برای جستجو بفرستید.")
    else:
        await update.message.reply_text("لطفاً فقط فایل اکسل با فرمت .xlsx ارسال کنید.")

# تابع جستجو در اکسل (فقط در "نام فارسی" و "نام انگلیسی")
def search_in_excel(keyword):
    if not os.path.exists(EXCEL_FILE_PATH):
        return None

    # باز کردن فایل اکسل
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb.active

    # شناسایی سرستون‌ها
    headers = {cell.value: idx for idx, cell in enumerate(sheet[1]) if cell.value}

    # بررسی وجود سرستون‌های موردنظر
    required_headers = ["نام فارسی", "نام انگلیسی", "لینک در کانال"]
    if not all(header in headers for header in required_headers):
        return "missing_headers"

    # لیست نتایج
    results = []

    # جستجو فقط در ستون‌های "نام فارسی" و "نام انگلیسی" (ردیف دوم به بعد)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # گرفتن مقادیر ستون‌های مورد نظر
        name_farsi = row[headers["نام فارسی"]]
        name_english = row[headers["نام انگلیسی"]]
        link = row[headers["لینک در کانال"]]
        
        # بررسی کلمه کلیدی در این دو ستون
        if (
            keyword.lower() in str(name_farsi).lower() or
            keyword.lower() in str(name_english).lower()
        ):
            # گرفتن مقادیر از ستون‌های مشخص
            result = {
                "نام فارسی": name_farsi,
                "نام انگلیسی": name_english,
                "لینک در کانال": link
            }
            results.append(result)

    wb.close()
    return results

# گرفتن کلمه و جستجو
async def search(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyword = update.message.text  # کلمه وارد شده توسط کاربر

    results = search_in_excel(keyword)

    if results is None:
        await update.message.reply_text("ابتدا یک فایل اکسل ارسال کنید.")
    elif results == "missing_headers":
        await update.message.reply_text("فایل اکسل باید شامل سرستون‌های 'نام فارسی'، 'نام انگلیسی' و 'لینک در کانال' باشد.")
    elif results:
        # ساختن متن خروجی
        output = "نتایج پیدا شده:\n"
        for idx, result in enumerate(results, start=1):
            output += (
                f"{idx} - : {result['نام فارسی']}\n"
                f"<a href='{result['لینک در کانال']}'>{result['نام انگلیسی']}</a>\n\n"
            )

        # تقسیم پیام‌ها اگر از 4096 کاراکتر بیشتر شد
        messages = []
        while len(output) > 4096:
            split_index = output.rfind("\n", 0, 4096)  # پیدا کردن آخرین خط قبل از 4096
            if split_index == -1:  # اگر خطی پیدا نشد، مجبوریم پیام را در نقطه‌ای ببریم
                split_index = 4096
            messages.append(output[:split_index])
            output = output[split_index:]

        # اضافه کردن پیام نهایی باقی‌مانده
        if output:
            messages.append(output)

        # ارسال پیام‌ها به صورت جداگانه
        for msg in messages:
            await update.message.reply_text(
                msg, parse_mode="HTML", disable_web_page_preview=True
            )
    else:
        await update.message.reply_text("متاسفم! هیچ نتیجه‌ای پیدا نشد.")

def main():
    load_dotenv()
    # توکن ربات خود را جایگزین کنید
    TOKEN = os.getenv('TOKEN')

    # ایجاد اپلیکیشن
    application = ApplicationBuilder().token(TOKEN).build()

    # افزودن دستور /start
    application.add_handler(CommandHandler("start", start))

    # افزودن MessageHandler برای دریافت فایل اکسل
    application.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), receive_excel))

    # افزودن MessageHandler برای جستجوی کلمه
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, search))

    # اجرا
    application.run_polling()

if __name__ == "__main__":
    main()
