import os
import tempfile
from collections import Counter
from pypdf import PdfReader, PdfWriter
from docx import Document
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "أهلاً بك! أرسل لي ملف PDF وسأحاول تنظيف النصوص المكررة وتحويله إلى ملف Word (.docx)."
    )


def remove_repeated_text_objects(reader: PdfReader) -> PdfWriter:
    writer = PdfWriter()
    total_pages = len(reader.pages)
    if total_pages == 0:
        return writer

    line_counter = Counter()

    for page in reader.pages:
        lines_in_page = set()
        raw_text = page.extract_text() or ""
        for line in raw_text.splitlines():
            clean_line = line.strip()
            if len(clean_line) >= 3:
                lines_in_page.add(clean_line)
        for item in lines_in_page:
            line_counter[item] += 1

    threshold = max(2, int(total_pages * 0.6)) if total_pages > 2 else 999999
    repeated_lines = {item for item, count in line_counter.items() if count >= threshold}

    for page in reader.pages:
        if repeated_lines:
            try:
                page.remove_text(
                    visitor_text=lambda text, cm, tm, font_dict, font_size: (
                        "" if (text or "").strip() in repeated_lines else text
                    )
                )
            except Exception:
                pass
        writer.add_page(page)

    return writer


def convert_pdf_to_word(pdf_file_path: str, docx_file_path: str):
    reader = PdfReader(pdf_file_path)
    writer = remove_repeated_text_objects(reader)

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_clean_pdf:
        cleaned_pdf_path = tmp_clean_pdf.name
        writer.write(cleaned_pdf_path)

    try:
        clean_reader = PdfReader(cleaned_pdf_path)
        word_doc = Document()

        for page_num, page in enumerate(clean_reader.pages, start=1):
            word_doc.add_heading(f"صفحة {page_num}", level=2)
            extracted_text = page.extract_text() or ""
            lines = [line.strip() for line in extracted_text.splitlines() if line.strip()]

            if not lines:
                word_doc.add_paragraph(
                    "لا يوجد نص قابل للاستخراج في هذه الصفحة (قد تكون صفحة مصورة أو ممسوحة ضوئياً)."
                )
            else:
                for line in lines:
                    paragraph = word_doc.add_paragraph(line)
                    if any("\u0600" <= ch <= "\u06FF" for ch in line):
                        paragraph.paragraph_format.bidi = True

            if page_num < len(clean_reader.pages):
                word_doc.add_page_break()

        word_doc.save(docx_file_path)
    finally:
        if os.path.exists(cleaned_pdf_path):
            os.remove(cleaned_pdf_path)


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_msg = update.message.document
    if not doc_msg:
        return

    original_name = doc_msg.file_name or "document.pdf"
    if not original_name.lower().endswith(".pdf"):
        await update.message.reply_text("من فضلك أرسل ملفاً بصيغة PDF فقط.")
        return

    wait_msg = await update.message.reply_text("جاري تحميل الملف وتنظيفه وتحويله إلى Word...")

    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, "input.pdf")
        base_name = os.path.splitext(original_name)[0]
        docx_path = os.path.join(temp_dir, f"{base_name}_converted.docx")

        tg_file = await doc_msg.get_file()
        await tg_file.download_to_drive(pdf_path)

        try:
            convert_pdf_to_word(pdf_path, docx_path)
            with open(docx_path, "rb") as docx_file:
                await update.message.reply_document(
                    document=docx_file,
                    filename=os.path.basename(docx_path),
                    caption="تم تحويل الملف وإرسال نسخة Word."
                )
        except Exception as err:
            await update.message.reply_text(f"حدث خطأ أثناء المعالجة: {err}")
        finally:
            await wait_msg.delete()


def run_bot():
    if not BOT_TOKEN:
        raise RuntimeError("لم يتم العثور على متغير البيئة BOT_TOKEN.")

    application = ApplicationBuilder().token(BOT_TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_pdf))
    print("Bot started successfully.")
    application.run_polling()


if __name__ == "__main__":
    run_bot()
