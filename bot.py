import asyncio
import glob
import os
import re
import tempfile
from collections import Counter
from docx import Document
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_ARABIC_SHAPER = True
except ImportError:
    HAS_ARABIC_SHAPER = False

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "أهلاً بك! أرسل لي:\n"
        "• ملف PDF لتحويله إلى Word (.docx) مع إزالة النصوص المكررة وتحسين النص العربي.\n"
        "• أو ملف Word (.docx) لتحويله إلى PDF (.pdf)."
    )


def has_arabic(text: str) -> bool:
    return any("\u0600" <= ch <= "\u06FF" for ch in text)


def fix_arabic_text(text: str) -> str:
    if not text:
        return ""

    replacements = {
        "\u0640": "",  # إزالة التطويل/الكشيدة الزائدة
        "\u200f": "",  # إزالة علامة RTL الخفية إن وجدت
        "\u200e": "",  # إزالة علامة LTR الخفية إن وجدت
        "ٌ": "",
        "ٍ": "",
        "َ": "",
        "ُ": "",
        "ِ": "",
        "ّ": "",
        "ْ": "",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)

    common_fixes = {
        "التى": "التي",
        "فى": "في",
        "الذى": "الذي",
        "الثانى": "الثاني",
        "الثالث": "الثالث",
        "الفصل الدراسى": "الفصل الدراسي",
    }
    for wrong, right in common_fixes.items():
        text = re.sub(rf"\b{re.escape(wrong)}\b", right, text)

    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def format_line_for_pdf(line: str) -> str:
    if has_arabic(line) and HAS_ARABIC_SHAPER:
        try:
            reshaped = arabic_reshaper.reshape(line)
            return get_display(reshaped)
        except Exception:
            return line
    return line


def find_ttf_font() -> str | None:
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path

    found = glob.glob("/usr/share/fonts/**/*.ttf", recursive=True)
    return found[0] if found else None


def is_explicit_watermark(line: str) -> bool:
    lower = line.lower()
    watermark_keywords = [
        "watermark",
        "confidential",
        "نسخة تجريبية للمراجعة",
    ]
    return any(kw in lower for kw in watermark_keywords)


def convert_pdf_to_word(pdf_file_path: str, docx_file_path: str):
    reader = PdfReader(pdf_file_path)
    total_pages = len(reader.pages)
    if total_pages == 0:
        word_doc = Document()
        word_doc.add_paragraph("الملف فارغ.")
        word_doc.save(docx_file_path)
        return

    pages_lines = []
    line_page_counter = Counter()

    for page in reader.pages:
        raw_text = page.extract_text() or ""
        page_unique = set()
        current_page_lines = []

        for raw_line in raw_text.splitlines():
            cleaned = fix_arabic_text(raw_line)
            if not cleaned:
                continue
            current_page_lines.append(cleaned)
            if len(cleaned) >= 4:
                page_unique.add(cleaned)

        pages_lines.append(current_page_lines)
        for item in page_unique:
            line_page_counter[item] += 1

    threshold = max(2, int(total_pages * 0.6)) if total_pages > 2 else 999999
    repeated_watermarks = {
        item for item, count in line_page_counter.items() if count >= threshold
    }

    word_doc = Document()
    normal_style = word_doc.styles["Normal"]
    normal_style.font.name = "Arial"

    for page_num, lines in enumerate(pages_lines, start=1):
        heading = word_doc.add_heading(f"صفحة {page_num}", level=2)
        heading.paragraph_format.bidi = True

        filtered_lines = []
        prev_line = None

        for line in lines:
            if is_explicit_watermark(line) or line in repeated_watermarks:
                continue
            if line == prev_line:
                continue
            filtered_lines.append(line)
            prev_line = line

        if not filtered_lines:
            p = word_doc.add_paragraph(
                "لا يوجد نص قابل للاستخراج في هذه الصفحة (قد تكون صفحة مصورة أو ممسوحة ضوئياً)."
            )
            p.paragraph_format.bidi = True
        else:
            for line in filtered_lines:
                if re.match(r"^(\d+[\.\-]|السؤال\s+\S+|أولاً|ثانياً|ثالثاً|رابعاً|خامساً|سادساً|سابعاً)", line):
                    word_doc.add_paragraph("")

                p = word_doc.add_paragraph(line)
                p.paragraph_format.space_after = 4
                if has_arabic(line):
                    p.paragraph_format.bidi = True

        if page_num < total_pages:
            word_doc.add_page_break()

    word_doc.save(docx_file_path)


def convert_word_to_pdf(docx_file_path: str, pdf_file_path: str):
    doc = Document(docx_file_path)
    c = canvas.Canvas(pdf_file_path, pagesize=A4)
    width, height = A4

    font_name = "Helvetica"
    ttf_path = find_ttf_font()
    if ttf_path:
        try:
            pdfmetrics.registerFont(TTFont("CustomFont", ttf_path))
            font_name = "CustomFont"
        except Exception:
            font_name = "Helvetica"

    margin_x = 45
    top_y = height - 50
    bottom_y = 50
    line_height = 20
    y = top_y

    c.setFont(font_name, 12)

    all_lines = []
    for p in doc.paragraphs:
        txt = fix_arabic_text(p.text or "")
        if txt:
            for subline in txt.splitlines():
                clean_sub = fix_arabic_text(subline)
                if clean_sub:
                    all_lines.append(clean_sub)

    for table in doc.tables:
        for row in table.rows:
            row_cells = [
                fix_arabic_text(cell.text.replace("\n", " "))
                for cell in row.cells
                if fix_arabic_text(cell.text)
            ]
            if row_cells:
                all_lines.append(" | ".join(row_cells))

    if not all_lines:
        all_lines = ["(الملف فارغ أو لا يحتوي على نصوص قابلة للتحويل)"]

    for raw_line in all_lines:
        max_chars = 85
        chunks = [raw_line[i:i + max_chars] for i in range(0, len(raw_line), max_chars)]

        for chunk in chunks:
            if y <= bottom_y:
                c.showPage()
                c.setFont(font_name, 12)
                y = top_y

            display_text = format_line_for_pdf(chunk)
            if has_arabic(chunk) and font_name != "Helvetica":
                c.drawRightString(width - margin_x, y, display_text)
            else:
                c.drawString(margin_x, y, display_text)
            y -= line_height

    c.save()


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_msg = update.message.document
    if not doc_msg:
        return

    original_name = doc_msg.file_name or "file"
    lower_name = original_name.lower()

    if not (lower_name.endswith(".pdf") or lower_name.endswith(".docx")):
        await update.message.reply_text("من فضلك أرسل ملفاً بصيغة PDF أو Word (.docx) فقط.")
        return

    wait_msg = await update.message.reply_text("جاري استلام الملف ومعالجته...")

    with tempfile.TemporaryDirectory() as temp_dir:
        base_name = os.path.splitext(original_name)[0]
        tg_file = await doc_msg.get_file()

        try:
            if lower_name.endswith(".pdf"):
                input_path = os.path.join(temp_dir, "input.pdf")
                output_path = os.path.join(temp_dir, f"{base_name}_converted.docx")
                await tg_file.download_to_drive(input_path)
                convert_pdf_to_word(input_path, output_path)
                caption = "تم تحويل PDF إلى Word (.docx)."
            else:
                input_path = os.path.join(temp_dir, "input.docx")
                output_path = os.path.join(temp_dir, f"{base_name}_converted.pdf")
                await tg_file.download_to_drive(input_path)
                convert_word_to_pdf(input_path, output_path)
                caption = "تم تحويل Word (.docx) إلى PDF."

            with open(output_path, "rb") as out_file:
                await update.message.reply_document(
                    document=out_file,
                    filename=os.path.basename(output_path),
                    caption=caption,
                )
        except Exception as err:
            await update.message.reply_text(f"حدث خطأ أثناء المعالجة: {err}")
        finally:
            await wait_msg.delete()


def run_bot():
    if not BOT_TOKEN:
        raise RuntimeError("لم يتم العثور على متغير البيئة BOT_TOKEN.")

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    application = ApplicationBuilder().token(BOT_TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    print("Bot started successfully.")
    application.run_polling()


if __name__ == "__main__":
    run_bot()
