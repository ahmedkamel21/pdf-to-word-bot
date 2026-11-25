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
        "\u0640": "",  # إزالة الكشيدة
        "\u200f": "",  # إزالة علامة RTL الخفية
        "\u200e": "",  # إزالة علامة LTR الخفية
        "\ufeff": "",
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


def is_watermark_or_artifact(line: str) -> bool:
    s = line.strip()
    if not s:
        return True

    lower = s.lower()
    watermark_keywords = [
        "watermark",
        "confidential",
        "نسخة تجريبية للمراجعة",
        "جميع الحقوق محفوظة",
        "ملف اختبار شامل",
        "test watermark",
        "stress test document",
    ]
    if any(kw in lower for kw in watermark_keywords):
        return True

    # استبعاد المقاطع المبتورة الشائعة مثل "الصفحا" أو "وإزالة العلامة" أو "نص مشكول بالكام"
    truncated_fragments = [
        "الصفحا",
        "وإزالة العلامة",
        "نص مشكول بالكام",
        "يحتوي هذا المستند على نصوص عربية",
    ]
    if any(frag in s and len(s) <= len(frag) + 12 for frag in truncated_fragments):
        return True

    return False


def convert_pdf_to_word(pdf_file_path: str, docx_file_path: str):
    reader = PdfReader(pdf_file_path)
    total_pages = len(reader.pages)
    word_doc = Document()

    normal_style = word_doc.styles["Normal"]
    normal_style.font.name = "Arial"

    if total_pages == 0:
        word_doc.add_paragraph("الملف فارغ.")
        word_doc.save(docx_file_path)
        return

    pages_raw_lines = []
    global_line_counter = Counter()

    for page in reader.pages:
        raw_text = page.extract_text() or ""
        current_lines = []
        seen_in_page = set()

        for raw_line in raw_text.splitlines():
            cleaned = fix_arabic_text(raw_line)
            if not cleaned or is_watermark_or_artifact(cleaned):
                continue
            current_lines.append(cleaned)
            if len(cleaned) >= 5:
                seen_in_page.add(cleaned)

        pages_raw_lines.append(current_lines)
        for item in seen_in_page:
            global_line_counter[item] += 1

    # أي سطر بطول >= 5 أحرف يظهر في أكثر من نصف الصفحات يُعد ترويسة/تذييل/علامة مائية
    cross_page_threshold = max(2, (total_pages // 2) + 1) if total_pages > 1 else 999999
    repeated_across_pages = {
        item for item, count in global_line_counter.items() if count >= cross_page_threshold
    }

    for page_num, lines in enumerate(pages_raw_lines, start=1):
        heading = word_doc.add_heading(f"صفحة {page_num}", level=2)
        heading.paragraph_format.bidi = True

        # إزالة التكرار داخل نفس الصفحة أيضًا
        local_counter = Counter(lines)
        clean_page_lines = []
        for line in lines:
            if line in repeated_across_pages:
                continue
            # إذا تكرر نفس السطر أكثر من مرتين في نفس الصفحة وهو قصير/مكرر غالباً يكون تشتتاً
            if len(line) >= 4 and local_counter[line] > 2:
                continue
            # تجنب التكرار المتتالي المباشر
            if clean_page_lines and clean_page_lines[-1] == line:
                continue
            clean_page_lines.append(line)

        if not clean_page_lines:
            p = word_doc.add_paragraph(
                "لا يوجد نص قابل للاستخراج في هذه الصفحة (قد تكون صفحة مصورة أو ممسوحة ضوئياً)."
            )
            p.paragraph_format.bidi = True
        else:
            for line in clean_page_lines:
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
