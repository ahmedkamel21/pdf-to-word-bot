import asyncio
import glob
import os
import re
import tempfile
from collections import Counter
from docx import Document
from docx.shared import Pt
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
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
        "• ملف PDF لتحويله إلى Word (.docx) مع إزالة النصوص المكررة وتنسيق الجداول والفقرات.\n"
        "• أو ملف Word (.docx) لتحويله إلى PDF (.pdf) مع الحفاظ على الجداول والتنسيق قدر الإمكان."
    )


def has_arabic(text: str) -> bool:
    return any("\u0600" <= ch <= "\u06FF" for ch in text)


def clean_noisy_fragments(text: str) -> str:
    if not text:
        return ""

    text = re.sub(r"(?:الصفحا)+", " ", text)

    noisy_patterns = [
        r"confidential\s*watermark\s*2026",
        r"test\s*watermark",
        r"stress\s*test\s*document\s*#?2026",
        r"نسخة\s*تجريبية\s*للمراجعة",
        r"جميع\s*الحقوق\s*محفوظة\s*2026",
        r"ملف\s*اختبار\s*شامل\s*-\s*علامة\s*مائية\s*نصية\s*متكررة\s*\|?",
        r"وإزالة\s*العلامة(?:\s*المائية)?",
        r"نص\s*مشكول\s*بالكام(?:ل)?",
        r"يحتوي\s*هذا\s*المستند\s*على\s*نصوص\s*عربية[^\n]*",
    ]
    for pat in noisy_patterns:
        text = re.sub(pat, " ", text, flags=re.IGNORECASE)

    return text


def fix_arabic_text(text: str) -> str:
    if not text:
        return ""

    replacements = {
        "\u0640": "",
        "\u200f": "",
        "\u200e": "",
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

    text = clean_noisy_fragments(text)

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


def shape_for_reportlab(text: str) -> str:
    if not text:
        return ""
    if has_arabic(text) and HAS_ARABIC_SHAPER:
        try:
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        except Exception:
            return text
    return text


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


def parse_table_row_if_any(line: str) -> list[str] | None:
    """
    يكتشف فقط السطور الواضحة التي تحتوي على أعمدة نصية مفصولة بـ '|'.
    """
    if "|" not in line:
        return None
    cells = [c.strip() for c in line.split("|") if c.strip()]
    if len(cells) < 2:
        return None
    # تجاهل الأسطر التي كلها رموز فقط بدون حروف أو أرقام مفيدة
    meaningful = [c for c in cells if re.search(r"[\w\u0600-\u06FF]", c)]
    if len(meaningful) < 2:
        return None
    return cells


def add_lines_with_tables_to_docx(word_doc: Document, lines: list[str]):
    """
    يضيف الأسطر لملف Word مع بناء جدول حقيقي فقط إذا وجد صفان متتاليان أو أكثر.
    """
    pending_table_rows: list[list[str]] = []

    def flush_table():
        nonlocal pending_table_rows
        if not pending_table_rows:
            return

        # لو صف واحد فقط أو أقل من صفين، نكتبه كفقرة عادية بدل جدول مبعثر
        if len(pending_table_rows) < 2:
            for r in pending_table_rows:
                joined = " | ".join(r)
                p = word_doc.add_paragraph(joined)
                p.paragraph_format.space_after = Pt(4)
                if has_arabic(joined):
                    p.paragraph_format.bidi = True
            pending_table_rows = []
            return

        max_cols = min(6, max(len(r) for r in pending_table_rows))
        tbl = word_doc.add_table(rows=len(pending_table_rows), cols=max_cols)
        tbl.style = "Table Grid"

        for r_idx, row_data in enumerate(pending_table_rows):
            for c_idx in range(max_cols):
                cell_text = row_data[c_idx] if c_idx < len(row_data) else ""
                cell = tbl.cell(r_idx, c_idx)
                cell.text = cell_text
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(2)
                    if has_arabic(cell_text):
                        p.paragraph_format.bidi = True

        word_doc.add_paragraph("")
        pending_table_rows = []

    for line in lines:
        row_cells = parse_table_row_if_any(line)
        if row_cells:
            pending_table_rows.append(row_cells)
            continue

        flush_table()

        is_heading = bool(
            re.match(
                r"^(\d+[\.\)]\s+.{2,60}|السؤال\s+\S+|أولاً|ثانياً|ثالثاً|رابعاً|خامساً|سادساً|سابعاً)",
                line,
            )
            and len(line) <= 90
        )

        if is_heading:
            h = word_doc.add_paragraph()
            run = h.add_run(line)
            run.bold = True
            run.font.size = Pt(12)
            h.paragraph_format.space_before = Pt(6)
            h.paragraph_format.space_after = Pt(3)
            if has_arabic(line):
                h.paragraph_format.bidi = True
        else:
            p = word_doc.add_paragraph(line)
            p.paragraph_format.space_after = Pt(4)
            if has_arabic(line):
                p.paragraph_format.bidi = True

    flush_table()


def is_junk_short_line(line: str) -> bool:
    s = line.strip()
    if not s:
        return True
    if re.fullmatch(r"[\-\—\_\.\|\s/\\]+", s):
        return True
    if s in {"Word DF", "TLFixERR-001_AR%"}:
        return True
    return False


def convert_pdf_to_word(pdf_file_path: str, docx_file_path: str):
    reader = PdfReader(pdf_file_path)
    total_pages = len(reader.pages)
    word_doc = Document()

    normal_style = word_doc.styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(11)

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
            if not cleaned or is_junk_short_line(cleaned):
                continue
            current_lines.append(cleaned)
            if len(cleaned) >= 6:
                seen_in_page.add(cleaned)

        pages_raw_lines.append(current_lines)
        for item in seen_in_page:
            global_line_counter[item] += 1

    cross_page_threshold = max(2, (total_pages // 2) + 1) if total_pages > 1 else 999999
    repeated_across_pages = {
        item for item, count in global_line_counter.items() if count >= cross_page_threshold
    }

    for page_num, lines in enumerate(pages_raw_lines, start=1):
        heading = word_doc.add_heading(f"صفحة {page_num}", level=2)
        heading.paragraph_format.bidi = True

        local_counter = Counter(lines)
        seen_long_in_page = set()
        clean_page_lines = []

        for line in lines:
            if line in repeated_across_pages:
                continue
            if len(line) >= 5 and local_counter[line] > 2:
                continue
            if clean_page_lines and clean_page_lines[-1] == line:
                continue
            # منع تكرار الجمل الطويلة داخل نفس الصفحة
            if len(line) >= 25 and line in seen_long_in_page:
                continue

            clean_page_lines.append(line)
            if len(line) >= 25:
                seen_long_in_page.add(line)

        if not clean_page_lines:
            p = word_doc.add_paragraph(
                "لا يوجد نص قابل للاستخراج في هذه الصفحة (قد تكون صفحة مصورة أو ممسوحة ضوئياً)."
            )
            p.paragraph_format.bidi = True
        else:
            add_lines_with_tables_to_docx(word_doc, clean_page_lines)

        if page_num < total_pages:
            word_doc.add_page_break()

    word_doc.save(docx_file_path)


def convert_word_to_pdf(docx_file_path: str, pdf_file_path: str):
    doc = Document(docx_file_path)

    font_name = "Helvetica"
    ttf_path = find_ttf_font()
    if ttf_path:
        try:
            pdfmetrics.registerFont(TTFont("CustomFont", ttf_path))
            font_name = "CustomFont"
        except Exception:
            font_name = "Helvetica"

    pdf_doc = SimpleDocTemplate(
        pdf_file_path,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=40,
        bottomMargin=40,
    )

    styles = getSampleStyleSheet()
    ar_style = ParagraphStyle(
        "ArabicStyle",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=11,
        leading=18,
        alignment=2 if font_name != "Helvetica" else 0,
        spaceAfter=6,
    )
    ltr_style = ParagraphStyle(
        "LtrStyle",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=10.5,
        leading=16,
        alignment=0,
        spaceAfter=6,
    )

    story = []
    has_content = False

    for p in doc.paragraphs:
        txt = fix_arabic_text(p.text or "")
        if not txt:
            continue
        has_content = True
        for subline in txt.splitlines():
            clean_sub = fix_arabic_text(subline)
            if not clean_sub:
                continue
            safe_text = (
                clean_sub.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )
            shaped = shape_for_reportlab(safe_text)
            style_to_use = ar_style if (has_arabic(clean_sub) and font_name != "Helvetica") else ltr_style
            story.append(Paragraph(shaped, style_to_use))

    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_cells = []
            for cell in row.cells:
                cell_txt = fix_arabic_text(cell.text.replace("\n", " "))
                safe_cell = (
                    cell_txt.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                )
                shaped_cell = shape_for_reportlab(safe_cell)
                cell_style = ar_style if (has_arabic(cell_txt) and font_name != "Helvetica") else ltr_style
                row_cells.append(Paragraph(shaped_cell, cell_style))
            if any(c.text.strip() for c in row_cells):
                table_data.append(row_cells)

        if table_data:
            has_content = True
            cols_count = max(len(r) for r in table_data)
            for r in table_data:
                while len(r) < cols_count:
                    r.append(Paragraph("", ltr_style))

            usable_width = A4[0] - 72
            col_w = usable_width / cols_count

            t = Table(table_data, colWidths=[col_w] * cols_count, repeatRows=1)
            t.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a5f")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f8fafc")),
                        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f9")]),
                        ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#cbd5e1")),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("TOPPADDING", (0, 0), (-1, -1), 6),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                        ("LEFTPADDING", (0, 0), (-1, -1), 6),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ]
                )
            )
            story.append(t)
            story.append(Spacer(1, 10))

    if not has_content:
        story.append(Paragraph(shape_for_reportlab("الملف فارغ أو لا يحتوي على نصوص قابلة للتحويل."), ar_style))

    pdf_doc.build(story)


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
