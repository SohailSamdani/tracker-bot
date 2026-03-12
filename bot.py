import os
import re
import logging
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import pytesseract
from PIL import Image
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TOKEN = os.environ.get("BOT_TOKEN", "")
EXCEL_FILE = "tracker_log.xlsx"

# ── Excel helpers ────────────────────────────────────────────────────────────

def get_or_create_workbook():
    if os.path.exists(EXCEL_FILE):
        return openpyxl.load_workbook(EXCEL_FILE)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tracker Log"
    headers = ["S.No", "Node ID", "Master", "Manual Cmd Angle", "Date Added", "Source Image"]
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="1F4E79")
    center = Alignment(horizontal="center", vertical="center")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 22
    ws.freeze_panes = "A2"
    wb.save(EXCEL_FILE)
    return wb


def get_existing_node_ids(wb):
    ws = wb["Tracker Log"]
    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1]:
            existing.add(str(row[1]).strip().upper())
    return existing


def append_nodes(nodes, image_name):
    wb = get_or_create_workbook()
    ws = wb["Tracker Log"]
    existing = get_existing_node_ids(wb)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    cell_font = Font(name="Arial", size=10)
    alt_fill = PatternFill("solid", start_color="DCE6F1")

    added = 0
    skipped = 0
    today = datetime.now().strftime("%d-%m-%Y")

    for node_id, master in nodes:
        node_id_upper = node_id.strip().upper()
        if node_id_upper in existing:
            skipped += 1
            continue
        row_num = ws.max_row + 1
        s_no = row_num - 1
        row_data = [s_no, node_id_upper, master, "", today, image_name]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col, value=val)
            cell.font = cell_font
            cell.alignment = center
            cell.border = border
            if s_no % 2 == 0:
                cell.fill = alt_fill
        existing.add(node_id_upper)
        added += 1

    wb.save(EXCEL_FILE)
    return added, skipped


# ── OCR & Parsing ────────────────────────────────────────────────────────────

def parse_genius_vision_image(image_bytes):
    img = Image.open(io.BytesIO(image_bytes))
    # Upscale for better OCR accuracy
    w, h = img.size
    img = img.resize((w * 2, h * 2), Image.LANCZOS)
    img = img.convert("L")  # grayscale

    text = pytesseract.image_to_string(img, config="--psm 6")
    logger.info(f"OCR raw text:\n{text[:500]}")

    nodes = []
    lines = text.splitlines()

    # Pattern: 4-char hex Node ID followed by Master XX on same or next line
    node_pattern = re.compile(r'\b([0-9A-Fa-f]{4})\b')
    master_pattern = re.compile(r'[Mm]aster\s*(\d+)')

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        node_match = node_pattern.search(line)
        if node_match:
            node_id = node_match.group(1).upper()
            # Look for master on same line or next 2 lines
            master = None
            search_text = " ".join(lines[i:i+3])
            m = master_pattern.search(search_text)
            if m:
                master = f"Master {m.group(1)}"
            if master:
                nodes.append((node_id, master))
        i += 1

    # Deduplicate while preserving order
    seen = set()
    unique = []
    for n in nodes:
        if n[0] not in seen:
            seen.add(n[0])
            unique.append(n)

    return unique


# ── Bot handlers ─────────────────────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = (
        "Tracker Node Bot\n\n"
        "Send me screenshots from GeniusVision and I will extract Node IDs and Masters into Excel.\n\n"
        "Commands:\n"
        "/export — Download the current Excel file\n"
        "/count  — See total nodes logged\n"
        "/clear  — Delete all data and start fresh\n"
        "/help   — Show this message"
    )
    await update.message.reply_text(msg)


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Processing image, please wait...")
    try:
        photo = update.message.photo[-1]  # highest resolution
        file = await context.bot.get_file(photo.file_id)
        image_bytes = await file.download_as_bytearray()

        nodes = parse_genius_vision_image(bytes(image_bytes))

        if not nodes:
            await update.message.reply_text(
                "Could not extract any Node IDs from this image.\n"
                "Make sure the screenshot clearly shows the Node ID and Master columns."
            )
            return

        image_name = f"photo_{datetime.now().strftime('%d%m%Y_%H%M%S')}"
        added, skipped = append_nodes(nodes, image_name)

        reply = (
            f"Extracted {len(nodes)} nodes from image.\n"
            f"New entries added: {added}\n"
            f"Duplicates skipped: {skipped}\n\n"
            f"Use /export to download the Excel file."
        )
        await update.message.reply_text(reply)

    except Exception as e:
        logger.error(f"Error processing image: {e}")
        await update.message.reply_text(f"Error processing image: {str(e)}")


async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("No data yet. Send some screenshots first.")
        return
    wb = get_or_create_workbook()
    ws = wb["Tracker Log"]
    count = ws.max_row - 1
    with open(EXCEL_FILE, "rb") as f:
        filename = f"Tracker_Log_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
        await update.message.reply_document(
            document=f,
            filename=filename,
            caption=f"Current log — {count} nodes total."
        )


async def count(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("No data yet.")
        return
    wb = get_or_create_workbook()
    ws = wb["Tracker Log"]
    total = ws.max_row - 1
    # Count per master
    masters = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[2]:
            masters[row[2]] = masters.get(row[2], 0) + 1
    lines = [f"Total nodes: {total}\n"]
    for m in sorted(masters.keys(), key=lambda x: int(x.split()[-1]) if x.split()[-1].isdigit() else 0):
        lines.append(f"  {m}: {masters[m]}")
    await update.message.reply_text("\n".join(lines))


async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
    get_or_create_workbook()
    await update.message.reply_text("All data cleared. Ready for fresh data.")


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await start(update, context)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if not TOKEN:
        raise ValueError("BOT_TOKEN environment variable not set.")
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("export", export))
    app.add_handler(CommandHandler("count", count))
    app.add_handler(CommandHandler("clear", clear))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    logger.info("Bot started.")
    app.run_polling()


if __name__ == "__main__":
    main()
