
import os
import io
import hmac
import logging
import tempfile
import textwrap

import requests
import fitz
import cv2
import numpy as np
import pytesseract

from PIL import Image, ImageDraw, ImageFont
from pytesseract import Output
from flask import Flask, request
from deep_translator import GoogleTranslator
from docx import Document
from openpyxl import load_workbook
from pypdf import PdfReader

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from xml.sax.saxutils import escape

TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET")
RENDER_URL = os.getenv("RENDER_EXTERNAL_URL", "").rstrip("/")
API = f"https://api.telegram.org/bot{TOKEN}/"

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

MAX_FILE_SIZE = 19 * 1024 * 1024
MAX_PAGES = 30


def send_message(chat_id, text):
    r = requests.post(
        API + "sendMessage",
        json={"chat_id": chat_id, "text": text},
        timeout=30
    )
    r.raise_for_status()


def send_document(chat_id, path, filename):
    with open(path, "rb") as f:
        r = requests.post(
            API + "sendDocument",
            data={"chat_id": chat_id},
            files={"document": (filename, f)},
            timeout=180
        )
    r.raise_for_status()


def translate(text):
    if not text or not text.strip():
        return text

    translator = GoogleTranslator(source="auto", target="uz")
    result = []

    for start in range(0, len(text), 3500):
        part = text[start:start + 3500]
        translated = translator.translate(part)
        result.append(translated or part)

    return "\n".join(result)


def download_file(file_id, destination):
    r = requests.get(
        API + "getFile",
        params={"file_id": file_id},
        timeout=30
    )
    r.raise_for_status()

    data = r.json()
    if not data.get("ok"):
        raise Exception("Telegram faylni topmadi")

    file_path = data["result"]["file_path"]
    url = f"https://api.telegram.org/file/bot{TOKEN}/{file_path}"

    r = requests.get(url, timeout=120)
    r.raise_for_status()

    with open(destination, "wb") as f:
        f.write(r.content)


def create_pdf(text, destination):
    styles = getSampleStyleSheet()
    font_name = "Helvetica"

    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont("DejaVu", font_path))
            font_name = "DejaVu"
        except Exception:
            pass

    styles["Normal"].fontName = font_name
    styles["Normal"].fontSize = 9

    doc = SimpleDocTemplate(destination, pagesize=A4)
    elements = []

    for line in text.splitlines():
        if not line.strip():
            elements.append(Spacer(1, 6))
        else:
            elements.append(
                Paragraph(escape(line), styles["Normal"])
            )
            elements.append(Spacer(1, 3))

    doc.build(elements)


def translate_text_pdf(source, destination):
    reader = PdfReader(source)
    pages = []

    for page in reader.pages:
        page_text = page.extract_text() or ""

        if page_text.strip():
            pages.append(translate(page_text))
        else:
            pages.append("[Matn topilmadi]")

    create_pdf("\n\n".join(pages), destination)


def load_font(size):
    paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf"
    ]

    for path in paths:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)

    return ImageFont.load_default()


def wrap_text(draw, text, font, max_width):
    words = text.split()
    lines = []
    current = ""

    for word in words:
        test = (current + " " + word).strip()
        box = draw.textbbox((0, 0), test, font=font)

        if box[2] - box[0] <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines


def draw_translation(image, translated, rect):
    draw = ImageDraw.Draw(image)

    x1, y1, x2, y2 = rect
    width = max(1, x2 - x1)
    height = max(1, y2 - y1)

    for size in range(25, 9, -1):
        font = load_font(size)
        lines = wrap_text(draw, translated, font, width - 8)

        spacing = max(2, size // 5)
        bbox = draw.multiline_textbbox(
            (0, 0),
            "\n".join(lines),
            font=font,
            spacing=spacing,
            align="center"
        )

        text_height = bbox[3] - bbox[1]

        if text_height <= height - 6:
            break

    draw.multiline_text(
        (x1 + width // 2, y1 + height // 2),
        "\n".join(lines),
        font=font,
        fill=(0, 0, 0),
        anchor="mm",
        align="center",
        spacing=spacing
    )


def translate_comic_page(image):
    image = image.convert("RGB")
    rgb = np.array(image)
    data = pytesseract.image_to_data(
        rgb,
        lang="eng",
        config="--psm 11",
        output_type=Output.DICT
    )

    groups = {}

    for i, word in enumerate(data["text"]):
        word = word.strip()

        try:
            confidence = float(data["conf"][i])
        except Exception:
            confidence = -1

        if not word or confidence < 35:
            continue

        key = (
            data["block_num"][i],
            data["par_num"][i]
        )

        groups.setdefault(key, []).append({
            "text": word,
            "x": data["left"][i],
            "y": data["top"][i],
            "w": data["width"][i],
            "h": data["height"][i]
        })

    if not groups:
        return image

    mask = np.zeros(rgb.shape[:2], dtype=np.uint8)
    blocks = []

    for words in groups.values():
        text = " ".join(w["text"] for w in words)

        x1 = min(w["x"] for w in words)
        y1 = min(w["y"] for w in words)
        x2 = max(w["x"] + w["w"] for w in words)
        y2 = max(w["y"] + w["h"] for w in words)

        padding = 5
        x1 = max(0, x1 - padding)
        y1 = max(0, y1 - padding)
        x2 = min(rgb.shape[1], x2 + padding)
        y2 = min(rgb.shape[0], y2 + padding)

        mask[y1:y2, x1:x2] = 255

        blocks.append((text, (x1, y1, x2, y2)))

    cleaned = cv2.inpaint(
        rgb,
        mask,
        5,
        cv2.INPAINT_TELEA
    )

    result = Image.fromarray(cleaned)

    for text, rect in blocks:
        try:
            translated = translate(text)
            draw_translation(result, translated, rect)
        except Exception:
            logger.exception("Matn tarjimasida xatolik")

    return result


def translate_comic_pdf(source, destination):
    pdf = fitz.open(source)

    if len(pdf) > MAX_PAGES:
        raise Exception("PDF sahifalari 30 tadan ko'p")

    output = fitz.open()

    for number, page in enumerate(pdf):
        logger.info("Komiks sahifasi: %s", number + 1)

        pix = page.get_pixmap(
            matrix=fitz.Matrix(2, 2),
            alpha=False
        )

        image = Image.frombytes(
            "RGB",
            [pix.width, pix.height],
            pix.samples
        )

        translated_image = translate_comic_page(image)

        buffer = io.BytesIO()
        translated_image.save(buffer, format="JPEG", quality=90)

        new_page = output.new_page(
            width=page.rect.width,
            height=page.rect.height
        )

        new_page.insert_image(
            new_page.rect,
            stream=buffer.getvalue()
        )

    output.save(destination)
    output.close()
    pdf.close()


def process_pdf(source, destination):
    reader = PdfReader(source)

    text_count = sum(
        len(page.extract_text() or "")
        for page in reader.pages
    )

    if text_count > 100:
        translate_text_pdf(source, destination)
    else:
        translate_comic_pdf(source, destination)


def translate_word(source, destination):
    doc = Document(source)

    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            paragraph.text = translate(paragraph.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    cell.text = translate(cell.text)

    doc.save(destination)


def translate_excel(source, destination):
    workbook = load_workbook(source)

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if (
                    isinstance(cell.value, str)
                    and cell.value.strip()
                    and not cell.value.startswith("=")
                ):
                    cell.value = translate(cell.value)

    workbook.save(destination)


def process_file(source, destination, extension):
    if extension == ".pdf":
        process_pdf(source, destination)
    elif extension == ".docx":
        translate_word(source, destination)
    elif extension == ".xlsx":
        translate_excel(source, destination)
    else:
        raise Exception("Noto'g'ri format")


@app.route("/", methods=["GET"])
def home():
    return "Uzbek Translator Bot ishlayapti!", 200


@app.route("/webhook", methods=["POST"])
def webhook():
    received = request.headers.get(
        "X-Telegram-Bot-Api-Secret-Token", ""
    )

    if not WEBHOOK_SECRET or not hmac.compare_digest(
        received, WEBHOOK_SECRET
    ):
        return "Forbidden", 403

    update = request.get_json(silent=True) or {}

    try:
        handle_update(update)
    except Exception:
        logger.exception("Xabarni qayta ishlashda xatolik")

    return "OK", 200


def handle_update(update):
    message = update.get("message", {})
    chat_id = message.get("chat", {}).get("id")

    if not chat_id:
        return

    if message.get("text") == "/start":
        send_message(
            chat_id,
            "Assalomu alaykum!\n\n"
            "Men PDF komikslarni ham o'zbekchaga tarjima qilaman.\n\n"
            "PDF, DOCX yoki XLSX fayl yuboring."
        )
        return

    document = message.get("document")

    if not document:
        send_message(chat_id, "Iltimos, fayl yuboring.")
        return

    filename = document.get("file_name", "file")
    extension = os.path.splitext(filename)[1].lower()

    if extension not in [".pdf", ".docx", ".xlsx"]:
        send_message(chat_id, "PDF, DOCX yoki XLSX yuboring.")
        return

    if document.get("file_size", 0) > MAX_FILE_SIZE:
        send_message(chat_id, "Fayl 19 MB dan katta.")
        return

    send_message(
        chat_id,
        "Faylingiz qabul qilindi. Tarjima boshlanmoqda. Kuting..."
    )

    try:
        with tempfile.TemporaryDirectory() as folder:
            source = os.path.join(folder, "original" + extension)
            destination = os.path.join(folder, "translated" + extension)

            download_file(document["file_id"], source)
            process_file(source, destination, extension)

            output_name = (
                os.path.splitext(filename)[0]
                + "_uzbek"
                + extension
            )

            send_document(chat_id, destination, output_name)
            send_message(chat_id, "Tarjima tugadi!")

    except Exception:
        logger.exception("Fayl tarjimasida xatolik")
        try:
            send_message(
                chat_id,
                "Tarjima vaqtida xatolik yuz berdi. "
                "Kichikroq PDF bilan urinib ko'ring."
            )
        except Exception:
            logger.exception("Xatolik xabarini yuborib bo'lmadi")


def register_webhook():
    if not TOKEN or not WEBHOOK_SECRET or not RENDER_URL:
        logger.error("BOT_TOKEN, WEBHOOK_SECRET yoki RENDER URL yo'q")
        return

    try:
        response = requests.post(
            API + "setWebhook",
            json={
                "url": RENDER_URL + "/webhook",
                "secret_token": WEBHOOK_SECRET,
                "allowed_updates": ["message"]
            },
            timeout=30
        )

        response.raise_for_status()

        if response.json().get("ok"):
            logger.info("Telegram webhook muvaffaqiyatli ulandi")
        else:
            logger.error("Webhook ulanmagan: %s", response.text)

    except Exception:
        logger.exception("Webhook ulashda xatolik")


register_webhook()
        
