
import os
import logging
import tempfile
import requests
import hmac
import textwrap

from flask import Flask, request
from deep_translator import GoogleTranslator
from docx import Document
from openpyxl import load_workbook
from pypdf import PdfReader

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from xml.sax.saxutils import escape


# =========================
# ASOSIY SOZLAMALAR
# =========================

TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET")
RENDER_URL = os.getenv("RENDER_EXTERNAL_URL", "").rstrip("/")

API = f"https://api.telegram.org/bot{TOKEN}/"

app = Flask(__name__)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

MAX_FILE_SIZE = 19 * 1024 * 1024


# =========================
# TELEGRAM XABAR YUBORISH
# =========================

def send_message(chat_id, text):
    response = requests.post(
        API + "sendMessage",
        json={
            "chat_id": chat_id,
            "text": text
        },
        timeout=30
    )
    response.raise_for_status()


def send_document(chat_id, path, filename):
    with open(path, "rb") as file:
        response = requests.post(
            API + "sendDocument",
            data={"chat_id": chat_id},
            files={
                "document": (filename, file)
            },
            timeout=180
        )

    response.raise_for_status()


# =========================
# MATNNI TARJIMA QILISH
# =========================

def translate(text):
    if not isinstance(text, str) or not text.strip():
        return text

    translator = GoogleTranslator(
        source="auto",
        target="uz"
    )

    result = []
    chunk = ""

    for paragraph in text.splitlines(keepends=True):
        if len(chunk) + len(paragraph) > 4000:
            if chunk.strip():
                result.append(translator.translate(chunk))
            chunk = ""

        if len(paragraph) > 4000:
            if chunk.strip():
                result.append(translator.translate(chunk))
                chunk = ""

            for part in textwrap.wrap(
                paragraph,
                width=3500,
                break_long_words=True,
                break_on_hyphens=False
            ):
                result.append(translator.translate(part))
        else:
            chunk += paragraph

    if chunk.strip():
        result.append(translator.translate(chunk))

    return "\n".join(result)


# =========================
# TELEGRAMDAN FAYL OLISH
# =========================

def download_file(file_id, destination):
    info = requests.get(
        API + "getFile",
        params={"file_id": file_id},
        timeout=30
    )

    info.raise_for_status()
    data = info.json()

    if not data.get("ok"):
        raise Exception("Telegram faylni topmadi")

    file_path = data["result"]["file_path"]

    url = (
        f"https://api.telegram.org/file/bot"
        f"{TOKEN}/{file_path}"
    )

    response = requests.get(url, timeout=120)
    response.raise_for_status()

    with open(destination, "wb") as file:
        file.write(response.content)


# =========================
# PDF YARATISH
# =========================

def create_pdf(text, destination):
    styles = getSampleStyleSheet()

    font_name = "Helvetica"

    font_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf"
    ]

    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(
                    TTFont("DejaVu", font_path)
                )
                font_name = "DejaVu"
                break
            except Exception:
                pass

    styles["Normal"].fontName = font_name
    styles["Normal"].fontSize = 9
    styles["Normal"].leading = 13

    document = SimpleDocTemplate(
        destination,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    elements = []

    for line in text.splitlines():
        if not line.strip():
            elements.append(Spacer(1, 6))
            continue

        safe_line = escape(line)

        if font_name == "Helvetica":
            safe_line = (
                safe_line.replace("ʻ", "'")
                .replace("’", "'")
                .replace("‘", "'")
            )

        elements.append(
            Paragraph(safe_line, styles["Normal"])
        )
        elements.append(Spacer(1, 3))

    document.build(elements)


# =========================
# PDF TARJIMASI
# =========================

def translate_pdf(source, destination):
    reader = PdfReader(source)

    translated_pages = []

    for page in reader.pages:
        page_text = page.extract_text() or ""

        if page_text.strip():
            translated_pages.append(
                translate(page_text)
            )
        else:
            translated_pages.append(
                "[Bu sahifada o'qiladigan matn topilmadi]"
            )

    create_pdf(
        "\n\n".join(translated_pages),
        destination
    )


# =========================
# WORD TARJIMASI
# =========================

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


# =========================
# EXCEL TARJIMASI
# =========================

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


# =========================
# FAYLNI QAYTA ISHLASH
# =========================

def process_file(source, destination, extension):

    if extension == ".pdf":
        translate_pdf(source, destination)

    elif extension == ".docx":
        translate_word(source, destination)

    elif extension == ".xlsx":
        translate_excel(source, destination)

    else:
        raise Exception("Noto'g'ri fayl formati")


# =========================
# TELEGRAM XABARLARINI
# QABUL QILISH
# =========================

def handle_update(update):

    message = update.get("message", {})

    chat = message.get("chat", {})
    chat_id = chat.get("id")

    if not chat_id:
        return

    text = message.get("text", "")

    if text == "/start":
        send_message(
            chat_id,
            "Assalomu alaykum!\n\n"
            "Men fayllarni o'zbek tiliga tarjima qiluvchi botman.\n\n"
            "Menga quyidagi fayllarni yuboring:\n"
            "PDF\n"
            "Word (DOCX)\n"
            "Excel (XLSX)\n\n"
            "Men fayl tilini avtomatik aniqlab, "
            "o'zbek tiliga tarjima qilaman."
        )
        return

    document = message.get("document")

    if not document:
        send_message(
            chat_id,
            "Iltimos, PDF, Word yoki Excel fayl yuboring."
        )
        return

    filename = document.get("file_name", "file")
    extension = os.path.splitext(filename)[1].lower()

    if extension not in [".pdf", ".docx", ".xlsx"]:
        send_message(
            chat_id,
            "Bu format qo'llab-quvvatlanmaydi.\n"
            "PDF, DOCX yoki XLSX yuboring."
        )
        return

    file_size = document.get("file_size", 0)

    if file_size > MAX_FILE_SIZE:
        send_message(
            chat_id,
            "Fayl hajmi 19 MB dan katta.\n"
            "Iltimos, kichikroq fayl yuboring."
        )
        return

    send_message(
        chat_id,
        "Faylingiz qabul qilindi.\n"
        "Tarjima boshlanmoqda, kuting..."
    )

    try:
        with tempfile.TemporaryDirectory() as folder:

            source = os.path.join(
                folder,
                "original" + extension
            )

            destination = os.path.join(
                folder,
                "translated" + extension
            )

            download_file(
                document["file_id"],
                source
            )

            process_file(
                source,
                destination,
                extension
            )

            output_name = (
                os.path.splitext(filename)[0]
                + "_uzbek"
                + extension
            )

            send_document(
                chat_id,
                destination,
                output_name
            )

            send_message(
                chat_id,
                "Tarjima muvaffaqiyatli tugadi!"
            )

    except Exception as error:
        logger.exception("Fayl tarjimasida xatolik")

        send_message(
            chat_id,
            "Faylni tarjima qilishda xatolik yuz berdi.\n"
            "Iltimos, boshqa fayl bilan urinib ko'ring."
        )


# =========================
# RENDER TEKSHIRUVI
# =========================

@app.route("/", methods=["GET"])
def home():
    return "Uzbek Translator Bot ishlayapti!", 200


# =========================
# TELEGRAM WEBHOOK
# =========================

@app.route("/webhook", methods=["POST"])
def webhook():

    received_secret = request.headers.get(
        "X-Telegram-Bot-Api-Secret-Token",
        ""
    )

    if not WEBHOOK_SECRET or not hmac.compare_digest(
        received_secret,
        WEBHOOK_SECRET
    ):
        return "Forbidden", 403

    update = request.get_json(silent=True) or {}

    try:
        handle_update(update)
    except Exception:
        logger.exception("Telegram xabarini qayta ishlashda xatolik")

    return "OK", 200


# =========================
# WEBHOOKNI TELEGRAMGA ULASH
# =========================

def register_webhook():

    if not TOKEN:
        logger.error("BOT_TOKEN kiritilmagan")
        return

    if not WEBHOOK_SECRET:
        logger.error("WEBHOOK_SECRET kiritilmagan")
        return

    if not RENDER_URL:
        logger.error("RENDER_EXTERNAL_URL topilmadi")
        return

    webhook_url = RENDER_URL + "/webhook"

    try:
        response = requests.post(
            API + "setWebhook",
            json={
                "url": webhook_url,
                "secret_token": WEBHOOK_SECRET,
                "allowed_updates": ["message"]
            },
            timeout=30
        )

        response.raise_for_status()

        result = response.json()

        if result.get("ok"):
            logger.info("Telegram webhook muvaffaqiyatli ulandi")
        else:
            logger.error("Webhook ulanmadi: %s", result)

    except Exception:
        logger.exception("Webhook ulashda xatolik")


# =========================
# DASturni ISHGA TUSHIRISH
# =========================

register_webhook()

if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))
