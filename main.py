
import os
import requests
import tempfile
from flask import Flask, request
from deep_translator import GoogleTranslator
from docx import Document
from openpyxl import load_workbook
from pypdf import PdfReader
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = Flask(__name__)

TOKEN = os.getenv("BOT_TOKEN")
API = f"https://api.telegram.org/bot{TOKEN}/"

def send_message(chat_id, text):
    requests.post(API + "sendMessage", json={
        "chat_id": chat_id,
        "text": text
    }, timeout=30)

def translate(text):
    if not text or not text.strip():
        return text
    try:
        return GoogleTranslator(
            source="auto", target="uz"
        ).translate(text[:4500])
    except Exception:
        return text

def download_file(file_id, path):
    info = requests.get(
        API + "getFile",
        params={"file_id": file_id},
        timeout=30
    ).json()

    if not info.get("ok"):
        raise Exception("Faylni olishda xatolik")

    file_path = info["result"]["file_path"]
    url = f"https://api.telegram.org/file/bot{TOKEN}/{file_path}"

    data = requests.get(url, timeout=60)
    data.raise_for_status()

    with open(path, "wb") as f:
        f.write(data.content)

def make_pdf(text, path):
    c = canvas.Canvas(path)
    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont("DejaVu", font_path))
        c.setFont("DejaVu", 10)
    else:
        c.setFont("Helvetica", 10)
        text = text.replace("‘", "'").replace("’", "'")

    y = 800
    for line in text.splitlines():
        words = line.split()
        current = ""
        for word in words:
            test = (current + " " + word).strip()
            if len(test) > 90:
                c.drawString(40, y, current)
                y -= 15
                current = word
            else:
                current = test

            if y < 50:
                c.showPage()
                y = 800
                if os.path.exists(font_path):
                    c.setFont("DejaVu", 10)
                else:
                    c.setFont("Helvetica", 10)

        if current:
            c.drawString(40, y, current)
            y -= 15

        if y < 50:
            c.showPage()
            y = 800
            if os.path.exists(font_path):
                c.setFont("DejaVu", 10)
            else:
                c.setFont("Helvetica", 10)

    c.save()

def process_file(source, output, ext):
    if ext == ".pdf":
        reader = PdfReader(source)
        pages = []
        for page in reader.pages:
            pages.append(translate(page.extract_text() or ""))
        make_pdf("\n\n".join(pages), output)

    elif ext == ".docx":
        doc = Document(source)

        for p in doc.paragraphs:
            if p.text.strip():
                p.text = translate(p.text)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        cell.text = translate(cell.text)

        doc.save(output)

    elif ext == ".xlsx":
        wb = load_workbook(source)

        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        if cell.value.strip() and not cell.value.startswith("="):
                            cell.value = translate(cell.value)

        wb.save(output)

def handle_update(update):
    message = update.get("message", {})
    chat_id = message.get("chat", {}).get("id")

    if not chat_id:
        return

    if message.get("text") == "/start":
        send_message(
            chat_id,
            "Assalomu alaykum!\n\n"
            "Men fayl tarjimon botman.\n"
            "PDF, Word yoki Excel fayl yuboring.\n"
            "Men uni avtomatik ravishda o'zbekchaga tarjima qilaman."
        )
        return

    doc = message.get("document")
    if not doc:
        send_message(chat_id, "Iltimos, PDF, DOCX yoki XLSX fayl yuboring.")
        return

    name = doc.get("file_name", "")
    ext = os.path.splitext(name)[1].lower()

    if ext not in [".pdf", ".docx", ".xlsx"]:
        send_message(chat_id, "Bu format qo'llab-quvvatlanmaydi.")
        return

    if doc.get("file_size", 0) > 19 * 1024 * 1024:
        send_message(chat_id, "Fayl hajmi juda katta. 19 MB dan kichik fayl yuboring.")
        return

    send_message(chat_id, "Faylingiz tarjima qilinmoqda. Biroz kuting...")

    try:
        with tempfile.TemporaryDirectory() as folder:
            source = os.path.join(folder, "input" + ext)
            output = os.path.join(folder, "translated" + ext)

            download_file(doc["file_id"], source)
            process_file(source, output, ext)

            with open(output, "rb") as f:
                requests.post(
                    API + "sendDocument",
                    data={"chat_id": chat_id},
                    files={"document": (f"tarjima{ext}", f)},
                    timeout=120
                )

    except Exception as e:
        send_message(
            chat_id,
            "Tarjima qilishda xatolik yuz berdi. "
            "Boshqa fayl bilan urinib ko'ring."
        )

@app.route("/", methods=["GET"])
def home():
    return "Uzbek Translator Bot ishlayapti!"

@app.route("/webhook", methods=["POST"])
def webhook():
    if request.headers.get("X-Telegram-Bot-Api-Secret-Token") != os.getenv("WEBHOOK_SECRET"):
        return "Forbidden", 403

    update = request.get_json(silent=True) or {}
    try:
        handle_update(update)
    except Exception:
        pass
    return "OK", 200


if TOKEN:
    webhook_url = os.getenv("RENDER_EXTERNAL_URL", "").rstrip("/") + "/webhook"

    requests.post(
        API + "setWebhook",
        json={
            "url": webhook_url,
            "secret_token": os.getenv("WEBHOOK_SECRET")
        },
        timeout=30
    )
    if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
  
