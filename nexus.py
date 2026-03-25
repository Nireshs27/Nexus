# MiniPrintAgent — silent local printing bridge for browser apps (Windows)
# v1.1  — adds global CORS preflight + hardened responses

import base64, json, os, sys, threading, subprocess
from pathlib import Path
from functools import wraps
from datetime import datetime

from flask import Flask, request, jsonify, make_response
from waitress import serve

import io
import re
import subprocess
from pathlib import Path
from datetime import datetime

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black

from reportlab.lib.utils import ImageReader
from urllib.request import urlopen

# --- Windows printing
import win32print

# escpos (optional, not required for Win32 RAW)
try:
    from escpos.printer import Win32Raw  # noqa
    HAVE_ESCPOS = True
except Exception:
    HAVE_ESCPOS = False

# Tray (optional)
TRY_TRAY = True
try:
    import pystray
    from PIL import Image, ImageDraw
except Exception:
    TRY_TRAY = False

APP_NAME = "MiniPrintAgent"
DEFAULT_PORT = 9979
CONFIG_DIR = Path(os.getenv("APPDATA", ".")) / APP_NAME
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "config.json"
LOG_FILE = CONFIG_DIR / "agent.log"


# --------------------------- utils ---------------------------

def log(msg: str):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def load_or_create_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        import secrets
        cfg = {
            "port": DEFAULT_PORT,
            "bind": "127.0.0.1",        # keep local-only for security
            "api_key": secrets.token_urlsafe(24),
            # add your Replit origin here for strict CORS; or use "*" to allow all
            "allowed_origins": ["*"]
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    return cfg


CFG = load_or_create_config()
API_KEY = CFG.get("api_key", "")


# --------------------------- Flask ---------------------------

app = Flask(__name__)


def _origin_allowed(origin: str | None) -> bool:
    allowed = CFG.get("allowed_origins", ["*"])
    return "*" in allowed or (origin and origin in allowed)


def _cors_headers():
    origin = request.headers.get("Origin")
    headers = {
        "Access-Control-Allow-Origin": origin if _origin_allowed(origin) else "*",
        "Vary": "Origin",
        "Access-Control-Allow-Headers": "Content-Type, X-Print-Key",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    }
    return headers


@app.before_request
def handle_preflight():
    # Global preflight: answer OPTIONS for any route with CORS headers
    if request.method == "OPTIONS":
        resp = make_response("", 204)
        for k, v in _cors_headers().items():
            resp.headers[k] = v
        return resp


@app.after_request
def add_cors(resp):
    # Add CORS to all responses
    for k, v in _cors_headers().items():
        resp.headers.setdefault(k, v)
    return resp


# def require_key(func):
#     @wraps(func)
#     def wrapper(*args, **kwargs):
#         k = request.headers.get("X-Print-Key", "")
#         if k != API_KEY:
#             return jsonify({"ok": False, "error": "unauthorized"}), 401
#         return func(*args, **kwargs)
#     return wrapper

# --------------------------- endpoints ---------------------------


@app.route("/status", methods=["GET"])
def status():
    try:
        locals_only = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
        printers = [p[2] for p in locals_only]
    except Exception as e:
        printers = []
        log(f"status printer enum error: {e}")

    return jsonify({
        "ok": True,
        "name": APP_NAME,
        "version": "1.1",
        "escpos_available": HAVE_ESCPOS,
        "printers": printers,
    })


@app.route("/printers", methods=["GET"])
def printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    items = win32print.EnumPrinters(flags)

    names = []
    seen = set()
    for p in items:
        name = p[2]
        if name and name not in seen:
            seen.add(name)
            names.append(name)

    return jsonify({"ok": True, "printers": names})


@app.route("/print-raw", methods=["POST"])
# @require_key
def print_raw():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    payload_b64 = data.get("dataB64")
    payload_hex = data.get("dataHex")
    payload_text = data.get("text")  # gets encoded to CP437 with CRLF

    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        if payload_b64:
            raw = base64.b64decode(payload_b64)
        elif payload_hex:
            raw = bytes.fromhex(payload_hex)
        elif payload_text is not None:
            raw = payload_text.replace("\n", "\r\n").encode("cp437", errors="replace")
        else:
            return jsonify({"ok": False, "error": "provide dataB64 or dataHex or text"}), 400

        _write_raw(printer, raw)
        return jsonify({"ok": True})
    except Exception as e:
        log(f"/print-raw error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/print-text", methods=["POST"])
# @require_key
def print_text():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    text = data.get("text", "")
    bold = bool(data.get("bold", False))
    cut = bool(data.get("cut", False))
    feed_lines = int(data.get("feedLines", 6))
    cut_mode = str(data.get("cutMode", "full")).lower()


    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    ESC = b"\x1b"; GS = b"\x1d"
    parts = [ESC + b"@"]
    if bold:
        parts.append(ESC + b"E" + b"\x01")
    parts.append(text.replace("\n", "\r\n").encode("cp437", errors="replace"))
    if bold:
        parts.append(ESC + b"E" + b"\x00")
    if cut:
        parts.append(b"\r\n" * feed_lines)
        if cut_mode == "partial":
            parts.append(GS + b"V" + b"\x01")
        else:
            parts.append(GS + b"V" + b"\x00")



    try:
        _write_raw(printer, b"".join(parts))
        return jsonify({"ok": True})
    except Exception as e:
        log(f"/print-text error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


import binascii, uuid

@app.route("/print-pdf", methods=["POST"])
# @require_key
def print_pdf():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer") or ""
    pdf_b64 = data.get("pdfB64") or ""
    sumatra = data.get("sumatraPath") or ""

    if not printer or not pdf_b64:
        return jsonify({"ok": False, "error": "need 'printer' and 'pdfB64'"}), 400
    if not sumatra or not Path(sumatra).exists():
        return jsonify({"ok": False, "error": "SumatraPDF not configured"}), 400

    # 1) Strip data URL prefix if present
    if isinstance(pdf_b64, str) and pdf_b64.startswith("data:"):
        # data:application/pdf;base64,JVBERi0x...
        if "," in pdf_b64:
            pdf_b64 = pdf_b64.split(",", 1)[1]

    # 2) Strict base64 decode + validate it's a PDF
    try:
        raw = base64.b64decode(pdf_b64, validate=True)
    except (binascii.Error, ValueError) as e:
        return jsonify({"ok": False, "error": f"Invalid base64: {e}"}), 400

    if not raw.startswith(b"%PDF-"):
        # helpful debug: show first few bytes as hex
        head = raw[:16].hex()
        return jsonify({"ok": False, "error": f"Invalid PDF bytes (missing %PDF-). First16={head}"}), 400

    # 3) Unique filename to avoid overlapping jobs
    pdf_path = CONFIG_DIR / f"job-{uuid.uuid4().hex}.pdf"
    pdf_path.write_bytes(raw)

    cmd = [sumatra, "-silent", "-print-to", printer, "-exit-on-print", str(pdf_path)]
    try:
        sp = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if sp.returncode != 0:
            raise RuntimeError(sp.stderr or sp.stdout or f"Sumatra exit {sp.returncode}")
        return jsonify({"ok": True})
    except Exception as e:
        log(f"/print-pdf error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        try:
            pdf_path.unlink(missing_ok=True)
        except Exception:
            pass

# --------------------------- slip helpers ---------------------------





# --------------------------- server /// tray ---------------------------

def server_thread():
    bind = CFG.get("bind", "127.0.0.1")
    port = int(CFG.get("port", DEFAULT_PORT))
    log(f"{APP_NAME} listening on http://{bind}:{port}")
    serve(app, host=bind, port=port, threads=6)


def make_icon_img():
    img = Image.new("RGB", (64, 64), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([8, 18, 56, 42], outline="black", width=2)
    d.rectangle([16, 8, 48, 18], fill="black")
    d.rectangle([16, 42, 48, 54], fill="black")
    d.ellipse([20, 24, 28, 32], fill="black")
    d.ellipse([30, 24, 38, 32], fill="black")
    d.ellipse([40, 24, 48, 32], fill="black")
    return img


def run_tray():
    if not TRY_TRAY:
        log("Tray disabled (missing pystray/Pillow).")
        return
    icon = pystray.Icon(APP_NAME)
    icon.title = APP_NAME
    icon.icon = make_icon_img()

    def open_log():
        os.startfile(LOG_FILE)

    def open_config():
        os.startfile(CONFIG_FILE)

    def quit_app():
        icon.stop()
        os._exit(0)

    icon.menu = pystray.Menu(
        pystray.MenuItem("Open log", lambda: open_log()),
        pystray.MenuItem("Open config", lambda: open_config()),
        pystray.MenuItem("Quit", lambda: quit_app()),
    )
    icon.run()


def ensure_autostart_shortcut():
    """Create a Startup shortcut for current user."""
    try:
        import win32com.client
        startup = Path(os.getenv("APPDATA")) / r"Microsoft\Windows\Start Menu\Programs\Startup"
        target = sys.executable
        script = Path(__file__).resolve()
        args = f'"{script}"'
        lnk = startup / f"{APP_NAME}.lnk"
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(str(lnk))
        shortcut.Targetpath = target
        shortcut.Arguments = args
        shortcut.WorkingDirectory = str(script.parent)
        shortcut.IconLocation = target
        shortcut.save()
        log(f"Autostart shortcut created at: {lnk}")
    except Exception as e:
        log(f"Autostart shortcut not created: {e}")

# -------------------------------------------- Thermal printer End --------------------------------------------










"""
Gold / Metal Slips (80mm Epson) — Raster printing (ESC/POS)

Includes:
1) Receiving Metal Slip
2) Provisional Delivery Slip
3) Packing & Delivery Slip
4) Packing List (with CODE128 barcode for Metal Account Code)

REQUIRES (for CODE128 image generation):
    pip install python-barcode pillow

Works on Windows RAW printing via win32print:
    pip install pywin32
"""

from pathlib import Path
from decimal import Decimal
from functools import lru_cache

from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageEnhance, ImageChops
from flask import Flask, request, jsonify

# --- Windows RAW print ---
try:
    import win32print
except Exception:
    win32print = None

# --- Code128 (python-barcode) ---
try:
    from barcode import Code128
    from barcode.writer import ImageWriter
except Exception:
    Code128 = None
    ImageWriter = None

# =========================
# GLOBAL CONFIG
# =========================

MAX_WIDTH_DOTS_80MM = 576

FONT_REGULAR_PATH = r"C:\Windows\Fonts\arial.ttf"
FONT_BOLD_PATH    = r"C:\Windows\Fonts\arialbd.ttf"

FALLBACK_REGULAR = r"C:\Windows\Fonts\arial.ttf"
FALLBACK_BOLD    = r"C:\Windows\Fonts\arialbd.ttf"


def log(msg: str):
    print(msg, flush=True)


def _pick_font(path: str, fallback: str) -> str:
    return path if Path(path).exists() else fallback


FONT_REGULAR_PATH = _pick_font(FONT_REGULAR_PATH, FALLBACK_REGULAR)
FONT_BOLD_PATH    = _pick_font(FONT_BOLD_PATH, FALLBACK_BOLD)

# Pre-load font objects once (fast + consistent) — these are the SOURCE OF TRUTH
F_HDR     = ImageFont.truetype(FONT_BOLD_PATH, 26)
F_TXT     = ImageFont.truetype(FONT_REGULAR_PATH, 24)
F_TXT_B   = ImageFont.truetype(FONT_BOLD_PATH, 24)
F_BIG     = ImageFont.truetype(FONT_BOLD_PATH, 50)
F_BIG_ROW = ImageFont.truetype(FONT_BOLD_PATH, 36)
F_UNIT    = ImageFont.truetype(FONT_REGULAR_PATH, 20)
# Job create slip — weight numerals (~2× F_TXT / F_TXT_B)
F_JC_WEIGHT   = ImageFont.truetype(FONT_REGULAR_PATH, 48)
F_JC_WEIGHT_B = ImageFont.truetype(FONT_BOLD_PATH, 48)

# =========================
# ESC/POS + RAW PRINT HELPERS
# =========================

def _esc_feed(n: int) -> bytes:
    return b"\r\n" * max(0, int(n))


def _esc_cut(mode: str = "full") -> bytes:
    GS = b"\x1d"
    m = (mode or "full").lower()
    return GS + b"V" + (b"\x01" if m in ("partial", "part", "p") else b"\x00")


def _img_to_escpos_raster(img_1bit: Image.Image) -> bytes:
    if img_1bit.mode != "1":
        img_1bit = img_1bit.convert("1")

    w, h = img_1bit.size
    wb = (w + 7) // 8

    xL = wb & 0xFF
    xH = (wb >> 8) & 0xFF
    yL = h & 0xFF
    yH = (h >> 8) & 0xFF

    px = img_1bit.load()
    data = bytearray()

    for yy in range(h):
        for xb in range(wb):
            b = 0
            for bit in range(8):
                xx = xb * 8 + bit
                if xx < w and px[xx, yy] == 0:
                    b |= (1 << (7 - bit))
            data.append(b)

    GS = b"\x1d"
    return GS + b"v0" + b"\x00" + bytes([xL, xH, yL, yH]) + bytes(data)


def _to_1bit(img: Image.Image, threshold: int = 160) -> Image.Image:
    gray = img.convert("L")
    return gray.point(lambda p: 0 if p < threshold else 255, mode="1")


def _to_1bit_floyd_steinberg(img: Image.Image) -> Image.Image:
    """
    Error-diffusion dither for 1-bit thermal output. Preserves facial mid-tones
    (halftone dots) instead of crushing them with a flat threshold.
    """
    gray = img.convert("L")
    try:
        dith = Image.Dither.FLOYDSTEINBERG
    except AttributeError:
        dith = Image.FLOYDSTEINBERG  # type: ignore[attr-defined]
    return gray.convert("1", dither=dith)


def _write_raw(printer_name: str, payload: bytes):
    """
    Robust Windows RAW printing.
    If you already have your own _write_raw, you can keep yours and remove this.
    """
    if win32print is None:
        raise RuntimeError("pywin32 is not installed. Install: pip install pywin32")

    h = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(h, 1, ("SlipPrint", None, "RAW"))
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, payload)
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)


# =========================
# DRAW HELPERS
# =========================

def _safe_str(v) -> str:
    return str(v or "").strip()


def _text_w(draw: ImageDraw.ImageDraw, s: str, font) -> float:
    return draw.textlength(s or "", font=font)


def _draw_text(draw, x, y, s, font):
    draw.text((x, y), s or "", font=font, fill=0)


def _draw_text_right(draw, x_right, y, s, font):
    w = _text_w(draw, s, font)
    draw.text((x_right - w, y), s or "", font=font, fill=0)


def _draw_text_center(draw, x_center, y, s, font):
    w = _text_w(draw, s, font)
    draw.text((x_center - w / 2, y), s or "", font=font, fill=0)


def _wrap_line_px(text: str, draw: ImageDraw.ImageDraw, font, max_w: int):
    words = (text or "").split()
    if not words:
        return [""]

    lines, cur = [], ""
    for w in words:
        test = w if not cur else f"{cur} {w}"
        if _text_w(draw, test, font) <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            # hard split long word
            if _text_w(draw, w, font) <= max_w:
                cur = w
            else:
                chunk = ""
                for ch in w:
                    t2 = chunk + ch
                    if _text_w(draw, t2, font) <= max_w:
                        chunk = t2
                    else:
                        if chunk:
                            lines.append(chunk)
                        chunk = ch
                cur = chunk
    if cur:
        lines.append(cur)
    return lines


def _dash_line(draw, x0, x1, y, dash=6, gap=4, width=1):
    x = x0
    while x < x1:
        draw.line((x, y, min(x + dash, x1), y), fill=0, width=width)
        x += dash + gap

def _font_line_h(font) -> int:
    a, d = font.getmetrics()
    return int(a + d)

def _draw_multiline(draw, x, y, text, font, line_gap=2):
    """
    Draw multiline text and return total height used.
    """
    lines = (text or "").split("\n")
    lh = _font_line_h(font)
    yy = y
    for i, ln in enumerate(lines):
        draw.text((x, yy), ln, font=font, fill=0)
        yy += lh + (line_gap if i < len(lines) - 1 else 0)
    return yy - y  # height used



# =========================
# GLOBAL CODE128 GENERATOR (PIL IMAGE)
# =========================

@lru_cache(maxsize=256)
def _code128_render_base(value: str, module_width: float, module_height: float, quiet_zone: float) -> Image.Image:
    """
    Render a grayscale base Code128 image (unscaled). Cached for speed.
    """
    if Code128 is None or ImageWriter is None:
        raise RuntimeError("python-barcode is not installed. Install: pip install python-barcode pillow")

    value = _safe_str(value) or "K0000"
    writer = ImageWriter()
    options = {
        "module_width": float(module_width),    # mm
        "module_height": float(module_height),  # mm
        "quiet_zone": float(quiet_zone),        # mm
        "write_text": False,                    # We print text ourselves using slip fonts
    }
    bc = Code128(value, writer=writer)
    return bc.render(writer_options=options).convert("L")


def code128_pil(value: str, target_w_px: int, target_h_px: int,
                *, module_width=0.40, module_height=18.0, quiet_zone=2.5, threshold=160) -> Image.Image:
    """
    Robust global CODE128 generator:
    - Generates barcode with python-barcode
    - Resizes to requested pixel size
    - Converts to 1-bit for thermal clarity
    """
    base = _code128_render_base(_safe_str(value), float(module_width), float(module_height), float(quiet_zone))
    img = base.resize((int(target_w_px), int(target_h_px)), Image.Resampling.LANCZOS)
    img = img.point(lambda p: 0 if p < threshold else 255, mode="1")
    return img

# =========================
# QUICK WEIGHT RECEIPT (80mm)
# Date & Time + centered Weight
# =========================

def _render_quick_weight_receipt_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))

    margin_x = 18
    top_pad = 20
    in_pad = 14

    box_x0 = margin_x
    box_x1 = W - margin_x
    box_w = box_x1 - box_x0

    date_time = _safe_str(data.get("dateTime") or data.get("date") or "")
    weight = _safe_str(data.get("weight") or data.get("value") or "")
    unit = _safe_str(data.get("unit") or "g")

    # Fonts
    f_dt = ImageFont.truetype(FONT_REGULAR_PATH, 24)
    f_lbl = ImageFont.truetype(FONT_REGULAR_PATH, 34)
    f_wt = ImageFont.truetype(FONT_BOLD_PATH, 72)
    f_unit = ImageFont.truetype(FONT_BOLD_PATH, 44)

    # Estimate height
    h_dt = _font_line_h(f_dt)
    h_lbl = _font_line_h(f_lbl)
    h_wt = _font_line_h(f_wt)
    h_unit = _font_line_h(f_unit)

    H = max(260, top_pad + h_dt + 22 + h_lbl + 16 + max(h_wt, h_unit) + 36)

    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)

    y = top_pad

    # Date/Time
    if date_time:
        _draw_text_center(draw, box_x0 + box_w / 2, y, date_time, f_dt)
        y += h_dt + 22

    # Weight label
    _draw_text_center(draw, box_x0 + box_w / 2, y, "Weight:", f_lbl)
    y += h_lbl + 16

    # Weight + unit centered as one group
    wt_text = weight or "0.00"
    gap = 10
    wt_w = _text_w(draw, wt_text, f_wt)
    unit_w = _text_w(draw, unit, f_unit)
    total_w = wt_w + gap + unit_w

    start_x = (W - total_w) / 2
    draw.text((start_x, y), wt_text, font=f_wt, fill=0)
    draw.text((start_x + wt_w + gap, y + max(0, (h_wt - h_unit) // 2 + 4)), unit, font=f_unit, fill=0)

    y += max(h_wt, h_unit) + 14

    img = img.crop((0, 0, W, min(H, y + 12)))
    return _to_1bit(img, threshold=160)


# =========================
# RECEIVING METAL SLIP
# =========================

def _render_receiving_metal_slip_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))

    margin_x = 18
    top_pad  = 14
    in_pad_x = 14

    box_w = W - (margin_x * 2)
    box_x0 = margin_x
    box_x1 = box_x0 + box_w

    y = top_pad
    H = 650
    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)

    box_y0 = y

    # Header
    y += 10
    _draw_text(draw, box_x0 + in_pad_x, y, _safe_str(data.get("accountCode")), F_HDR)
    _draw_text_right(draw, box_x1 - in_pad_x, y, _safe_str(data.get("date")), F_TXT)
    y += 36

    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 12

    cust = _safe_str(data.get("customerCode"))
    if cust:
        _draw_text(draw, box_x0 + in_pad_x, y, cust, F_TXT_B)
        y += 32

    rno = _safe_str(data.get("receiptNo"))
    if rno:
        _draw_text(draw, box_x0 + in_pad_x, y, f"Receipt No: {rno}", F_TXT)
        y += 30

    uw = _safe_str(data.get("unitWeight"))
    item = _safe_str(data.get("itemName"))
    combined = f"{uw}-{item}" if uw else item
    _draw_text(draw, box_x0 + in_pad_x, y, combined, F_TXT_B)
    y += 32

    y += 6
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 10

    gw = _safe_str(data.get("grossWt"))
    nw = _safe_str(data.get("netWt"))
    same = (gw == nw) and gw != ""

    label_w = int(box_w * 0.42)
    split_x = box_x0 + label_w

    if same:
        row_h = 120
        draw.line((split_x, y, split_x, y + row_h), fill=0, width=1)

        _draw_text_center(draw, box_x0 + label_w/2, y + 28, "Gross Wt", F_TXT_B)
        _draw_text_center(draw, box_x0 + label_w/2, y + 58, "/Net Wt", F_TXT_B)

        right_center = split_x + (box_x1 - split_x)/2
        _draw_text_center(draw, right_center, y + 22, gw, F_BIG)
        _draw_text_center(draw, right_center, y + 82, "gms", F_UNIT)

        y += row_h
        draw.line((box_x0, y, box_x1, y), fill=0, width=1)
        y += 10
    else:
        row_h = 60

        def draw_weight_row(label, val):
            nonlocal y
            _draw_text_center(draw, box_x0 + label_w/2, y + 16, label, F_TXT_B)
            right_center = split_x + (box_x1 - split_x)/2
            _draw_text_center(draw, right_center, y + 6, val, F_BIG_ROW)
            _draw_text_center(draw, right_center, y + 38, "gms", F_UNIT)
            draw.line((split_x, y, split_x, y + row_h), fill=0, width=1)
            y += row_h
            draw.line((box_x0, y, box_x1, y), fill=0, width=1)

        draw_weight_row("Gross Wt", gw)
        draw_weight_row("Net Wt", nw)
        y += 10

        try:
            eh = str((Decimal(gw) - Decimal(nw)).copy_abs())
        except Exception:
            eh = ""

        if eh:
            _draw_text(draw, box_x0 + in_pad_x, y + 6, "EH/Stone Wt", F_TXT)
            _draw_text_right(draw, box_x1 - in_pad_x, y + 6, f"{eh} gms", F_TXT)
            y += 34
            draw.line((box_x0, y, box_x1, y), fill=0, width=1)
            y += 10

    qty = _safe_str(data.get("qty"))
    _draw_text(draw, box_x0 + in_pad_x, y + 6, f"Qty: {qty}", F_TXT_B)
    y += 42

    box_y1 = y + 8
    draw.rectangle((box_x0, box_y0, box_x1, box_y1), outline=0, width=1)

    img = img.crop((0, 0, W, min(H, box_y1 + 20)))
    return _to_1bit(img, threshold=160)


# =========================
# PROVISIONAL DELIVERY SLIP
# =========================

def _render_provisional_delivery_slip_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))

    margin_x = 18
    top_pad  = 14
    in_pad_x = 14

    box_w = W - (margin_x * 2)
    box_x0 = margin_x
    box_x1 = box_x0 + box_w

    H = 700
    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)

    y = top_pad
    box_y0 = y

    y += 10
    _draw_text(draw, box_x0 + in_pad_x, y, _safe_str(data.get("accountCode")), F_HDR)
    _draw_text_right(draw, box_x1 - in_pad_x, y, _safe_str(data.get("date")), F_TXT)
    y += 36

    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 12

    cust = _safe_str(data.get("customerCode"))
    if cust:
        _draw_text(draw, box_x0 + in_pad_x, y, cust, F_TXT_B)
        y += 32

    pbc = _safe_str(data.get("partyBranchCode"))
    if pbc:
        _draw_text(draw, box_x0 + in_pad_x, y, pbc, F_TXT)
        y += 30

    uw = _safe_str(data.get("unitWeight"))
    item = _safe_str(data.get("itemName"))
    combined = f"{uw}-{item}" if uw else item

    usable_w = box_w - (in_pad_x * 2)
    lines = _wrap_line_px(combined, draw, F_TXT_B, usable_w)
    for ln in lines:
        _draw_text(draw, box_x0 + in_pad_x, y, ln, F_TXT_B)
        y += 30

    po = _safe_str(data.get("poNumber"))
    if po:
        _draw_text(draw, box_x0 + in_pad_x, y, po, F_TXT)
        y += 30

    y += 6
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 10

    gw = _safe_str(data.get("grossWt"))
    nw = _safe_str(data.get("netWt"))
    same = (gw == nw) and gw != ""

    label_w = int(box_w * 0.42)
    split_x = box_x0 + label_w

    if same:
        row_h = 120
        draw.line((split_x, y, split_x, y + row_h), fill=0, width=1)

        _draw_text_center(draw, box_x0 + label_w/2, y + 28, "Gross Wt", F_TXT_B)
        _draw_text_center(draw, box_x0 + label_w/2, y + 58, "/Net Wt", F_TXT_B)

        right_center = split_x + (box_x1 - split_x)/2
        _draw_text_center(draw, right_center, y + 22, gw, F_BIG)
        _draw_text_center(draw, right_center, y + 82, "gms", F_UNIT)

        y += row_h
        draw.line((box_x0, y, box_x1, y), fill=0, width=1)
        y += 10
    else:
        row_h = 60

        def draw_weight_row(label, val):
            nonlocal y
            _draw_text_center(draw, box_x0 + label_w/2, y + 16, label, F_TXT_B)
            right_center = split_x + (box_x1 - split_x)/2
            _draw_text_center(draw, right_center, y + 6, val, F_BIG_ROW)
            _draw_text_center(draw, right_center, y + 38, "gms", F_UNIT)
            draw.line((split_x, y, split_x, y + row_h), fill=0, width=1)
            y += row_h
            draw.line((box_x0, y, box_x1, y), fill=0, width=1)

        draw_weight_row("Net Wt", nw)
        draw_weight_row("Gross Wt", gw)
        y += 10

    qty = _safe_str(data.get("qty"))
    _draw_text(draw, box_x0 + in_pad_x, y + 6, f"Qty: {qty}", F_TXT_B)
    y += 42

    box_y1 = y + 8
    draw.rectangle((box_x0, box_y0, box_x1, box_y1), outline=0, width=1)

    img = img.crop((0, 0, W, min(H, box_y1 + 20)))
    return _to_1bit(img, threshold=160)


# =============================
# PACKING & DELIVERY PRINT SLIP
# =============================

def _draw_value_with_unit_centered(
    draw, x_left, x_right, y, val_text, val_font,
    unit_text="gms", unit_font=None, gap=8
):
    val_text = _safe_str(val_text)
    unit_text = _safe_str(unit_text)
    unit_font = unit_font or val_font

    w_val = draw.textlength(val_text, font=val_font)
    w_unit = draw.textlength(unit_text, font=unit_font)
    total = w_val + gap + w_unit

    cx = (x_left + x_right) / 2
    x = cx - (total / 2)

    draw.text((x, y), val_text, font=val_font, fill=0)
    draw.text((x + w_val + gap, y), unit_text, font=unit_font, fill=0)


def _render_packing_delivery_slip_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))

    margin_x = 10
    top_pad  = 8
    in_pad_x = 10

    box_w = W - (margin_x * 2)
    box_x0 = margin_x
    box_x1 = box_x0 + box_w

    gap_after_header   = 26
    gap_after_divider  = 10
    line_gap           = 24
    item_line_gap      = 24
    gap_before_weights = 6
    gap_after_weights  = 6

    same_row_h = 86
    diff_row_h = 54
    qty_row_h  = 28

    # -------------------------
    # Optional Delivery ID barcode
    # -------------------------
    delivery_id = _safe_str(
        data.get("deliveryId")
        or data.get("deliveryID")
        or data.get("deliveryNo")
        or data.get("deliveryNumber")
    )

    show_delivery_barcode = bool(delivery_id)

    # extra space only when barcode exists
    delivery_bar_h = 40
    delivery_bar_w = 260  # or 280 if space allows
    delivery_text_gap = 6
    delivery_block_gap = 20
    delivery_total_h = 0
    if show_delivery_barcode:
        delivery_total_h = delivery_bar_h + _font_line_h(F_UNIT) + delivery_text_gap + delivery_block_gap

    H = 520 + delivery_total_h + 20
    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)

    y = top_pad
    box_y0 = y

    # Outer top border
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 8

    # -------------------------
    # Top header line
    # -------------------------
    account_code = _safe_str(data.get("accountCode"))
    date_str = _safe_str(data.get("date"))

    if account_code:
        _draw_text(draw, box_x0 + in_pad_x, y, account_code, F_TXT)
    if date_str:
        # _draw_text_right(draw, box_x1 - in_pad_x, y, date_str, F_TXT)
        _draw_text_right(draw, box_x1 - in_pad_x - 80, y, date_str, F_TXT)

    y += gap_after_header

    # -------------------------
    # Delivery ID barcode block
    # -------------------------
    if show_delivery_barcode:
        barcode_left = box_x0 + in_pad_x + 6
        barcode_right = box_x1 - in_pad_x - 6
        avail_w = barcode_right - barcode_left

        delivery_bar_w = min(int(avail_w * 0.80), 300)
        paste_x = int(barcode_left + (avail_w - delivery_bar_w) / 2)

        delivery_barcode_img = code128_pil(
            delivery_id,
            delivery_bar_w,
            delivery_bar_h,
            module_width=0.40,
            module_height=14.0,
            quiet_zone=2.5,
            threshold=160,
        )

        if delivery_barcode_img.mode != "1":
            delivery_barcode_img = delivery_barcode_img.convert("1")

        img.paste(delivery_barcode_img, (paste_x, y))
        y += delivery_bar_h + delivery_text_gap

        _draw_text_center(draw, box_x0 + box_w / 2, y, f"ID: {delivery_id}", F_UNIT)
        y += _font_line_h(F_UNIT) + delivery_block_gap

    # Divider below header / delivery barcode block
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += gap_after_divider

    # -------------------------
    # Customer + branch/PO/order line
    # -------------------------
    customer = _safe_str(data.get("customerCode") or data.get("customer"))
    branch_code = _safe_str(data.get("partyBranchCode"))
    po_no = _safe_str(data.get("poNumber"))
    order_no = _safe_str(data.get("orderNo"))

    if customer:
        _draw_text(draw, box_x0 + in_pad_x, y, customer, F_TXT)

    right_text = ""
    if branch_code and po_no:
        right_text = f"{branch_code}: {po_no}"
    elif branch_code and order_no:
        right_text = f"{branch_code}: {order_no}"
    elif po_no:
        right_text = po_no
    elif order_no:
        right_text = order_no
    elif branch_code:
        right_text = branch_code

    if right_text:
        # _draw_text_right(draw, box_x1 - in_pad_x, y, right_text, F_TXT)
        _draw_text_right(draw, box_x1 - in_pad_x - 80, y, right_text, F_TXT)

    y += line_gap

    # -------------------------
    # Item / product line
    # -------------------------
    unitwt = _safe_str(data.get("unitWt") or data.get("unitWeight"))
    item_name = _safe_str(
        data.get("itemShortName")
        or data.get("productShortName")
        or data.get("itemName")
        or data.get("productName")
        or data.get("product")
    )
    item_text = f"{unitwt}-{item_name}".strip("-").strip()

    usable_w = box_w - (in_pad_x * 2)
    lines = _wrap_line_px(item_text, draw, F_TXT_B, usable_w)
    for ln in lines:
        _draw_text(draw, box_x0 + in_pad_x, y, ln, F_TXT_B)
        y += item_line_gap

    y += gap_before_weights
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 6

    # -------------------------
    # Weights
    # -------------------------
    gw = _safe_str(data.get("grossWt"))
    nw = _safe_str(data.get("netWt"))
    same = (gw == nw) and gw != ""

    label_w = int(box_w * 0.42)
    split_x = box_x0 + label_w

    val_font  = ImageFont.truetype(FONT_BOLD_PATH, 38)
    unit_font = F_TXT

    def _centered_y(row_top: int, row_h: int, val_font, unit_font) -> int:
        val_h = _font_line_h(val_font)
        unit_h = _font_line_h(unit_font)
        content_h = max(val_h, unit_h)
        return row_top + int((row_h - content_h) / 2) - 2

    if same:
        row_h = same_row_h
        draw.line((split_x, y, split_x, y + row_h), fill=0, width=1)

        _draw_text_center(draw, box_x0 + label_w / 2, y + 18, "Gross Wt", F_TXT_B)
        _draw_text_center(draw, box_x0 + label_w / 2, y + 42, "/Net Wt", F_TXT_B)

        val_y = _centered_y(y, row_h, val_font, unit_font)
        # _draw_value_with_unit_centered(draw, split_x, box_x1, val_y, gw, val_font, "gms", unit_font, gap=8)
        _draw_value_with_unit_centered(draw, split_x, box_x1 - 80, val_y, gw, val_font, "gms", unit_font, gap=8)

        y += row_h
        draw.line((box_x0, y, box_x1, y), fill=0, width=1)
        y += gap_after_weights
    else:
        row_h = diff_row_h

        def draw_weight_row(label: str, val: str):
            nonlocal y
            _draw_text_center(draw, box_x0 + label_w / 2, y + (row_h // 2) - 10, label, F_TXT_B)
            val_y = _centered_y(y, row_h, val_font, unit_font)
            # _draw_value_with_unit_centered(draw, split_x, box_x1, val_y, val, val_font, "gms", unit_font, gap=8)
            _draw_value_with_unit_centered(draw, split_x, box_x1 - 80, val_y, val, val_font, "gms", unit_font, gap=8)
            draw.line((split_x, y, split_x, y + row_h), fill=0, width=1)
            y += row_h
            draw.line((box_x0, y, box_x1, y), fill=0, width=1)

        draw_weight_row("Gross Wt", gw)
        draw_weight_row("Net Wt", nw)

        y += gap_after_weights

    # -------------------------
    # Qty
    # -------------------------
    qty = _safe_str(data.get("qty"))
    _draw_text(draw, box_x0 + in_pad_x, y + 6, f"Qty: {qty}", F_TXT_B)
    y += qty_row_h

    box_y1 = y
    draw.rectangle((box_x0, box_y0, box_x1, box_y1), outline=0, width=1)

    img = img.crop((0, 0, W, min(H, box_y1 + 12)))
    return _to_1bit(img, threshold=160)


# =============================
# MULTIROW PACKING & DELIVERY PRINT SLIP
# =============================
def _render_multirow_packing_list_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))

    margin_x = 18
    top_pad = 14
    in_pad = 12

    box_x0 = margin_x
    box_x1 = W - margin_x
    box_w = box_x1 - box_x0

    rows = data.get("rows") or []
    if not isinstance(rows, list):
        rows = []

    # -------------------------
    # spacing / sizing
    # -------------------------
    pad_after_title = 12
    pad_after_top_line = 10
    pad_after_info = 10
    table_header_gap = 8
    row_line_gap = 18
    row_block_gap = 6
    totals_gap = 6
    barcode_top_gap = 12
    bottom_pad = 14

    tmp = Image.new("L", (10, 10), 255)
    draw_tmp = ImageDraw.Draw(tmp)

    def wrap_text(s: str, max_px: int, font) -> list[str]:
        lines = _wrap_line_px(_safe_str(s), draw_tmp, font, max_px)
        return lines if lines else [""]

    # -------------------------
    # columns (match screenshot)
    # -------------------------
    col1_x = box_x0 + in_pad                 # Sh Code / Or no
    col2_x = box_x0 + int(box_w * 0.20)      # Unit Wt / P Sh Name -> moved LEFT
    col3_r = box_x0 + int(box_w * 0.66)      # Gross Wt
    col4_r = box_x0 + int(box_w * 0.84)      # Net Wt
    col5_r = box_x1 - in_pad                 # Qty

    left_max_px = (col2_x - 10) - col1_x
    prod_max_px = (col3_r - 22) - col2_x

    h_title = _font_line_h(F_PL_TITLE)
    h_info = _font_line_h(F_PL_INFO)
    h_hdr = _font_line_h(F_PL_ROW_B)
    h_row = _font_line_h(F_PL_ROW)

    # -------------------------
    # height estimation
    # -------------------------
    rows_h = 0
    for r in rows:
        sh_code = _safe_str(r.get("shCode") or r.get("partyShortCode") or r.get("partyAccountShortCode") or r.get("code"))
        order_no = _safe_str(r.get("orderNo"))
        branch_code = _safe_str(r.get("branchCode"))

        left_lines = []
        if sh_code:
            left_lines.append(sh_code)

        if order_no:
            left_lines.append(f"/ {order_no}")
        elif branch_code:
            left_lines.append(f"/ {branch_code}")

        if not left_lines:
            left_lines = [""]

        unitwt = _safe_str(r.get("unitWt") or r.get("unitWeight"))
        shortn = _safe_str(r.get("itemShortName") or r.get("productShortName") or r.get("productName") or r.get("product"))
        prod_text = f"{unitwt} {shortn}".strip()

        left_wrapped = []
        for ln in left_lines:
            left_wrapped.extend(wrap_text(ln, left_max_px, F_PL_ROW))

        prod_wrapped = wrap_text(prod_text, prod_max_px, F_PL_ROW)

        line_count = max(len(left_wrapped), len(prod_wrapped), 1)
        rows_h += (line_count * row_line_gap) + row_block_gap

    header_h = (
        10 + h_title + pad_after_title +
        1 + pad_after_top_line +
        (h_info * 2) + pad_after_info +
        (h_hdr * 2) + table_header_gap +
        1 + 8
    )

    totals_h = (h_info + totals_gap) * 3 + 12
    barcode_h = 96 + 10 + h_info + bottom_pad

    H = max(top_pad + header_h + rows_h + totals_h + barcode_h + 40, 900)

    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)

    y = top_pad
    box_y0 = y

    # -------------------------
    # title
    # -------------------------
    y += 10
    _draw_text_center(draw, box_x0 + box_w / 2, y, "Multirow Packing List", F_PL_TITLE)
    y += h_title + pad_after_title

    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += pad_after_top_line

    # -------------------------
    # top info (2 rows exactly like screenshot)
    # -------------------------
    date_str = _safe_str(data.get("date"))
    account_code = _safe_str(data.get("accountCode"))
    supplier = _safe_str(data.get("supplier"))
    recipient = _safe_str(data.get("recipient"))

    _draw_text(draw, box_x0 + in_pad, y, f"Date: {date_str}" if date_str else "Date:", F_PL_INFO)
    _draw_text_right(draw, box_x1 - in_pad, y, f"Supplier:{supplier}" if supplier else "Supplier:", F_PL_INFO)
    y += h_info + 4

    _draw_text(draw, box_x0 + in_pad, y, f"A/c Code: {account_code}" if account_code else "A/c Code:", F_PL_INFO)
    _draw_text_right(draw, box_x1 - in_pad, y, f"Recipient: {recipient}" if recipient else "Recipient:", F_PL_INFO)
    y += h_info + 8

    # TOP BORDER LINE ABOVE TABLE HEADER
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 8

    # table header starts here
    h1 = _draw_multiline(draw, col1_x, y, "Sh Code\n/ Or no", F_PL_ROW_B, line_gap=2)
    h2 = _draw_multiline(draw, col2_x - 4, y, "Unit Wt\nP Sh Name", F_PL_ROW_B, line_gap=2)
    h3 = _draw_multiline(draw, col3_r - int(_text_w(draw, "Gross Wt", F_PL_ROW_B)), y, "Gross Wt\n(gms)", F_PL_ROW_B, line_gap=2)
    h4 = _draw_multiline(draw, col4_r - int(_text_w(draw, "Net Wt", F_PL_ROW_B)), y, "Net Wt\n(gms)", F_PL_ROW_B, line_gap=2)

    qty_y = y + int((max(h1, h2, h3, h4) - h_hdr) / 2)
    _draw_text_right(draw, col5_r, qty_y, "Qty", F_PL_ROW_B)

    y += max(h1, h2, h3, h4) + table_header_gap
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 8

    # -------------------------
    # rows
    # -------------------------
    for r in rows:
        sh_code = _safe_str(r.get("shCode") or r.get("partyShortCode") or r.get("partyAccountShortCode") or r.get("code"))
        order_no = _safe_str(r.get("orderNo"))
        branch_code = _safe_str(r.get("branchCode"))

        left_lines = []
        if sh_code:
            left_lines.append(sh_code)

        if order_no:
            left_lines.append(f"/ {order_no}")
        elif branch_code:
            left_lines.append(f"/ {branch_code}")

        if not left_lines:
            left_lines = [""]

        unitwt = _safe_str(r.get("unitWt") or r.get("unitWeight"))
        shortn = _safe_str(r.get("itemShortName") or r.get("productShortName") or r.get("productName") or r.get("product"))
        prod_text = f"{unitwt} {shortn}".strip()

        gross_wt = _safe_str(r.get("grossWt") or r.get("grossWeight"))
        net_wt = _safe_str(r.get("netWt") or r.get("netWeight"))
        qty = _safe_str(r.get("qty"))

        left_wrapped = []
        for ln in left_lines:
            left_wrapped.extend(wrap_text(ln, left_max_px, F_PL_ROW))

        prod_wrapped = wrap_text(prod_text, prod_max_px, F_PL_ROW)

        line_count = max(len(left_wrapped), len(prod_wrapped), 1)
        row_top = y

        for i in range(line_count):
            yy = row_top + i * row_line_gap

            if i < len(left_wrapped):
                _draw_text(draw, col1_x, yy, left_wrapped[i], F_PL_ROW)

            if i < len(prod_wrapped):
                _draw_text(draw, col2_x, yy, prod_wrapped[i], F_PL_ROW)

            if i == 0:
                _draw_text_right(draw, col3_r, yy, gross_wt, F_PL_ROW)
                _draw_text_right(draw, col4_r, yy, net_wt, F_PL_ROW)
                _draw_text_right(draw, col5_r, yy, qty, F_PL_ROW)

        y += (line_count * row_line_gap) + row_block_gap

    # divider
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 12

    # -------------------------
    # totals
    # -------------------------
    total_net = _safe_str(data.get("totalNetWt"))
    total_gross = _safe_str(data.get("totalGrossWt"))
    total_qty = _safe_str(data.get("totalQty"))
    row_count = _safe_str(data.get("rowCount")) or str(len(rows))

    if total_net:
        _draw_text_center(draw, box_x0 + box_w / 2, y, f"Total Net Weight (gms):  {total_net}", F_PL_INFO)
        y += h_info + totals_gap

    if total_gross:
        _draw_text_center(draw, box_x0 + box_w / 2, y, f"Total Gross Weight (gms):  {total_gross}", F_PL_INFO)
        y += h_info + totals_gap

    _draw_text_center(draw, box_x0 + box_w / 2, y, f"Total Qty: {total_qty} Row Count: {row_count}", F_PL_INFO)
    y += h_info + barcode_top_gap

    # -------------------------
    # barcode
    # -------------------------
    barcode_val = account_code or "K0000"

    bx0 = box_x0 + in_pad
    bx1 = box_x1 - in_pad
    avail_w = bx1 - bx0

    bar_h = 84
    bar_w = int(avail_w * 0.72)

    barcode_img = code128_pil(
        barcode_val,
        bar_w,
        bar_h,
        module_width=0.22,
        module_height=12.0,
        quiet_zone=1.2,
        threshold=220,
    )

    if barcode_img.mode != "1":
        barcode_img = barcode_img.convert("1")

    paste_x = int(bx0 + (avail_w - bar_w) / 2)
    img.paste(barcode_img, (paste_x, y))

    y += bar_h + 8
    _draw_text_center(draw, box_x0 + box_w / 2, y, barcode_val, F_PL_INFO)
    y += h_info

    box_y1 = y + 12
    draw.rectangle((box_x0, box_y0, box_x1, box_y1), outline=0, width=1)

    img = img.crop((0, 0, W, box_y1 + 12))
    return img

# =========================
# PACKING LIST FONTS
# =========================
F_PL_TITLE  = F_HDR
F_PL_INFO   = F_TXT
F_PL_INFO_B = F_TXT_B
F_PL_ROW    = F_UNIT
F_PL_ROW_B  = ImageFont.truetype(FONT_BOLD_PATH, 20)


# =========================
# ROBUST HEIGHT + MULTILINE HELPERS (ADD ONCE)
# =========================
def _font_line_h(font) -> int:
    a, d = font.getmetrics()
    return int(a + d)

def _draw_multiline(draw, x, y, text, font, line_gap=2) -> int:
    """
    Draw multiline text and return total height used.
    """
    lines = (text or "").split("\n")
    lh = _font_line_h(font)
    yy = y
    for i, ln in enumerate(lines):
        draw.text((x, yy), ln, font=font, fill=0)
        yy += lh + (line_gap if i < len(lines) - 1 else 0)
    return yy - y


def _render_packing_list_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))

    margin_x = 18
    top_pad  = 14
    in_pad   = 14

    box_x0 = margin_x
    box_x1 = W - margin_x
    box_w  = box_x1 - box_x0

    rows = data.get("rows") or []
    if not isinstance(rows, list):
        rows = []

    # ---- spacing ----
    pad_after_title_line = 12
    pad_after_solid_line = 10
    pad_before_dash_line = 8
    pad_after_dash_line  = 12
    info_gap_y           = 6
    table_row_h          = 28
    wrapped_row_extra    = 18
    totals_gap_y         = 6
    bottom_pad           = 16

    tmp = Image.new("L", (10, 10), 255)
    draw_tmp = ImageDraw.Draw(tmp)

    def wrap2(s: str, max_px: int, font) -> list[str]:
        lines = _wrap_line_px(_safe_str(s), draw_tmp, font, max_px)
        if not lines:
            return [""]
        return lines[:2]

    # -------------------------
    # COLUMN POSITIONS
    # Shift product column to the right so it won't collide with Gross Wt
    # -------------------------
    col1_x = box_x0 + in_pad                    # Sh code / Or no
    col2_x = box_x0 + int(box_w * 0.20)         # Unit Wt + Product Short Name (moved right)
    col3_r = box_x0 + int(box_w * 0.70)         # Gross Wt right edge
    col4_r = box_x0 + int(box_w * 0.87)         # Net Wt right edge
    col5_r = box_x1 - in_pad                    # Qty right edge

    product_max_px = (col3_r - 14) - col2_x

    # ---- height estimate ----
    h_title = _font_line_h(F_PL_TITLE)
    h_info  = _font_line_h(F_PL_INFO)
    h_hdr   = _font_line_h(F_PL_ROW_B)

    extra_wrap = 0
    for r in rows:
        unitwt = _safe_str(r.get("unitWt") or r.get("unitWeight"))
        shortn = _safe_str(r.get("itemShortName") or r.get("productShortName") or r.get("productName") or r.get("product"))
        prod = f"{unitwt} {shortn}".strip()
        prod_lines = wrap2(prod, product_max_px, F_PL_ROW)
        if len(prod_lines) > 1:
            extra_wrap += wrapped_row_extra

    table_header_h = (h_hdr * 2) + 2
    header_h  = 10 + h_title + pad_after_title_line + 1 + pad_after_solid_line
    header_h += (h_info + info_gap_y) * 3
    header_h += pad_before_dash_line + 1 + pad_after_dash_line

    table_h   = table_header_h + 8 + (len(rows) * table_row_h) + extra_wrap + 14
    totals_h  = (h_info + totals_gap_y) * 4 + 20   # net + gross + qty + row count
    barcode_h = 78 + 10 + h_info + bottom_pad + 20

    H = max(top_pad + header_h + table_h + totals_h + barcode_h + 200, 1600)

    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)

    y = top_pad
    box_y0 = y

    # -------------------------
    # Title
    # -------------------------
    y += 10
    _draw_text_center(draw, box_x0 + box_w / 2, y, "Packing List", F_PL_TITLE)
    y += h_title + pad_after_title_line

    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += pad_after_solid_line

    # -------------------------
    # Top info
    # -------------------------
    date_str  = _safe_str(data.get("date"))
    dc_no     = _safe_str(data.get("dcNo"))
    ac_code   = _safe_str(data.get("accountCode"))
    packet_no = _safe_str(data.get("packetNo"))
    supplier  = _safe_str(data.get("supplier"))
    recipient = _safe_str(data.get("recipient"))

    _draw_text(draw, box_x0 + in_pad, y, f"Date: {date_str}" if date_str else "Date:", F_PL_INFO)
    _draw_text_right(draw, box_x1 - in_pad, y, f"DC No: {dc_no}" if dc_no else "DC No:", F_PL_INFO)
    y += h_info + info_gap_y

    _draw_text(draw, box_x0 + in_pad, y, f"A/c Code: {ac_code}" if ac_code else "A/c Code:", F_PL_INFO)
    _draw_text_right(draw, box_x1 - in_pad, y, f"Packet No: {packet_no}" if packet_no else "Packet No:", F_PL_INFO)
    y += h_info + info_gap_y

    _draw_text(draw, box_x0 + in_pad, y, f"Supplier: {supplier}" if supplier else "Supplier:", F_PL_INFO)
    _draw_text_right(draw, box_x1 - in_pad, y, f"Recipient: {recipient}" if recipient else "Recipient:", F_PL_INFO)
    y += h_info + pad_before_dash_line

    _dash_line(draw, box_x0 + 2, box_x1 - 2, y, dash=6, gap=4, width=1)
    y += pad_after_dash_line

    # -------------------------
    # Table header
    # -------------------------
    h1 = _draw_multiline(draw, col1_x, y, "Sh code\n/ Or no", F_PL_ROW_B, line_gap=2)
    h2 = _draw_multiline(draw, col2_x, y, "Unit Wt\nP Sh Name", F_PL_ROW_B, line_gap=2)
    h3 = _draw_multiline(draw, col3_r - int(_text_w(draw, "Gross Wt", F_PL_ROW_B)), y, "Gross Wt\n(gms)", F_PL_ROW_B, line_gap=2)
    h4 = _draw_multiline(draw, col4_r - int(_text_w(draw, "Net Wt", F_PL_ROW_B)), y, "Net Wt\n(gms)", F_PL_ROW_B, line_gap=2)

    qty_y = y + int((max(h1, h2, h3, h4) - _font_line_h(F_PL_ROW_B)) / 2)
    _draw_text_right(draw, col5_r, qty_y, "Qty", F_PL_ROW_B)

    y += max(h1, h2, h3, h4) + 8
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 8

    # -------------------------
    # Rows
    # -------------------------
    for r in rows:
        party_short_code = _safe_str(
            r.get("partyShortCode") or
            r.get("partyAccountShortCode") or
            r.get("shCode") or
            r.get("code")
        )
        order_no = _safe_str(r.get("orderNo"))

        left_code = f"{party_short_code} /\n{order_no}" if order_no else party_short_code

        unitwt = _safe_str(r.get("unitWt") or r.get("unitWeight"))
        shortn = _safe_str(
            r.get("itemShortName") or
            r.get("productShortName") or
            r.get("productName") or
            r.get("product")
        )
        prod = f"{unitwt} {shortn}".strip()

        gross_wt = _safe_str(r.get("grossWt") or r.get("grossWeight"))
        net_wt   = _safe_str(r.get("netWt") or r.get("netWeight"))
        qty      = _safe_str(r.get("qty"))

        left_lines = left_code.split("\n")
        prod_lines = wrap2(prod, product_max_px, F_PL_ROW)

        # first line
        _draw_text(draw, col1_x, y, left_lines[0], F_PL_ROW)
        _draw_text(draw, col2_x, y, prod_lines[0], F_PL_ROW)
        _draw_text_right(draw, col3_r, y, gross_wt, F_PL_ROW)
        _draw_text_right(draw, col4_r, y, net_wt, F_PL_ROW)
        _draw_text_right(draw, col5_r, y, qty, F_PL_ROW)

        y += table_row_h

        # second line for OR number
        if len(left_lines) > 1:
            _draw_text(draw, col1_x, y - 10, left_lines[1], F_PL_ROW)

        # wrapped second product line
        if len(prod_lines) > 1:
            _draw_text(draw, col2_x, y - 10, prod_lines[1], F_PL_ROW)

        if len(left_lines) > 1 or len(prod_lines) > 1:
            y += wrapped_row_extra

    y += 4
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 12

    # -------------------------
    # Totals
    # -------------------------
    total_net   = _safe_str(data.get("totalNetWt"))
    total_gross = _safe_str(data.get("totalGrossWt"))
    total_qty   = _safe_str(data.get("totalQty"))
    row_count   = _safe_str(data.get("rowCount")) or str(len(rows))

    if total_net:
        _draw_text_center(draw, box_x0 + box_w / 2, y, f"Total Net Weight (gms):  {total_net}", F_PL_INFO)
        y += h_info + totals_gap_y

    if total_gross:
        _draw_text_center(draw, box_x0 + box_w / 2, y, f"Total Gross Weight (gms):  {total_gross}", F_PL_INFO)
        y += h_info + totals_gap_y

    if total_qty:
        _draw_text_center(draw, box_x0 + box_w / 2, y, f"Total Qty: {total_qty}", F_PL_INFO)
        y += h_info + totals_gap_y

    if row_count:
        _draw_text_center(draw, box_x0 + box_w / 2, y, f"Row Count: {row_count}", F_PL_INFO)
        y += h_info + 8
    # -------------------------
    # Barcode
    # -------------------------
    barcode_val = ac_code or "K0000"

    bx0 = box_x0 + in_pad
    bx1 = box_x1 - in_pad
    avail_w = bx1 - bx0

    bar_h = 78
    bar_w = int(avail_w * 0.68)

    barcode_img = code128_pil(barcode_val, bar_w, bar_h, module_width=0.40, quiet_zone=2.5)
    paste_x = int(bx0 + (avail_w - bar_w) / 2)
    img.paste(barcode_img, (paste_x, y))

    y += bar_h + 8
    _draw_text_center(draw, box_x0 + box_w / 2, y, barcode_val, F_PL_INFO)
    y += h_info

    box_y1 = y + 12
    draw.rectangle((box_x0, box_y0, box_x1, box_y1), outline=0, width=1)

    img = img.crop((0, 0, W, box_y1 + 12))
    return _to_1bit(img, threshold=160)


# =========================
# RATE SLIP FONTS (same sizing style as Packing List)
# =========================
def _draw_multiline_center(draw, x_center, y, text, font, line_gap=2) -> int:
    lines = (text or "").split("\n")
    lh = _font_line_h(font)
    yy = y
    for i, ln in enumerate(lines):
        _draw_text_center(draw, x_center, yy, ln, font)
        yy += lh + (line_gap if i < len(lines) - 1 else 0)
    return yy - y


F_RS_TITLE = F_HDR                 # 26 bold
F_RS_INFO  = F_TXT                 # 24 regular
F_RS_HDR   = ImageFont.truetype(FONT_BOLD_PATH, 20)
F_RS_ROW   = F_UNIT                # 20 regular


def _render_rate_slip_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))

    # Same layout style as Packing List
    margin_x = 18
    top_pad  = 14
    in_pad   = 14

    box_x0 = margin_x
    box_x1 = W - margin_x
    box_w  = box_x1 - box_x0

    rows = data.get("rows") or []

    def _has_any(key: str) -> bool:
        return any(_safe_str(r.get(key)) for r in rows)

    # If you already have a flag, it wins
    if data.get("isPureAccount") is not None:
        is_pure = bool(data.get("isPureAccount"))
    elif _has_any("touchPercent"):
        is_pure = True
    elif _has_any("wastagePercent"):
        is_pure = False
    else:
        # fallback (choose what you want)
        is_pure = True


    # title = "Rate Slip Format - For Pure Account" if is_pure else "Rate Slip Format - For Non-Pure Account"
    title = "Rate Slip Format"

    # Fonts (reuse packing list style)
    F_RS_TITLE = F_HDR           # 26 bold
    F_RS_INFO  = F_TXT           # 24 regular
    F_RS_INFO_B = F_TXT_B        # 24 bold
    F_RS_ROW   = F_UNIT          # 20 regular
    F_RS_ROW_B = ImageFont.truetype(FONT_BOLD_PATH, 20)

    # Spacing (same robustness as Packing List)
    pad_after_title_line = 14
    pad_after_solid_line = 12
    pad_before_dash_line = 10
    pad_after_dash_line  = 14

    info_gap_y   = 8
    table_row_h  = 28
    totals_gap_y = 6

    def wrap2(s: str, max_px: int, font) -> list[str]:
        s = _safe_str(s)
        if not s:
            return [""]
        words = s.split()
        out, cur = [], ""
        for w in words:
            t = w if not cur else f"{cur} {w}"
            if _text_w(draw, t, font) <= max_px:
                cur = t
            else:
                if cur:
                    out.append(cur)
                cur = w
        if cur:
            out.append(cur)
        return out[:2]

    # Estimate height (safe)
    extra_wrap = 0
    for r in rows:
        prod = _safe_str(r.get("productName")) or _safe_str(r.get("product"))
        unitwt = _safe_str(r.get("unitWt")) or _safe_str(r.get("unitWeight"))
        if unitwt and prod:
            prod = f"{unitwt} {prod}".strip()
        if len(prod) > 22:
            extra_wrap += (table_row_h - 8)

    h_info = _font_line_h(F_RS_INFO)
    h_hdr  = _font_line_h(F_RS_ROW_B)

    table_header_h = (h_hdr * 2) + 2

    header_h  = 10 + _font_line_h(F_RS_TITLE) + pad_after_title_line + 1 + pad_after_solid_line
    header_h += (h_info + info_gap_y) * 3
    header_h += pad_before_dash_line + 1 + pad_after_dash_line

    table_h   = table_header_h + 8 + (len(rows) * table_row_h) + extra_wrap + 16
    totals_h  = (h_info + totals_gap_y) * 5 + 18
    barcode_h = 78 + 10 + (h_info + 18)
    bottom_pad = 18

    H = top_pad + header_h + table_h + totals_h + barcode_h + bottom_pad
    H = max(H, 560)

    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)

    y = top_pad
    box_y0 = y

    # -------------------------
    # Title
    # -------------------------
    y += 10
    _draw_text_center(draw, box_x0 + box_w/2, y, title, F_RS_TITLE)
    y += _font_line_h(F_RS_TITLE) + pad_after_title_line

    # Top solid line
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += pad_after_solid_line

    # -------------------------
    # Top Info
    # -------------------------
    date_str  = _safe_str(data.get("date"))
    dc_no     = _safe_str(data.get("dcNo"))
    ac_code   = _safe_str(data.get("accountCode"))
    packet_no = _safe_str(data.get("packetNo"))
    supplier  = _safe_str(data.get("supplier"))
    recipient = _safe_str(data.get("recipient"))

    _draw_text(draw, box_x0 + in_pad, y, f"Date: {date_str}" if date_str else "Date:", F_RS_INFO)
    _draw_text_right(draw, box_x1 - in_pad, y, f"DC No: {dc_no}" if dc_no else "DC No:", F_RS_INFO)
    y += h_info + info_gap_y

    _draw_text(draw, box_x0 + in_pad, y, f"A/c Code: {ac_code}" if ac_code else "A/c Code:", F_RS_INFO)
    _draw_text_right(draw, box_x1 - in_pad, y, f"Packet No: {packet_no}" if packet_no else "Packet No:", F_RS_INFO)
    y += h_info + info_gap_y

    _draw_text(draw, box_x0 + in_pad, y, f"Supplier: {supplier}" if supplier else "Supplier:", F_RS_INFO)
    _draw_text_right(draw, box_x1 - in_pad, y, f"Recipient: {recipient}" if recipient else "Recipient:", F_RS_INFO)
    y += h_info + pad_before_dash_line

    # dashed separator
    _dash_line(draw, box_x0 + 2, box_x1 - 2, y, dash=6, gap=4, width=1)
    y += pad_after_dash_line

    # -------------------------
    # Table Columns (FIXED columns + right aligned numbers)
    # -------------------------
    # Column boundaries (RIGHT edges) — tuned to stop collapse
    col_prod_x   = box_x0 + in_pad
    col_qty_r    = box_x0 + int(box_w * 0.42)   # Qty
    col_net_r    = box_x0 + int(box_w * 0.62)   # Net Wt
    col_mid_r    = box_x0 + int(box_w * 0.80)   # Touch/Wastage
    col_pure_r   = box_x1 - in_pad              # Pure Wt

    # Header positions (left for labels, right for numeric headers)
    h_prod = _draw_multiline(draw, col_prod_x, y, "Unit Wt\nProduct Name", F_RS_ROW_B, line_gap=2)

    qty_y = y + int((max(h_prod, table_header_h) - _font_line_h(F_RS_ROW_B)) / 2)
    _draw_text_right(draw, col_qty_r, qty_y, "Qty", F_RS_ROW_B)

    h_net = _draw_multiline(draw, col_net_r - int(_text_w(draw, "Net Wt", F_RS_ROW_B)), y, "Net Wt\n(gms)", F_RS_ROW_B, line_gap=2)

    mid_title = "Touch\n%" if is_pure else "Wastage\n%"
    h_mid = _draw_multiline(draw, col_mid_r - int(_text_w(draw, "Wastage", F_RS_ROW_B)), y, mid_title, F_RS_ROW_B, line_gap=2)

    h_pure = _draw_multiline(draw, col_pure_r - int(_text_w(draw, "Pure Wt", F_RS_ROW_B)), y, "Pure Wt\n(gms)", F_RS_ROW_B, line_gap=2)

    y += max(h_prod, h_net, h_mid, h_pure) + 8
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 10

    # -------------------------
    # Rows
    # -------------------------
    max_prod_px = (col_qty_r - 14) - col_prod_x

    for r in rows:
        prod = _safe_str(r.get("productName")) or _safe_str(r.get("product"))
        unitwt = _safe_str(r.get("unitWt")) or _safe_str(r.get("unitWeight"))
        if unitwt and prod:
            prod = f"{unitwt} {prod}".strip()

        qty = _safe_str(r.get("qty"))
        net = _safe_str(r.get("netWt")) or _safe_str(r.get("netWeight"))

        if is_pure:
            mid = (
                _safe_str(r.get("touchPercent")) or   # ✅ your payload
                _safe_str(r.get("touchPct")) or
                _safe_str(r.get("touch"))             # fallback
            )
        else:
            mid = (
                _safe_str(r.get("wastagePercent")) or # ✅ your payload
                _safe_str(r.get("wastagePct")) or
                _safe_str(r.get("wastage"))           # fallback
            )

        pure = _safe_str(r.get("pureWt")) or _safe_str(r.get("pureWeight"))

        prod_lines = wrap2(prod, max_prod_px, F_RS_ROW)

        _draw_text(draw, col_prod_x, y, prod_lines[0], F_RS_ROW)
        _draw_text_right(draw, col_qty_r, y, qty, F_RS_ROW)
        _draw_text_right(draw, col_net_r, y, net, F_RS_ROW)
        _draw_text_right(draw, col_mid_r, y, mid, F_RS_ROW)
        _draw_text_right(draw, col_pure_r, y, pure, F_RS_ROW)

        y += table_row_h

        if len(prod_lines) > 1:
            _draw_text(draw, col_prod_x, y - 8, prod_lines[1], F_RS_ROW)
            y += (table_row_h - 8)

    y += 6
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 14

    # -------------------------
    # Totals (centered like your printed sample)
    # -------------------------
    total_net   = _safe_str(data.get("totalNetWt"))
    total_gross = _safe_str(data.get("totalGrossWt"))
    total_pure  = _safe_str(data.get("totalPureWt"))
    total_qty   = _safe_str(data.get("totalQty"))
    row_count   = _safe_str(data.get("rowCount")) or str(len(rows))

    if total_net:
        _draw_text_center(draw, box_x0 + box_w/2, y, f"Total Net Weight (gms):  {total_net}", F_RS_INFO)
        y += h_info + totals_gap_y
    if total_gross:
        _draw_text_center(draw, box_x0 + box_w/2, y, f"Total Gross Weight (gms):  {total_gross}", F_RS_INFO)
        y += h_info + totals_gap_y
    if total_pure:
        _draw_text_center(draw, box_x0 + box_w/2, y, f"Total Pure Weight (gms):  {total_pure}", F_RS_INFO)
        y += h_info + totals_gap_y

    if total_qty:
        _draw_text_center(draw, box_x0 + box_w/2, y, f"Total Qty:  {total_qty}", F_RS_INFO)
        y += h_info + totals_gap_y

    _draw_text_center(draw, box_x0 + box_w/2, y, f"Row Count:  {row_count}", F_RS_INFO)
    y += h_info - 2

    # -------------------------
    # Barcode (same as Packing List)
    # -------------------------
    y += 10
    barcode_val = ac_code or "K0000"

    bx0 = box_x0 + in_pad
    bx1 = box_x1 - in_pad
    avail_w = bx1 - bx0

    bar_h = 78
    bar_w = int(avail_w * 0.92)

    barcode_img = code128_pil(barcode_val, bar_w, bar_h, module_width=0.40, quiet_zone=2.5)
    paste_x = int(bx0 + (avail_w - bar_w) / 2)
    img.paste(barcode_img, (paste_x, y))

    y += bar_h + 10
    _draw_text_center(draw, box_x0 + box_w/2, y, barcode_val, F_RS_INFO)
    y += h_info

    # Outer border
    box_y1 = y + 12
    draw.rectangle((box_x0, box_y0, box_x1, box_y1), outline=0, width=1)

    img = img.crop((0, 0, W, min(H, box_y1 + 12)))
    return _to_1bit(img, threshold=160) 

# =========================
# BALANCE SUMMARY (80mm) — Customer/Vendor
# Adds column: Metal A/c
# =========================

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

def _render_balance_summary_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))

    # --- layout (same style as Packing List) ---
    margin_x = 18
    top_pad  = 14
    in_pad   = 14

    box_x0 = margin_x
    box_x1 = W - margin_x
    box_w  = box_x1 - box_x0

    rows = data.get("rows") or []
    if not isinstance(rows, list):
        rows = []

    # =========================
    # TITLE LOGIC (AUTO HEADINGS)
    # =========================
    # You can pass any of these keys:
    # - balanceType: "customer_pure" | "customer_mb" | "vendor_mb"
    # - reportType / slipType / titleKey (same accepted values)
    # - sourceTitle: original UI title like "Customer Accounts Balance (Pure Weight)"
    #
    # Or you can still pass "title" directly to override everything.
    explicit_title = _safe_str(data.get("title"))
    if explicit_title:
        title = explicit_title
    else:
        # 1) normalized type key
        tkey = (
            _safe_str(data.get("balanceType")) or
            _safe_str(data.get("reportType")) or
            _safe_str(data.get("slipType")) or
            _safe_str(data.get("titleKey"))
        ).strip().lower()

        # 2) sometimes UI sends full text; map it too
        source_title = (_safe_str(data.get("sourceTitle")) or _safe_str(data.get("uiTitle"))).strip().lower()

        # map by type-key first
        if tkey in ("customer_pure_weight"):
            title = "Customer Pure Balance Summary"
        elif tkey in ("customer_mb"):
            title = "Customer Balance Summary"
        elif tkey in ("vendor_pure_weight"):
            title = "Vendor Pure Balance Summary"
        elif tkey in ("vendor_mb"):
            title = "Vendor Balance Summary"
        else:
            # map by sourceTitle text (what you wrote)
            if "customer accounts balance" in source_title and "pure" in source_title:
                title = "Customer Pure Balance Summary"
            elif "customer accounts balance" in source_title and ("mb" in source_title or "purity" in source_title):
                title = "Customer Balance Summary"
            elif "vendor accounts balance" in source_title and "pure" in source_title:
                title = "Vendor Pure Balance Summary"
            elif "vendor accounts balance" in source_title and ("mb" in source_title or "purity" in source_title):
                title = "Vendor Balance Summary"
            else:
                # fallback flags if present
                if data.get("isVendor") is True:
                    title = "Vendor Balance Summary"
                elif data.get("isCustomer") is True:
                    title = "Customer Balance Summary"
                else:
                    title = "Balance Summary"

    print_dt = _safe_str(data.get("printDateTime"))
    total_balance = _safe_str(data.get("totalBalance")) or _safe_str(data.get("total"))

    # --- fonts (reuse global ones; create a small bold for table header) ---
    F_BS_TITLE  = F_HDR
    F_BS_INFO   = F_TXT
    F_BS_INFO_B = F_TXT_B
    F_BS_ROW    = F_UNIT

    try:
        F_BS_ROW_B = ImageFont.truetype(FONT_BOLD_PATH, 20)
    except Exception:
        F_BS_ROW_B = F_TXT_B

    # --- spacing constants (robust) ---
    pad_after_title_line = 14
    pad_after_solid_line = 12
    header_bottom_pad    = 8
    row_h_base           = 28
    row_wrap_extra       = 20
    after_table_line_gap = 12
    totals_gap_y         = 6
    bottom_pad           = 18

    # --- column anchors ---
    col_metal_x = box_x0 + in_pad
    col_ac_x    = box_x0 + int(box_w * 0.26)
    col_party_x = box_x0 + int(box_w * 0.46)
    col_bal_r   = box_x1 - in_pad

    party_max_px = (col_bal_r - 12) - col_party_x

    # --- pre-calc height safely ---
    h_title = _font_line_h(F_BS_TITLE)
    h_info  = _font_line_h(F_BS_INFO)
    h_hdr   = _font_line_h(F_BS_ROW_B)

    extra_wrap = 0
    tmp = Image.new("L", (10, 10), 255)
    dtmp = ImageDraw.Draw(tmp)

    def wrap2(text: str, max_px: int, font) -> list[str]:
        lines = _wrap_line_px(text, dtmp, font, max_px)
        if not lines:
            return [""]
        return lines[:2]

    for r in rows:
        if not isinstance(r, dict):
            continue
        party = _safe_str(r.get("party")) or _safe_str(r.get("name"))
        plines = wrap2(party, party_max_px, F_BS_ROW)
        if len(plines) > 1:
            extra_wrap += row_wrap_extra

    totals_lines = 0
    if total_balance:
        totals_lines += 1
    if print_dt:
        totals_lines += max(1, len(print_dt.splitlines()))
    if totals_lines == 0:
        totals_lines = 1

    header_h = (
        10 + h_title + pad_after_title_line
        + 1 + pad_after_solid_line
        + max(h_hdr, row_h_base) + header_bottom_pad
        + 1 + 10
    )

    table_h  = (len(rows) * row_h_base) + extra_wrap + 1 + after_table_line_gap
    totals_h = totals_lines * (h_info + totals_gap_y) + bottom_pad + 20

    H = top_pad + header_h + table_h + totals_h
    H = max(H, 380)

    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)

    y = top_pad
    box_y0 = y

    # Title
    y += 10
    _draw_text_center(draw, box_x0 + box_w / 2, y, title, F_BS_TITLE)
    y += h_title + pad_after_title_line

    # top line
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += pad_after_solid_line

    # table header
    _draw_text(draw, col_metal_x, y, "Metal A/c", F_BS_ROW_B)
    _draw_text(draw, col_ac_x,    y, "A/c",      F_BS_ROW_B)
    _draw_text(draw, col_party_x, y, "Party",    F_BS_ROW_B)
    _draw_text_right(draw, col_bal_r, y, "Balance", F_BS_ROW_B)

    y += max(row_h_base, h_hdr) + header_bottom_pad
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 10

    # rows
    for r in rows:
        if not isinstance(r, dict):
            continue

        metal_ac = _safe_str(r.get("metalAc")) or _safe_str(r.get("metalAccount")) or _safe_str(r.get("metal_ac"))
        ac       = _safe_str(r.get("ac")) or _safe_str(r.get("account")) or _safe_str(r.get("acCode"))
        party    = _safe_str(r.get("party")) or _safe_str(r.get("name"))
        bal      = _safe_str(r.get("balance")) or _safe_str(r.get("bal")) or _safe_str(r.get("closingBalance"))

        party_lines = wrap2(party, party_max_px, F_BS_ROW)

        _draw_text(draw, col_metal_x, y, metal_ac, F_BS_ROW)
        _draw_text(draw, col_ac_x,    y, ac,       F_BS_ROW)
        _draw_text(draw, col_party_x, y, party_lines[0], F_BS_ROW)
        _draw_text_right(draw, col_bal_r, y, bal, F_BS_ROW)

        y += row_h_base

        if len(party_lines) > 1:
            _draw_text(draw, col_party_x, y - 8, party_lines[1], F_BS_ROW)
            y += row_wrap_extra

    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += after_table_line_gap

    # totals
    if total_balance:
        _draw_text(draw, col_party_x, y, "Total", F_BS_INFO_B)
        _draw_text_right(draw, col_bal_r, y, total_balance, F_BS_INFO_B)
        y += h_info + totals_gap_y

    # print time
    if print_dt:
        for ln in print_dt.splitlines():
            _draw_text_center(draw, box_x0 + box_w / 2, y, ln, F_BS_INFO)
            y += h_info + 2

    # border
    box_y1 = y + 12
    draw.rectangle((box_x0, box_y0, box_x1, box_y1), outline=0, width=1)

    img = img.crop((0, 0, W, min(H, box_y1 + 12)))
    return _to_1bit(img, threshold=160)


# =========================
# JOB CREATE SLIP (80mm)
# =========================

def _fmt_job_slip_weight(v) -> str:
    try:
        return f"{max(0.0, float(v)):.3f}g"
    except Exception:
        return "0.000g"


def _decode_worker_thumb_b64(b64: str):
    if not b64 or not isinstance(b64, str):
        return None
    s = b64.strip()
    if s.startswith("data:") and "," in s:
        s = s.split(",", 1)[1]
    try:
        raw = base64.b64decode(s)
        return Image.open(io.BytesIO(raw)).convert("RGBA")
    except Exception:
        return None


def _paste_worker_thumb_job_create_slip(canvas: Image.Image, gx: int, gy: int, thumb_rgba: Image.Image, size: int):
    """
    Paste employee photo for JOB CREATE slip: composite on white, stretch contrast,
    gamma-brighten shadows (reduces 'solid black face' on thermal), then paste with alpha.
    Final slip uses Floyd–Steinberg dither so tones survive 1-bit printing.
    """
    t = thumb_rgba.resize((size, size), Image.Resampling.LANCZOS)
    if t.mode != "RGBA":
        t = t.convert("RGBA")
    bg = Image.new("RGBA", (size, size), (255, 255, 255, 255))
    comp = Image.alpha_composite(bg, t)
    L = comp.convert("RGB").convert("L")
    L = ImageOps.autocontrast(L, cutoff=1)
    # Gamma < 1 lifts dark skin / shadow areas before 1-bit conversion
    gamma = 0.52
    lut = [min(255, int(round(255.0 * ((i / 255.0) ** gamma)))) for i in range(256)]
    L = L.point(lut)
    L = ImageEnhance.Brightness(L).enhance(1.14)
    L = ImageEnhance.Contrast(L).enhance(1.06)
    a = t.split()[3]
    canvas.paste(L, (gx, gy), a)


def _paste_worker_thumb_circle_job_create_slip(canvas: Image.Image, gx: int, gy: int, thumb_rgba: Image.Image, size: int):
    """Same processing as square paste, but composite onto canvas with a circular mask."""
    t = thumb_rgba.resize((size, size), Image.Resampling.LANCZOS)
    if t.mode != "RGBA":
        t = t.convert("RGBA")
    bg = Image.new("RGBA", (size, size), (255, 255, 255, 255))
    comp = Image.alpha_composite(bg, t)
    L = comp.convert("RGB").convert("L")
    L = ImageOps.autocontrast(L, cutoff=1)
    gamma = 0.52
    lut = [min(255, int(round(255.0 * ((i / 255.0) ** gamma)))) for i in range(256)]
    L = L.point(lut)
    L = ImageEnhance.Brightness(L).enhance(1.14)
    L = ImageEnhance.Contrast(L).enhance(1.06)
    a = t.split()[3]
    circle = Image.new("L", (size, size), 0)
    cd = ImageDraw.Draw(circle)
    cd.ellipse((0, 0, size - 1, size - 1), fill=255)
    mask = ImageChops.multiply(a, circle)
    canvas.paste(L, (gx, gy), mask)


def _draw_worker_initials_circle_job_create_slip(draw: ImageDraw.ImageDraw, gx: int, gy: int, size: int, name: str, font_ini):
    draw.ellipse((gx, gy, gx + size - 1, gy + size - 1), outline=0, width=1)
    ini = (_safe_str(name) or "?")[:2].upper()
    cx = gx + size // 2
    cy = gy + size // 2
    lh = _font_line_h(font_ini)
    _draw_text_center(draw, cx, cy - lh // 2, ini, font_ini)


def _render_job_create_slip_image_80mm(data: dict) -> Image.Image:
    W = int(data.get("maxWidthDots", MAX_WIDTH_DOTS_80MM))
    margin_x = 18
    top_pad = 14
    in_pad_x = 14
    box_x0 = margin_x
    box_x1 = W - margin_x
    box_w = box_x1 - box_x0

    job_code = _safe_str(data.get("jobCode")) or "-"
    date_time = _safe_str(data.get("dateTime")) or "-"
    # Default process when a finished row omits processName (job-level from client)
    process_name_default = _safe_str(data.get("processName")) or "-"

    workers_in = data.get("workers")
    if not isinstance(workers_in, list) or len(workers_in) == 0:
        workers_in = [{"name": "NA"}]

    enable_giving = bool(data.get("enableGiving"))
    giving_rows = data.get("givingRows") or []
    if not isinstance(giving_rows, list):
        giving_rows = []

    giving_net = 0.0
    try:
        giving_net = float(data.get("givingNet") or 0)
    except Exception:
        giving_net = 0.0

    has_giving = enable_giving and (giving_net > 0 or len(giving_rows) > 0)

    finished_rows = data.get("finishedRows") or []
    if not isinstance(finished_rows, list):
        finished_rows = []

    form_item = _safe_str(data.get("formItemName")) or "-"
    try:
        form_qty = int(data.get("formQuantity") or 1)
    except Exception:
        form_qty = 1

    if len(finished_rows) == 0:
        finished_rows = [{"itemName": form_item, "quantity": form_qty, "unitWeight": data.get("formUnitWeight")}]

    # Tall enough for many workers + many giving rows (crop to content at end)
    H = 4000
    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)
    y = top_pad
    box_y0 = y

    # Header
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 10
    _draw_text_center(draw, box_x0 + box_w / 2, y, "JOB CREATE SLIP", F_HDR)
    y += _font_line_h(F_HDR) + 8
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 14

    _draw_text(draw, box_x0 + in_pad_x, y, job_code, F_TXT)
    _draw_text_right(draw, box_x1 - in_pad_x, y, date_time, F_TXT)
    y += _font_line_h(F_TXT) + 10

    # Workers: sketch layout — 1 worker = centered circle + name below; 2+ = overlapping circles + comma names
    worker_entries = []
    for w in workers_in:
        if not isinstance(w, dict):
            continue
        nm = _safe_str(w.get("name")) or "-"
        b64 = w.get("imageB64")
        worker_entries.append({"name": nm, "b64": b64 if isinstance(b64, str) else None})

    if not worker_entries:
        worker_entries = [{"name": "NA", "b64": None}]

    thumb_base = int(round(56 * 1.6))
    gap_after_workers = 10
    F_NAME = F_TXT
    lh_name = _font_line_h(F_NAME)
    cx_page = box_x0 + box_w / 2
    wrap_w_names = max(60, int(box_w - 2 * in_pad_x))

    n_workers = len(worker_entries)

    if n_workers == 1:
        thumb_size = thumb_base
        ent = worker_entries[0]
        gx = int(cx_page - thumb_size / 2)
        im = _decode_worker_thumb_b64(ent["b64"] or "")
        if im:
            _paste_worker_thumb_circle_job_create_slip(img, gx, y, im, thumb_size)
        else:
            _draw_worker_initials_circle_job_create_slip(draw, gx, y, thumb_size, ent["name"], F_TXT_B)
        y += thumb_size + gap_after_workers
        name_plain = _safe_str(ent.get("name")) or "-"
        lines = _wrap_line_px(name_plain, draw, F_NAME, wrap_w_names) or ["-"]
        for ln in lines:
            _draw_text_center(draw, cx_page, y, ln, F_NAME)
            y += lh_name + 2
    else:
        thumb_m = thumb_base
        step = max(1, int(thumb_m * 0.9))
        max_inner = int(box_w - 2 * in_pad_x)
        while thumb_m > 32:
            total_w = thumb_m + (n_workers - 1) * step
            if total_w <= max_inner:
                break
            thumb_m = max(32, int(thumb_m * 0.9))
            step = max(1, int(thumb_m * 0.9))
        total_w = thumb_m + (n_workers - 1) * step
        start_gx = int(cx_page - total_w / 2)
        row_top = y
        for i, ent in enumerate(worker_entries):
            gx_i = start_gx + i * step
            im = _decode_worker_thumb_b64(ent["b64"] or "")
            if im:
                _paste_worker_thumb_circle_job_create_slip(img, gx_i, row_top, im, thumb_m)
            else:
                _draw_worker_initials_circle_job_create_slip(draw, gx_i, row_top, thumb_m, ent["name"], F_TXT_B)
        y = row_top + thumb_m + gap_after_workers
        combined = ", ".join(_safe_str(e.get("name")) or "-" for e in worker_entries)
        lines = _wrap_line_px(combined, draw, F_NAME, wrap_w_names) or ["-"]
        for ln in lines:
            _draw_text_center(draw, cx_page, y, ln, F_NAME)
            y += lh_name + 2

    y += 4

    # INPUTS
    if has_giving:
        draw.line((box_x0, y, box_x1, y), fill=0, width=1)
        y += 10
        t_in = "INPUTS"
        tw = _text_w(draw, t_in, F_TXT_B)
        _draw_text_center(draw, box_x0 + box_w / 2, y, t_in, F_TXT_B)
        y += _font_line_h(F_TXT_B) + 4
        draw.line((box_x0 + box_w / 2 - tw / 2, y, box_x0 + box_w / 2 + tw / 2, y), fill=0, width=1)
        y += 12

        total_net = 0.0
        lh_lbl = _font_line_h(F_TXT)
        lh_wt = _font_line_h(F_JC_WEIGHT)

        def _draw_input_row(label: str, wt_str: str):
            nonlocal y
            rh = max(lh_lbl, lh_wt)
            _draw_text(draw, box_x0 + in_pad_x, y + (rh - lh_lbl) // 2, label[:48], F_TXT)
            _draw_text_right(draw, box_x1 - in_pad_x, y + (rh - lh_wt) // 2, wt_str, F_JC_WEIGHT)
            y += rh + 6

        if len(giving_rows) > 0:
            for r in giving_rows:
                if not isinstance(r, dict):
                    continue
                item = _safe_str(r.get("itemName")) or "-"
                try:
                    nw = float(r.get("netWeight") or 0)
                except Exception:
                    nw = 0.0
                total_net += nw
                _draw_input_row(item, _fmt_job_slip_weight(nw))
        else:
            i_name = _safe_str(data.get("givingItemName")) or "Gold Bar"
            try:
                gg = float(data.get("givingGross") or 0)
            except Exception:
                gg = 0.0
            try:
                st = float(data.get("stone") or 0)
            except Exception:
                st = 0.0
            try:
                eh = float(data.get("eh") or 0)
            except Exception:
                eh = 0.0
            _draw_input_row(i_name, _fmt_job_slip_weight(gg))
            if st > 0:
                _draw_input_row("Stone", _fmt_job_slip_weight(st))
            if eh > 0:
                _draw_input_row("EH/Enamel", _fmt_job_slip_weight(eh))
            _draw_input_row("Net", _fmt_job_slip_weight(giving_net))
            total_net = giving_net

        lh_tot_l = _font_line_h(F_TXT_B)
        lh_tot_w = _font_line_h(F_JC_WEIGHT_B)
        rht = max(lh_tot_l, lh_tot_w)
        _draw_text(draw, box_x0 + in_pad_x, y + (rht - lh_tot_l) // 2, "Total weight:", F_TXT_B)
        _draw_text_right(draw, box_x1 - in_pad_x, y + (rht - lh_tot_w) // 2, _fmt_job_slip_weight(total_net), F_JC_WEIGHT_B)
        y += rht + 10

    # FINISHED ITEM
    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 10
    t_fi = "FINISHED ITEM"
    twf = _text_w(draw, t_fi, F_TXT_B)
    _draw_text_center(draw, box_x0 + box_w / 2, y, t_fi, F_TXT_B)
    y += _font_line_h(F_TXT_B) + 4
    draw.line((box_x0 + box_w / 2 - twf / 2, y, box_x0 + box_w / 2 + twf / 2, y), fill=0, width=1)
    y += 12

    # FINISHED ITEM: left-to-right — Unit weight | Item name | Qty | Process (wrap item + process; small F_TXT)
    x_l = box_x0 + in_pad_x
    x_r = box_x1 - in_pad_x
    w_wt = 96
    w_qty = 28
    gap_c = 8
    # [wt][item][qty][proc] with gaps between each
    inner = (x_r - x_l) - w_wt - w_qty - 3 * gap_c
    w_item = max(60, int(inner * 0.55))
    w_proc = inner - w_item
    if w_proc < 60:
        w_proc = 60
        w_item = max(60, inner - w_proc)
    x_item = x_l + w_wt + gap_c
    x_qty = x_item + w_item + gap_c
    x_proc = x_qty + w_qty + gap_c
    w_proc_px = x_r - x_proc
    lh_fi = _font_line_h(F_TXT)
    line_step = lh_fi + 2
    # Hide process column when it repeats the previous row (same string as line above)
    _prev_finished_proc = None

    for r in finished_rows:
        if not isinstance(r, dict):
            continue
        unit_wt = _fmt_job_slip_weight(r.get("unitWeight"))
        try:
            q = int(r.get("quantity") or 1)
        except Exception:
            q = 1
        q_disp = str(q) if q <= 99 else "99"
        nm = _safe_str(r.get("itemName")) or "-"
        row_proc = _safe_str(r.get("processName")) if isinstance(r, dict) else ""
        proc_text = row_proc if row_proc else process_name_default
        item_lines = _wrap_line_px(nm, draw, F_TXT, w_item)
        if _prev_finished_proc is not None and proc_text == _prev_finished_proc:
            proc_lines = []
        else:
            proc_lines = _wrap_line_px(proc_text, draw, F_TXT, w_proc_px)
            if not proc_lines:
                proc_lines = ["-"]
            _prev_finished_proc = proc_text
        if not item_lines:
            item_lines = ["-"]
        n_lines = max(len(item_lines), len(proc_lines), 1)

        row_top = y
        for i in range(n_lines):
            yy = row_top + i * line_step
            if i == 0:
                _draw_text_right(draw, x_l + w_wt, yy, unit_wt, F_TXT)
                _draw_text(draw, x_qty, yy, q_disp, F_TXT)
            if i < len(item_lines):
                _draw_text(draw, x_item, yy, item_lines[i], F_TXT)
            if i < len(proc_lines):
                _draw_text(draw, x_proc, yy, proc_lines[i], F_TXT)
        y = row_top + n_lines * line_step + 8

    draw.line((box_x0, y, box_x1, y), fill=0, width=1)
    y += 8
    box_y1 = y + 8
    draw.rectangle((box_x0, box_y0, box_x1, box_y1), outline=0, width=1)

    img = img.crop((0, 0, W, min(H, box_y1 + 14)))
    return _to_1bit_floyd_steinberg(img)


# ==========================
# SLIPS PRINTING ENDPOINTS
# ==========================

@app.route("/print-quick-weight-receipt-raster", methods=["POST"])
def print_quick_weight_receipt_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_quick_weight_receipt_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({
            "ok": True,
            "widthDots": img1.size[0],
            "heightDots": img1.size[1]
        })
    except Exception as e:
        log(f"/print-quick-weight-receipt-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/print-receiving-metal-slip-raster", methods=["POST"])
def print_receiving_metal_slip_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_receiving_metal_slip_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({"ok": True, "widthDots": img1.size[0], "heightDots": img1.size[1]})
    except Exception as e:
        log(f"/print-receiving-metal-slip-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/print-job-create-slip-raster", methods=["POST"])
def print_job_create_slip_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_job_create_slip_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({"ok": True, "widthDots": img1.size[0], "heightDots": img1.size[1]})
    except Exception as e:
        log(f"/print-job-create-slip-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/print-provisional-delivery-slip-raster", methods=["POST"])
def print_provisional_delivery_slip_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_provisional_delivery_slip_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({"ok": True, "widthDots": img1.size[0], "heightDots": img1.size[1]})
    except Exception as e:
        log(f"/print-provisional-delivery-slip-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/print-packing-delivery-slip-raster", methods=["POST"])
def print_packing_delivery_slip_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_packing_delivery_slip_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({"ok": True, "widthDots": img1.size[0], "heightDots": img1.size[1]})
    except Exception as e:
        log(f"/print-packing-delivery-slip-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/print-multirow-packing-list-raster", methods=["POST"])
def print_multirow_packing_list_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_multirow_packing_list_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({
            "ok": True,
            "widthDots": img1.size[0],
            "heightDots": img1.size[1]
        })
    except Exception as e:
        log(f"/print-multirow-packing-list-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/print-packing-list-raster", methods=["POST"])
def print_packing_list_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_packing_list_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({"ok": True, "widthDots": img1.size[0], "heightDots": img1.size[1]})
    except Exception as e:
        log(f"/print-packing-list-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500
    
@app.route("/print-rate-slip-raster", methods=["POST"])
def print_rate_slip_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_rate_slip_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({"ok": True, "widthDots": img1.size[0], "heightDots": img1.size[1]})
    except Exception as e:
        log(f"/print-rate-slip-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/print-balance-summary-raster", methods=["POST"])
def print_balance_summary_raster():
    data = request.get_json(force=True, silent=True) or {}
    printer = data.get("printer")
    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400

    try:
        img1 = _render_balance_summary_image_80mm(data)
        parts = [b"\x1b@", _img_to_escpos_raster(img1), b"\r\n"]

        feed_lines = int(data.get("feedLines", 3))
        cut = bool(data.get("cut", True))
        cut_mode = str(data.get("cutMode", "full")).lower()

        if cut:
            parts.append(_esc_feed(feed_lines))
            parts.append(_esc_cut(cut_mode))

        _write_raw(printer, b"".join(parts))
        return jsonify({"ok": True, "widthDots": img1.size[0], "heightDots": img1.size[1]})
    except Exception as e:
        log(f"/print-balance-summary-raster error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500



# --------------------------- scale (METTLER TOLEDO / USB-Serial) ---------------------------

import time, re
from typing import Optional
import serial
from serial.tools import list_ports

SCALE_LOCK = threading.Lock()
SCALE_STATE = {
    "connected": False,
    "port": None,
    "baud": None,
    "mode": None,
    "last_raw": "",
    "last_value": None,
    "last_unit": "",
    "last_ts": None,
    "error": "scale not configured",
}

# simple pubsub for SSE
_SCALE_SUBSCRIBERS: list["queue.Queue[str]"] = []
try:
    import queue
except Exception:
    queue = None  # should not happen

# Runtime-only config: GoldTrackerPro is the source of truth
RUNTIME_SCALE_CFG = {
    "port": None,
    "baud": None,
    "mode": None,
    "poll_ms": None,
    "cmd": None,
}


def _clean_str_or_none(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def _clean_int_or_none(v):
    if v in (None, "", "null", "None"):
        return None
    try:
        return int(v)
    except Exception:
        return None


def _scale_cfg():
    """
    Runtime-only scale config.
    Do not cast None to int/str.
    """
    port = _clean_str_or_none(RUNTIME_SCALE_CFG.get("port"))
    baud = _clean_int_or_none(RUNTIME_SCALE_CFG.get("baud"))
    mode = _clean_str_or_none(RUNTIME_SCALE_CFG.get("mode"))
    poll_ms = _clean_int_or_none(RUNTIME_SCALE_CFG.get("poll_ms"))
    cmd = RUNTIME_SCALE_CFG.get("cmd")

    if cmd is None:
        cmd = None
    else:
        cmd = str(cmd)

    return {
        "port": port,
        "baud": baud,
        "mode": mode.lower() if mode else None,
        "poll_ms": poll_ms,
        "cmd": cmd,
        "configured": bool(port and baud and mode),
    }


def _validate_scale_cfg(cfg: dict):
    errors = []

    port = cfg.get("port")
    baud = cfg.get("baud")
    mode = cfg.get("mode")
    poll_ms = cfg.get("poll_ms")

    if not port:
        errors.append("port missing")

    if baud is None:
        errors.append("baud missing")
    elif not isinstance(baud, int) or baud <= 0:
        errors.append("baud invalid")

    if not mode:
        errors.append("mode missing")
    elif mode not in ("auto", "poll", "stream"):
        errors.append("mode invalid")

    if poll_ms is not None:
        if not isinstance(poll_ms, int) or poll_ms < 50 or poll_ms > 5000:
            errors.append("poll_ms invalid")

    return errors


def _scale_publish(payload_json: str):
    # best-effort broadcast (SSE)
    for q in list(_SCALE_SUBSCRIBERS):
        try:
            q.put_nowait(payload_json)
        except Exception:
            pass


def _parse_weight(line: str):
    """
    Generic parser:
    - extracts first number with optional decimals
    - tries to detect unit token (g, kg, etc.)
    """
    s = (line or "").strip()
    if not s:
        return None, ""

    unit = ""
    for u in ("g", "kg", "lb", "oz", "ct"):
        if re.search(rf"\b{re.escape(u)}\b", s, flags=re.IGNORECASE):
            unit = u
            break

    m = re.search(r"([-+]?\d+(?:\.\d+)?)", s)
    if not m:
        return None, unit

    try:
        val = float(m.group(1))
        return val, unit
    except Exception:
        return None, unit


class ScaleReader:
    def __init__(self):
        self._t: Optional[threading.Thread] = None
        self._stop = threading.Event()

    def start(self):
        if self._t and self._t.is_alive():
            return
        self._stop = threading.Event()
        self._t = threading.Thread(target=self._run, daemon=True)
        self._t.start()

    def stop(self, wait: float = 2.0):
        self._stop.set()
        t = self._t
        if t and t.is_alive():
            t.join(timeout=wait)

    def _set_state(self, **kwargs):
        with SCALE_LOCK:
            SCALE_STATE.update(kwargs)

    def _run(self):
        while not self._stop.is_set():
            cfg = _scale_cfg()

            if not cfg.get("configured"):
                self._set_state(
                    connected=False,
                    port=cfg.get("port"),
                    baud=cfg.get("baud"),
                    mode=cfg.get("mode"),
                    error="scale not configured",
                )
                time.sleep(0.2)
                continue

            port = cfg["port"]
            baud = cfg["baud"]
            mode = cfg["mode"]
            poll_ms = cfg["poll_ms"] if cfg["poll_ms"] is not None else 200
            poll_ms = max(50, min(5000, poll_ms))
            cmd = (cfg["cmd"] or "SI\r\n").encode("ascii", errors="ignore")

            self._set_state(
                connected=False,
                port=port,
                baud=baud,
                mode=mode,
                error="",
            )

            try:
                ser = serial.Serial(
                    port=port,
                    baudrate=baud,
                    timeout=0.05,
                    write_timeout=0.2,
                    bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                )
            except Exception as e:
                self._set_state(
                    connected=False,
                    port=port,
                    baud=baud,
                    mode=mode,
                    error=f"open failed: {e}",
                )
                time.sleep(0.2)
                continue

            self._set_state(connected=True, error="")
            log(f"[SCALE] Connected {port} @ {baud} mode={mode}")

            last_poll = 0.0
            last_rx = time.time()

            try:
                time.sleep(0.05)
                try:
                    ser.reset_input_buffer()
                except Exception:
                    pass

                while not self._stop.is_set():
                    try:
                        raw = ser.readline()
                    except Exception as e:
                        raise RuntimeError(f"read failed: {e}")

                    if raw:
                        last_rx = time.time()
                        line = raw.decode("utf-8", errors="ignore").strip()
                        val, unit = _parse_weight(line)

                        with SCALE_LOCK:
                            SCALE_STATE["last_raw"] = line
                            SCALE_STATE["last_ts"] = datetime.now().isoformat(timespec="seconds")
                            if val is not None:
                                SCALE_STATE["last_value"] = val
                            if unit:
                                SCALE_STATE["last_unit"] = unit

                            payload = json.dumps({
                                "ok": True,
                                "value": SCALE_STATE["last_value"],
                                "unit": SCALE_STATE["last_unit"],
                                "raw": SCALE_STATE["last_raw"],
                                "ts": SCALE_STATE["last_ts"],
                                "connected": SCALE_STATE["connected"],
                                "error": SCALE_STATE["error"],
                            })

                        _scale_publish(payload)

                    now = time.time()

                    if mode == "poll":
                        if (now - last_poll) * 1000.0 >= poll_ms:
                            try:
                                ser.write(cmd)
                                last_poll = now
                            except Exception as e:
                                raise RuntimeError(f"write failed: {e}")

                    elif mode == "auto":
                        # Read stream if available; otherwise fallback to polling quickly
                        if (now - last_poll) * 1000.0 >= poll_ms:
                            if (now - last_rx) > 0.25:
                                try:
                                    ser.write(cmd)
                                    last_poll = now
                                except Exception as e:
                                    raise RuntimeError(f"write failed: {e}")

                    elif mode == "stream":
                        # No polling; only read incoming data
                        pass

                    time.sleep(0.005)

            except Exception as e:
                self._set_state(connected=False, error=str(e))
                log(f"[SCALE] error: {e}")
            finally:
                try:
                    ser.close()
                except Exception:
                    pass

            time.sleep(0.1)

SCALE_READER = ScaleReader()
SCALE_READER.start()


def _auth_or_query_key_ok() -> bool:
    # For SSE/EventSource cases where headers are hard, allow ?key=... too
    k = request.headers.get("X-Print-Key", "")
    if k and k == API_KEY:
        return True
    qk = request.args.get("key", "")
    return bool(qk) and qk == API_KEY


@app.route("/scales/", methods=["GET"])
def list_scales_like_devices():
    ports = list_ports.comports()
    out = []
    for p in ports:
        out.append({
            "device": p.device,
            "name": p.name,
            "description": p.description,
            "hwid": p.hwid,
            "vid": p.vid,
            "pid": p.pid,
            "serial_number": p.serial_number,
            "manufacturer": p.manufacturer,
            "product": p.product,
            "location": p.location,
        })
    return jsonify({
        "ok": True,
        "devices": out,
    })


@app.route("/scale/status", methods=["GET"])
def scale_status():
    ports = []
    try:
        ports = [p.device for p in list_ports.comports()]
    except Exception:
        pass

    with SCALE_LOCK:
        st = dict(SCALE_STATE)

    return jsonify({
        "ok": True,
        "ports": ports,
        "state": st,      # actual live state
        "config": _scale_cfg(),  # selected runtime config from GoldTrackerPro
    })


@app.route("/scale/latest", methods=["GET"])
def scale_latest():
    if not _auth_or_query_key_ok():
        return jsonify({"ok": False, "error": "unauthorized"}), 401

    with SCALE_LOCK:
        st = dict(SCALE_STATE)

    return jsonify({
        "ok": True,
        "value": st.get("last_value"),
        "unit": st.get("last_unit"),
        "raw": st.get("last_raw"),
        "ts": st.get("last_ts"),
        "connected": st.get("connected"),
        "error": st.get("error"),
    })


@app.route("/scale/config", methods=["POST"])
def scale_config():
    data = request.get_json(force=True, silent=True) or {}

    new_cfg = {
        "port": _clean_str_or_none(data.get("port")) if "port" in data else RUNTIME_SCALE_CFG.get("port"),
        "baud": _clean_int_or_none(data.get("baud")) if "baud" in data else RUNTIME_SCALE_CFG.get("baud"),
        "mode": (_clean_str_or_none(data.get("mode")).lower() if _clean_str_or_none(data.get("mode")) else None) if "mode" in data else RUNTIME_SCALE_CFG.get("mode"),
        "poll_ms": _clean_int_or_none(data.get("poll_ms")) if "poll_ms" in data else RUNTIME_SCALE_CFG.get("poll_ms"),
        "cmd": (str(data.get("cmd")) if data.get("cmd") is not None else None) if "cmd" in data else RUNTIME_SCALE_CFG.get("cmd"),
    }

    cfg_for_validation = {
        "port": _clean_str_or_none(new_cfg["port"]),
        "baud": _clean_int_or_none(new_cfg["baud"]),
        "mode": _clean_str_or_none(new_cfg["mode"]),
        "poll_ms": _clean_int_or_none(new_cfg["poll_ms"]),
        "cmd": new_cfg["cmd"],
    }

    errors = _validate_scale_cfg(cfg_for_validation)
    if errors:
        return jsonify({
            "ok": False,
            "error": "invalid scale config",
            "details": errors,
            "config": {
                "port": cfg_for_validation["port"],
                "baud": cfg_for_validation["baud"],
                "mode": cfg_for_validation["mode"],
                "poll_ms": cfg_for_validation["poll_ms"],
                "cmd": cfg_for_validation["cmd"],
                "configured": False,
            }
        }), 400

    RUNTIME_SCALE_CFG["port"] = cfg_for_validation["port"]
    RUNTIME_SCALE_CFG["baud"] = cfg_for_validation["baud"]
    RUNTIME_SCALE_CFG["mode"] = cfg_for_validation["mode"]
    RUNTIME_SCALE_CFG["poll_ms"] = cfg_for_validation["poll_ms"] if cfg_for_validation["poll_ms"] is not None else 200
    RUNTIME_SCALE_CFG["cmd"] = cfg_for_validation["cmd"] if cfg_for_validation["cmd"] is not None else "SI\r\n"

    # Runtime only. Intentionally do NOT persist scale config into config.json.
    SCALE_READER.stop(wait=2.0)
    SCALE_READER.start()

    return jsonify({
        "ok": True,
        "config": _scale_cfg(),
        "source": "runtime"
    })

@app.route("/scale/test", methods=["POST"])
def scale_test():
    """
    Test a port+baud without disrupting the live ScaleReader.
    - If ScaleReader is already connected to the same port+baud → return live state immediately.
    - Otherwise → stop ScaleReader, open port, send poll cmd, read, close, restart ScaleReader.
    """
    data = request.get_json(force=True, silent=True) or {}
    port = _clean_str_or_none(data.get("port"))
    baud = _clean_int_or_none(data.get("baud")) or 9600
    cmd  = str(data.get("cmd") or "SI\r\n")
    timeout_s = min(float(data.get("timeout", 3.0)), 10.0)

    if not port:
        return jsonify({"ok": False, "error": "port required"}), 400

    # If ScaleReader is already live on the exact port+baud, return live state — no port fight.
    with SCALE_LOCK:
        st  = dict(SCALE_STATE)
        cfg = _scale_cfg()

    if (st.get("connected")
            and st.get("port") == port
            and cfg.get("baud") == baud
            and st.get("last_value") is not None):
        return jsonify({
            "ok":    True,
            "port":  port,
            "baud":  baud,
            "raw":   st.get("last_raw", ""),
            "value": st.get("last_value"),
            "unit":  st.get("last_unit") or "g",
            "source": "live",
        })

    # Different port or baud — pause ScaleReader so we can open the port exclusively.
    SCALE_READER.stop(wait=2.0)
    ser    = None
    result = None
    try:
        ser = serial.Serial(port, baud, timeout=0.5, write_timeout=0.5)
        time.sleep(0.3)
        try:
            ser.reset_input_buffer()
        except Exception:
            pass
        ser.write(cmd.encode("ascii", errors="ignore"))

        deadline = time.time() + timeout_s
        while time.time() < deadline:
            line_raw = ser.readline()
            if line_raw:
                line = line_raw.decode("utf-8", errors="ignore").strip()
                if line:
                    val, unit = _parse_weight(line)
                    result = {"ok": True, "port": port, "baud": baud,
                              "raw": line, "value": val, "unit": unit or "g", "source": "test"}
                    break

        if not result:
            result = {"ok": False, "error": "No data received from scale", "port": port, "baud": baud}

    except serial.SerialException as e:
        err = str(e)
        if "Access is denied" in err or "PermissionError" in err:
            result = {"ok": False, "error": f"Port {port} is in use by another process"}
        else:
            result = {"ok": False, "error": f"Cannot open port: {err}"}
    except Exception as e:
        result = {"ok": False, "error": str(e)}
    finally:
        if ser:
            try:
                ser.close()
            except Exception:
                pass
        SCALE_READER.start()  # always restart after test

    return jsonify(result)


@app.route("/scale/autobaud", methods=["POST"])
def scale_autobaud():
    """
    Try each baud rate in sequence via the same stop+test+restart approach.
    Returns on first baud that yields a valid scale reading.
    """
    data            = request.get_json(force=True, silent=True) or {}
    port            = _clean_str_or_none(data.get("port"))
    baud_candidates = data.get("baud_candidates") or [9600, 4800, 19200, 38400, 115200]
    cmd             = str(data.get("cmd") or "SI\r\n")

    if not port:
        return jsonify({"ok": False, "error": "port required"}), 400

    SCALE_READER.stop(wait=2.0)
    found = None

    for baud in baud_candidates:
        ser = None
        try:
            ser = serial.Serial(port, baud, timeout=0.5, write_timeout=0.5)
            time.sleep(0.2)
            try:
                ser.reset_input_buffer()
            except Exception:
                pass
            ser.write(cmd.encode("ascii", errors="ignore"))

            deadline = time.time() + 1.5   # 1.5 s per candidate
            while time.time() < deadline:
                line_raw = ser.readline()
                if line_raw:
                    line = line_raw.decode("utf-8", errors="ignore").strip()
                    if line:
                        val, unit = _parse_weight(line)
                        found = {"ok": True, "port": port, "baud": baud,
                                 "raw": line, "value": val, "unit": unit or "g"}
                        break
        except Exception:
            pass
        finally:
            if ser:
                try:
                    ser.close()
                except Exception:
                    pass

        if found:
            break

    SCALE_READER.start()  # always restart after detection

    return jsonify(found if found else {"ok": False, "error": "No readable data on any baud rate"})


@app.route("/scale/clear", methods=["POST"])
def scale_clear():
    RUNTIME_SCALE_CFG["port"] = None
    RUNTIME_SCALE_CFG["baud"] = None
    RUNTIME_SCALE_CFG["mode"] = None
    RUNTIME_SCALE_CFG["poll_ms"] = None
    RUNTIME_SCALE_CFG["cmd"] = None

    SCALE_READER.stop(wait=2.0)
    SCALE_READER.start()

    with SCALE_LOCK:
        SCALE_STATE["connected"] = False
        SCALE_STATE["port"] = None
        SCALE_STATE["baud"] = None
        SCALE_STATE["mode"] = None
        SCALE_STATE["error"] = "scale not configured"

    return jsonify({
        "ok": True,
        "config": _scale_cfg(),
        "source": "runtime",
        "cleared": True
    })


@app.route("/scale/stream", methods=["GET"])
def scale_stream():
    if queue is None:
        return jsonify({"ok": False, "error": "queue module unavailable"}), 500

    def gen():
        q = queue.Queue(maxsize=50)
        _SCALE_SUBSCRIBERS.append(q)

        with SCALE_LOCK:
            snap = json.dumps({
                "ok": True,
                "value": SCALE_STATE.get("last_value"),
                "unit": SCALE_STATE.get("last_unit"),
                "raw": SCALE_STATE.get("last_raw"),
                "ts": SCALE_STATE.get("last_ts"),
                "connected": SCALE_STATE.get("connected"),
                "error": SCALE_STATE.get("error"),
            })
        yield f"data: {snap}\n\n"

        last_keepalive = time.time()

        try:
            while True:
                try:
                    msg = q.get(timeout=1.0)
                    yield f"data: {msg}\n\n"
                except Exception:
                    if time.time() - last_keepalive > 15:
                        yield "event: ping\ndata: {}\n\n"
                        last_keepalive = time.time()
        finally:
            try:
                _SCALE_SUBSCRIBERS.remove(q)
            except Exception:
                pass

    resp = app.response_class(gen(), mimetype="text/event-stream")
    resp.headers["Cache-Control"] = "no-cache"
    resp.headers["X-Accel-Buffering"] = "no"
    return resp

# --------------------------- METTLER TOLEDO SLIP ENDS ---------------------------

from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame,
    Table, TableStyle, Paragraph, Spacer, FrameBreak
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def _ensure_reportlab_fonts():
    try:
        pdfmetrics.registerFont(TTFont("Arial", r"C:\Windows\Fonts\arial.ttf"))
        pdfmetrics.registerFont(TTFont("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"))
        return "Arial", "Arial-Bold"
    except Exception:
        return "Helvetica", "Helvetica-Bold"


def _safe_str(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return str(v)
    return str(v).strip()


def _join_addr(*parts):
    parts = [_safe_str(p) for p in parts if _safe_str(p)]
    return "\n".join(parts)

def _load_bg_image(bg_source: str):
    """
    Supports:
    - local absolute file path
    - http/https URL
    Returns ImageReader or None
    """
    src = _safe_str(bg_source)
    if not src:
        return None

    try:
        if src.lower().startswith(("http://", "https://")):
            with urlopen(src, timeout=15) as resp:
                raw = resp.read()
            return ImageReader(io.BytesIO(raw))

        p = Path(src)
        if p.exists():
            return ImageReader(str(p))

    except Exception as e:
        log(f"background image load failed: {e}")

    return None


def _render_delivery_challan_pdf_reportlab(
    data: dict,
    output_pdf_path: Path,
    two_per_page: bool = False,  # True => two challans on same A4 page (top/bottom)
) -> Path:
    base_font, bold_font = _ensure_reportlab_fonts()
    styles = getSampleStyleSheet()

    # Current font sizes:
    # title: 12, normal: 9, small: 8.5
    s_normal = ParagraphStyle("n", parent=styles["Normal"], fontName=base_font, fontSize=9, leading=10)
    s_bold   = ParagraphStyle("b", parent=styles["Normal"], fontName=bold_font, fontSize=9, leading=10)
    s_title  = ParagraphStyle("t", parent=styles["Normal"], fontName=bold_font, fontSize=12, leading=14, alignment=1)

    s_small  = ParagraphStyle("s", parent=styles["Normal"], fontName=base_font, fontSize=8.5, leading=9.5)
    s_small_b = ParagraphStyle("sb", parent=styles["Normal"], fontName=bold_font, fontSize=8.5, leading=9.5)
    s_small_b_center = ParagraphStyle("sbc", parent=s_small_b, alignment=1)

    def P(txt, style=s_normal):
        return Paragraph((_safe_str(txt)).replace("\n", "<br/>"), style)

    output_pdf_path.parent.mkdir(parents=True, exist_ok=True)

    doc = BaseDocTemplate(
        str(output_pdf_path),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
    )

    total_w = doc.width

    # Keep own LEFT and party RIGHT, but add more spacing between blocks by widening the gutter.
    # (Instead of 50/50 split, use a larger center gap)
    gutter = 10 * mm  # ✅ increase space between left & right blocks (UI/UX)
    left_w = (total_w - gutter) * 0.50
    right_w = (total_w - gutter) - left_w

    # Columns for remarks/signatures (keep as before proportions but based on left/right)
    sig_left_w = left_w + (right_w * 0.35)
    sig_right_w = right_w * 0.65

    bg_source = (
        _safe_str(data.get("backgroundImageUrl")) or
        _safe_str(data.get("backgroundImagePath")) or
        _safe_str(data.get("logo")) or
        _safe_str(data.get("backgroundImage"))
    )
    bg_image = _load_bg_image(bg_source)

    def _draw_page_bg(canvas_obj, doc_obj):
        if not bg_image:
            return

        try:
            page_w, page_h = doc.pagesize
            canvas_obj.saveState()

            try:
                canvas_obj.setFillAlpha(0.30)
            except Exception:
                pass

            def draw_bg_in_area(area_x, area_y, area_w, area_h):
                img_w = min(area_w * 0.42, 120 * mm)
                img_h = min(area_h * 0.42, 120 * mm)

                x = area_x + ((area_w - img_w) / 2)
                y = area_y + ((area_h - img_h) / 2)

                canvas_obj.drawImage(
                    bg_image,
                    x,
                    y,
                    width=img_w,
                    height=img_h,
                    preserveAspectRatio=True,
                    mask='auto'
                )

            if two_per_page:
                gap = 1 * mm
                half_h = (page_h - doc.topMargin - doc.bottomMargin - gap) / 2.0

                content_x = doc.leftMargin
                content_w = page_w - doc.leftMargin - doc.rightMargin

                bottom_y = doc.bottomMargin
                top_y = doc.bottomMargin + half_h + gap

                draw_bg_in_area(content_x, top_y, content_w, half_h)
                draw_bg_in_area(content_x, bottom_y, content_w, half_h)
            else:
                content_x = doc.leftMargin
                content_y = doc.bottomMargin
                content_w = page_w - doc.leftMargin - doc.rightMargin
                content_h = page_h - doc.topMargin - doc.bottomMargin

                draw_bg_in_area(content_x, content_y, content_w, content_h)

            canvas_obj.restoreState()

        except Exception as e:
            log(f"background draw failed: {e}")

    def _get_additional_infos():
        # Support both flat keys and nested object
        add = data.get("additionalInfos") or data.get("additionalInfo") or {}
        if not isinstance(add, dict):
            add = {}

        transport_mode = _safe_str(add.get("transportMode") or data.get("transportMode"))
        transporter_name = _safe_str(add.get("transporterName") or data.get("transporterName"))
        vehicle_no = _safe_str(add.get("vehicleNo") or data.get("vehicleNo") or add.get("vehicleNumber") or data.get("vehicleNumber"))

        return transport_mode, transporter_name, vehicle_no

    def build_one_challan_flowables():
        story = []

        # TITLE
        title_tbl = Table([[P("DELIVERY CHALLAN", s_title)]], colWidths=[total_w])
        title_tbl.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        story.append(title_tbl)
        story.append(Spacer(1, 2 * mm))

        # DATA
        date_str = _safe_str(data.get("date"))
        dc_no = _safe_str(data.get("dcNo"))

        own = data.get("own") or {}
        party = data.get("party") or {}

        own_name = _safe_str(own.get("name") or data.get("ownName") or "KCK JEWELL WORKS")
        own_addr = _join_addr(
            own.get("address1") or data.get("ownAddress1"),
            own.get("address2") or data.get("ownAddress2"),
            own.get("address3") or data.get("ownAddress3"),
            data.get("ownAddress"),
        )
        own_gst = _safe_str(own.get("gst") or data.get("ownGstNo"))
        own_email = _safe_str(own.get("email") or data.get("ownEmail"))
        own_phone = _safe_str(own.get("phone") or data.get("ownPhone"))

        party_heading = "Details of Receiver (Billed to):"
        party_name = _safe_str(party.get("name") or data.get("partyName"))
        party_addr = _join_addr(
            party.get("address1") or data.get("partyAddress1"),
            party.get("address2") or data.get("partyAddress2"),
            party.get("address3") or data.get("partyAddress3"),
            data.get("partyAddress"),
        )
        party_gst = _safe_str(party.get("gst") or data.get("partyGstNo"))
        party_email = _safe_str(party.get("email") or data.get("partyEmail"))
        party_phone = _safe_str(party.get("phone") or data.get("partyPhone"))

        # LEFT block (Date + Own)  ✅ keep OWN on LEFT
        left_block = []
        if date_str:
            left_block.append(P(f"<b>Date :</b>&nbsp;&nbsp;{date_str}", s_small))
            left_block.append(Spacer(1, 2 * mm))  # ✅ space below Date
        left_block.append(P(f"<b>{own_name}</b>", s_bold))
        left_block.append(P(own_addr, s_normal))
        if own_gst:
            left_block.append(P(f"GST No: {own_gst}", s_normal))
        if own_email:
            left_block.append(P(f"Email id: {own_email}", s_normal))
        if own_phone:
            left_block.append(P(f"Phone : {own_phone}", s_normal))

        # RIGHT block (DC No + Party) ✅ keep PARTY on RIGHT
        right_block = []
        if dc_no:
            right_block.append(P(f"<b>DC No:</b>&nbsp;&nbsp;{dc_no}", s_small))
            right_block.append(Spacer(1, 2 * mm))  # ✅ space below DC No
        right_block.append(P(f"<b><u>{party_heading}</u></b>", s_bold))
        right_block.append(P(party_name, s_normal))
        right_block.append(P(party_addr, s_normal))
        if party_gst:
            right_block.append(P(f"GST No: {party_gst}", s_normal))
        if party_email:
            right_block.append(P(f"Email id: {party_email}", s_normal))
        if party_phone:
            right_block.append(P(f"Phone : {party_phone}", s_normal))

        # Tables for blocks
        left_tbl = Table([[x] for x in left_block], colWidths=[left_w])
        left_tbl.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))

        right_tbl = Table([[x] for x in right_block], colWidths=[right_w])
        right_tbl.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),  # party block reads better left-aligned
        ]))

        # ✅ add gutter column in middle to increase spacing (no swap!)
        top_tbl = Table([[left_tbl, "", right_tbl]], colWidths=[left_w, gutter, right_w])
        top_tbl.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        story.append(top_tbl)
        story.append(Spacer(1, 3 * mm))

        # ITEMS TABLE (scale to exact total_w)
        rows = data.get("rows") or []
        if not isinstance(rows, list):
            rows = []

        base_cols = [30, 90, 25, 25, 20]
        scale = float(total_w) / (sum(base_cols) * mm)
        colWidths = [(c * mm) * scale for c in base_cols]

        table_data = [[
            P("HSN/SAC Code", s_small_b_center),
            P("DESCRIPTION", s_small_b_center),
            P("GROSS WEIGHT<br/>(gms)", s_small_b_center),
            P("NET WEIGHT<br/>(gms)", s_small_b_center),
            P("QUANTITY<br/>(pcs)", s_small_b_center),
        ]]

        for r in rows:
            if not isinstance(r, dict):
                continue
            hsn = _safe_str(r.get("hsnSacCode") or r.get("hsn"))
            desc = _safe_str(r.get("description") or r.get("desc"))
            gross = _safe_str(r.get("grossWeight") or r.get("grossWt"))
            net = _safe_str(r.get("netWeight") or r.get("netWt"))
            qty = _safe_str(r.get("quantity") or r.get("qty"))
            table_data.append([
                P(hsn, s_small),
                P(desc, s_small),
                P(gross, s_small),
                P(net, s_small),
                P(qty, s_small),
            ])

        items_tbl = Table(table_data, colWidths=colWidths, repeatRows=1)
        items_tbl.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("ALIGN", (2, 1), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(items_tbl)
        story.append(Spacer(1, 2 * mm))

        # ✅ Additional Infos line (below table, single line)
        transport_mode, transporter_name, vehicle_no = _get_additional_infos()
        if transport_mode or transporter_name or vehicle_no:
            parts = []
            if transport_mode:
                parts.append(f"<b>Transport Mode:</b> {transport_mode}")
            if transporter_name:
                parts.append(f"<b>Transporter Name:</b> {transporter_name}")
            if vehicle_no:
                parts.append(f"<b>Vehicle No:</b> {vehicle_no}")

            add_line = " &nbsp;&nbsp;&nbsp;&nbsp; ".join(parts)
            add_tbl = Table([[P(add_line, s_small)]], colWidths=[total_w])
            add_tbl.setStyle(TableStyle([
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]))
            story.append(add_tbl)
            story.append(Spacer(1, 4 * mm))
        else:
            story.append(Spacer(1, 6 * mm))  # keep consistent spacing if no additional infos

        # REMARKS + SIGNATURES
        remarks = _safe_str(data.get("remarks"))
        packs = _safe_str(data.get("noOfPacks"))

        rem_tbl = Table(
            [[P(f"Remarks : {remarks}", s_small), P(f"No. of Packs: {packs}", s_small)]],
            colWidths=[sig_left_w, sig_right_w],
        )
        rem_tbl.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        story.append(rem_tbl)

        sig_tbl = Table(
            [[P("<b><u>Signature of Recipient</u></b>", s_small_b),
              P("<b><u>Signature of Supplier</u></b>", s_small_b)]],
            colWidths=[sig_left_w, sig_right_w],
        )
        sig_tbl.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(sig_tbl)

        # ✅ more space ABOVE Authorised Signatory
        story.append(Spacer(1, 16 * mm))

        auth_tbl = Table(
            [[P("Authorised Signatory", s_small), P("Authorised Signatory", s_small)]],
            colWidths=[sig_left_w, sig_right_w],
        )
        auth_tbl.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("ALIGN", (0, 0), (0, 0), "LEFT"),
            ("ALIGN", (1, 0), (1, 0), "LEFT"),
        ]))
        story.append(auth_tbl)
        story.append(Spacer(1, 2 * mm))

        # NOTES
        notes_lines = data.get("notes") or []
        if isinstance(notes_lines, str):
            notes_lines = [notes_lines]
        if not isinstance(notes_lines, list):
            notes_lines = []

        story.append(P("<b>Notes:</b>", s_small_b))
        story.append(Spacer(1, 1 * mm))

        numbered = []
        for i, n in enumerate(notes_lines, start=1):
            n = _safe_str(n)
            if n:
                numbered.append(f"{i}. {n}")
        if not numbered:
            numbered = ["1."]

        notes_para = P("<br/>".join(numbered), s_small)
        notes_tbl = Table([[notes_para]], colWidths=[total_w])
        notes_tbl.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 1, colors.black),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(notes_tbl)

        return story

    if not two_per_page:
        frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="single")
        doc.addPageTemplates([
            PageTemplate(id="p1", frames=[frame], onPage=_draw_page_bg)
        ])
        story = build_one_challan_flowables()
        doc.build(story)
        return output_pdf_path

    # Two challans on same A4 page (Top/Bottom)
    gap = 1 * mm  # ✅ reduced gap between 2 challans (was 6mm)
    half_h = (doc.height - gap) / 2.0

    bottom_frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, half_h, id="bottom")
    top_frame = Frame(doc.leftMargin, doc.bottomMargin + half_h + gap, doc.width, half_h, id="top")

    doc.addPageTemplates([
        PageTemplate(id="two_up", frames=[top_frame, bottom_frame], onPage=_draw_page_bg)
    ])

    story = []
    story.extend(build_one_challan_flowables())
    story.append(FrameBreak())
    story.extend(build_one_challan_flowables())

    doc.build(story)
    return output_pdf_path

@app.route("/print-delivery-challan-a4", methods=["POST"])
def print_delivery_challan_a4():
    data = request.get_json(force=True, silent=True) or {}

    printer = data.get("printer")
    save_dir = Path(data.get("saveDir") or "")
    file_name = data.get("fileName") or ""
    sumatra = data.get("sumatraPath") or ""

    if not printer:
        return jsonify({"ok": False, "error": "missing 'printer'"}), 400
    if not str(save_dir).strip():
        return jsonify({"ok": False, "error": "missing 'saveDir' (output folder path)"}), 400
    if not sumatra or not Path(sumatra).exists():
        return jsonify({"ok": False, "error": "SumatraPDF not configured"}), 400

    try:
        save_dir.mkdir(parents=True, exist_ok=True)

        if not file_name:
            file_name = f"DC-{uuid.uuid4().hex}.pdf"
        if not file_name.lower().endswith(".pdf"):
            file_name += ".pdf"

        out_path = save_dir / file_name

        # ✅ Read from JSON (1 = normal, 2 = two challans on same page)
        copies_per_page = int(data.get("copiesPerPage") or 1)
        two_per_page = (copies_per_page == 2)

        # ✅ 1) Generate PDF (ONE PAGE if two_per_page=True => top+bottom challans)
        _render_delivery_challan_pdf_reportlab(data, out_path, two_per_page=two_per_page)

        # ✅ 2) Print ONLY ONCE (no nup, no copies)
        cmd = [sumatra, "-silent", "-print-to", printer, "-exit-on-print", str(out_path)]
        sp = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if sp.returncode != 0:
            raise RuntimeError(sp.stderr or sp.stdout or f"Sumatra exit {sp.returncode}")

        return jsonify({
            "ok": True,
            "savedPath": str(out_path),
            "printed": True,
            "copiesPerPage": copies_per_page,
            "twoPerPage": two_per_page
        })
    except Exception as e:
        log(f"/print-delivery-challan-a4 error: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500


# --------------------------- main ---------------------------

if __name__ == "__main__":
    log(f"{APP_NAME} starting… Config: {CONFIG_FILE}")
    log(f"API KEY (use header X-Print-Key): {API_KEY}")
    ensure_autostart_shortcut()

    t = threading.Thread(target=server_thread, daemon=True)
    t.start()

    if TRY_TRAY:
        run_tray()
    else:
        t.join()