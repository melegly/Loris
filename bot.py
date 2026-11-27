import asyncio
from collections import Counter
from datetime import datetime, timedelta, timezone
import io
import json
import logging
import os
from pathlib import Path
import re
import zlib

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.types import BufferedInputFile, InlineKeyboardButton, InlineKeyboardMarkup
import openpyxl
from openpyxl.styles import Alignment, Font
import pdfplumber
from PIL import Image, ImageChops, ImageDraw, ImageFont

# --- НАСТРОЙКИ (задаются в "Переменных окружения" на хостинге) ---
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ALLOWED_USER_ID = int(os.getenv("ALLOWED_USER_ID", "0").strip() or "0")
ENV_EXCLUDED = {
    item.strip()
    for item in os.getenv("EXCLUDED_CONTRACTS", "").split(",")
    if item.strip()
}

EXCLUDED_FILE = Path("excluded_contracts.json")
UI_SETTINGS_FILE = Path("ui_settings.json")
PROFILE_FILE = Path("profile_settings.json")

DEFAULT_UI_SETTINGS = {
    "domain": "1wujym.life",
    "status_time": "auto",  # "auto" или строка вида "22:21"
    "battery": 76,
}

DEFAULT_PROFILE = {
    "name": "",
    "dob": "",
    "country": "Россия",
    "phone": "",
    "email": "",
}

# Кэш последней разобранной выписки и состояние ввода данных профиля (по user_id)
LAST_TRANSFERS: dict[int, list[dict]] = {}
WAITING_PROFILE: set[int] = set()

PHONE_RE = re.compile(
    r"(?:\+?7|8)[\s\-\(\)]*\d{3}[\s\-\(\)]*\d{3}[\s\-\(\)]*\d{2}[\s\-\(\)]*\d{2}"
)
MASKED_CARD_RE = re.compile(
    r"(?:"
    r"\b\d{4}[\s\-]?\d{2}[\*•xXхХ]{2}[\s\-]?[\*•xXхХ]{4}[\s\-]?\d{4}\b"
    r"|\b\d{4,6}[\*•xXхХ]{2,}\d{4}\b"
    r"|(?<!\d)[\*•xXхХ]{1,4}\s?\d{4}\b(?![\*•xXхХ])"
    r")"
)

logging.basicConfig(level=logging.INFO)
dp = Dispatcher()


# ---------------- ХРАНЕНИЕ НАСТРОЕК ----------------

def load_excluded() -> set[str]:
    if EXCLUDED_FILE.exists():
        try:
            data = json.loads(EXCLUDED_FILE.read_text(encoding="utf-8"))
            return set(data)
        except Exception:
            pass
    return set(ENV_EXCLUDED)


def save_excluded(contracts: set[str]) -> None:
    EXCLUDED_FILE.write_text(
        json.dumps(sorted(contracts), ensure_ascii=False),
        encoding="utf-8",
    )


def load_ui_settings() -> dict:
    settings = dict(DEFAULT_UI_SETTINGS)
    if UI_SETTINGS_FILE.exists():
        try:
            data = json.loads(UI_SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                settings.update(data)
        except Exception:
            pass
    return settings


def save_ui_settings(settings: dict) -> None:
    UI_SETTINGS_FILE.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_profile() -> dict:
    profile = dict(DEFAULT_PROFILE)
    if PROFILE_FILE.exists():
        try:
            data = json.loads(PROFILE_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                profile.update(data)
        except Exception:
            pass
    return profile


def save_profile(profile: dict) -> None:
    PROFILE_FILE.write_text(
        json.dumps(profile, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def is_allowed_user(user: types.User | None) -> bool:
    if ALLOWED_USER_ID == 0:
        return True
    return user is not None and user.id == ALLOWED_USER_ID


# ---------------- РАЗБОР PDF Т-БАНКА ----------------

def extract_client_name(first_page_text: str) -> str:
    """Извлекает ФИО клиента из шапки PDF-выписки Т-Банка."""
    stop_words = (
        "справка",
        "акционерное",
        "тбанк",
        "россия",
        "адрес",
        "паспорт",
        "продукт",
        "договор",
        "движение",
        "уважением",
        "руководитель",
        "москва",
        "область",
        "улица",
        "выдан",
        "серия",
        "номер",
        "дата",
        "подразделения",
        "операции",
        "сумма",
    )
    name_re = re.compile(r"^[А-ЯЁ][а-яё\-]+(?:\s+[А-ЯЁ][а-яё\-]+){1,2}$")
    for raw_line in first_page_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if "движение средств" in line.lower():
            break
        line = re.split(
            r"\s+(?:АКЦИОНЕРНОЕ|РОССИЯ,|ТЕЛ\.:|АО\s+«ТБанк»)",
            line,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0].strip()
        if not line:
            continue
        low = line.lower()
        if any(sw in low for sw in stop_words):
            continue
        if name_re.fullmatch(line):
            return line
    return ""


def strip_own_card(desc: str) -> tuple[str, str]:
    protected_spans: list[tuple[int, int]] = []
    for pattern in (PHONE_RE, MASKED_CARD_RE):
        for match in pattern.finditer(desc):
            protected_spans.append(match.span())

    for match in re.finditer(r"(?<![\d+])\b(\d{4})\b(?!\d)", desc):
        start, end = match.span()
        if any(p_start < end and start < p_end for p_start, p_end in protected_spans):
            continue
        before = desc[:start].lower()
        if re.search(r"(?:на карту|с карты|номеру карты)\s+$", before):
            continue
        cleaned = desc[:start] + " " + desc[end:]
        return re.sub(r"\s+", " ", cleaned).strip(), match.group(1)

    return re.sub(r"\s+", " ", desc).strip(), ""


def extract_counterparty(desc: str) -> tuple[str | None, str]:
    phone_match = PHONE_RE.search(desc)
    if phone_match:
        digits = re.sub(r"\D", "", phone_match.group(0))
        if len(digits) == 11 and digits.startswith("8"):
            digits = "7" + digits[1:]
        return f"phone:{digits}", f"+{digits}"

    contract_match = re.search(r"договор\w*\s+(\d{6,})", desc, re.IGNORECASE)
    if contract_match:
        num = contract_match.group(1)
        return f"contract:{num}", f"Договор {num}"

    ext_card_match = MASKED_CARD_RE.search(desc)
    if ext_card_match:
        norm_mask = re.sub(r"[\s\-]+", "", ext_card_match.group(0))
        norm_mask = re.sub(r"[•xXхХ]", "*", norm_mask)
        return f"card:{norm_mask}", f"Карта {norm_mask}"

    plain_card_match = re.search(
        r"(?:на карту|с карты|номеру карты)\s+(\d{4,16})\b",
        desc,
        re.IGNORECASE,
    )
    if plain_card_match:
        num = plain_card_match.group(1)
        return f"card:{num}", f"Карта {num}"

    return None, ""


def detect_method(desc: str, counterparty_key: str | None) -> str:
    desc_low = desc.lower()
    if any(
        k in desc_low
        for k in ("сбп", "система быстрых платежей", "номеру телефона")
    ):
        return "СБП"
    if (
        (counterparty_key and counterparty_key.startswith("card:"))
        or any(
            k in desc_low
            for k in (
                "карт",
                "card2card",
                "c2c",
                "p2p",
                "moneysend",
                "visa direct",
                "vtb",
                "втб",
                "sbol",
                "сбол",
                "yoomoney",
                "юмани",
                "transfer",
            )
        )
    ):
        return "По карте"
    if "договор" in desc_low:
        return "По договору"
    return "Перевод"


def parse_tbank_pdf(
    pdf_bytes: bytes, excluded_contracts: set[str]
) -> tuple[list[dict], int, int, str]:
    skip_prefixes = (
        "АКЦИОНЕРНОЕ ОБЩЕСТВО",
        "РОССИЯ,",
        "ТЕЛ.:",
        "Справка о движении",
        "Исх. №",
        "Адрес места",
        "О продукте",
        "Дата заключения",
        "Номер договора",
        "Номер лицевого",
        "Движение средств",
        "Дата и время",
        "операции",
        "АО «ТБанк»",
        "БИК ",
    )
    stop_markers = ("Пополнения:", "С уважением,")

    raw_lines: list[str] = []
    client_name = ""

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if page_idx == 0:
                client_name = extract_client_name(text)
            for line in text.splitlines():
                clean_line = line.strip()
                if not clean_line:
                    continue
                if any(clean_line.startswith(m) for m in stop_markers):
                    break
                if any(clean_line.startswith(p) for p in skip_prefixes):
                    continue
                if re.fullmatch(r"\d+", clean_line):
                    continue
                raw_lines.append(clean_line)

    blocks: list[list[str]] = []
    date_start_re = re.compile(r"^\d{2}\.\d{2}\.\d{4}\b")
    for line in raw_lines:
        if date_start_re.match(line):
            blocks.append([line])
        elif blocks:
            blocks[-1].append(line)

    amount_re = re.compile(
        r"[+-]\s*\d[\d\s\u00a0]*[.,]\d{2}(?:\s*(?:₽|RUB|USD|EUR|Р|P))?"
    )

    card_transfer_keywords = (
        "перевод с карты",
        "перевод на карту",
        "перевод по номеру карты",
        "пополнение с карты",
        "пополнение на карту",
        "пополнение карты",
        "пополнение по номеру карты",
        "зачисление на карту",
        "с карты на карту",
        "между картами",
        "card2card",
        "c2c",
        "p2p",
        "moneysend",
        "visa direct",
        "perevod",
    )

    include_keywords = (
        "перевод",
        "система быстрых платежей",
        "сбп",
        "пополнение",
        "на карту",
        "с карты",
        "по номеру карты",
        "зачисление",
        *card_transfer_keywords,
    )

    hard_exclude_keywords = (
        "кэшбэк",
        "комиссия",
        "проценты",
        "перевод себе",
        "между своими",
        "со своего счета",
        "на свой счет",
        "инвесткопилка",
        "накопительн",
        "вклад",
        "брокер",
        "банкомат",
        "наличн",
        "терминал",
        "возврат",
        "atm",
    )
    purchase_exclude_keywords = ("оплата", "покупка")

    candidates: list[dict] = []

    for block in blocks:
        full_text = " ".join(block)

        dates = re.findall(r"\b\d{2}\.\d{2}\.\d{4}\b", full_text)
        times = re.findall(r"\b\d{2}:\d{2}\b", full_text)
        amounts = amount_re.findall(full_text)

        if not dates or not amounts:
            continue

        raw_amount = amounts[-1]
        numeric_str = re.sub(r"[^\d.,+-]", "", raw_amount).replace(",", ".")
        try:
            amount = float(numeric_str)
        except ValueError:
            continue

        desc = amount_re.sub(" ", full_text)
        desc = re.sub(r"\b\d{2}\.\d{2}\.\d{4}\b", " ", desc)
        desc = re.sub(r"\b\d{2}:\d{2}\b", " ", desc)
        desc = desc.replace("|", " ")

        desc, card = strip_own_card(desc)
        desc_low = desc.lower()

        if any(bad in desc_low for bad in hard_exclude_keywords):
            continue
        if any(contract in desc for contract in excluded_contracts):
            continue

        is_card_transfer = any(k in desc_low for k in card_transfer_keywords)
        if not is_card_transfer and any(
            bad in desc_low for bad in purchase_exclude_keywords
        ):
            continue

        if not any(good in desc_low for good in include_keywords):
            continue

        direction = "Входящий" if amount > 0 else "Исходящий"
        counterparty_key, counterparty_label = extract_counterparty(desc)
        method = detect_method(desc, counterparty_key)

        candidates.append(
            {
                "date": dates[0],
                "time": times[0] if times else "",
                "direction": direction,
                "method": method,
                "amount": amount,
                "counterparty_key": counterparty_key,
                "counterparty": counterparty_label,
                "description": desc,
                "card": card,
            }
        )

    total_recognized = len(candidates)

    key_counts = Counter(
        item["counterparty_key"]
        for item in candidates
        if item["counterparty_key"] is not None
    )

    results: list[dict] = []
    skipped_repeated = 0

    for item in candidates:
        key = item["counterparty_key"]
        if key is not None and key_counts[key] > 1:
            skipped_repeated += 1
            continue
        results.append(item)

    return results, total_recognized, skipped_repeated, client_name


def parse_profile_input(text: str) -> dict[str, str]:
    """Разбирает сообщение пользователя на почту, телефон, дату рождения и имя."""
    parsed: dict[str, str] = {}
    remaining = text.strip()

    email_m = re.search(
        r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}",
        remaining,
    )
    if email_m:
        parsed["email"] = email_m.group(0).strip()
        remaining = remaining[: email_m.start()] + " " + remaining[email_m.end() :]

    dob_m = re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", remaining)
    if dob_m:
        parsed["dob"] = dob_m.group(0)
        remaining = remaining[: dob_m.start()] + " " + remaining[dob_m.end() :]

    phone_m = re.search(r"\+?\d[\d\s\-\(\)]{7,}\d", remaining)
    if phone_m:
        parsed["phone"] = re.sub(r"\s+", " ", phone_m.group(0)).strip()
        remaining = remaining[: phone_m.start()] + " " + remaining[phone_m.end() :]

    leftover = re.sub(
        r"\b(?:телефон|тел|почта|email|e-mail|mail|имя|фио|дата\s+рождения|др)\s*[:\-]?",
        " ",
        remaining,
        flags=re.IGNORECASE,
    )
    leftover = re.sub(r"[,;|\n\r]+", " ", leftover)
    leftover = re.sub(r"\s+", " ", leftover).strip()
    if leftover and re.search(r"[А-Яа-яЁёA-Za-z]", leftover):
        parsed["name"] = leftover

    return parsed


# ---------------- EXCEL ----------------

def build_excel(rows: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Переводы"

    headers = [
        "Дата",
        "Время",
        "Тип",
        "Способ",
        "Сумма (₽)",
        "Номер / Договор / Карта",
        "Описание",
        "Своя карта",
    ]
    ws.append(headers)

    bold_font = Font(bold=True)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append(
            [
                row["date"],
                row["time"],
                row["direction"],
                row["method"],
                row["amount"],
                row["counterparty"],
                row["description"],
                row["card"],
            ]
        )
        ws.cell(row=ws.max_row, column=5).number_format = "#,##0.00"

    col_widths = [12, 8, 12, 14, 14, 24, 46, 11]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ---------------- ПОИСК РЕСУРСОВ И ГЕНЕРАЦИЯ PNG ----------------

def find_asset_ci(
    exact_names: tuple[str, ...],
    substrings: tuple[str, ...] = (),
    exts: tuple[str, ...] = (),
) -> Path | None:
    search_dirs = (Path("assets"), Path("."))
    exact_lower = {name.lower() for name in exact_names}
    exts_lower = {ext.lower() for ext in exts}

    all_files: list[Path] = []
    for d in search_dirs:
        if d.exists() and d.is_dir():
            for p in d.iterdir():
                if p.is_file():
                    all_files.append(p)

    for p in all_files:
        if p.name.lower() in exact_lower:
            return p

    if substrings:
        for p in all_files:
            if exts_lower and p.suffix.lower() not in exts_lower:
                continue
            stem_low = p.stem.lower()
            if any(sub in stem_low for sub in substrings):
                return p

    return None


def get_sbp_asset_path() -> Path | None:
    return find_asset_ci(
        exact_names=("sbp.png", "sbp_icon.png", "sbp.webp", "sbp.jpg", "sbp.jpeg"),
        substrings=("sbp", "сбп"),
        exts=(".png", ".webp", ".jpg", ".jpeg"),
    )


def get_card_asset_path() -> Path | None:
    return find_asset_ci(
        exact_names=("card.png", "card_icon.png", "card.webp"),
        substrings=("card", "карт"),
        exts=(".png", ".webp", ".jpg", ".jpeg"),
    )


def get_font_asset_path(bold: bool = False) -> Path | None:
    if bold:
        return find_asset_ci(
            exact_names=(
                "font_bold.ttf",
                "inter-semibold.ttf",
                "inter-bold.ttf",
                "roboto-bold.ttf",
                "roboto-medium.ttf",
            ),
            substrings=("semibold", "bold"),
            exts=(".ttf", ".otf"),
        )
    return find_asset_ci(
        exact_names=(
            "font_regular.ttf",
            "inter-regular.ttf",
            "roboto-regular.ttf",
        ),
        substrings=("regular",),
        exts=(".ttf", ".otf"),
    )


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    asset_path = get_font_asset_path(bold=bold) or get_font_asset_path(bold=not bold)
    if asset_path:
        try:
            return ImageFont.truetype(str(asset_path), size)
        except Exception:
            pass

    sys_candidates = (
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
        )
        if bold
        else (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        )
    )
    for sys_path in sys_candidates:
        if Path(sys_path).exists():
            try:
                return ImageFont.truetype(sys_path, size)
            except Exception:
                pass

    return ImageFont.load_default()


def format_rub(amount_abs: float, sign: str) -> str:
    formatted = f"{amount_abs:,.2f}".replace(",", " ").replace(".", ",")
    return f"{sign}{formatted} ₽"


def make_seconds(row: dict) -> str:
    base_time = row.get("time") or "12:00"
    if len(base_time) == 8:
        return base_time
    seed = f"{row['date']}|{base_time}|{row['amount']}|{row['description']}"
    sec = 10 + (zlib.crc32(seed.encode("utf-8")) % 47)
    return f"{base_time}:{sec:02d}"


def draw_default_card_icon(width: int = 108, height: int = 76) -> Image.Image:
    scale = 4
    w, h = width * scale, height * scale
    card = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(card, "RGBA")

    radius = int(h * 0.16)
    d.rounded_rectangle([0, 0, w - 1, h - 1], radius=radius, fill=(34, 44, 86, 255))
    d.rectangle([0, int(h * 0.20), w - 1, int(h * 0.37)], fill=(20, 27, 56, 255))
    d.rounded_rectangle(
        [int(w * 0.11), int(h * 0.56), int(w * 0.33), int(h * 0.80)],
        radius=int(h * 0.065),
        fill=(238, 192, 86, 255),
    )

    r = int(h * 0.165)
    cy = int(h * 0.68)
    cx1 = int(w * 0.65)
    cx2 = int(w * 0.80)
    d.ellipse([cx1 - r, cy - r, cx1 + r, cy + r], fill=(234, 42, 42, 255))
    d.ellipse([cx2 - r, cy - r, cx2 + r, cy + r], fill=(242, 160, 44, 240))

    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w - 1, h - 1], radius=radius, fill=255)
    card.putalpha(mask)

    return card.resize((width, height), Image.Resampling.LANCZOS)


def draw_default_sbp_icon() -> Image.Image:
    w, h = 280, 120
    icon = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(icon)
    d.polygon([(12, 16), (68, 60), (12, 104)], fill=(91, 45, 144, 255))
    d.polygon([(34, 8), (90, 50), (46, 60)], fill=(250, 175, 24, 255))
    d.polygon([(34, 112), (90, 70), (46, 60)], fill=(0, 166, 81, 255))
    f = load_font(56, bold=True)
    d.text((108, 24), "сбп", fill=(20, 24, 35, 255), font=f)
    return icon


def trim_icon_margins(raw: Image.Image) -> Image.Image:
    raw = raw.convert("RGBA")

    alpha_bbox = raw.getchannel("A").getbbox()
    if alpha_bbox:
        raw = raw.crop(alpha_bbox)

    white_bg = Image.new("RGBA", raw.size, (255, 255, 255, 255))
    flat = Image.alpha_composite(white_bg, raw).convert("RGB")
    diff = ImageChops.difference(
        flat, Image.new("RGB", flat.size, (255, 255, 255))
    ).convert("L")
    mask = diff.point(lambda p: 255 if p > 12 else 0)
    content_bbox = mask.getbbox()
    if content_bbox:
        raw = raw.crop(content_bbox)

    return raw


def load_icon(kind: str, max_w: int = 142, max_h: int = 90) -> Image.Image:
    if kind == "card":
        p = get_card_asset_path()
        if p:
            raw = trim_icon_margins(Image.open(p).convert("RGBA"))
            raw.thumbnail((108, 76), Image.Resampling.LANCZOS)
            return raw
        return draw_default_card_icon(108, 76)

    p = get_sbp_asset_path()
    raw = Image.open(p).convert("RGBA") if p else draw_default_sbp_icon()
    raw = trim_icon_margins(raw)
    raw.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
    return raw


def draw_ios_signal(img: Image.Image, left_x: int, bottom_y: int) -> None:
    scale = 4
    w, h = 56, 40
    layer = Image.new("RGBA", (w * scale, h * scale), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    bar_w = 9 * scale
    gap = 6 * scale
    heights = (13, 20, 28, 36)
    base_y = (h - 2) * scale

    for i, bh in enumerate(heights):
        x0 = i * (bar_w + gap)
        y0 = base_y - bh * scale
        d.rounded_rectangle(
            [x0, y0, x0 + bar_w, base_y],
            radius=bar_w // 2,
            fill=(255, 255, 255, 255),
        )

    small = layer.resize((w, h), Image.Resampling.LANCZOS)
    img.paste(small, (left_x, bottom_y - (h - 2)), small)


def draw_ios_wifi(img: Image.Image, center_x: int, bottom_y: int) -> None:
    scale = 6
    w, h = 56, 42
    W_s, H_s = w * scale, h * scale
    cx = W_s / 2
    cy = (h - 4.0) * scale

    mask = Image.new("L", (W_s, H_s), 0)
    md = ImageDraw.Draw(mask)

    a1, a2 = 226, 314

    r0 = 10.6 * scale
    md.pieslice([cx - r0, cy - r0, cx + r0, cy + r0], start=a1, end=a2, fill=255)
    r_tip = 2.0 * scale
    md.ellipse(
        [cx - r_tip, cy - r_tip * 1.4, cx + r_tip, cy + r_tip * 0.6],
        fill=255,
    )

    for r_in_base, r_out_base in ((16.2, 23.0), (28.6, 35.4)):
        r_in = r_in_base * scale
        r_out = r_out_base * scale
        arc_layer = Image.new("L", (W_s, H_s), 0)
        ad = ImageDraw.Draw(arc_layer)
        ad.pieslice(
            [cx - r_out, cy - r_out, cx + r_out, cy + r_out],
            start=a1,
            end=a2,
            fill=255,
        )
        ad.ellipse(
            [cx - r_in, cy - r_in, cx + r_in, cy + r_in],
            fill=0,
        )
        mask = ImageChops.lighter(mask, arc_layer)

    wifi_rgba = Image.new("RGBA", (W_s, H_s), (255, 255, 255, 255))
    wifi_rgba.putalpha(mask)
    wifi_small = wifi_rgba.resize((w, h), Image.Resampling.LANCZOS)

    img.paste(wifi_small, (int(center_x - w / 2), int(bottom_y - 39)), wifi_small)


def draw_ios_battery(img: Image.Image, left_x: int, bottom_y: int, battery: int) -> None:
    """Классическая иконка батареи iOS без цифр процентов."""
    scale = 4
    bat_w, bat_h = 72, 36
    cw, ch = 84, 42
    W_s, H_s = cw * scale, ch * scale
    top_pad = 3 * scale

    pct = max(5, min(100, int(battery)))

    bat_img = Image.new("RGBA", (W_s, H_s), (0, 0, 0, 0))
    bd = ImageDraw.Draw(bat_img, "RGBA")

    outline_color = (142, 145, 152, 255)
    fill_color = (255, 255, 255, 255) if pct > 20 else (255, 69, 58, 255)

    x0 = 1 * scale
    y0 = top_pad
    x1 = x0 + bat_w * scale
    y1 = y0 + bat_h * scale
    bd.rounded_rectangle(
        [x0, y0, x1, y1],
        radius=10 * scale,
        outline=outline_color,
        width=3 * scale,
    )

    pad = int(5.5 * scale)
    max_fill_w = (x1 - x0) - 2 * pad
    fill_w = max(int(8 * scale), int(max_fill_w * pct / 100))
    bd.rounded_rectangle(
        [x0 + pad, y0 + pad, x0 + pad + fill_w, y1 - pad],
        radius=int(5.2 * scale),
        fill=fill_color,
    )

    nub_x0 = x1 + int(3.0 * scale)
    nub_x1 = x1 + int(7.5 * scale)
    nub_h = 13 * scale
    nub_y0 = int(y0 + (bat_h * scale - nub_h) / 2)
    nub_y1 = nub_y0 + nub_h
    bd.pieslice(
        [nub_x0 - int(4.5 * scale), nub_y0, nub_x1, nub_y1],
        start=270,
        end=90,
        fill=outline_color,
    )

    small = bat_img.resize((cw, ch), Image.Resampling.LANCZOS)
    img.paste(small, (left_x, bottom_y - bat_h - 3), small)


def draw_status_bar(img: Image.Image, draw: ImageDraw.ImageDraw, W: int, time_str: str, battery: int) -> None:
    font_time = load_font(45, bold=True)
    tb = draw.textbbox((0, 0), time_str, font=font_time)
    th = tb[3] - tb[1]
    draw.text((135, 62 - th / 2 - tb[1]), time_str, fill=(255, 255, 255), font=font_time)

    bottom_y = 80
    draw_ios_signal(img, left_x=W - 296, bottom_y=bottom_y)
    draw_ios_wifi(img, center_x=W - 199, bottom_y=bottom_y)
    draw_ios_battery(img, left_x=W - 156, bottom_y=bottom_y, battery=battery)


def draw_browser_bar(draw: ImageDraw.ImageDraw, W: int, domain: str) -> None:
    x0, y0, x1, y1 = 36, 136, W - 36, 246
    draw.rounded_rectangle([x0, y0, x1, y1], radius=55, fill=(58, 60, 66))

    lx, ly = x0 + 48, (y0 + y1) // 2
    draw.rounded_rectangle([lx - 18, ly - 18, lx + 18, ly + 18], radius=7, outline=(215, 218, 224), width=4)
    draw.ellipse([lx - 6, ly - 6, lx + 6, ly + 6], fill=(215, 218, 224))

    font_dom = load_font(37, bold=False)
    tb = draw.textbbox((0, 0), domain, font=font_dom)
    tw, th = tb[2] - tb[0], tb[3] - tb[1]
    draw.text(((W - tw) / 2, (y0 + y1 - th) / 2 - tb[1]), domain, fill=(228, 231, 236), font=font_dom)

    sx, sy = x1 - 52, (y0 + y1) // 2
    draw.rounded_rectangle([sx - 16, sy - 6, sx + 16, sy + 20], radius=4, outline=(215, 218, 224), width=4)
    draw.line([(sx, sy + 6), (sx, sy - 20)], fill=(215, 218, 224), width=4)
    draw.line([(sx - 9, sy - 11), (sx, sy - 20), (sx + 9, sy - 11)], fill=(215, 218, 224), width=4)


def draw_checkmark(draw: ImageDraw.ImageDraw, x: int, y: int, color=(46, 158, 91)) -> None:
    draw.line([(x, y + 16), (x + 11, y + 27), (x + 30, y + 7)], fill=color, width=6, joint="curve")


def draw_verified_badge(size: int = 40) -> Image.Image:
    """Зелёный кружок с белой галочкой подтверждения."""
    scale = 4
    s = size * scale
    badge = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(badge, "RGBA")
    pad = 2 * scale
    d.ellipse([pad, pad, s - pad, s - pad], fill=(42, 175, 102, 255))
    d.line(
        [
            (int(s * 0.28), int(s * 0.52)),
            (int(s * 0.44), int(s * 0.67)),
            (int(s * 0.72), int(s * 0.35)),
        ],
        fill=(255, 255, 255, 255),
        width=int(3.6 * scale),
        joint="curve",
    )
    return badge.resize((size, size), Image.Resampling.LANCZOS)


def draw_ru_flag(width: int = 46, height: int = 32) -> Image.Image:
    """Флаг России со скруглёнными углами."""
    scale = 4
    w, h = width * scale, height * scale
    flag = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(flag)
    h1 = h // 3
    h2 = (2 * h) // 3
    d.rectangle([0, 0, w, h1], fill=(255, 255, 255, 255))
    d.rectangle([0, h1, w, h2], fill=(28, 87, 184, 255))
    d.rectangle([0, h2, w, h], fill=(213, 43, 30, 255))

    radius = int(5 * scale)
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w - 1, h - 1], radius=radius, fill=255)
    flag.putalpha(mask)

    border = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(border).rounded_rectangle(
        [0, 0, w - 1, h - 1],
        radius=radius,
        outline=(218, 222, 230, 255),
        width=2 * scale,
    )
    flag = Image.alpha_composite(flag, border)
    return flag.resize((width, height), Image.Resampling.LANCZOS)


def draw_chevron_right(img: Image.Image, right_x: int, center_y: int) -> None:
    """Тонкая серая стрелка '>' со сглаживанием."""
    scale = 4
    w, h = 22, 34
    layer = Image.new("RGBA", (w * scale, h * scale), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    d.line(
        [
            (4 * scale, 4 * scale),
            (17 * scale, 17 * scale),
            (4 * scale, 30 * scale),
        ],
        fill=(136, 144, 158, 255),
        width=int(3.2 * scale),
        joint="curve",
    )
    small = layer.resize((w, h), Image.Resampling.LANCZOS)
    img.paste(small, (right_x - w, int(center_y - h / 2)), small)


def get_status_time_str(ui_settings: dict) -> str:
    status_time = ui_settings.get("status_time", "auto")
    if status_time == "auto":
        msk = timezone(timedelta(hours=3))
        return datetime.now(msk).strftime("%H:%M")
    return status_time


def draw_sheet_header(
    img: Image.Image,
    draw: ImageDraw.ImageDraw,
    W: int,
    H: int,
    title: str,
    ui_settings: dict,
) -> int:
    domain = ui_settings.get("domain", "1wujym.life")
    battery = int(ui_settings.get("battery", 76))
    status_time = get_status_time_str(ui_settings)

    draw_status_bar(img, draw, W, status_time, battery)
    draw_browser_bar(draw, W, domain)

    sheet_top = 285
    draw.rounded_rectangle([0, sheet_top, W, H + 80], radius=64, fill=(255, 255, 255))

    font_back = load_font(45, bold=True)
    font_title = load_font(72, bold=True)

    bx, by = 62, sheet_top + 64
    draw.line([(bx + 20, by), (bx, by + 20), (bx + 20, by + 40)], fill=(27, 102, 224), width=6, joint="curve")
    draw.text((bx + 44, by - 7), "Назад", fill=(27, 102, 224), font=font_back)

    cx, cy = W - 82, sheet_top + 84
    draw.line([(cx - 19, cy - 19), (cx + 19, cy + 19)], fill=(45, 52, 66), width=5)
    draw.line([(cx + 19, cy - 19), (cx - 19, cy + 19)], fill=(45, 52, 66), width=5)

    draw.text((54, sheet_top + 172), title, fill=(17, 24, 39), font=font_title)
    return sheet_top


def build_screenshots(
    transfers: list[dict],
    mode: str,
    ui_settings: dict,
    per_page: int = 7,
) -> list[bytes]:
    if mode == "deposits":
        filtered = [t for t in transfers if t["amount"] < 0]
        active_tab = "Депозиты"
    elif mode == "withdrawals":
        filtered = [t for t in transfers if t["amount"] > 0]
        active_tab = "Выводы"
    else:
        filtered = list(transfers)
        active_tab = "Все"

    if not filtered:
        return []

    W, H = 1170, 2532

    sbp_icon = load_icon("sbp", max_w=142, max_h=90)
    card_icon = load_icon("card", max_w=108, max_h=76)

    font_tab = load_font(44, bold=True)
    font_row_title = load_font(47, bold=True)
    font_status = load_font(38, bold=False)
    font_amount = load_font(46, bold=True)
    font_meta = load_font(37, bold=False)

    pages: list[bytes] = []

    for start_idx in range(0, len(filtered), per_page):
        chunk = filtered[start_idx : start_idx + per_page]

        img = Image.new("RGB", (W, H), (14, 16, 20))
        draw = ImageDraw.Draw(img)
        sheet_top = draw_sheet_header(img, draw, W, H, "История платежей", ui_settings)

        tab_y0, tab_y1 = sheet_top + 308, sheet_top + 436
        tab_x0, tab_x1 = 54, W - 54
        draw.rounded_rectangle([tab_x0, tab_y0, tab_x1, tab_y1], radius=28, fill=(246, 247, 249))

        tab_names = ("Все", "Депозиты", "Выводы")
        tab_w = (tab_x1 - tab_x0) / 3
        for idx, t_name in enumerate(tab_names):
            tx0 = tab_x0 + idx * tab_w
            tx1 = tx0 + tab_w
            if t_name == active_tab:
                draw.rounded_rectangle(
                    [tx0 + 8, tab_y0 + 8, tx1 - 8, tab_y1 - 8],
                    radius=22,
                    fill=(235, 237, 241),
                )
            tb = draw.textbbox((0, 0), t_name, font=font_tab)
            tw, th = tb[2] - tb[0], tb[3] - tb[1]
            draw.text(
                (tx0 + (tab_w - tw) / 2, tab_y0 + (tab_y1 - tab_y0 - th) / 2 - tb[1]),
                t_name,
                fill=(17, 24, 39),
                font=font_tab,
            )

        list_top = tab_y1 + 38
        row_h = 252

        for i, row in enumerate(chunk):
            ry = list_top + i * row_h

            is_deposit = row["amount"] < 0
            is_card = row["method"] == "По карте"

            icon = card_icon if is_card else sbp_icon
            ix = 44 + (142 - icon.width) // 2
            iy = ry + 48 + (92 - icon.height) // 2
            img.paste(icon, (ix, iy), icon)

            if is_deposit:
                title_lines = (
                    ("Перевод по номеру", "карты")
                    if is_card
                    else ("Перевод по номеру", "телефона")
                )
                amount_str = format_rub(abs(row["amount"]), "+")
            else:
                title_lines = (
                    ("Вывод по номеру", "карты")
                    if is_card
                    else ("Вывод по номеру", "телефона")
                )
                amount_str = format_rub(abs(row["amount"]), "-")

            tx = 206
            draw.text((tx, ry + 18), title_lines[0], fill=(17, 24, 39), font=font_row_title)
            draw.text((tx, ry + 78), title_lines[1], fill=(17, 24, 39), font=font_row_title)

            draw_checkmark(draw, tx + 2, ry + 148)
            draw.text((tx + 46, ry + 143), "Выполнено", fill=(46, 158, 91), font=font_status)

            right_x = W - 52

            ab = draw.textbbox((0, 0), amount_str, font=font_amount)
            draw.text((right_x - (ab[2] - ab[0]), ry + 20), amount_str, fill=(17, 24, 39), font=font_amount)

            date_str = row["date"]
            db = draw.textbbox((0, 0), date_str, font=font_meta)
            draw.text((right_x - (db[2] - db[0]), ry + 92), date_str, fill=(142, 149, 160), font=font_meta)

            time_str = make_seconds(row)
            tb_t = draw.textbbox((0, 0), time_str, font=font_meta)
            draw.text((right_x - (tb_t[2] - tb_t[0]), ry + 146), time_str, fill=(142, 149, 160), font=font_meta)

            if i < len(chunk) - 1:
                sep_y = ry + row_h - 16
                draw.line([(tx, sep_y), (right_x, sep_y)], fill=(242, 244, 247), width=2)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        pages.append(buf.getvalue())

    return pages


def build_profile_screenshot(profile: dict, ui_settings: dict) -> bytes:
    """Генерирует PNG-скриншот страницы 'Настройки' профиля."""
    W, H = 1170, 2532
    img = Image.new("RGB", (W, H), (14, 16, 20))
    draw = ImageDraw.Draw(img)

    sheet_top = draw_sheet_header(img, draw, W, H, "Настройки", ui_settings)

    font_section = load_font(42, bold=False)
    font_label = load_font(42, bold=True)
    font_helper = load_font(34, bold=False)

    text_dark = (17, 24, 39)
    text_grey = (138, 145, 158)
    text_blue = (36, 114, 232)
    card_bg = (245, 246, 249)
    divider_color = (233, 236, 241)

    badge_icon = draw_verified_badge(40)
    ru_flag = draw_ru_flag(46, 32)

    card_x0 = 54
    card_x1 = W - 54
    card_radius = 36

    def draw_row_content(
        y_center: int,
        label: str,
        value: str = "",
        verified: bool = False,
        show_flag: bool = False,
    ) -> None:
        # Подпись слева
        lb = draw.textbbox((0, 0), label, font=font_label)
        lw, lh = lb[2] - lb[0], lb[3] - lb[1]
        draw.text(
            (card_x0 + 44, y_center - lh / 2 - lb[1]),
            label,
            fill=text_dark,
            font=font_label,
        )

        # Стрелка '>' справа
        chev_right = card_x1 - 40
        draw_chevron_right(img, chev_right, y_center)
        right_cursor = chev_right - 46

        if verified:
            img.paste(
                badge_icon,
                (right_cursor - badge_icon.width, int(y_center - badge_icon.height / 2)),
                badge_icon,
            )
            right_cursor -= badge_icon.width + 20

        if show_flag:
            img.paste(
                ru_flag,
                (right_cursor - ru_flag.width, int(y_center - ru_flag.height / 2)),
                ru_flag,
            )
            right_cursor -= ru_flag.width + 18

        if value:
            max_val_w = right_cursor - (card_x0 + 44 + lw + 30)
            val_font_size = 40
            val_font = load_font(val_font_size, bold=False)
            vb = draw.textbbox((0, 0), value, font=val_font)
            while (vb[2] - vb[0]) > max_val_w and val_font_size > 26:
                val_font_size -= 2
                val_font = load_font(val_font_size, bold=False)
                vb = draw.textbbox((0, 0), value, font=val_font)
            vw, vh = vb[2] - vb[0], vb[3] - vb[1]
            draw.text(
                (right_cursor - vw, y_center - vh / 2 - vb[1]),
                value,
                fill=text_grey,
                font=val_font,
            )

    # --- Блок 1: Основные данные ---
    sec1_y = sheet_top + 296
    draw.text((card_x0, sec1_y), "Основные данные", fill=text_grey, font=font_section)

    row_h = 156
    c1_y0 = sec1_y + 74
    c1_y1 = c1_y0 + row_h * 3
    draw.rounded_rectangle([card_x0, c1_y0, card_x1, c1_y1], radius=card_radius, fill=card_bg)

    draw.line(
        [(card_x0 + 42, c1_y0 + row_h), (card_x1 - 42, c1_y0 + row_h)],
        fill=divider_color,
        width=2,
    )
    draw.line(
        [(card_x0 + 42, c1_y0 + row_h * 2), (card_x1 - 42, c1_y0 + row_h * 2)],
        fill=divider_color,
        width=2,
    )

    draw_row_content(c1_y0 + row_h // 2, "Имя", profile.get("name", ""))
    draw_row_content(c1_y0 + row_h + row_h // 2, "Дата рождения", profile.get("dob", ""))
    draw_row_content(
        c1_y0 + row_h * 2 + row_h // 2,
        "Страна регистрации",
        profile.get("country", "Россия"),
        show_flag=True,
    )

    # --- Блок 2: Контактные данные ---
    sec2_y = c1_y1 + 82
    draw.text((card_x0, sec2_y), "Контактные данные", fill=text_grey, font=font_section)

    single_h = 152
    c2_y0 = sec2_y + 72
    c2_y1 = c2_y0 + single_h
    draw.rounded_rectangle([card_x0, c2_y0, card_x1, c2_y1], radius=card_radius, fill=card_bg)
    draw_row_content(
        c2_y0 + single_h // 2,
        "Телефон",
        profile.get("phone", ""),
        verified=bool(profile.get("phone")),
    )

    c3_y0 = c2_y1 + 32
    c3_y1 = c3_y0 + single_h
    draw.rounded_rectangle([card_x0, c3_y0, card_x1, c3_y1], radius=card_radius, fill=card_bg)
    draw_row_content(
        c3_y0 + single_h // 2,
        "Почта",
        profile.get("email", ""),
        verified=bool(profile.get("email")),
    )

    # Подпись под контактами (автоподбор размера, чтобы не вылезла за край)
    h1_part1 = "Для изменения подтвержденных данных "
    h1_part2 = "свяжитесь с нами"
    h1_size = 34
    f_h1 = load_font(h1_size, bold=False)
    full_w = draw.textbbox((0, 0), h1_part1 + h1_part2, font=f_h1)[2]
    while full_w > (card_x1 - card_x0) and h1_size > 26:
        h1_size -= 1
        f_h1 = load_font(h1_size, bold=False)
        full_w = draw.textbbox((0, 0), h1_part1 + h1_part2, font=f_h1)[2]

    h1_y = c3_y1 + 36
    draw.text((card_x0, h1_y), h1_part1, fill=text_grey, font=f_h1)
    p1_w = draw.textbbox((0, 0), h1_part1, font=f_h1)[2]
    draw.text((card_x0 + p1_w, h1_y), h1_part2, fill=text_blue, font=f_h1)

    # --- Блок 3: Безопасность ---
    sec3_y = h1_y + 114
    draw.text((card_x0, sec3_y), "Безопасность", fill=text_grey, font=font_section)

    c4_y0 = sec3_y + 72
    c4_y1 = c4_y0 + single_h
    draw.rounded_rectangle([card_x0, c4_y0, card_x1, c4_y1], radius=card_radius, fill=card_bg)
    draw_row_content(c4_y0 + single_h // 2, "Пароль")

    h2_y = c4_y1 + 36
    draw.text(
        (card_x0, h2_y),
        "Для изменения нужно ввести текущий пароль",
        fill=text_grey,
        font=font_helper,
    )

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def build_action_keyboard(transfers: list[dict]) -> InlineKeyboardMarkup:
    dep_cnt = sum(1 for t in transfers if t["amount"] < 0)
    wd_cnt = sum(1 for t in transfers if t["amount"] > 0)
    all_cnt = len(transfers)

    buttons = []
    row1 = []
    if dep_cnt:
        row1.append(
            InlineKeyboardButton(
                text=f"Скрины: Депозиты ({dep_cnt})", callback_data="scr:deposits"
            )
        )
    if wd_cnt:
        row1.append(
            InlineKeyboardButton(
                text=f"Скрины: Выводы ({wd_cnt})", callback_data="scr:withdrawals"
            )
        )
    if row1:
        buttons.append(row1)
    if all_cnt:
        buttons.append(
            [
                InlineKeyboardButton(
                    text=f"Скрины: Все ({all_cnt})", callback_data="scr:all"
                )
            ]
        )
    buttons.append(
        [
            InlineKeyboardButton(
                text="Скриншот профиля", callback_data="scr:profile"
            )
        ]
    )
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def profile_prompt_text(profile: dict) -> str:
    cur_name = profile.get("name") or "не указано"
    cur_dob = profile.get("dob") or "не указана"
    return (
        "Введи **телефон** и **почту** для скриншота профиля "
        "(одной строкой или с новой строки).\n\n"
        f"Сейчас в профиле:\n"
        f"• Имя: `{cur_name}`\n"
        f"• Дата рождения: `{cur_dob}`\n\n"
        "Если нужно сразу поменять имя или дату рождения, просто допиши их туда же:\n"
        "`+79991234567 mail@gmail.com`\n"
        "или\n"
        "`+79991234567 mail@gmail.com 15.05.1995 Татьяна Берлева`\n\n"
        "Для отмены нажми /cancel"
    )


# ---------------- КОМАНДЫ И ОБРАБОТЧИКИ ----------------

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    if message.from_user:
        WAITING_PROFILE.discard(message.from_user.id)
    ui = load_ui_settings()
    sbp_p = get_sbp_asset_path()
    f_reg = get_font_asset_path(bold=False)
    f_bold = get_font_asset_path(bold=True)

    assets_info = (
        f"Файлы на сервере:\n"
        f"• СБП: {sbp_p.name if sbp_p else 'НЕ НАЙДЕН (используется заглушка)'}\n"
        f"• Шрифты: {f_reg.name if f_reg else 'системный'} / {f_bold.name if f_bold else 'системный'}"
    )
    await message.answer(
        "Кидай PDF-выписку из Т-Банка — верну Excel и кнопки для генерации PNG-скриншотов.\n"
        "• Исходящие из выписки → Депозиты (+)\n"
        "• Входящие из выписки → Выводы (-)\n\n"
        f"Настройки скрина: домен `{ui['domain']}`, часы `{ui['status_time']}`, заряд `{ui['battery']}%`\n"
        f"{assets_info}\n\n"
        "Команды:\n"
        "/profile — сделать скриншот профиля (Настройки)\n"
        "/domain 1wujym.life — сменить домен в строке браузера\n"
        "/status 22:21 76 — задать часы и уровень батареи (или /status auto 65)\n"
        "/excluded — список игнорируемых договоров\n"
        "/exclude 1234567890 — добавить договор в исключения\n"
        "/unexclude 1234567890 — убрать из исключений",
        parse_mode="Markdown",
    )


@dp.message(Command("profile"))
async def cmd_profile(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    if message.from_user:
        WAITING_PROFILE.add(message.from_user.id)
    profile = load_profile()
    await message.answer(profile_prompt_text(profile), parse_mode="Markdown")


@dp.message(Command("cancel"))
async def cmd_cancel(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    if message.from_user:
        WAITING_PROFILE.discard(message.from_user.id)
    await message.answer("Отменено.")


@dp.message(Command("domain"))
async def cmd_domain(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    if message.from_user:
        WAITING_PROFILE.discard(message.from_user.id)
    parts = (message.text or "").split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Формат: /domain 1wujym.life")
        return
    ui = load_ui_settings()
    ui["domain"] = parts[1].strip()
    save_ui_settings(ui)
    await message.answer(f"Домен обновлён: {ui['domain']}")


@dp.message(Command("status"))
async def cmd_status(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    if message.from_user:
        WAITING_PROFILE.discard(message.from_user.id)
    parts = (message.text or "").split()
    if len(parts) < 2:
        await message.answer("Формат: /status 22:21 76 (или /status auto 54)")
        return
    ui = load_ui_settings()
    t_val = parts[1].strip()
    if t_val != "auto" and not re.fullmatch(r"\d{1,2}:\d{2}", t_val):
        await message.answer("Время укажи в формате ЧЧ:ММ (например, 22:21) или auto.")
        return
    ui["status_time"] = t_val
    if len(parts) >= 3 and parts[2].isdigit():
        ui["battery"] = max(5, min(100, int(parts[2])))
    save_ui_settings(ui)
    await message.answer(
        f"Статус-бар обновлён: время {ui['status_time']}, батарея {ui['battery']}%"
    )


@dp.message(Command("excluded"))
async def cmd_excluded(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    contracts = sorted(load_excluded())
    text = "\n".join(contracts) if contracts else "Список пуст."
    await message.answer("Игнорируемые договоры:\n" + text)


@dp.message(Command("exclude"))
async def cmd_exclude(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    parts = (message.text or "").split()
    if len(parts) < 2 or not parts[1].isdigit():
        await message.answer("Формат: /exclude 1234567890")
        return
    contracts = load_excluded()
    contracts.add(parts[1])
    save_excluded(contracts)
    await message.answer(f"Договор {parts[1]} добавлен в исключения.")


@dp.message(Command("unexclude"))
async def cmd_unexclude(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    parts = (message.text or "").split()
    if len(parts) < 2:
        await message.answer("Формат: /unexclude 1234567890")
        return
    contracts = load_excluded()
    contracts.discard(parts[1])
    save_excluded(contracts)
    await message.answer(f"Договор {parts[1]} убран из исключений.")


@dp.message(F.document)
async def handle_pdf(message: types.Message, bot: Bot):
    if not is_allowed_user(message.from_user):
        return
    if message.from_user:
        WAITING_PROFILE.discard(message.from_user.id)

    doc = message.document
    if not doc or not (doc.file_name or "").lower().endswith(".pdf"):
        await message.answer("Пришли именно PDF-файл выписки из Т-Банка.")
        return

    wait_msg = await message.answer("Разбираю выписку...")

    try:
        file_info = await bot.get_file(doc.file_id)
        buf = io.BytesIO()
        await bot.download_file(file_info.file_path, buf)
        pdf_bytes = buf.getvalue()

        excluded = load_excluded()
        transfers, total_recognized, skipped_repeated, client_name = parse_tbank_pdf(
            pdf_bytes, excluded
        )

        if client_name:
            profile = load_profile()
            profile["name"] = client_name
            save_profile(profile)

        if not transfers:
            await wait_msg.edit_text(
                f"Распознано переводов: {total_recognized}\n"
                f"Отсечено повторяющихся: {skipped_repeated}\n"
                "Разовых переводов для таблицы не осталось."
            )
            return

        if message.from_user:
            LAST_TRANSFERS[message.from_user.id] = transfers

        incoming = [t for t in transfers if t["amount"] > 0]
        outgoing = [t for t in transfers if t["amount"] < 0]

        total_in = sum(t["amount"] for t in incoming)
        total_out = sum(t["amount"] for t in outgoing)
        balance = total_in + total_out

        method_counts = Counter(t["method"] for t in transfers)
        methods_line = ", ".join(
            f"{name} — {cnt}" for name, cnt in method_counts.most_common()
        )

        excel_bytes = build_excel(transfers)
        out_name = f"transfers_{Path(doc.file_name).stem}.xlsx"

        caption = (
            f"Распознано переводов (всего): {total_recognized}\n"
            f"Отсечено повторяющихся: {skipped_repeated}\n"
            f"В таблице (разовых): {len(transfers)} ({methods_line})\n\n"
            f"Входящие (→ Выводы): {len(incoming)} шт. (+{total_in:,.2f} ₽)\n"
            f"Исходящие (→ Депозиты): {len(outgoing)} шт. ({total_out:,.2f} ₽)\n"
            f"Разница: {balance:+,.2f} ₽"
        ).replace(",", " ")

        await message.answer_document(
            BufferedInputFile(excel_bytes, filename=out_name),
            caption=caption,
            reply_markup=build_action_keyboard(transfers),
        )
        await wait_msg.delete()

    except Exception as e:
        logging.exception("Ошибка при разборе PDF")
        await wait_msg.edit_text(f"Не удалось разобрать файл: {e}")


@dp.callback_query(F.data == "scr:profile")
async def handle_profile_button(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    WAITING_PROFILE.add(callback.from_user.id)
    profile = load_profile()
    await callback.answer()
    await callback.message.answer(profile_prompt_text(profile), parse_mode="Markdown")


@dp.callback_query(F.data.startswith("scr:"))
async def handle_screenshots(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    user_id = callback.from_user.id
    transfers = LAST_TRANSFERS.get(user_id)
    if not transfers:
        await callback.answer(
            "Сначала закинь PDF-выписку заново (бот перезапускался).",
            show_alert=True,
        )
        return

    mode = callback.data.split(":", 1)[1]
    labels = {"deposits": "Депозиты", "withdrawals": "Выводы", "all": "Все"}
    label = labels.get(mode, "История")

    await callback.answer(f"Генерирую скриншоты: {label}...")

    try:
        ui = load_ui_settings()
        pages = build_screenshots(transfers, mode=mode, ui_settings=ui, per_page=7)
        if not pages:
            await callback.message.answer(f"Операций для вкладки «{label}» нет.")
            return

        total_pages = len(pages)
        for idx, png_bytes in enumerate(pages, start=1):
            await callback.message.answer_photo(
                BufferedInputFile(png_bytes, filename=f"{mode}_{idx}.png"),
                caption=f"{label} — скрин {idx}/{total_pages}",
            )
    except Exception as e:
        logging.exception("Ошибка генерации PNG")
        await callback.message.answer(f"Ошибка при создании PNG: {e}")


@dp.message(F.text)
async def handle_profile_text(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    if not message.from_user or message.from_user.id not in WAITING_PROFILE:
        return

    parsed = parse_profile_input(message.text or "")
    if not any(k in parsed for k in ("phone", "email", "dob", "name")):
        await message.answer(
            "Не удалось распознать телефон или почту. Пришли их в одном сообщении "
            "(например: `+79991234567 mail@gmail.com`) или нажми /cancel.",
            parse_mode="Markdown",
        )
        return

    profile = load_profile()
    profile.update(parsed)
    save_profile(profile)
    WAITING_PROFILE.discard(message.from_user.id)

    try:
        ui = load_ui_settings()
        png_bytes = build_profile_screenshot(profile, ui)
        await message.answer_photo(
            BufferedInputFile(png_bytes, filename="profile.png"),
            caption="Скриншот профиля готов (нажми /profile, если нужно переделать).",
        )
    except Exception as e:
        logging.exception("Ошибка генерации скриншота профиля")
        await message.answer(f"Ошибка при создании скриншота профиля: {e}")


async def main():
    if not BOT_TOKEN:
        raise ValueError("Не задан BOT_TOKEN в переменных окружения!")
    bot = Bot(token=BOT_TOKEN)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
