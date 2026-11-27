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
from PIL import Image, ImageDraw, ImageFont

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

DEFAULT_UI_SETTINGS = {
    "domain": "1wujym.life",
    "status_time": "auto",  # "auto" или строка вида "22:21"
    "battery": 47,
}

# Кэш последней разобранной выписки в памяти (по user_id)
LAST_TRANSFERS: dict[int, list[dict]] = {}

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


def is_allowed_user(user: types.User | None) -> bool:
    if ALLOWED_USER_ID == 0:
        return True
    return user is not None and user.id == ALLOWED_USER_ID


# ---------------- РАЗБОР PDF Т-БАНКА ----------------

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
) -> tuple[list[dict], int, int]:
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
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
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

    return results, total_recognized, skipped_repeated


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


# ---------------- ГЕНЕРАЦИЯ PNG-СКРИНШОТОВ ----------------

def find_asset(*names: str) -> Path | None:
    base_dirs = (Path("assets"), Path("."))
    for d in base_dirs:
        for name in names:
            p = d / name
            if p.exists() and p.is_file():
                return p
    return None


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    if bold:
        candidates = (
            "font_bold.ttf",
            "Inter-SemiBold.ttf",
            "Inter-Bold.ttf",
            "Roboto-Bold.ttf",
            "Roboto-Medium.ttf",
            "font_regular.ttf",
        )
        sys_candidates = (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
        )
    else:
        candidates = (
            "font_regular.ttf",
            "Inter-Regular.ttf",
            "Roboto-Regular.ttf",
            "font_bold.ttf",
        )
        sys_candidates = (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        )

    asset_path = find_asset(*candidates)
    if asset_path:
        try:
            return ImageFont.truetype(str(asset_path), size)
        except Exception:
            pass

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


def draw_default_card_icon(size: int = 104) -> Image.Image:
    icon = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(icon)
    # Синий корпус банковской карты
    d.rounded_rectangle([4, 18, size - 4, size - 18], radius=14, fill=(32, 46, 82, 255))
    # Светлая полоска / чип
    d.rounded_rectangle([14, 34, 42, 46], radius=4, fill=(80, 102, 150, 255))
    # Два перекрывающихся круга (красный и оранжевый)
    r = 16
    cy = size - 38
    cx1 = size - 42
    cx2 = size - 26
    d.ellipse([cx1 - r, cy - r, cx1 + r, cy + r], fill=(235, 38, 42, 255))
    d.ellipse([cx2 - r, cy - r, cx2 + r, cy + r], fill=(247, 158, 27, 235))
    return icon


def draw_default_sbp_icon(size: int = 104) -> Image.Image:
    icon = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(icon)
    d.rounded_rectangle([6, 16, size - 6, size - 16], radius=18, fill=(242, 244, 248, 255))
    # Простые цветные треугольники в стиле СБП, если sbp.png ещё не загружен
    d.polygon([(18, 28), (46, 52), (18, 76)], fill=(91, 45, 144, 255))
    d.polygon([(28, 24), (56, 46), (34, 52)], fill=(250, 175, 24, 255))
    d.polygon([(28, 80), (56, 58), (34, 52)], fill=(0, 166, 81, 255))
    f = load_font(22, bold=True)
    d.text((54, 39), "сбп", fill=(20, 24, 35, 255), font=f)
    return icon


def load_icon(kind: str, max_w: int = 114, max_h: int = 84) -> Image.Image:
    if kind == "card":
        p = find_asset("card.png", "card_icon.png")
        if p:
            raw = Image.open(p).convert("RGBA")
        else:
            raw = draw_default_card_icon()
    else:
        p = find_asset("sbp.png", "sbp_icon.png")
        if p:
            raw = Image.open(p).convert("RGBA")
        else:
            raw = draw_default_sbp_icon()

    bbox = raw.getbbox()
    if bbox:
        raw = raw.crop(bbox)
    raw.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
    return raw


def draw_status_bar(draw: ImageDraw.ImageDraw, W: int, time_str: str, battery: int) -> None:
    font_time = load_font(42, bold=True)
    draw.text((160, 38), time_str, fill=(255, 255, 255), font=font_time)

    # Индикатор сигнала (4 столбика)
    bx = W - 340
    by_bottom = 80
    heights = (14, 22, 30, 38)
    for i, h in enumerate(heights):
        x0 = bx + i * 16
        fill = (255, 255, 255) if i < 3 else (110, 114, 122)
        draw.rounded_rectangle([x0, by_bottom - h, x0 + 10, by_bottom], radius=3, fill=fill)

    # Иконка Wi-Fi
    wx, wy = W - 235, 76
    draw.ellipse([wx - 5, wy - 5, wx + 5, wy + 5], fill=(255, 255, 255))
    for r in (16, 28, 40):
        draw.arc([wx - r, wy - r, wx + r, wy + r], start=225, end=315, fill=(255, 255, 255), width=5)

    # Иконка батареи с процентом
    bat_x, bat_y, bat_w, bat_h = W - 165, 42, 76, 38
    draw.rounded_rectangle(
        [bat_x, bat_y, bat_x + bat_w, bat_y + bat_h],
        radius=10,
        outline=(155, 158, 165),
        width=3,
        fill=(45, 48, 54),
    )
    # Контакт справа
    draw.rounded_rectangle(
        [bat_x + bat_w + 2, bat_y + 11, bat_x + bat_w + 8, bat_y + bat_h - 11],
        radius=3,
        fill=(155, 158, 165),
    )
    pct = max(5, min(100, int(battery)))
    fill_w = int((bat_w - 8) * pct / 100)
    fill_color = (255, 255, 255) if pct > 20 else (239, 68, 68)
    draw.rounded_rectangle(
        [bat_x + 4, bat_y + 4, bat_x + 4 + fill_w, bat_y + bat_h - 4],
        radius=7,
        fill=fill_color,
    )
    font_bat = load_font(26, bold=True)
    bat_text = str(pct)
    tb = draw.textbbox((0, 0), bat_text, font=font_bat)
    tw, th = tb[2] - tb[0], tb[3] - tb[1]
    txt_color = (18, 20, 24) if pct >= 40 else (255, 255, 255)
    draw.text(
        (bat_x + (bat_w - tw) / 2, bat_y + (bat_h - th) / 2 - tb[1]),
        bat_text,
        fill=txt_color,
        font=font_bat,
    )


def draw_browser_bar(draw: ImageDraw.ImageDraw, W: int, domain: str) -> None:
    # Серый овал адресной строки
    x0, y0, x1, y1 = 36, 136, W - 36, 246
    draw.rounded_rectangle([x0, y0, x1, y1], radius=55, fill=(58, 60, 66))

    # Иконка камеры/объектива слева
    lx, ly = x0 + 48, (y0 + y1) // 2
    draw.rounded_rectangle([lx - 18, ly - 18, lx + 18, ly + 18], radius=7, outline=(215, 218, 224), width=4)
    draw.ellipse([lx - 6, ly - 6, lx + 6, ly + 6], fill=(215, 218, 224))

    # Домен по центру
    font_dom = load_font(34, bold=False)
    tb = draw.textbbox((0, 0), domain, font=font_dom)
    tw, th = tb[2] - tb[0], tb[3] - tb[1]
    draw.text(((W - tw) / 2, (y0 + y1 - th) / 2 - tb[1]), domain, fill=(225, 228, 232), font=font_dom)

    # Иконка "Поделиться" справа
    sx, sy = x1 - 52, (y0 + y1) // 2
    draw.rounded_rectangle([sx - 16, sy - 6, sx + 16, sy + 20], radius=4, outline=(215, 218, 224), width=4)
    draw.line([(sx, sy + 6), (sx, sy - 20)], fill=(215, 218, 224), width=4)
    draw.line([(sx - 9, sy - 11), (sx, sy - 20), (sx + 9, sy - 11)], fill=(215, 218, 224), width=4)


def draw_checkmark(draw: ImageDraw.ImageDraw, x: int, y: int, color=(46, 158, 91)) -> None:
    draw.line([(x, y + 14), (x + 10, y + 24), (x + 26, y + 6)], fill=color, width=5, joint="curve")


def build_screenshots(
    transfers: list[dict],
    mode: str,
    ui_settings: dict,
    per_page: int = 7,
) -> list[bytes]:
    """
    mode:
      - 'deposits': исходящие из выписки (amount < 0) -> показываются как '+' Депозиты
      - 'withdrawals': входящие из выписки (amount > 0) -> показываются как '-' Выводы
      - 'all': все операции (исходящие -> '+', входящие -> '-')
    """
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
    domain = ui_settings.get("domain", "1wujym.life")
    battery = int(ui_settings.get("battery", 47))
    status_time = ui_settings.get("status_time", "auto")
    if status_time == "auto":
        msk = timezone(timedelta(hours=3))
        status_time = datetime.now(msk).strftime("%H:%M")

    sbp_icon = load_icon("sbp")
    card_icon = load_icon("card")

    font_back = load_font(40, bold=False)
    font_title = load_font(62, bold=True)
    font_tab = load_font(38, bold=True)
    font_row_title = load_font(41, bold=True)
    font_status = load_font(33, bold=False)
    font_amount = load_font(41, bold=True)
    font_meta = load_font(32, bold=False)

    pages: list[bytes] = []

    for start_idx in range(0, len(filtered), per_page):
        chunk = filtered[start_idx : start_idx + per_page]

        img = Image.new("RGB", (W, H), (14, 16, 20))
        draw = ImageDraw.Draw(img)

        draw_status_bar(draw, W, status_time, battery)
        draw_browser_bar(draw, W, domain)

        # Белая подложка с закруглёнными верхними углами
        sheet_top = 285
        draw.rounded_rectangle([0, sheet_top, W, H + 80], radius=64, fill=(255, 255, 255))

        # Кнопка "< Назад"
        bx, by = 64, sheet_top + 66
        draw.line([(bx + 18, by), (bx, by + 18), (bx + 18, by + 36)], fill=(27, 102, 224), width=6, joint="curve")
        draw.text((bx + 40, by - 6), "Назад", fill=(27, 102, 224), font=font_back)

        # Крестик справа
        cx, cy = W - 82, sheet_top + 84
        draw.line([(cx - 18, cy - 18), (cx + 18, cy + 18)], fill=(45, 52, 66), width=5)
        draw.line([(cx + 18, cy - 18), (cx - 18, cy + 18)], fill=(45, 52, 66), width=5)

        # Заголовок "История платежей"
        draw.text((56, sheet_top + 180), "История платежей", fill=(17, 24, 39), font=font_title)

        # Переключатель вкладок: Все / Депозиты / Выводы
        tab_y0, tab_y1 = sheet_top + 305, sheet_top + 425
        tab_x0, tab_x1 = 56, W - 56
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

        # Отрисовка операций (до 7 на экран)
        list_top = tab_y1 + 48
        row_h = 248

        for i, row in enumerate(chunk):
            ry = list_top + i * row_h

            # Инверсия: исходящий из выписки (<0) = Депозит (+), входящий (>0) = Вывод (-)
            is_deposit = row["amount"] < 0
            is_card = row["method"] == "По карте"

            icon = card_icon if is_card else sbp_icon
            ix = 56 + (116 - icon.width) // 2
            iy = ry + 64 + (84 - icon.height) // 2
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

            # Название (2 строки)
            tx = 204
            draw.text((tx, ry + 24), title_lines[0], fill=(17, 24, 39), font=font_row_title)
            draw.text((tx, ry + 78), title_lines[1], fill=(17, 24, 39), font=font_row_title)

            # Статус "✓ Выполнено"
            draw_checkmark(draw, tx + 2, ry + 144)
            draw.text((tx + 42, ry + 140), "Выполнено", fill=(46, 158, 91), font=font_status)

            # Правая колонка: сумма, дата, время с секундами
            right_x = W - 56

            ab = draw.textbbox((0, 0), amount_str, font=font_amount)
            draw.text((right_x - (ab[2] - ab[0]), ry + 26), amount_str, fill=(17, 24, 39), font=font_amount)

            date_str = row["date"]
            db = draw.textbbox((0, 0), date_str, font=font_meta)
            draw.text((right_x - (db[2] - db[0]), ry + 94), date_str, fill=(156, 163, 175), font=font_meta)

            time_str = make_seconds(row)
            tb_t = draw.textbbox((0, 0), time_str, font=font_meta)
            draw.text((right_x - (tb_t[2] - tb_t[0]), ry + 142), time_str, fill=(156, 163, 175), font=font_meta)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        pages.append(buf.getvalue())

    return pages


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
    return InlineKeyboardMarkup(inline_keyboard=buttons)


# ---------------- КОМАНДЫ И ОБРАБОТЧИКИ ----------------

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    ui = load_ui_settings()
    await message.answer(
        "Кидай PDF-выписку из Т-Банка — верну Excel и кнопки для генерации PNG-скриншотов.\n"
        "• Исходящие из выписки → Депозиты (+)\n"
        "• Входящие из выписки → Выводы (-)\n\n"
        f"Текущие настройки скрина: домен `{ui['domain']}`, часы `{ui['status_time']}`, заряд `{ui['battery']}%`\n\n"
        "Команды:\n"
        "/domain 1wujym.life — сменить домен в строке браузера\n"
        "/status 22:21 47 — задать часы и % батареи (или /status auto 65)\n"
        "/excluded — список игнорируемых договоров\n"
        "/exclude 1234567890 — добавить договор в исключения\n"
        "/unexclude 1234567890 — убрать из исключений",
        parse_mode="Markdown",
    )


@dp.message(Command("domain"))
async def cmd_domain(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
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
    parts = (message.text or "").split()
    if len(parts) < 2:
        await message.answer("Формат: /status 22:21 47 (или /status auto 54)")
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
        transfers, total_recognized, skipped_repeated = parse_tbank_pdf(
            pdf_bytes, excluded
        )

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


async def main():
    if not BOT_TOKEN:
        raise ValueError("Не задан BOT_TOKEN в переменных окружения!")
    bot = Bot(token=BOT_TOKEN)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
