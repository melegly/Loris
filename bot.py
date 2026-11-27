import asyncio
from collections import Counter
import copy
from datetime import datetime, timedelta, timezone
import html
import io
import json
import logging
import os
from pathlib import Path
import re
import secrets
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
EP_SETTINGS_FILE = Path("ep_settings.json")
AGENTS_FILE = Path("agents_templates.json")

DEFAULT_UI_SETTINGS = {
    "domain": "1wujym.life",
    "status_time": "auto",  # "auto" или строка вида "22:21"
    "battery": 76,
}

DEFAULT_PROFILE = {
    "name": "",
    "dob": "",
    "passport": "",
    "country": "Россия",
    "phone": "",
    "email": "",
}

DEFAULT_EP = {
    "owner": "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ «СИСТЕМА ПБО»",
    "cert": "03861D760091B449A9496EFF4FE4488A",
    "ca": "Федеральная налоговая служба",
    "valid_from": "24.01.2026",
    "valid_to": "24.10.2027",
    "signed_date": "auto",  # "auto" (дата бланка/сегодня) или "ДД.ММ.ГГГГ"
    "position": "center",  # "left", "center", "right"
}

# Кэш операций, бланков и активных режимов ввода (по user_id)
LAST_TRANSFERS: dict[int, list[dict]] = {}
LAST_TRANSFERS_ORIG: dict[int, list[dict]] = {}
LAST_PDF_STEM: dict[int, str] = {}
LAST_BLANK_PAGES: dict[int, list[bytes]] = {}
LAST_BLANK_META: dict[int, dict] = {}
LAST_SIGNED_PNG: dict[int, bytes] = {}

WAITING_PROFILE: set[int] = set()
WAITING_AGENT_ADD: set[int] = set()
EDIT_MODE: dict[int, str] = {}
BLANK_SESSIONS: dict[int, dict] = {}
EP_SESSIONS: dict[int, dict] = {}

BLANK_STEPS = (
    "year_date",
    "agent_name",
    "agent_details",
    "fio",
    "person_inn",
    "person_meta",
    "passport",
    "tax_rate",
    "incomes",
    "deductions",
    "totals",
)

EP_STEPS = (
    "owner",
    "cert",
    "ca",
    "dates",
    "position",
)

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


# ---------------- ХРАНЕНИЕ НАСТРОЕК И ШАБЛОНОВ ----------------

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


def load_ep_settings() -> dict:
    ep = dict(DEFAULT_EP)
    if EP_SETTINGS_FILE.exists():
        try:
            data = json.loads(EP_SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                ep.update(data)
        except Exception:
            pass
    if not ep.get("cert"):
        ep["cert"] = generate_cert_hex()
    return ep


def save_ep_settings(ep: dict) -> None:
    EP_SETTINGS_FILE.write_text(
        json.dumps(ep, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_agents() -> list[dict]:
    if AGENTS_FILE.exists():
        try:
            data = json.loads(AGENTS_FILE.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []


def save_agents(agents: list[dict]) -> None:
    AGENTS_FILE.write_text(
        json.dumps(agents, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def upsert_agent_template(agent_data: dict) -> dict:
    agents = load_agents()
    name = (agent_data.get("agent_name") or "").strip()
    name_low = name.lower()
    existing = next(
        (a for a in agents if (a.get("agent_name") or "").strip().lower() == name_low),
        None,
    )
    if existing:
        existing.update(agent_data)
        save_agents(agents)
        return existing

    new_id = secrets.token_hex(3)
    entry = dict(agent_data)
    entry["id"] = new_id
    agents.append(entry)
    save_agents(agents)
    return entry


def generate_cert_hex() -> str:
    return secrets.token_hex(16).upper()


def clear_active_inputs(user_id: int | None) -> None:
    if user_id is None:
        return
    WAITING_PROFILE.discard(user_id)
    WAITING_AGENT_ADD.discard(user_id)
    EDIT_MODE.pop(user_id, None)
    BLANK_SESSIONS.pop(user_id, None)
    EP_SESSIONS.pop(user_id, None)


def is_allowed_user(user: types.User | None) -> bool:
    if ALLOWED_USER_ID == 0:
        return True
    return user is not None and user.id == ALLOWED_USER_ID


def now_msk() -> datetime:
    return datetime.now(timezone(timedelta(hours=3)))


# ---------------- РАЗБОР PDF Т-БАНКА И ОПРЕДЕЛЕНИЕ ТИПА PDF ----------------

def is_tbank_statement_pdf(pdf_bytes: bytes) -> bool:
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if not pdf.pages:
                return False
            text = (pdf.pages[0].extract_text() or "").lower()
            if "справка о движении" in text:
                return True
            if "движение средств" in text and any(
                k in text for k in ("тбанк", "тинькофф", "tbank")
            ):
                return True
    except Exception:
        pass
    return False


def render_pdf_to_png_pages(pdf_bytes: bytes, resolution: int = 200) -> list[bytes]:
    pages_png: list[bytes] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            pil_img = page.to_image(resolution=resolution).original.convert("RGB")
            buf = io.BytesIO()
            pil_img.save(buf, format="PNG")
            pages_png.append(buf.getvalue())
    return pages_png


def images_to_pdf_bytes(png_list: list[bytes], resolution: float = 200.0) -> bytes:
    imgs: list[Image.Image] = []
    for b in png_list:
        raw = Image.open(io.BytesIO(b))
        if raw.mode in ("RGBA", "LA") or (
            raw.mode == "P" and "transparency" in raw.info
        ):
            bg = Image.new("RGB", raw.size, (255, 255, 255))
            rgba = raw.convert("RGBA")
            bg.paste(rgba, (0, 0), rgba)
            imgs.append(bg)
        else:
            imgs.append(raw.convert("RGB"))
    out = io.BytesIO()
    imgs[0].save(
        out,
        format="PDF",
        resolution=resolution,
        save_all=True,
        append_images=imgs[1:],
    )
    return out.getvalue()


def extract_client_name(first_page_text: str) -> str:
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


def extract_client_passport(first_page_text: str) -> str:
    """Извлекает серию и номер паспорта из шапки PDF-выписки Т-Банка."""
    header = re.split(
        r"движение\s+средств", first_page_text, maxsplit=1, flags=re.IGNORECASE
    )[0]
    ser_m = re.search(r"серия\s*:\s*(\d{2}\s?\d{2})\b", header, re.IGNORECASE)
    num_m = re.search(r"номер\s*:\s*(\d{6})\b", header, re.IGNORECASE)
    if ser_m and num_m:
        ser = re.sub(r"\s+", "", ser_m.group(1))
        return f"{ser} {num_m.group(1)}"
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
) -> tuple[list[dict], int, int, str, str]:
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
    client_passport = ""

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if page_idx == 0:
                client_name = extract_client_name(text)
                client_passport = extract_client_passport(text)
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

    return results, total_recognized, skipped_repeated, client_name, client_passport


def parse_profile_input(text: str) -> dict[str, str]:
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


def build_transfers_caption(
    transfers: list[dict],
    total_recognized: int | None = None,
    skipped_repeated: int | None = None,
) -> str:
    incoming = [t for t in transfers if t["amount"] > 0]
    outgoing = [t for t in transfers if t["amount"] < 0]
    total_in = sum(t["amount"] for t in incoming)
    total_out = sum(t["amount"] for t in outgoing)
    balance = total_in + total_out

    method_counts = Counter(t["method"] for t in transfers)
    methods_line = ", ".join(
        f"{name} — {cnt}" for name, cnt in method_counts.most_common()
    )

    lines: list[str] = []
    if total_recognized is not None and skipped_repeated is not None:
        lines.append(f"Распознано переводов (всего): {total_recognized}")
        lines.append(f"Отсечено повторяющихся: {skipped_repeated}")
    else:
        lines.append("Обновлённый список операций")

    lines.append(f"В таблице: {len(transfers)} ({methods_line})")
    lines.append("")
    lines.append(f"Входящие (→ Выводы): {len(incoming)} шт. (+{total_in:,.2f} ₽)")
    lines.append(f"Исходящие (→ Депозиты): {len(outgoing)} шт. ({total_out:,.2f} ₽)")
    lines.append(f"Разница: {balance:+,.2f} ₽")
    return "\n".join(lines).replace(",", " ")


# ---------------- ПОИСК РЕСУРСОВ И ОБЩИЕ ХЕЛПЕРЫ PNG ----------------

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


def draw_home_indicator(draw: ImageDraw.ImageDraw, W: int, H: int) -> None:
    """Нижняя полоса Home Indicator в стиле iPhone."""
    bar_w = 420
    bar_h = 15
    x0 = (W - bar_w) // 2
    y0 = H - 38
    draw.rounded_rectangle(
        [x0, y0, x0 + bar_w, y0 + bar_h],
        radius=bar_h // 2,
        fill=(22, 24, 28),
    )


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


def get_status_time_str(ui_settings: dict, minute_offset: int = 0) -> str:
    status_time = ui_settings.get("status_time", "auto")
    if status_time == "auto":
        dt = now_msk() + timedelta(minutes=minute_offset)
        return dt.strftime("%H:%M")
    m = re.fullmatch(r"(\d{1,2}):(\d{2})", str(status_time).strip())
    if m:
        hh, mm = int(m.group(1)), int(m.group(2))
        total_min = (hh * 60 + mm + minute_offset) % (24 * 60)
        return f"{total_min // 60:02d}:{total_min % 60:02d}"
    return str(status_time)


def draw_sheet_header(
    img: Image.Image,
    draw: ImageDraw.ImageDraw,
    W: int,
    H: int,
    title: str,
    ui_settings: dict,
    minute_offset: int = 0,
) -> int:
    domain = ui_settings.get("domain", "1wujym.life")
    battery = int(ui_settings.get("battery", 76))
    status_time = get_status_time_str(ui_settings, minute_offset=minute_offset)

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


# ---------------- ГЕНЕРАЦИЯ СКРИНШОТОВ ОПЕРАЦИЙ И ПРОФИЛЯ ----------------

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

    for page_idx, start_idx in enumerate(range(0, len(filtered), per_page)):
        chunk = filtered[start_idx : start_idx + per_page]

        img = Image.new("RGB", (W, H), (14, 16, 20))
        draw = ImageDraw.Draw(img)
        # Автосмещение времени на +1 мин на каждый следующий скрин в серии
        sheet_top = draw_sheet_header(
            img,
            draw,
            W,
            H,
            "История платежей",
            ui_settings,
            minute_offset=page_idx,
        )

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

        list_top = tab_y1 + 32
        row_h = 244

        for i, row in enumerate(chunk):
            ry = list_top + i * row_h

            is_deposit = row["amount"] < 0
            is_card = row["method"] == "По карте"

            icon = card_icon if is_card else sbp_icon
            ix = 44 + (142 - icon.width) // 2
            iy = ry + 46 + (92 - icon.height) // 2
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
            draw.text((tx, ry + 16), title_lines[0], fill=(17, 24, 39), font=font_row_title)
            draw.text((tx, ry + 75), title_lines[1], fill=(17, 24, 39), font=font_row_title)

            draw_checkmark(draw, tx + 2, ry + 144)
            draw.text((tx + 46, ry + 139), "Выполнено", fill=(46, 158, 91), font=font_status)

            right_x = W - 52

            ab = draw.textbbox((0, 0), amount_str, font=font_amount)
            draw.text((right_x - (ab[2] - ab[0]), ry + 18), amount_str, fill=(17, 24, 39), font=font_amount)

            date_str = row["date"]
            db = draw.textbbox((0, 0), date_str, font=font_meta)
            draw.text((right_x - (db[2] - db[0]), ry + 90), date_str, fill=(142, 149, 160), font=font_meta)

            time_str = make_seconds(row)
            tb_t = draw.textbbox((0, 0), time_str, font=font_meta)
            draw.text((right_x - (tb_t[2] - tb_t[0]), ry + 142), time_str, fill=(142, 149, 160), font=font_meta)

            if i < len(chunk) - 1:
                sep_y = ry + row_h - 14
                draw.line([(tx, sep_y), (right_x, sep_y)], fill=(242, 244, 247), width=2)

        draw_home_indicator(draw, W, H)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        pages.append(buf.getvalue())

    return pages


def build_profile_screenshot(profile: dict, ui_settings: dict) -> bytes:
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
        lb = draw.textbbox((0, 0), label, font=font_label)
        lw, lh = lb[2] - lb[0], lb[3] - lb[1]
        draw.text(
            (card_x0 + 44, y_center - lh / 2 - lb[1]),
            label,
            fill=text_dark,
            font=font_label,
        )

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

    draw_home_indicator(draw, W, H)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ---------------- ГЕНЕРАЦИЯ БЛАНКА КНД 1175018 (A4 PNG/PDF) ----------------

def fmt_money_blank(val: float | str) -> str:
    if isinstance(val, str):
        return val.strip()
    return f"{val:,.2f}".replace(",", " ")


def fmt_int_blank(val: float | int | str) -> str:
    if isinstance(val, str):
        return val.strip()
    return f"{int(round(val)):,}".replace(",", " ")


def draw_dashed_line(
    draw: ImageDraw.ImageDraw,
    x0: int,
    x1: int,
    y: int,
    color=(25, 25, 25),
    width: int = 2,
    dash: int = 7,
    gap: int = 4,
) -> None:
    x = x0
    while x < x1:
        end = min(x + dash, x1)
        draw.line([(x, y), (end, y)], fill=color, width=width)
        x += dash + gap


def draw_text_with_sup(
    draw: ImageDraw.ImageDraw,
    x: int,
    y: int,
    text: str,
    sup: str,
    font_main,
    font_sup,
    color=(15, 15, 15),
) -> int:
    draw.text((x, y), text, fill=color, font=font_main)
    tw = draw.textbbox((0, 0), text, font=font_main)[2]
    if sup:
        draw.text((x + tw + 1, y - 5), sup, fill=color, font=font_sup)
        sw = draw.textbbox((0, 0), sup, font=font_sup)[2]
        return x + tw + sw + 2
    return x + tw


def draw_fit_text_on_line(
    draw: ImageDraw.ImageDraw,
    text: str,
    x0: int,
    x1: int,
    line_y: int,
    base_size: int = 22,
    bold: bool = False,
    align: str = "center",
    color=(15, 15, 15),
) -> None:
    if not text:
        return
    max_w = max(20, (x1 - x0) - 8)
    size = base_size
    f = load_font(size, bold=bold)
    bb = draw.textbbox((0, 0), text, font=f)
    while (bb[2] - bb[0]) > max_w and size > 13:
        size -= 1
        f = load_font(size, bold=bold)
        bb = draw.textbbox((0, 0), text, font=f)
    tw = bb[2] - bb[0]
    th = bb[3] - bb[1]
    if align == "left":
        tx = x0 + 6
    elif align == "right":
        tx = x1 - tw - 6
    else:
        tx = x0 + (x1 - x0 - tw) / 2
    ty = line_y - th - 5 - bb[1]
    draw.text((tx, ty), text, fill=color, font=f)


def draw_cell_text(
    draw: ImageDraw.ImageDraw,
    lines: list[str],
    x0: int,
    y0: int,
    x1: int,
    y1: int,
    base_size: int = 20,
    bold: bool = False,
    align: str = "center",
    pad_x: int = 8,
    color=(15, 15, 15),
) -> None:
    non_empty = [ln for ln in lines if ln]
    if not non_empty:
        return
    max_w = max(20, (x1 - x0) - 2 * pad_x)
    size = base_size
    f = load_font(size, bold=bold)
    while size > 12:
        widest = max(draw.textbbox((0, 0), ln, font=f)[2] for ln in non_empty)
        if widest <= max_w:
            break
        size -= 1
        f = load_font(size, bold=bold)

    line_gap = 4
    heights = []
    bboxes = []
    for ln in non_empty:
        bb = draw.textbbox((0, 0), ln, font=f)
        bboxes.append(bb)
        heights.append(bb[3] - bb[1])

    total_h = sum(heights) + line_gap * (len(non_empty) - 1)
    cur_y = y0 + (y1 - y0 - total_h) / 2

    for ln, bb, lh in zip(non_empty, bboxes, heights):
        tw = bb[2] - bb[0]
        if align == "left":
            tx = x0 + pad_x
        elif align == "right":
            tx = x1 - pad_x - tw
        else:
            tx = x0 + (x1 - x0 - tw) / 2
        draw.text((tx, cur_y - bb[1]), ln, fill=color, font=f)
        cur_y += lh + line_gap


def build_knd_1175018_png(data: dict) -> bytes:
    """Отрисовывает заполненный бланк КНД 1175018 (A4, 1654x2339 px)."""
    W, H = 1654, 2339
    img = Image.new("RGB", (W, H), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    black = (18, 18, 18)
    ML = 90
    MR = 1564

    f_app = load_font(17, bold=False)
    f_knd = load_font(26, bold=True)
    f_title = load_font(30, bold=True)
    f_sub = load_font(23, bold=True)
    f_sec = load_font(24, bold=True)
    f_lbl = load_font(21, bold=False)
    f_sup = load_font(13, bold=False)

    app_x = 1235
    draw.text((app_x, 78), "Приложение № 4", fill=black, font=f_app)
    draw.text((app_x, 102), "к приказу ФНС России", fill=black, font=f_app)
    draw.text((app_x, 126), "от 09.01.2024 № ЕД-7-11/1@", fill=black, font=f_app)

    draw.text((125, 174), "Форма по КНД 1175018", fill=black, font=f_knd)

    title = "СПРАВКА О ДОХОДАХ И СУММАХ НАЛОГА ФИЗИЧЕСКОГО ЛИЦА"
    tb = draw.textbbox((0, 0), title, font=f_title)
    draw.text(((W - (tb[2] - tb[0])) / 2, 224), title, fill=black, font=f_title)

    year_str = data.get("year", str(now_msk().year))
    issue_date = data.get("issue_date", now_msk().strftime("%d.%m.%Y"))
    d_parts = issue_date.split(".") if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", issue_date) else ["", "", ""]

    sy = 278
    s_line_y = sy + 26
    draw.text((590, sy), "за", fill=black, font=f_sub)
    draw.line([(628, s_line_y), (722, s_line_y)], fill=black, width=2)
    draw_fit_text_on_line(draw, year_str, 628, 722, s_line_y, base_size=22, bold=True)

    draw.text((734, sy), "год", fill=black, font=f_sub)
    draw.text((796, sy), "от", fill=black, font=f_sub)

    draw.line([(836, s_line_y), (886, s_line_y)], fill=black, width=2)
    draw_fit_text_on_line(draw, d_parts[0], 836, 886, s_line_y, base_size=22, bold=True)
    draw.text((890, sy), ".", fill=black, font=f_sub)

    draw.line([(902, s_line_y), (952, s_line_y)], fill=black, width=2)
    draw_fit_text_on_line(draw, d_parts[1], 902, 952, s_line_y, base_size=22, bold=True)
    draw.text((956, sy), ".", fill=black, font=f_sub)

    draw.line([(968, s_line_y), (1064, s_line_y)], fill=black, width=2)
    draw_fit_text_on_line(draw, d_parts[2], 968, 1064, s_line_y, base_size=22, bold=True)

    # ================= РАЗДЕЛ 1 =================
    draw.text((ML, 336), "1. Сведения о налоговом агенте", fill=black, font=f_sec)

    y = 380
    ly = y + 25
    x_end = draw_text_with_sup(draw, ML, y, "Код по ОКТМО", "1", f_lbl, f_sup, black)
    draw_dashed_line(draw, x_end + 8, 492, ly)
    draw_fit_text_on_line(draw, data.get("oktmo", ""), x_end + 8, 492, ly, base_size=21)

    draw.text((512, y), "Телефон", fill=black, font=f_lbl)
    draw_dashed_line(draw, 610, 855, ly)
    draw_fit_text_on_line(draw, data.get("agent_phone", ""), 610, 855, ly, base_size=21)

    x_end = draw_text_with_sup(draw, 875, y, "ИНН", "2", f_lbl, f_sup, black)
    draw_dashed_line(draw, x_end + 8, 1190, ly)
    draw_fit_text_on_line(draw, data.get("agent_inn", ""), x_end + 8, 1190, ly, base_size=21)

    x_end = draw_text_with_sup(draw, 1210, y, "КПП", "3", f_lbl, f_sup, black)
    draw_dashed_line(draw, x_end + 8, MR, ly)
    draw_fit_text_on_line(draw, data.get("agent_kpp", ""), x_end + 8, MR, ly, base_size=21)

    y = 424
    ly = y + 25
    draw.text((ML, y), "Налоговый агент", fill=black, font=f_lbl)
    draw_dashed_line(draw, 282, MR, ly)
    draw_fit_text_on_line(
        draw, data.get("agent_name", ""), 282, MR, ly, base_size=22, align="left"
    )

    y = 468
    ly = y + 25
    draw.text((ML, y), "Форма реорганизации (ликвидация) (код)", fill=black, font=f_lbl)
    draw_dashed_line(draw, 542, 650, ly)
    draw_fit_text_on_line(draw, data.get("reorg_code", ""), 542, 650, ly, base_size=21)

    y = 512
    ly = y + 25
    x1 = draw_text_with_sup(draw, ML, y, "ИНН", "2", f_lbl, f_sup, black)
    x2 = draw_text_with_sup(draw, x1 + 2, y, "/КПП", "3", f_lbl, f_sup, black)
    draw.text((x2 + 6, y), "реорганизованной организации", fill=black, font=f_lbl)
    draw_dashed_line(draw, 542, 775, ly)
    draw.text((783, y), "/", fill=black, font=f_lbl)
    draw_dashed_line(draw, 800, 1020, ly)

    # ================= РАЗДЕЛ 2 =================
    draw.text(
        (ML, 568),
        "2. Сведения о физическом лице - получателе дохода",
        fill=black,
        font=f_sec,
    )

    y = 612
    ly = y + 25
    x_end = draw_text_with_sup(draw, ML, y, "ИНН", "2", f_lbl, f_sup, black)
    draw_dashed_line(draw, x_end + 10, 445, ly)
    draw_fit_text_on_line(draw, data.get("person_inn", ""), x_end + 10, 445, ly, base_size=21)

    y = 656
    ly = y + 25
    draw.text((ML, y), "Фамилия", fill=black, font=f_lbl)
    draw_dashed_line(draw, 198, 610, ly)
    draw_fit_text_on_line(draw, data.get("last_name", ""), 198, 610, ly, base_size=22)

    draw.text((626, y), "Имя", fill=black, font=f_lbl)
    draw_dashed_line(draw, 682, 1082, ly)
    draw_fit_text_on_line(draw, data.get("first_name", ""), 682, 1082, ly, base_size=22)

    x_end = draw_text_with_sup(draw, 1098, y, "Отчество", "4", f_lbl, f_sup, black)
    draw_dashed_line(draw, x_end + 8, MR, ly)
    draw_fit_text_on_line(draw, data.get("middle_name", ""), x_end + 8, MR, ly, base_size=22)

    y = 700
    ly = y + 25
    draw.text((ML, y), "Статус налогоплательщика", fill=black, font=f_lbl)
    draw_dashed_line(draw, 382, 466, ly)
    draw_fit_text_on_line(draw, data.get("tax_status", ""), 382, 466, ly, base_size=21)

    draw.text((494, y), "Дата рождения", fill=black, font=f_lbl)
    dob = data.get("dob", "")
    dob_parts = dob.split(".") if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", dob) else ["", "", ""]

    draw_dashed_line(draw, 668, 718, ly)
    draw_fit_text_on_line(draw, dob_parts[0], 668, 718, ly, base_size=21)
    draw.text((721, y), ".", fill=black, font=f_lbl)

    draw_dashed_line(draw, 732, 782, ly)
    draw_fit_text_on_line(draw, dob_parts[1], 732, 782, ly, base_size=21)
    draw.text((785, y), ".", fill=black, font=f_lbl)

    draw_dashed_line(draw, 796, 892, ly)
    draw_fit_text_on_line(draw, dob_parts[2], 796, 892, ly, base_size=21)

    draw.text((922, y), "Гражданство (код страны)", fill=black, font=f_lbl)
    draw_dashed_line(draw, 1212, 1345, ly)
    draw_fit_text_on_line(draw, data.get("country_code", ""), 1212, 1345, ly, base_size=21)

    y = 744
    ly = y + 25
    draw.text((ML, y), "Код документа, удостоверяющего личность:", fill=black, font=f_lbl)
    draw_dashed_line(draw, 558, 678, ly)
    draw_fit_text_on_line(draw, data.get("doc_code", ""), 558, 678, ly, base_size=21)

    draw.text((706, y), "Серия и номер документа", fill=black, font=f_lbl)
    draw_dashed_line(draw, 992, MR, ly)
    draw_fit_text_on_line(draw, data.get("doc_number", ""), 992, MR, ly, base_size=21)

    # ================= РАЗДЕЛ 3 =================
    y3 = 802
    draw.text((ML, y3), "3. Доходы, облагаемые по ставке", fill=black, font=f_sec)
    rate_ly = y3 + 27
    draw_dashed_line(draw, 540, 622, rate_ly)
    draw_fit_text_on_line(draw, str(data.get("tax_rate", "13")), 540, 622, rate_ly, base_size=22, bold=True)
    draw.text((632, y3), "%", fill=black, font=f_sec)

    tbl3_top = 846
    hdr3_h = 62
    row3_h = 37
    rows_per_side = 12
    col3_widths = (94, 106, 200, 106, 209)
    col3_headers = (
        ["Месяц"],
        ["Код", "дохода"],
        ["Сумма дохода"],
        ["Код", "вычета"],
        ["Сумма вычета"],
    )

    incomes: list[dict] = data.get("incomes", [])
    table_starts = (ML, ML + 715 + 44)

    for side_idx, tx0 in enumerate(table_starts):
        tx1 = tx0 + sum(col3_widths)
        ty_bottom = tbl3_top + hdr3_h + row3_h * rows_per_side

        draw.rectangle([tx0, tbl3_top, tx1, ty_bottom], outline=black, width=2)
        draw.line([(tx0, tbl3_top + hdr3_h), (tx1, tbl3_top + hdr3_h)], fill=black, width=2)
        for r in range(1, rows_per_side):
            ry = tbl3_top + hdr3_h + r * row3_h
            draw.line([(tx0, ry), (tx1, ry)], fill=black, width=2)

        cx = tx0
        col_bounds: list[tuple[int, int]] = []
        for c_idx, cw in enumerate(col3_widths):
            col_bounds.append((cx, cx + cw))
            if c_idx > 0:
                draw.line([(cx, tbl3_top), (cx, ty_bottom)], fill=black, width=2)
            draw_cell_text(
                draw,
                col3_headers[c_idx],
                cx,
                tbl3_top,
                cx + cw,
                tbl3_top + hdr3_h,
                base_size=20,
                bold=False,
            )
            cx += cw

        for r in range(rows_per_side):
            rec_idx = side_idx * rows_per_side + r
            if rec_idx >= len(incomes):
                break
            rec = incomes[rec_idx]
            ry0 = tbl3_top + hdr3_h + r * row3_h
            ry1 = ry0 + row3_h
            vals = (
                rec.get("month", ""),
                rec.get("code", ""),
                fmt_money_blank(rec["amount"]) if rec.get("amount") not in (None, "") else "",
                rec.get("ded_code", ""),
                fmt_money_blank(rec["ded_amount"]) if rec.get("ded_amount") not in (None, "", 0) else "",
            )
            aligns = ("center", "center", "right", "center", "right")
            for c_idx, (bx0, bx1) in enumerate(col_bounds):
                if vals[c_idx]:
                    draw_cell_text(
                        draw,
                        [vals[c_idx]],
                        bx0,
                        ry0,
                        bx1,
                        ry1,
                        base_size=20,
                        align=aligns[c_idx],
                        pad_x=10,
                    )

    # ================= РАЗДЕЛ 4 =================
    y4 = 1392
    draw.text(
        (ML, y4),
        "4. Стандартные, социальные и имущественные налоговые вычеты",
        fill=black,
        font=f_sec,
    )

    tbl4_top = y4 + 42
    hdr4_h = 60
    row4_h = 42
    tbl4_bottom = tbl4_top + hdr4_h + row4_h
    col4_widths = (118, 250, 118, 250, 118, 250, 118, 252)

    draw.rectangle([ML, tbl4_top, MR, tbl4_bottom], outline=black, width=2)
    draw.line([(ML, tbl4_top + hdr4_h), (MR, tbl4_top + hdr4_h)], fill=black, width=2)

    cx = ML
    col4_bounds: list[tuple[int, int]] = []
    for c_idx, cw in enumerate(col4_widths):
        col4_bounds.append((cx, cx + cw))
        if c_idx > 0:
            draw.line([(cx, tbl4_top), (cx, tbl4_bottom)], fill=black, width=2)
        hdr_lines = ["Код", "вычета"] if c_idx % 2 == 0 else ["Сумма вычета"]
        draw_cell_text(draw, hdr_lines, cx, tbl4_top, cx + cw, tbl4_top + hdr4_h, base_size=20)
        cx += cw

    deductions: list[dict] = data.get("deductions", [])
    for d_idx, ded in enumerate(deductions[:4]):
        c_code = d_idx * 2
        c_sum = d_idx * 2 + 1
        bx0, bx1 = col4_bounds[c_code]
        draw_cell_text(draw, [ded.get("code", "")], bx0, tbl4_top + hdr4_h, bx1, tbl4_bottom, base_size=20)
        sx0, sx1 = col4_bounds[c_sum]
        val_s = fmt_money_blank(ded["amount"]) if ded.get("amount") not in (None, "") else ""
        draw_cell_text(draw, [val_s], sx0, tbl4_top + hdr4_h, sx1, tbl4_bottom, base_size=20, align="right", pad_x=12)

    # ================= РАЗДЕЛ 5 =================
    y5 = 1574
    draw.text((ML, y5), "5. Общая сумма дохода и сумма налога", fill=black, font=f_sec)

    tbl5_top = y5 + 42
    row5_h = 68
    tbl5_bottom = tbl5_top + row5_h * 4
    c5_x = (ML, ML + 420, ML + 690, ML + 1204, MR)

    draw.rectangle([ML, tbl5_top, MR, tbl5_bottom], outline=black, width=2)
    for r in range(1, 4):
        ry = tbl5_top + r * row5_h
        draw.line([(ML, ry), (MR, ry)], fill=black, width=2)
    for vx in c5_x[1:-1]:
        draw.line([(vx, tbl5_top), (vx, tbl5_bottom)], fill=black, width=2)

    totals = data.get("totals", {})
    sec5_rows = (
        (
            ["Общая сумма дохода"],
            totals.get("total_income", ""),
            ["Налоговая база"],
            totals.get("tax_base", ""),
        ),
        (
            ["Сумма налога исчисленная"],
            totals.get("tax_calc", ""),
            ["Сумма фиксированных авансовых", "платежей"],
            totals.get("fixed_adv", ""),
        ),
        (
            ["Сумма налога на прибыль", "организаций, подлежащая зачету"],
            totals.get("profit_tax", ""),
            ["Сумма налога, исчисленная и уплаченная в", "иностранном государстве"],
            totals.get("foreign_tax", ""),
        ),
        (
            ["Сумма налога удержанная"],
            totals.get("tax_withheld", ""),
            ["Сумма налога, излишне удержанная", "налоговым агентом"],
            totals.get("tax_over", ""),
        ),
    )

    for r_idx, (l_lbl, l_val, r_lbl, r_val) in enumerate(sec5_rows):
        ry0 = tbl5_top + r_idx * row5_h
        ry1 = ry0 + row5_h
        draw_cell_text(draw, l_lbl, c5_x[0], ry0, c5_x[1], ry1, base_size=20, align="left", pad_x=14)
        if l_val != "":
            draw_cell_text(draw, [str(l_val)], c5_x[1], ry0, c5_x[2], ry1, base_size=21, align="right", pad_x=14)
        draw_cell_text(draw, r_lbl, c5_x[2], ry0, c5_x[3], ry1, base_size=20, align="left", pad_x=14)
        if r_val != "":
            draw_cell_text(draw, [str(r_val)], c5_x[3], ry0, c5_x[4], ry1, base_size=21, align="right", pad_x=14)

    # ================= РАЗДЕЛ 6 =================
    y6 = 1928
    draw.text(
        (ML, y6),
        "6. Сумма дохода, с которого не удержан налог налоговым агентом, и сумма неудержанного налога",
        fill=black,
        font=f_sec,
    )

    tbl6_top = y6 + 42
    row6_h = 44
    tbl6_bottom = tbl6_top + row6_h * 2
    c6_mid = ML + 940

    draw.rectangle([ML, tbl6_top, MR, tbl6_bottom], outline=black, width=2)
    draw.line([(ML, tbl6_top + row6_h), (MR, tbl6_top + row6_h)], fill=black, width=2)
    draw.line([(c6_mid, tbl6_top), (c6_mid, tbl6_bottom)], fill=black, width=2)

    draw_cell_text(
        draw,
        ["Сумма дохода, с которого не удержан налог налоговым агентом"],
        ML,
        tbl6_top,
        c6_mid,
        tbl6_top + row6_h,
        base_size=20,
        align="left",
        pad_x=14,
    )
    draw_cell_text(
        draw,
        ["Сумма неудержанного налога"],
        ML,
        tbl6_top + row6_h,
        c6_mid,
        tbl6_bottom,
        base_size=20,
        align="left",
        pad_x=14,
    )
    if totals.get("unwithheld_income"):
        draw_cell_text(
            draw,
            [str(totals["unwithheld_income"])],
            c6_mid,
            tbl6_top,
            MR,
            tbl6_top + row6_h,
            base_size=21,
            align="right",
            pad_x=14,
        )
    if totals.get("unwithheld_tax"):
        draw_cell_text(
            draw,
            [str(totals["unwithheld_tax"])],
            c6_mid,
            tbl6_top + row6_h,
            MR,
            tbl6_bottom,
            base_size=21,
            align="right",
            pad_x=14,
        )

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ---------------- ШТАМП ЭЛЕКТРОННОЙ ПОДПИСИ ----------------

def resolve_signed_date(ep: dict, blank_meta: dict | None = None) -> str:
    sd = (ep.get("signed_date") or "auto").strip()
    if sd.lower() == "auto":
        if blank_meta and blank_meta.get("issue_date"):
            return blank_meta["issue_date"]
        return now_msk().strftime("%d.%m.%Y")
    return sd


def build_ep_stamp_image(
    ep: dict,
    target_width: int,
    blank_meta: dict | None = None,
) -> Image.Image:
    scale = 2
    ref_w = 1140
    ref_h = 204
    w, h = ref_w * scale, ref_h * scale

    stamp = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    d = ImageDraw.Draw(stamp)

    border_color = (24, 24, 24, 255)
    d.rectangle([0, 0, w - 1, h - 1], outline=border_color, width=2 * scale)

    pad_x = 26 * scale
    inner_w = w - 2 * pad_x
    text_color = (18, 18, 18, 255)

    owner = (ep.get("owner") or "").strip()
    cert = (ep.get("cert") or generate_cert_hex()).strip().upper()
    ca = (ep.get("ca") or "Федеральная налоговая служба").strip()
    valid_from = (ep.get("valid_from") or "").strip()
    valid_to = (ep.get("valid_to") or "").strip()
    signed_date = resolve_signed_date(ep, blank_meta)

    lines_spec = (
        ("ДОКУМЕНТ ПОДПИСАН ЭЛЕКТРОННОЙ ПОДПИСЬЮ", 26, True, 18),
        (f"Владелец:  {owner}", 23, False, 58),
        (f"Сертификат: {cert}", 23, False, 92),
        (f"УЦ:  {ca}", 23, False, 126),
        (
            f"Действителен:  с {valid_from} по {valid_to},  подписан: {signed_date}",
            23,
            False,
            160,
        ),
    )

    for line_text, base_sz, is_bold, y_base in lines_spec:
        sz = base_sz * scale
        f = load_font(sz, bold=is_bold)
        bb = d.textbbox((0, 0), line_text, font=f)
        while (bb[2] - bb[0]) > inner_w and sz > 13 * scale:
            sz -= scale
            f = load_font(sz, bold=is_bold)
            bb = d.textbbox((0, 0), line_text, font=f)
        d.text((pad_x, y_base * scale), line_text, fill=text_color, font=f)

    out_w = max(320, int(target_width))
    out_h = max(60, int(round(out_w * (ref_h / ref_w))))
    return stamp.resize((out_w, out_h), Image.Resampling.LANCZOS)


def apply_ep_stamp_to_image(
    img_bytes: bytes,
    ep: dict,
    blank_meta: dict | None = None,
) -> bytes:
    raw = Image.open(io.BytesIO(img_bytes))
    if raw.mode in ("RGBA", "LA") or (raw.mode == "P" and "transparency" in raw.info):
        base = Image.new("RGB", raw.size, (255, 255, 255))
        rgba = raw.convert("RGBA")
        base.paste(rgba, (0, 0), rgba)
    else:
        base = raw.convert("RGB")

    W, H = base.size

    if W == 1654 and H == 2339:
        stamp_w = 1120
        margin_x = 90
        margin_bottom = 44
    else:
        ratio = 0.68 if W >= 1100 else 0.80
        stamp_w = int(W * ratio)
        margin_x = max(20, int(W * 0.055))
        margin_bottom = max(20, int(H * 0.024))

    stamp = build_ep_stamp_image(ep, target_width=stamp_w, blank_meta=blank_meta)

    pos = ep.get("position", "center")
    if pos == "left":
        px = margin_x
    elif pos == "right":
        px = max(margin_x, W - margin_x - stamp.width)
    else:
        px = (W - stamp.width) // 2

    py = max(10, H - margin_bottom - stamp.height)
    base.paste(stamp, (px, py), stamp)

    out = io.BytesIO()
    base.save(out, format="PNG")
    return out.getvalue()


def normalize_uploaded_image(raw_bytes: bytes) -> bytes:
    raw = Image.open(io.BytesIO(raw_bytes))
    if raw.mode in ("RGBA", "LA") or (raw.mode == "P" and "transparency" in raw.info):
        base = Image.new("RGB", raw.size, (255, 255, 255))
        rgba = raw.convert("RGBA")
        base.paste(rgba, (0, 0), rgba)
    else:
        base = raw.convert("RGB")
    out = io.BytesIO()
    base.save(out, format="PNG")
    return out.getvalue()


# ---------------- ПРАВКА СПИСКА ОПЕРАЦИЙ ----------------

def transfer_sort_key(row: dict) -> tuple[int, int, int, int, int]:
    try:
        d, m, y = row.get("date", "01.01.2000").split(".")
        t_parts = (row.get("time") or "00:00").split(":")
        hh = int(t_parts[0])
        mm = int(t_parts[1]) if len(t_parts) > 1 else 0
        return (int(y), int(m), int(d), hh, mm)
    except Exception:
        return (0, 0, 0, 0, 0)


def parse_date_obj(s: str) -> datetime | None:
    try:
        return datetime.strptime(s.strip(), "%d.%m.%Y")
    except Exception:
        return None


def build_edit_menu(user_id: int) -> tuple[str, InlineKeyboardMarkup]:
    transfers = LAST_TRANSFERS.get(user_id, [])
    if not transfers:
        return (
            "Список операций пуст. Отправь PDF-выписку или нажми «Сбросить правки».",
            InlineKeyboardMarkup(
                inline_keyboard=[
                    [InlineKeyboardButton(text="Сбросить правки", callback_data="edit:reset")]
                ]
            ),
        )

    lines = [f"<b>Редактирование списка ({len(transfers)} шт.):</b>\n"]
    max_show = 45
    for idx, t in enumerate(transfers[:max_show], start=1):
        role = "Вывод" if t["amount"] > 0 else "Деп"
        amt = f"{abs(t['amount']):,.0f}".replace(",", " ")
        lines.append(
            f"{idx}. <code>{t['date']} {t['time']}</code> | {role} ({t['method']}) | <b>{amt} ₽</b>"
        )
    if len(transfers) > max_show:
        lines.append(f"<i>... и ещё {len(transfers) - max_show} шт.</i>")

    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="Удалить строки", callback_data="edit:del"),
                InlineKeyboardButton(text="+ Добавить", callback_data="edit:add"),
            ],
            [
                InlineKeyboardButton(text="Фильтр (дата / сумма)", callback_data="edit:filter"),
                InlineKeyboardButton(text="Сбросить правки", callback_data="edit:reset"),
            ],
            [
                InlineKeyboardButton(text="Готово — обновить Excel и кнопки", callback_data="edit:done"),
            ],
        ]
    )
    return "\n".join(lines), kb


def apply_edit_delete(transfers: list[dict], text: str) -> tuple[list[dict], int]:
    to_remove: set[int] = set()
    for token in re.split(r"[\s,;]+", text.strip()):
        if not token:
            continue
        rm = re.fullmatch(r"(\d+)\s*[-–]\s*(\d+)", token)
        if rm:
            a, b = int(rm.group(1)), int(rm.group(2))
            for i in range(min(a, b), max(a, b) + 1):
                if 1 <= i <= len(transfers):
                    to_remove.add(i)
        elif token.isdigit():
            i = int(token)
            if 1 <= i <= len(transfers):
                to_remove.add(i)
    new_list = [t for idx, t in enumerate(transfers, start=1) if idx not in to_remove]
    return new_list, len(to_remove)


def apply_edit_add(text: str) -> dict | None:
    """
    Форматы:
      деп 7000 25.11.2026 18:10 сбп
      выв 5000 26.11.2026 14:30 карта
      -7000 25.11.2026 18:10 сбп
      +5000 26.11.2026 14:30 карта
    """
    raw = text.strip()
    low = raw.lower()

    date_m = re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", raw)
    time_m = re.search(r"\b\d{1,2}:\d{2}\b", raw)

    rem = raw
    date_str = date_m.group(0) if date_m else now_msk().strftime("%d.%m.%Y")
    if date_m:
        rem = rem[: date_m.start()] + " " + rem[date_m.end() :]

    if time_m:
        hh, mm = time_m.group(0).split(":")
        time_str = f"{int(hh):02d}:{int(mm):02d}"
        rem = re.sub(r"\b\d{1,2}:\d{2}\b", " ", rem, count=1)
    else:
        time_str = now_msk().strftime("%H:%M")

    method = "По карте" if ("карт" in low or "card" in low) else "СБП"

    amt_m = re.search(r"([+-]?\s*\d[\d\s]*[.,]?\d*)", rem)
    if not amt_m:
        return None
    try:
        raw_num = amt_m.group(1).replace(" ", "").replace(",", ".")
        val = float(raw_num)
    except ValueError:
        return None
    if val == 0:
        return None

    if any(k in low for k in ("деп", "пополн", "исход")):
        amount = -abs(val)
    elif any(k in low for k in ("выв", "вход")):
        amount = abs(val)
    else:
        amount = val

    direction = "Входящий" if amount > 0 else "Исходящий"
    return {
        "date": date_str,
        "time": time_str,
        "direction": direction,
        "method": method,
        "amount": amount,
        "counterparty_key": None,
        "counterparty": "",
        "description": "Добавлено вручную",
        "card": "",
    }


def apply_edit_filter(transfers: list[dict], text: str) -> tuple[list[dict], str] | None:
    low = text.lower()
    dates = re.findall(r"\b\d{2}\.\d{2}\.\d{4}\b", text)
    min_m = re.search(r"от\s+(\d[\d\s.,]*)", low)
    max_m = re.search(r"до\s+(\d[\d\s.,]*)", low)

    if not dates and not min_m and not max_m:
        return None

    d_from = parse_date_obj(dates[0]) if len(dates) >= 1 else None
    d_to = parse_date_obj(dates[1]) if len(dates) >= 2 else d_from
    if d_from and d_to and d_from > d_to:
        d_from, d_to = d_to, d_from

    min_amt = parse_float_val(min_m.group(1)) if min_m else None
    max_amt = parse_float_val(max_m.group(1)) if max_m else None

    result: list[dict] = []
    for t in transfers:
        a = abs(t["amount"])
        if min_amt is not None and a < min_amt:
            continue
        if max_amt is not None and a > max_amt:
            continue
        if d_from and d_to:
            td = parse_date_obj(t.get("date", ""))
            if not td or not (d_from <= td <= d_to):
                continue
        result.append(t)

    parts = []
    if d_from and d_to:
        parts.append(
            f"даты {d_from.strftime('%d.%m.%Y')}–{d_to.strftime('%d.%m.%Y')}"
        )
    if min_amt is not None:
        parts.append(f"от {min_amt:,.0f} ₽".replace(",", " "))
    if max_amt is not None:
        parts.append(f"до {max_amt:,.0f} ₽".replace(",", " "))
    return result, ", ".join(parts)


# ---------------- КЛАВИАТУРЫ И ТЕКСТЫ ----------------

def build_start_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="Бланк (КНД 1175018)", callback_data="scr:blank"),
                InlineKeyboardButton(text="Скриншот профиля", callback_data="scr:profile"),
            ],
            [
                InlineKeyboardButton(text="Настройка ЭП", callback_data="ep:menu"),
                InlineKeyboardButton(text="Шаблоны компаний", callback_data="agt:menu"),
            ],
        ]
    )


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
                ),
                InlineKeyboardButton(
                    text="Править список", callback_data="edit:menu"
                ),
            ]
        )
    buttons.append(
        [
            InlineKeyboardButton(text="Скриншот профиля", callback_data="scr:profile"),
            InlineKeyboardButton(text="Бланк", callback_data="scr:blank"),
        ]
    )
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def build_blank_ready_keyboard(has_agent: bool = False) -> InlineKeyboardMarkup:
    rows = [
        [
            InlineKeyboardButton(text="Подписать (ЭП)", callback_data="ep:menu"),
            InlineKeyboardButton(text="Скачать PNG", callback_data="blk:get_png"),
        ]
    ]
    row2 = []
    if has_agent:
        row2.append(InlineKeyboardButton(text="Сохранить компанию", callback_data="agt:save_last"))
    row2.append(InlineKeyboardButton(text="Новый бланк", callback_data="scr:blank"))
    rows.append(row2)
    return InlineKeyboardMarkup(inline_keyboard=rows)


def build_agents_menu() -> tuple[str, InlineKeyboardMarkup]:
    agents = load_agents()
    rows: list[list[InlineKeyboardButton]] = []
    if not agents:
        text = (
            "<b>Шаблоны налоговых агентов (компаний)</b>\n\n"
            "Пока пусто. Ты можешь добавить компанию сейчас или нажать "
            "<b>«Сохранить компанию»</b> после заполнения любого бланка."
        )
    else:
        lines = ["<b>Сохранённые компании (нажми, чтобы применить к ЭП или удалить):</b>\n"]
        for idx, a in enumerate(agents, start=1):
            inn = a.get("agent_inn") or "—"
            kpp = a.get("agent_kpp") or "—"
            lines.append(
                f"{idx}. <b>{html.escape(a.get('agent_name', ''))}</b> (ИНН {html.escape(inn)} / КПП {html.escape(kpp)})"
            )
            rows.append(
                [
                    InlineKeyboardButton(
                        text=f"ЭП: {a.get('agent_name', '')[:22]}",
                        callback_data=f"agt:use_ep:{a['id']}",
                    ),
                    InlineKeyboardButton(
                        text="Удалить",
                        callback_data=f"agt:del:{a['id']}",
                    ),
                ]
            )
        text = "\n".join(lines)

    rows.append([InlineKeyboardButton(text="+ Добавить компанию", callback_data="agt:add")])
    return text, InlineKeyboardMarkup(inline_keyboard=rows)


def build_ep_menu(user_id: int) -> tuple[str, InlineKeyboardMarkup]:
    ep = load_ep_settings()
    meta = LAST_BLANK_META.get(user_id)
    pages = LAST_BLANK_PAGES.get(user_id)
    has_blank = bool(pages)
    pos_cur = ep.get("position", "center")
    pos_labels = {"left": "Снизу слева", "center": "Снизу по центру", "right": "Снизу справа"}
    signed_preview = resolve_signed_date(ep, meta)
    signed_note = " (авто)" if (ep.get("signed_date") or "auto").lower() == "auto" else ""

    text = (
        "<b>Электронная подпись (ЭП)</b>\n\n"
        f"• Владелец: <code>{html.escape(ep.get('owner', ''))}</code>\n"
        f"• Сертификат: <code>{html.escape(ep.get('cert', ''))}</code>\n"
        f"• УЦ: <code>{html.escape(ep.get('ca', ''))}</code>\n"
        f"• Действителен: <code>с {html.escape(ep.get('valid_from', ''))} "
        f"по {html.escape(ep.get('valid_to', ''))}, подписан: {html.escape(signed_preview)}</code>{signed_note}\n"
        f"• Позиция: <code>{pos_labels.get(pos_cur, 'Снизу по центру')}</code>\n"
    )

    rows: list[list[InlineKeyboardButton]] = []
    if has_blank:
        p_cnt = len(pages or [])
        text += f"\nБланк загружен ({p_cnt} стр.) — можно подписать сразу или изменить реквизиты ЭП."
        rows.append([InlineKeyboardButton(text="Подписать текущими данными", callback_data="ep:apply")])
        if meta and meta.get("agent_name"):
            short_agent = meta["agent_name"][:28]
            rows.append(
                [
                    InlineKeyboardButton(
                        text=f"Подписать от: {short_agent}",
                        callback_data="ep:apply_from_blank",
                    )
                ]
            )
    else:
        text += "\nОтправь боту бланк (PDF, фото или PNG/JPG) либо заполни новый через /blank."

    rows.append(
        [
            InlineKeyboardButton(
                text=("• Слева" if pos_cur == "left" else "Слева"),
                callback_data="ep:pos_left",
            ),
            InlineKeyboardButton(
                text=("• По центру" if pos_cur == "center" else "По центру"),
                callback_data="ep:pos_center",
            ),
            InlineKeyboardButton(
                text=("• Справа" if pos_cur == "right" else "Справа"),
                callback_data="ep:pos_right",
            ),
        ]
    )
    rows.append([InlineKeyboardButton(text="Настроить реквизиты ЭП", callback_data="ep:edit")])
    return text, InlineKeyboardMarkup(inline_keyboard=rows)


def profile_prompt_text(profile: dict) -> str:
    cur_name = html.escape(profile.get("name") or "не указано")
    cur_dob = html.escape(profile.get("dob") or "не указана")
    return (
        "Введи <b>телефон</b> и <b>почту</b> для скриншота профиля "
        "(одной строкой или с новой строки).\n\n"
        f"Сейчас в профиле:\n"
        f"• Имя: <code>{cur_name}</code>\n"
        f"• Дата рождения: <code>{cur_dob}</code>\n\n"
        "Если нужно сразу поменять имя или дату рождения, просто допиши их туда же:\n"
        "<code>+79991234567 mail@gmail.com</code>\n"
        "или\n"
        "<code>+79991234567 mail@gmail.com 15.05.1995 Татьяна Берлева</code>\n\n"
        "Для отмены нажми /cancel"
    )


# ---------------- ЛОГИКА ОПРОСНИКА ЭП ----------------

def get_ep_step_prompt(step_idx: int, user_id: int) -> tuple[str, InlineKeyboardMarkup]:
    step = EP_STEPS[step_idx]
    total = len(EP_STEPS)
    ep = load_ep_settings()
    meta = LAST_BLANK_META.get(user_id)
    header = f"<b>Настройка ЭП — шаг {step_idx + 1}/{total}</b>\n\n"
    rows: list[list[InlineKeyboardButton]] = []

    if step == "owner":
        cur = ep.get("owner", "")
        text = (
            header
            + "<b>Владелец подписи</b> (название организации или ФИО).\n"
            + f"Текущее значение:\n<code>{html.escape(cur)}</code>"
        )
        if meta and meta.get("agent_name"):
            ag = meta["agent_name"]
            rows.append([InlineKeyboardButton(text=f"Из бланка: {ag[:26]}", callback_data="epw:from_blank")])
        rows.append([InlineKeyboardButton(text="Оставить текущее", callback_data="epw:keep")])

    elif step == "cert":
        cur = ep.get("cert", "")
        text = (
            header
            + "<b>Номер сертификата</b> (32 символа, например <code>03861D760091B449A9496EFF4FE4488A</code>).\n"
            + f"Текущий:\n<code>{html.escape(cur)}</code>"
        )
        rows.append([InlineKeyboardButton(text="Сгенерировать новый", callback_data="epw:gen")])
        rows.append([InlineKeyboardButton(text="Оставить текущий", callback_data="epw:keep")])

    elif step == "ca":
        cur = ep.get("ca", "")
        text = (
            header
            + "<b>УЦ (удостоверяющий центр).</b>\n"
            + f"Текущий: <code>{html.escape(cur)}</code>"
        )
        rows.append([InlineKeyboardButton(text="Федеральная налоговая служба", callback_data="epw:default")])
        rows.append([InlineKeyboardButton(text="Оставить текущий", callback_data="epw:keep")])

    elif step == "dates":
        cur_from = ep.get("valid_from", "")
        cur_to = ep.get("valid_to", "")
        cur_signed = resolve_signed_date(ep, meta)
        text = (
            header
            + "<b>Даты действия и подписания.</b>\n"
            + f"Сейчас: <code>с {html.escape(cur_from)} по {html.escape(cur_to)}, подписан: {html.escape(cur_signed)}</code>\n\n"
            + "Введи через пробел:\n"
            + "• <b>3 даты</b> — с, по, подписан:\n"
            + "  <code>24.01.2026 24.10.2027 24.11.2026</code>\n"
            + "• или <b>1 дату</b> — поменяется только дата подписания:\n"
            + "  <code>24.11.2026</code>"
        )
        rows.append([InlineKeyboardButton(text="Авто (срок 15 мес / сегодня)", callback_data="epw:default")])
        rows.append([InlineKeyboardButton(text="Оставить текущие", callback_data="epw:keep")])

    else:  # position
        text = header + "<b>Где разместить штамп ЭП внизу бланка?</b>"
        rows.append(
            [
                InlineKeyboardButton(text="Слева", callback_data="epw:pos_left"),
                InlineKeyboardButton(text="По центру", callback_data="epw:pos_center"),
                InlineKeyboardButton(text="Справа", callback_data="epw:pos_right"),
            ]
        )

    rows.append([InlineKeyboardButton(text="Отмена", callback_data="epw:cancel")])
    return text, InlineKeyboardMarkup(inline_keyboard=rows)


async def send_signed_blank(message: types.Message, user_id: int, ep_override: dict | None = None) -> None:
    pages = LAST_BLANK_PAGES.get(user_id)
    if not pages:
        await message.answer(
            "Настройки ЭП сохранены. Теперь отправь бланк (PDF, фото или PNG/JPG) "
            "или заполни новый через /blank — и нажми «Подписать»."
        )
        return

    ep = ep_override if ep_override is not None else load_ep_settings()
    meta = LAST_BLANK_META.get(user_id)
    wait_msg = await message.answer("Накладываю электронную подпись и собираю PDF...")
    try:
        signed_last_png = apply_ep_stamp_to_image(pages[-1], ep, blank_meta=meta)
        LAST_SIGNED_PNG[user_id] = signed_last_png
        signed_pages = pages[:-1] + [signed_last_png]
        signed_pdf = images_to_pdf_bytes(signed_pages, resolution=200.0)

        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(text="Скачать PNG", callback_data="ep:get_png"),
                    InlineKeyboardButton(text="Изменить ЭП", callback_data="ep:menu"),
                ]
            ]
        )
        await message.answer_photo(
            BufferedInputFile(signed_last_png, filename="signed_blank.png"),
            caption="Подписанный бланк готов (ниже готовый PDF-документ).",
            reply_markup=kb,
        )
        await message.answer_document(
            BufferedInputFile(signed_pdf, filename="signed_blank.pdf"),
            caption="Подписанный PDF. Можно прислать другой бланк (PDF или картинку) либо нажать /blank.",
        )
        await wait_msg.delete()
    except Exception as e:
        logging.exception("Ошибка наложения ЭП")
        await wait_msg.edit_text(f"Не удалось наложить подпись: {e}")


async def send_ep_step_or_finish(message: types.Message, user_id: int) -> None:
    session = EP_SESSIONS.get(user_id)
    if not session:
        return

    step_idx = session["step_idx"]
    if step_idx >= len(EP_STEPS):
        ep = load_ep_settings()
        ep.update(session["data"])
        save_ep_settings(ep)
        EP_SESSIONS.pop(user_id, None)
        await send_signed_blank(message, user_id, ep_override=ep)
        return

    text, kb = get_ep_step_prompt(step_idx, user_id)
    await message.answer(text, reply_markup=kb, parse_mode="HTML")


def apply_ep_step(session: dict, user_id: int, raw_text: str | None, action: str = "text") -> str | None:
    step = EP_STEPS[session["step_idx"]]
    ep = load_ep_settings()
    meta = LAST_BLANK_META.get(user_id)
    data = session["data"]
    text = (raw_text or "").strip()

    if step == "owner":
        if action == "keep":
            data["owner"] = ep.get("owner", "")
        elif action == "from_blank" and meta and meta.get("agent_name"):
            data["owner"] = meta["agent_name"]
        else:
            if not text:
                return "Введи владельца ЭП или нажми «Оставить текущее»."
            data["owner"] = text

    elif step == "cert":
        if action == "keep":
            data["cert"] = ep.get("cert", "") or generate_cert_hex()
        elif action == "gen":
            data["cert"] = generate_cert_hex()
        else:
            cleaned = re.sub(r"\s+", "", text).upper()
            if not cleaned:
                return "Введи номер сертификата или нажми «Сгенерировать новый»."
            data["cert"] = cleaned

    elif step == "ca":
        if action == "keep":
            data["ca"] = ep.get("ca", "Федеральная налоговая служба")
        elif action == "default":
            data["ca"] = "Федеральная налоговая служба"
        else:
            if not text:
                return "Введи название УЦ или выбери «Федеральная налоговая служба»."
            data["ca"] = text

    elif step == "dates":
        if action == "keep":
            data["valid_from"] = ep.get("valid_from", "24.01.2026")
            data["valid_to"] = ep.get("valid_to", "24.10.2027")
            data["signed_date"] = ep.get("signed_date", "auto")
        elif action == "default":
            today = now_msk()
            data["valid_from"] = (today - timedelta(days=90)).strftime("%d.%m.%Y")
            data["valid_to"] = (today + timedelta(days=365)).strftime("%d.%m.%Y")
            data["signed_date"] = "auto"
        else:
            found = re.findall(r"\b\d{2}\.\d{2}\.\d{4}\b", text)
            if len(found) >= 3:
                data["valid_from"] = found[0]
                data["valid_to"] = found[1]
                data["signed_date"] = found[2]
            elif len(found) == 2:
                data["valid_from"] = found[0]
                data["valid_to"] = found[1]
                data["signed_date"] = ep.get("signed_date", "auto")
            elif len(found) == 1:
                data["valid_from"] = ep.get("valid_from", "24.01.2026")
                data["valid_to"] = ep.get("valid_to", "24.10.2027")
                data["signed_date"] = found[0]
            else:
                return (
                    "Не найдено дат в формате ДД.ММ.ГГГГ.\n"
                    "Пример: <code>24.01.2026 24.10.2027 24.11.2026</code>"
                )

    elif step == "position":
        if action == "pos_left" or text.lower() in ("слева", "лево", "left"):
            data["position"] = "left"
        elif action == "pos_right" or text.lower() in ("справа", "право", "right"):
            data["position"] = "right"
        else:
            data["position"] = "center"

    session["step_idx"] += 1
    return None


# ---------------- ЛОГИКА ОПРОСНИКА БЛАНКА КНД 1175018 ----------------

def parse_float_val(token: str) -> float:
    cleaned = token.replace(" ", "").replace("\u00a0", "").replace(",", ".")
    return float(cleaned)


def parse_agent_details(text: str) -> dict[str, str]:
    res = {"oktmo": "", "agent_phone": "", "agent_inn": "", "agent_kpp": ""}
    raw_tokens = [t.strip() for t in re.split(r"[\n,;]+", text.strip()) if t.strip()]
    if len(raw_tokens) == 1:
        raw_tokens = raw_tokens[0].split()

    if len(raw_tokens) == 4:
        res["oktmo"] = "" if raw_tokens[0] == "-" else raw_tokens[0]
        res["agent_phone"] = "" if raw_tokens[1] == "-" else raw_tokens[1]
        res["agent_inn"] = "" if raw_tokens[2] == "-" else raw_tokens[2]
        res["agent_kpp"] = "" if raw_tokens[3] == "-" else raw_tokens[3]
        return res

    for tok in raw_tokens:
        if tok == "-":
            continue
        digits = re.sub(r"\D", "", tok)
        if ("+" in tok or "(" in tok or "-" in tok) and len(digits) >= 10:
            res["agent_phone"] = tok
        elif len(digits) == 8 and not res["oktmo"]:
            res["oktmo"] = digits
        elif len(digits) == 9 and not res["agent_kpp"]:
            res["agent_kpp"] = digits
        elif len(digits) in (10, 12) and not res["agent_inn"]:
            res["agent_inn"] = digits
        elif len(digits) == 11 and digits.startswith(("7", "8")) and not res["agent_phone"]:
            res["agent_phone"] = tok
        elif len(digits) == 11 and not res["oktmo"]:
            res["oktmo"] = digits
        elif not res["oktmo"]:
            res["oktmo"] = tok
        elif not res["agent_phone"]:
            res["agent_phone"] = tok
        elif not res["agent_inn"]:
            res["agent_inn"] = tok
        elif not res["agent_kpp"]:
            res["agent_kpp"] = tok

    return res


def parse_incomes_input(text: str) -> list[dict]:
    records: list[dict] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        parts = line.replace(";", " ").split()
        if len(parts) < 3:
            continue
        m_part = parts[0]
        code = parts[1]
        try:
            amount = parse_float_val(parts[2])
        except ValueError:
            continue

        ded_code = parts[3] if len(parts) >= 4 else ""
        ded_amount: float | str = ""
        if len(parts) >= 5:
            try:
                ded_amount = parse_float_val(parts[4])
            except ValueError:
                ded_amount = parts[4]

        range_m = re.fullmatch(r"(\d{1,2})\s*[-–]\s*(\d{1,2})", m_part)
        if range_m:
            m1, m2 = int(range_m.group(1)), int(range_m.group(2))
            if 1 <= m1 <= m2 <= 12:
                for m in range(m1, m2 + 1):
                    records.append(
                        {
                            "month": f"{m:02d}",
                            "code": code,
                            "amount": amount,
                            "ded_code": ded_code,
                            "ded_amount": ded_amount,
                        }
                    )
        elif m_part.isdigit() and 1 <= int(m_part) <= 12:
            records.append(
                {
                    "month": f"{int(m_part):02d}",
                    "code": code,
                    "amount": amount,
                    "ded_code": ded_code,
                    "ded_amount": ded_amount,
                }
            )
    return records[:24]


def parse_deductions_input(text: str) -> list[dict]:
    items: list[dict] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        parts = line.replace(";", " ").split()
        if len(parts) >= 2:
            try:
                val = parse_float_val(parts[1])
            except ValueError:
                val = parts[1]
            items.append({"code": parts[0], "amount": val})
    return items[:4]


def compute_auto_totals(data: dict) -> dict[str, str]:
    incomes = data.get("incomes", [])
    deductions = data.get("deductions", [])
    try:
        rate = float(str(data.get("tax_rate", "13")).replace(",", "."))
    except ValueError:
        rate = 13.0

    total_income = sum(
        float(r["amount"]) for r in incomes if isinstance(r.get("amount"), (int, float))
    )
    row_deds = sum(
        float(r["ded_amount"])
        for r in incomes
        if isinstance(r.get("ded_amount"), (int, float))
    )
    sec4_deds = sum(
        float(d["amount"])
        for d in deductions
        if isinstance(d.get("amount"), (int, float))
    )
    tax_base = max(0.0, total_income - row_deds - sec4_deds)
    tax_calc = round(tax_base * rate / 100.0)

    return {
        "total_income": fmt_money_blank(total_income),
        "tax_base": fmt_money_blank(tax_base),
        "tax_calc": fmt_int_blank(tax_calc),
        "tax_withheld": fmt_int_blank(tax_calc),
        "fixed_adv": "",
        "profit_tax": "",
        "foreign_tax": "",
        "tax_over": "",
        "unwithheld_income": "",
        "unwithheld_tax": "",
    }


def get_blank_step_prompt(step_idx: int, data: dict) -> tuple[str, InlineKeyboardMarkup]:
    step = BLANK_STEPS[step_idx]
    total_steps = len(BLANK_STEPS)
    header = f"<b>Бланк КНД 1175018 — шаг {step_idx + 1}/{total_steps}</b>\n\n"
    profile = load_profile()
    today_str = now_msk().strftime("%d.%m.%Y")
    cur_year = str(now_msk().year)

    rows: list[list[InlineKeyboardButton]] = []

    if step == "year_date":
        text = (
            header
            + "Укажи <b>год</b> и <b>дату справки</b>.\n"
            + f"Например: <code>{cur_year} {today_str}</code> или просто <code>{cur_year}</code>."
        )
        rows.append([InlineKeyboardButton(text=f"По умолчанию ({cur_year}, {today_str})", callback_data="blk:default")])

    elif step == "agent_name":
        agents = load_agents()
        text = (
            header
            + "<b>Налоговый агент</b> (название организации / ИП).\n"
            + 'Например: <code>ООО «Ромашка»</code>'
        )
        if agents:
            text += "\n\nИли выбери сохранённый шаблон (реквизиты ОКТМО/ИНН/КПП подставятся сами):"
            for a in agents[:6]:
                rows.append(
                    [
                        InlineKeyboardButton(
                            text=f"Шаблон: {a.get('agent_name', '')[:30]}",
                            callback_data=f"blk:tpl:{a['id']}",
                        )
                    ]
                )
        rows.append([InlineKeyboardButton(text="Пропустить", callback_data="blk:skip")])

    elif step == "agent_details":
        text = (
            header
            + "<b>Реквизиты налогового агента:</b> ОКТМО, Телефон, ИНН, КПП.\n"
            + "Введи через пробел или с новой строки (можно указать не все):\n"
            + "<code>20701000 +74732000000 3662123456 366201001</code>"
        )
        rows.append([InlineKeyboardButton(text="Пропустить", callback_data="blk:skip")])

    elif step == "fio":
        saved_name = profile.get("name", "").strip()
        text = (
            header
            + "<b>ФИО получателя дохода</b> (Фамилия Имя Отчество).\n"
            + "Например: <code>Иванов Иван Иванович</code>"
        )
        if saved_name:
            text += f"\n\nИз профиля/выписки: <code>{html.escape(saved_name)}</code>"
            rows.append([InlineKeyboardButton(text=f"Оставить: {saved_name[:28]}", callback_data="blk:default")])

    elif step == "person_inn":
        text = (
            header
            + "<b>ИНН физического лица</b> (12 цифр).\n"
            + "Например: <code>366212345678</code>"
        )
        rows.append([InlineKeyboardButton(text="Пропустить", callback_data="blk:skip")])

    elif step == "person_meta":
        saved_dob = profile.get("dob", "").strip()
        text = (
            header
            + "<b>Статус налогоплательщика, дата рождения, код страны.</b>\n"
            + "Например: <code>1 27.03.2006 643</code>\n"
            + "(Можно ввести просто дату <code>27.03.2006</code> — статус <code>1</code> и код страны <code>643</code> подставятся сами.)"
        )
        if saved_dob:
            rows.append([InlineKeyboardButton(text=f"1 / {saved_dob} / 643", callback_data="blk:default")])
        else:
            rows.append([InlineKeyboardButton(text="Только 1 и 643 (без даты)", callback_data="blk:default")])

    elif step == "passport":
        saved_pass = profile.get("passport", "").strip()
        text = (
            header
            + "<b>Код документа и серия/номер паспорта.</b>\n"
            + "Например: <code>21 2022 720045</code>\n"
            + "(Можно ввести просто <code>2022 720045</code> — код <code>21</code> подставится сам.)"
        )
        if saved_pass:
            text += f"\n\nИз последней выписки: <code>21 {html.escape(saved_pass)}</code>"
            rows.append([InlineKeyboardButton(text=f"Подставить: 21 {saved_pass}", callback_data="blk:default")])
        rows.append([InlineKeyboardButton(text="Пропустить", callback_data="blk:skip")])

    elif step == "tax_rate":
        text = (
            header
            + "<b>Ставка налога (%)</b> для раздела 3.\n"
            + "Введи число (например, <code>13</code> или <code>15</code>)."
        )
        rows.append([InlineKeyboardButton(text="13%", callback_data="blk:default")])

    elif step == "incomes":
        text = (
            header
            + "<b>Доходы по месяцам (раздел 3).</b>\n"
            + "Формат строки: <code>Месяц КодДохода Сумма [КодВычета СуммаВычета]</code>\n\n"
            + "Примеры:\n"
            + "• Одинаковый доход с января по ноябрь:\n"
            + "<code>01-11 2000 85000</code>\n"
            + "• Разные суммы по месяцам (в одном сообщении):\n"
            + "<code>01 2000 82000\n02 2000 85500\n03 2000 90000</code>"
        )
        rows.append([InlineKeyboardButton(text="Пропустить (пустая таблица)", callback_data="blk:skip")])

    elif step == "deductions":
        text = (
            header
            + "<b>Вычеты (раздел 4)</b> — если нужны.\n"
            + "Формат: <code>КодВычета Сумма</code> (до 4 строк), например:\n"
            + "<code>126 1400</code>"
        )
        rows.append([InlineKeyboardButton(text="Пропустить", callback_data="blk:skip")])

    else:  # totals
        auto_t = compute_auto_totals(data)
        text = (
            header
            + "<b>Раздел 5 (общие суммы дохода и налога).</b>\n"
            + "Авторасчёт по разделу 3:\n"
            + f"• Общая сумма дохода: <code>{auto_t['total_income']}</code>\n"
            + f"• Налоговая база: <code>{auto_t['tax_base']}</code>\n"
            + f"• Налог исчисленный / удержанный: <code>{auto_t['tax_calc']}</code>\n\n"
            + "Нажми <b>«Авторасчёт»</b> или введи вручную 4 числа через пробел:\n"
            + "<code>ОбщийДоход НалоговаяБаза Исчисленный Удержанный</code>"
        )
        rows.append([InlineKeyboardButton(text="Авторасчёт (готово)", callback_data="blk:default")])

    rows.append([InlineKeyboardButton(text="Отмена", callback_data="blk:cancel")])
    return text, InlineKeyboardMarkup(inline_keyboard=rows)


async def send_blank_step_or_finish(message: types.Message, user_id: int) -> None:
    session = BLANK_SESSIONS.get(user_id)
    if not session:
        return

    step_idx = session["step_idx"]
    data = session["data"]

    if step_idx >= len(BLANK_STEPS):
        BLANK_SESSIONS.pop(user_id, None)
        wait_msg = await message.answer("Формирую бланк КНД 1175018 (PNG + A4 PDF)...")
        try:
            png_bytes = build_knd_1175018_png(data)
            pdf_bytes = images_to_pdf_bytes([png_bytes], resolution=200.0)
            LAST_BLANK_PAGES[user_id] = [png_bytes]
            LAST_BLANK_META[user_id] = {
                "agent_name": (data.get("agent_name") or "").strip(),
                "oktmo": (data.get("oktmo") or "").strip(),
                "agent_phone": (data.get("agent_phone") or "").strip(),
                "agent_inn": (data.get("agent_inn") or "").strip(),
                "agent_kpp": (data.get("agent_kpp") or "").strip(),
                "issue_date": (data.get("issue_date") or "").strip(),
                "tpl_ep": data.get("_tpl_ep"),
            }
            has_agent = bool(LAST_BLANK_META[user_id]["agent_name"])
            await message.answer_photo(
                BufferedInputFile(png_bytes, filename="knd_1175018.png"),
                caption="Готовый бланк КНД 1175018 (ниже файл A4 PDF). Можешь сразу подписать его ЭП.",
                reply_markup=build_blank_ready_keyboard(has_agent=has_agent),
            )
            await message.answer_document(
                BufferedInputFile(pdf_bytes, filename="knd_1175018.pdf"),
                caption="Бланк в формате PDF. Для подписи нажми «Подписать (ЭП)» выше или /sign.",
            )
            await wait_msg.delete()
        except Exception as e:
            logging.exception("Ошибка генерации бланка")
            await wait_msg.edit_text(f"Ошибка при создании бланка: {e}")
        return

    text, kb = get_blank_step_prompt(step_idx, data)
    await message.answer(text, reply_markup=kb, parse_mode="HTML")


def apply_blank_step(session: dict, raw_text: str | None, action: str = "text") -> str | None:
    step = BLANK_STEPS[session["step_idx"]]
    data = session["data"]
    profile = load_profile()
    today_str = now_msk().strftime("%d.%m.%Y")
    cur_year = str(now_msk().year)
    text = (raw_text or "").strip()

    if action == "skip" or text == "-":
        if step == "year_date":
            data["year"] = cur_year
            data["issue_date"] = today_str
        elif step == "tax_rate":
            data["tax_rate"] = "13"
        elif step == "totals":
            data["totals"] = compute_auto_totals(data)
        session["step_idx"] += 1
        return None

    if step == "agent_name" and action.startswith("tpl:"):
        tpl_id = action.split(":", 1)[1]
        agents = load_agents()
        tpl = next((a for a in agents if a.get("id") == tpl_id), None)
        if not tpl:
            return "Шаблон не найден, введи название компании вручную."
        data["agent_name"] = tpl.get("agent_name", "")
        data["oktmo"] = tpl.get("oktmo", "")
        data["agent_phone"] = tpl.get("agent_phone", "")
        data["agent_inn"] = tpl.get("agent_inn", "")
        data["agent_kpp"] = tpl.get("agent_kpp", "")
        if tpl.get("ep_owner"):
            data["_tpl_ep"] = {
                "owner": tpl.get("ep_owner", ""),
                "cert": tpl.get("ep_cert", ""),
                "ca": tpl.get("ep_ca", "Федеральная налоговая служба"),
                "valid_from": tpl.get("ep_valid_from", ""),
                "valid_to": tpl.get("ep_valid_to", ""),
            }
        # Пропускаем и agent_name, и agent_details
        session["step_idx"] += 2
        return None

    if step == "year_date":
        if action == "default":
            data["year"] = cur_year
            data["issue_date"] = today_str
        else:
            date_m = re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", text)
            year_m = re.search(r"\b(19|20)\d{2}\b", text)
            if not date_m and not year_m:
                return "Укажи год (например, 2026) или год и дату (2026 27.11.2026)."
            data["issue_date"] = date_m.group(0) if date_m else today_str
            if year_m:
                data["year"] = year_m.group(0)
            else:
                data["year"] = data["issue_date"].split(".")[-1]

    elif step == "agent_name":
        data["agent_name"] = text

    elif step == "agent_details":
        data.update(parse_agent_details(text))

    elif step == "fio":
        full_name = profile.get("name", "").strip() if action == "default" else text
        parts = full_name.split()
        if not parts:
            return "Введи хотя бы фамилию и имя."
        data["last_name"] = parts[0]
        data["first_name"] = parts[1] if len(parts) >= 2 else ""
        data["middle_name"] = " ".join(parts[2:]) if len(parts) >= 3 else ""

    elif step == "person_inn":
        digits = re.sub(r"\D", "", text)
        if not digits:
            return "Введи ИНН цифрами или нажми «Пропустить»."
        data["person_inn"] = digits

    elif step == "person_meta":
        if action == "default":
            data["tax_status"] = "1"
            data["dob"] = profile.get("dob", "").strip()
            data["country_code"] = "643"
        else:
            dob_m = re.search(r"\b\d{2}\.\d{2}\.\d{4}\b", text)
            rem = text
            if dob_m:
                data["dob"] = dob_m.group(0)
                rem = rem[: dob_m.start()] + " " + rem[dob_m.end() :]
            tokens = rem.split()
            status = "1"
            country = "643"
            for tok in tokens:
                d = re.sub(r"\D", "", tok)
                if len(d) == 1:
                    status = d
                elif len(d) == 3:
                    country = d
            data["tax_status"] = status
            data["country_code"] = country

    elif step == "passport":
        if action == "default":
            saved_pass = profile.get("passport", "").strip()
            if saved_pass:
                data["doc_code"] = "21"
                data["doc_number"] = saved_pass
        else:
            m_code = re.match(r"^(\d{2})\s+(.+)$", text)
            if m_code:
                data["doc_code"] = m_code.group(1)
                data["doc_number"] = m_code.group(2).strip()
            else:
                data["doc_code"] = "21"
                data["doc_number"] = text

    elif step == "tax_rate":
        if action == "default":
            data["tax_rate"] = "13"
        else:
            cleaned = re.sub(r"[^\d.,]", "", text)
            if not cleaned:
                return "Введи ставку числом (например, 13)."
            data["tax_rate"] = cleaned

    elif step == "incomes":
        recs = parse_incomes_input(text)
        if not recs:
            return (
                "Не удалось распознать строки доходов.\n"
                "Пример: <code>01-11 2000 85000</code> или <code>01 2000 85000</code>"
            )
        data["incomes"] = recs

    elif step == "deductions":
        deds = parse_deductions_input(text)
        if not deds:
            return "Формат: <code>126 1400</code> (код и сумма) или нажми «Пропустить»."
        data["deductions"] = deds

    elif step == "totals":
        if action == "default" or text.lower() in ("авто", "auto"):
            data["totals"] = compute_auto_totals(data)
        else:
            parts = text.replace(";", " ").split()
            if len(parts) < 4:
                return "Введи 4 числа (общий доход, база, исчисленный, удержанный) или нажми «Авторасчёт»."
            try:
                ti = fmt_money_blank(parse_float_val(parts[0]))
                tb = fmt_money_blank(parse_float_val(parts[1]))
                tc = fmt_int_blank(parse_float_val(parts[2]))
                tw = fmt_int_blank(parse_float_val(parts[3]))
            except ValueError:
                return "Проверь числа — должно быть 4 значения через пробел."
            data["totals"] = {
                "total_income": ti,
                "tax_base": tb,
                "tax_calc": tc,
                "tax_withheld": tw,
            }

    session["step_idx"] += 1
    return None


# ---------------- КОМАНДЫ И ОБРАБОТЧИКИ ----------------

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    clear_active_inputs(message.from_user.id if message.from_user else None)
    ui = load_ui_settings()
    sbp_p = get_sbp_asset_path()
    f_reg = get_font_asset_path(bold=False)
    f_bold = get_font_asset_path(bold=True)

    assets_info = (
        "Файлы на сервере:\n"
        f"• СБП: {html.escape(sbp_p.name) if sbp_p else 'НЕ НАЙДЕН (используется заглушка)'}\n"
        f"• Шрифты: {html.escape(f_reg.name) if f_reg else 'системный'} / "
        f"{html.escape(f_bold.name) if f_bold else 'системный'}"
    )
    await message.answer(
        "• Кидай <b>PDF-выписку из Т-Банка</b> — верну Excel, правку списка и PNG-скриншоты.\n"
        "• Кидай <b>готовый бланк (PDF, фото или PNG/JPG)</b> — предложу подписать его настраиваемой ЭП.\n\n"
        f"Настройки скрина: домен <code>{html.escape(ui['domain'])}</code>, "
        f"часы <code>{html.escape(str(ui['status_time']))}</code>, заряд <code>{ui['battery']}%</code>\n"
        f"{assets_info}\n\n"
        "Команды:\n"
        "/blank — заполнить бланк справки о доходах (PDF + PNG)\n"
        "/sign — настроить ЭП и подписать бланк\n"
        "/agents — шаблоны налоговых агентов (компаний)\n"
        "/edit — править список операций из последней выписки\n"
        "/profile — сделать скриншот профиля (Настройки)\n"
        "/domain 1wujym.life — сменить домен в строке браузера\n"
        "/status 22:21 76 — задать часы и уровень батареи (или /status auto 65)\n"
        "/excluded — список игнорируемых договоров\n"
        "/exclude 1234567890 — добавить договор в исключения\n"
        "/unexclude 1234567890 — убрать из исключений",
        parse_mode="HTML",
        reply_markup=build_start_keyboard(),
    )


@dp.message(Command("blank"))
async def cmd_blank(message: types.Message):
    if not is_allowed_user(message.from_user) or not message.from_user:
        return
    user_id = message.from_user.id
    clear_active_inputs(user_id)
    BLANK_SESSIONS[user_id] = {"step_idx": 0, "data": {}}
    await send_blank_step_or_finish(message, user_id)


@dp.message(Command("sign"))
async def cmd_sign(message: types.Message):
    if not is_allowed_user(message.from_user) or not message.from_user:
        return
    user_id = message.from_user.id
    clear_active_inputs(user_id)
    text, kb = build_ep_menu(user_id)
    await message.answer(text, reply_markup=kb, parse_mode="HTML")


@dp.message(Command("agents"))
async def cmd_agents(message: types.Message):
    if not is_allowed_user(message.from_user) or not message.from_user:
        return
    clear_active_inputs(message.from_user.id)
    text, kb = build_agents_menu()
    await message.answer(text, reply_markup=kb, parse_mode="HTML")


@dp.message(Command("edit"))
async def cmd_edit(message: types.Message):
    if not is_allowed_user(message.from_user) or not message.from_user:
        return
    user_id = message.from_user.id
    clear_active_inputs(user_id)
    if user_id not in LAST_TRANSFERS:
        await message.answer("Сначала отправь PDF-выписку из Т-Банка.")
        return
    text, kb = build_edit_menu(user_id)
    await message.answer(text, reply_markup=kb, parse_mode="HTML")


@dp.message(Command("profile"))
async def cmd_profile(message: types.Message):
    if not is_allowed_user(message.from_user) or not message.from_user:
        return
    user_id = message.from_user.id
    clear_active_inputs(user_id)
    WAITING_PROFILE.add(user_id)
    profile = load_profile()
    await message.answer(profile_prompt_text(profile), parse_mode="HTML")


@dp.message(Command("cancel"))
async def cmd_cancel(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    clear_active_inputs(message.from_user.id if message.from_user else None)
    await message.answer("Отменено.")


@dp.message(Command("domain"))
async def cmd_domain(message: types.Message):
    if not is_allowed_user(message.from_user):
        return
    clear_active_inputs(message.from_user.id if message.from_user else None)
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
    clear_active_inputs(message.from_user.id if message.from_user else None)
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


@dp.message(F.photo)
async def handle_uploaded_photo_blank(message: types.Message, bot: Bot):
    if not is_allowed_user(message.from_user) or not message.from_user or not message.photo:
        return
    user_id = message.from_user.id
    clear_active_inputs(user_id)

    try:
        photo = message.photo[-1]
        file_info = await bot.get_file(photo.file_id)
        buf = io.BytesIO()
        await bot.download_file(file_info.file_path, buf)
        LAST_BLANK_PAGES[user_id] = [normalize_uploaded_image(buf.getvalue())]
        LAST_BLANK_META.pop(user_id, None)

        text, kb = build_ep_menu(user_id)
        await message.answer(
            "Бланк (фото) получен! Выбери действие:\n\n" + text,
            reply_markup=kb,
            parse_mode="HTML",
        )
    except Exception as e:
        logging.exception("Ошибка загрузки фото бланка")
        await message.answer(f"Не удалось прочитать изображение: {e}")


@dp.message(F.document)
async def handle_document(message: types.Message, bot: Bot):
    if not is_allowed_user(message.from_user) or not message.from_user:
        return
    user_id = message.from_user.id
    clear_active_inputs(user_id)

    doc = message.document
    if not doc:
        return
    fname = (doc.file_name or "").lower()

    if fname.endswith((".png", ".jpg", ".jpeg", ".webp")):
        try:
            file_info = await bot.get_file(doc.file_id)
            buf = io.BytesIO()
            await bot.download_file(file_info.file_path, buf)
            LAST_BLANK_PAGES[user_id] = [normalize_uploaded_image(buf.getvalue())]
            LAST_BLANK_META.pop(user_id, None)

            text, kb = build_ep_menu(user_id)
            await message.answer(
                "Бланк (изображение) получен! Выбери действие:\n\n" + text,
                reply_markup=kb,
                parse_mode="HTML",
            )
        except Exception as e:
            logging.exception("Ошибка загрузки файла бланка")
            await message.answer(f"Не удалось открыть изображение: {e}")
        return

    if not fname.endswith(".pdf"):
        await message.answer(
            "Пришли PDF-выписку из Т-Банка для разбора или бланк (PDF / фото / PNG / JPG) для подписи ЭП."
        )
        return

    wait_msg = await message.answer("Обрабатываю PDF...")

    try:
        file_info = await bot.get_file(doc.file_id)
        buf = io.BytesIO()
        await bot.download_file(file_info.file_path, buf)
        pdf_bytes = buf.getvalue()

        # Если это не выписка Т-Банка — считаем, что прислали PDF-бланк для подписи ЭП
        if not is_tbank_statement_pdf(pdf_bytes):
            await wait_msg.edit_text("Готовлю PDF-бланк к подписи...")
            pages_png = render_pdf_to_png_pages(pdf_bytes, resolution=200)
            if not pages_png:
                await wait_msg.edit_text("Не удалось прочитать страницы PDF.")
                return
            LAST_BLANK_PAGES[user_id] = pages_png
            LAST_BLANK_META.pop(user_id, None)

            text, kb = build_ep_menu(user_id)
            await wait_msg.delete()
            await message.answer(
                f"PDF-бланк получен ({len(pages_png)} стр.)! Выбери действие:\n\n" + text,
                reply_markup=kb,
                parse_mode="HTML",
            )
            return

        excluded = load_excluded()
        transfers, total_recognized, skipped_repeated, client_name, client_passport = (
            parse_tbank_pdf(pdf_bytes, excluded)
        )

        if client_name or client_passport:
            profile = load_profile()
            if client_name:
                profile["name"] = client_name
            if client_passport:
                profile["passport"] = client_passport
            save_profile(profile)

        if not transfers:
            await wait_msg.edit_text(
                f"Распознано переводов: {total_recognized}\n"
                f"Отсечено повторяющихся: {skipped_repeated}\n"
                "Разовых переводов для таблицы не осталось."
            )
            return

        LAST_TRANSFERS[user_id] = transfers
        LAST_TRANSFERS_ORIG[user_id] = copy.deepcopy(transfers)
        LAST_PDF_STEM[user_id] = Path(doc.file_name or "statement").stem

        excel_bytes = build_excel(transfers)
        out_name = f"transfers_{LAST_PDF_STEM[user_id]}.xlsx"
        caption = build_transfers_caption(transfers, total_recognized, skipped_repeated)

        await message.answer_document(
            BufferedInputFile(excel_bytes, filename=out_name),
            caption=caption,
            reply_markup=build_action_keyboard(transfers),
        )
        await wait_msg.delete()

    except Exception as e:
        logging.exception("Ошибка при разборе PDF")
        await wait_msg.edit_text(f"Не удалось обработать PDF: {e}")


# ---------------- CALLBACK-ОБРАБОТЧИКИ ----------------

@dp.callback_query(F.data == "scr:profile")
async def handle_profile_button(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    user_id = callback.from_user.id
    clear_active_inputs(user_id)
    WAITING_PROFILE.add(user_id)
    profile = load_profile()
    await callback.answer()
    await callback.message.answer(profile_prompt_text(profile), parse_mode="HTML")


@dp.callback_query(F.data == "scr:blank")
async def handle_blank_button(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    user_id = callback.from_user.id
    clear_active_inputs(user_id)
    BLANK_SESSIONS[user_id] = {"step_idx": 0, "data": {}}
    await callback.answer()
    await send_blank_step_or_finish(callback.message, user_id)


@dp.callback_query(F.data.startswith("agt:"))
async def handle_agents_callback(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    user_id = callback.from_user.id
    parts = callback.data.split(":")
    action = parts[1]

    if action == "menu":
        clear_active_inputs(user_id)
        await callback.answer()
        text, kb = build_agents_menu()
        await callback.message.answer(text, reply_markup=kb, parse_mode="HTML")
        return

    if action == "add":
        clear_active_inputs(user_id)
        WAITING_AGENT_ADD.add(user_id)
        await callback.answer()
        await callback.message.answer(
            "<b>Добавление шаблона компании</b>\n\n"
            "Пришли в одном сообщении с новой строки:\n"
            "1-я строка: <b>Название компании</b>\n"
            "2-я строка (необязательно): <b>ОКТМО Телефон ИНН КПП</b>\n\n"
            "Пример:\n"
            "<code>ООО «СИСТЕМА ПБО»\n"
            "20701000 +74732000000 3662123456 366201001</code>\n\n"
            "Для отмены нажми /cancel",
            parse_mode="HTML",
        )
        return

    if action == "save_last":
        meta = LAST_BLANK_META.get(user_id)
        if not meta or not meta.get("agent_name"):
            await callback.answer("Нет данных компании из последнего бланка.", show_alert=True)
            return
        ep = load_ep_settings()
        tpl = upsert_agent_template(
            {
                "agent_name": meta.get("agent_name", ""),
                "oktmo": meta.get("oktmo", ""),
                "agent_phone": meta.get("agent_phone", ""),
                "agent_inn": meta.get("agent_inn", ""),
                "agent_kpp": meta.get("agent_kpp", ""),
                "ep_owner": ep.get("owner") if ep.get("owner") else meta.get("agent_name", ""),
                "ep_cert": ep.get("cert", ""),
                "ep_ca": ep.get("ca", "Федеральная налоговая служба"),
                "ep_valid_from": ep.get("valid_from", ""),
                "ep_valid_to": ep.get("valid_to", ""),
            }
        )
        await callback.answer(f"Сохранено: {tpl['agent_name'][:25]}", show_alert=True)
        return

    if action == "del" and len(parts) >= 3:
        tpl_id = parts[2]
        agents = [a for a in load_agents() if a.get("id") != tpl_id]
        save_agents(agents)
        await callback.answer("Шаблон удалён")
        text, kb = build_agents_menu()
        try:
            await callback.message.edit_text(text, reply_markup=kb, parse_mode="HTML")
        except Exception:
            await callback.message.answer(text, reply_markup=kb, parse_mode="HTML")
        return

    if action == "use_ep" and len(parts) >= 3:
        tpl_id = parts[2]
        tpl = next((a for a in load_agents() if a.get("id") == tpl_id), None)
        if not tpl:
            await callback.answer("Шаблон не найден", show_alert=True)
            return
        ep = load_ep_settings()
        ep["owner"] = tpl.get("ep_owner") or tpl.get("agent_name") or ep["owner"]
        if tpl.get("ep_cert"):
            ep["cert"] = tpl["ep_cert"]
        if tpl.get("ep_ca"):
            ep["ca"] = tpl["ep_ca"]
        if tpl.get("ep_valid_from"):
            ep["valid_from"] = tpl["ep_valid_from"]
        if tpl.get("ep_valid_to"):
            ep["valid_to"] = tpl["ep_valid_to"]
        save_ep_settings(ep)
        await callback.answer("Реквизиты ЭП обновлены из шаблона")
        text, kb = build_ep_menu(user_id)
        await callback.message.answer(text, reply_markup=kb, parse_mode="HTML")
        return


@dp.callback_query(F.data.startswith("edit:"))
async def handle_edit_callback(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    user_id = callback.from_user.id
    action = callback.data.split(":", 1)[1]

    if user_id not in LAST_TRANSFERS and action != "reset":
        await callback.answer("Сначала отправь PDF-выписку из Т-Банка.", show_alert=True)
        return

    if action == "menu":
        clear_active_inputs(user_id)
        await callback.answer()
        text, kb = build_edit_menu(user_id)
        await callback.message.answer(text, reply_markup=kb, parse_mode="HTML")
        return

    if action == "del":
        clear_active_inputs(user_id)
        EDIT_MODE[user_id] = "del"
        await callback.answer()
        await callback.message.answer(
            "Введи <b>номера строк</b> для удаления через пробел или диапазоном.\n"
            "Например: <code>1 4 7</code> или <code>2-5</code>\n\nДля отмены: /cancel",
            parse_mode="HTML",
        )
        return

    if action == "add":
        clear_active_inputs(user_id)
        EDIT_MODE[user_id] = "add"
        await callback.answer()
        await callback.message.answer(
            "Введи операцию для добавления:\n"
            "• <code>деп 7000 25.11.2026 18:10 сбп</code> (попадёт в Депозиты +)\n"
            "• <code>выв 5000 26.11.2026 14:30 карта</code> (попадёт в Выводы -)\n\n"
            "Дату, время и способ можно не писать — подставятся текущие и СБП.\n"
            "Для отмены: /cancel",
            parse_mode="HTML",
        )
        return

    if action == "filter":
        clear_active_inputs(user_id)
        EDIT_MODE[user_id] = "filter"
        await callback.answer()
        await callback.message.answer(
            "Введи условие фильтра (можно комбинировать):\n"
            "• По сумме: <code>от 3000</code> или <code>от 2000 до 15000</code>\n"
            "• По датам: <code>25.11.2026-26.11.2026</code> или <code>26.11.2026</code>\n"
            "• Вместе: <code>25.11.2026-26.11.2026 от 5000</code>\n\n"
            "Для отмены: /cancel",
            parse_mode="HTML",
        )
        return

    if action == "reset":
        clear_active_inputs(user_id)
        orig = LAST_TRANSFERS_ORIG.get(user_id)
        if not orig:
            await callback.answer("Исходный список не найден.", show_alert=True)
            return
        LAST_TRANSFERS[user_id] = copy.deepcopy(orig)
        await callback.answer("Правки сброшены")
        text, kb = build_edit_menu(user_id)
        await callback.message.answer(text, reply_markup=kb, parse_mode="HTML")
        return

    if action == "done":
        clear_active_inputs(user_id)
        transfers = LAST_TRANSFERS.get(user_id, [])
        if not transfers:
            await callback.answer("В списке не осталось операций.", show_alert=True)
            return
        await callback.answer("Пересобираю Excel...")
        excel_bytes = build_excel(transfers)
        stem = LAST_PDF_STEM.get(user_id, "edited")
        caption = build_transfers_caption(transfers)
        await callback.message.answer_document(
            BufferedInputFile(excel_bytes, filename=f"transfers_{stem}_edited.xlsx"),
            caption=caption,
            reply_markup=build_action_keyboard(transfers),
        )
        return


@dp.callback_query(F.data.startswith("ep:"))
async def handle_ep_action(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    user_id = callback.from_user.id
    action = callback.data.split(":", 1)[1]

    if action == "menu":
        clear_active_inputs(user_id)
        await callback.answer()
        text, kb = build_ep_menu(user_id)
        await callback.message.answer(text, reply_markup=kb, parse_mode="HTML")
        return

    if action == "get_png":
        png_bytes = LAST_SIGNED_PNG.get(user_id)
        if not png_bytes:
            await callback.answer("Подписанный PNG не найден.", show_alert=True)
            return
        await callback.answer()
        await callback.message.answer_document(
            BufferedInputFile(png_bytes, filename="signed_blank.png"),
            caption="Подписанный бланк в PNG без сжатия.",
        )
        return

    if action in ("pos_left", "pos_center", "pos_right"):
        ep = load_ep_settings()
        ep["position"] = action.split("_", 1)[1]
        save_ep_settings(ep)
        await callback.answer("Позиция обновлена")
        text, kb = build_ep_menu(user_id)
        try:
            await callback.message.edit_text(text, reply_markup=kb, parse_mode="HTML")
        except Exception:
            await callback.message.answer(text, reply_markup=kb, parse_mode="HTML")
        return

    if action == "edit":
        clear_active_inputs(user_id)
        EP_SESSIONS[user_id] = {"step_idx": 0, "data": {}}
        await callback.answer()
        await send_ep_step_or_finish(callback.message, user_id)
        return

    if action == "apply":
        await callback.answer("Подписываю...")
        await send_signed_blank(callback.message, user_id)
        return

    if action == "apply_from_blank":
        ep = load_ep_settings()
        meta = LAST_BLANK_META.get(user_id) or {}
        tpl_ep = meta.get("tpl_ep")
        if isinstance(tpl_ep, dict):
            for k in ("owner", "cert", "ca", "valid_from", "valid_to"):
                if tpl_ep.get(k):
                    ep[k] = tpl_ep[k]
        elif meta.get("agent_name"):
            ep["owner"] = meta["agent_name"]
        if meta.get("issue_date"):
            ep["signed_date"] = meta["issue_date"]
        save_ep_settings(ep)
        await callback.answer("Подписываю данными из бланка...")
        await send_signed_blank(callback.message, user_id, ep_override=ep)
        return


@dp.callback_query(F.data.startswith("epw:"))
async def handle_ep_wizard_callback(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    user_id = callback.from_user.id
    action = callback.data.split(":", 1)[1]

    if action == "cancel":
        EP_SESSIONS.pop(user_id, None)
        await callback.answer("Отменено")
        await callback.message.answer("Настройка ЭП отменена.")
        return

    session = EP_SESSIONS.get(user_id)
    if not session:
        await callback.answer("Сессия истекла, нажми /sign заново.", show_alert=True)
        return

    await callback.answer()
    err = apply_ep_step(session, user_id, None, action=action)
    if err:
        await callback.message.answer(err, parse_mode="HTML")
        return
    await send_ep_step_or_finish(callback.message, user_id)


@dp.callback_query(F.data.startswith("blk:"))
async def handle_blank_callback(callback: types.CallbackQuery):
    if not is_allowed_user(callback.from_user):
        await callback.answer()
        return

    user_id = callback.from_user.id
    action = callback.data.split(":", 1)[1]

    if action == "get_png":
        pages = LAST_BLANK_PAGES.get(user_id)
        if not pages:
            await callback.answer("Бланк не найден.", show_alert=True)
            return
        await callback.answer()
        await callback.message.answer_document(
            BufferedInputFile(pages[-1], filename="knd_1175018.png"),
            caption="Бланк КНД 1175018 в PNG без сжатия.",
        )
        return

    if action == "cancel":
        BLANK_SESSIONS.pop(user_id, None)
        await callback.answer("Отменено")
        await callback.message.answer("Заполнение бланка отменено.")
        return

    session = BLANK_SESSIONS.get(user_id)
    if not session:
        await callback.answer("Сессия истекла, нажми /blank заново.", show_alert=True)
        return

    await callback.answer()
    err = apply_blank_step(session, None, action=action)
    if err:
        await callback.message.answer(err, parse_mode="HTML")
        return
    await send_blank_step_or_finish(callback.message, user_id)


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
async def handle_text_input(message: types.Message):
    if not is_allowed_user(message.from_user) or not message.from_user:
        return
    user_id = message.from_user.id
    raw_text = (message.text or "").strip()

    # 1. Добавление шаблона налогового агента
    if user_id in WAITING_AGENT_ADD:
        lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
        if not lines:
            await message.answer("Пришли хотя бы название компании или /cancel.")
            return
        agent_name = lines[0]
        details = parse_agent_details("\n".join(lines[1:])) if len(lines) > 1 else {}
        ep = load_ep_settings()
        tpl = upsert_agent_template(
            {
                "agent_name": agent_name,
                "oktmo": details.get("oktmo", ""),
                "agent_phone": details.get("agent_phone", ""),
                "agent_inn": details.get("agent_inn", ""),
                "agent_kpp": details.get("agent_kpp", ""),
                "ep_owner": agent_name.upper(),
                "ep_cert": generate_cert_hex(),
                "ep_ca": ep.get("ca", "Федеральная налоговая служба"),
                "ep_valid_from": ep.get("valid_from", "24.01.2026"),
                "ep_valid_to": ep.get("valid_to", "24.10.2027"),
            }
        )
        WAITING_AGENT_ADD.discard(user_id)
        text, kb = build_agents_menu()
        await message.answer(
            f"Шаблон <b>{html.escape(tpl['agent_name'])}</b> добавлен!\n\n" + text,
            reply_markup=kb,
            parse_mode="HTML",
        )
        return

    # 2. Правка списка операций
    if user_id in EDIT_MODE:
        mode = EDIT_MODE[user_id]
        transfers = LAST_TRANSFERS.get(user_id, [])

        if mode == "del":
            new_list, removed = apply_edit_delete(transfers, raw_text)
            if removed == 0:
                await message.answer("Не найдено подходящих номеров строк. Попробуй снова или /cancel.")
                return
            LAST_TRANSFERS[user_id] = new_list
            EDIT_MODE.pop(user_id, None)
            text, kb = build_edit_menu(user_id)
            await message.answer(f"Удалено строк: {removed}.\n\n" + text, reply_markup=kb, parse_mode="HTML")
            return

        if mode == "add":
            item = apply_edit_add(raw_text)
            if not item:
                await message.answer(
                    "Не удалось разобрать строку.\nПример: <code>деп 7000 25.11.2026 18:10 сбп</code>",
                    parse_mode="HTML",
                )
                return
            transfers.append(item)
            transfers.sort(key=transfer_sort_key, reverse=True)
            LAST_TRANSFERS[user_id] = transfers
            EDIT_MODE.pop(user_id, None)
            text, kb = build_edit_menu(user_id)
            await message.answer("Операция добавлена.\n\n" + text, reply_markup=kb, parse_mode="HTML")
            return

        if mode == "filter":
            res = apply_edit_filter(transfers, raw_text)
            if res is None:
                await message.answer(
                    "Не распознано условие.\nПример: <code>25.11.2026-26.11.2026 от 3000</code>",
                    parse_mode="HTML",
                )
                return
            new_list, desc_filter = res
            LAST_TRANSFERS[user_id] = new_list
            EDIT_MODE.pop(user_id, None)
            text, kb = build_edit_menu(user_id)
            await message.answer(
                f"Применён фильтр ({html.escape(desc_filter)}). Осталось: {len(new_list)} шт.\n\n" + text,
                reply_markup=kb,
                parse_mode="HTML",
            )
            return

    # 3. Активная настройка ЭП
    if user_id in EP_SESSIONS:
        session = EP_SESSIONS[user_id]
        err = apply_ep_step(session, user_id, raw_text, action="text")
        if err:
            await message.answer(err, parse_mode="HTML")
            return
        await send_ep_step_or_finish(message, user_id)
        return

    # 4. Активное заполнение бланка КНД 1175018
    if user_id in BLANK_SESSIONS:
        session = BLANK_SESSIONS[user_id]
        err = apply_blank_step(session, raw_text, action="text")
        if err:
            await message.answer(err, parse_mode="HTML")
            return
        await send_blank_step_or_finish(message, user_id)
        return

    # 5. Ввод данных для скриншота профиля
    if user_id not in WAITING_PROFILE:
        return

    parsed = parse_profile_input(raw_text)
    if not any(k in parsed for k in ("phone", "email", "dob", "name")):
        await message.answer(
            "Не удалось распознать телефон или почту. Пришли их в одном сообщении "
            "(например: <code>+79991234567 mail@gmail.com</code>) или нажми /cancel.",
            parse_mode="HTML",
        )
        return

    profile = load_profile()
    profile.update(parsed)
    save_profile(profile)
    WAITING_PROFILE.discard(user_id)

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
