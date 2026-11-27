import asyncio
from collections import Counter
import io
import json
import logging
import os
from pathlib import Path
import re

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.types import BufferedInputFile
import openpyxl
from openpyxl.styles import Alignment, Font
import pdfplumber

# --- НАСТРОЙКИ (задаются в "Переменных окружения" на хостинге) ---
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ALLOWED_USER_ID = int(os.getenv("ALLOWED_USER_ID", "0").strip() or "0")
ENV_EXCLUDED = {
    item.strip()
    for item in os.getenv("EXCLUDED_CONTRACTS", "").split(",")
    if item.strip()
}

EXCLUDED_FILE = Path("excluded_contracts.json")

logging.basicConfig(level=logging.INFO)
dp = Dispatcher()


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


def is_allowed(message: types.Message) -> bool:
    if ALLOWED_USER_ID == 0:
        return True
    return message.from_user is not None and message.from_user.id == ALLOWED_USER_ID


def extract_counterparty(desc: str) -> tuple[str | None, str]:
    """
    Возвращает (ключ_для_поиска_повторов, читаемое_значение_для_таблицы).
    Если в описании нет номера телефона, договора или карты (например, просто СБП),
    возвращает (None, ""), чтобы такие операции не отсекались как повторы.
    """
    phone_match = re.search(
        r"(?:\+?7|8)[\s\-\(\)]*\d{3}[\s\-\(\)]*\d{3}[\s\-\(\)]*\d{2}[\s\-\(\)]*\d{2}",
        desc,
    )
    if phone_match:
        digits = re.sub(r"\D", "", phone_match.group(0))
        if len(digits) == 11 and digits.startswith("8"):
            digits = "7" + digits[1:]
        return f"phone:{digits}", f"+{digits}"

    contract_match = re.search(r"договор\w*\s+(\d{6,})", desc, re.IGNORECASE)
    if contract_match:
        num = contract_match.group(1)
        return f"contract:{num}", f"Договор {num}"

    ext_card_match = re.search(r"(?:\d{4,6}\*+\d{4}|\*+\d{4})", desc)
    if ext_card_match:
        mask = ext_card_match.group(0)
        return f"card:{mask}", f"Карта {mask}"

    return None, ""


def parse_tbank_pdf(
    pdf_bytes: bytes, excluded_contracts: set[str]
) -> tuple[list[dict], int]:
    """
    Извлекает из PDF-выписки Т-Банка переводы людям и пополнения переводом на карту,
    полностью удаляя все операции по номерам/договорам/картам, встретившимся > 1 раза.
    """
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

    include_keywords = (
        "перевод",
        "система быстрых платежей",
        "сбп",
        "пополнение",
        "на карту",
        "с карты",
        "зачисление",
        "card2card",
        "c2c",
    )
    exclude_keywords = (
        "оплата",
        "покупка",
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
    )

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

        card = ""
        card_match = re.search(r"(?<![\d*+])\b(\d{4})\b(?![\d*])", desc)
        if card_match:
            card = card_match.group(1)
            desc = desc[: card_match.start()] + " " + desc[card_match.end() :]

        desc = re.sub(r"\s+", " ", desc).strip()
        desc_low = desc.lower()

        if any(bad in desc_low for bad in exclude_keywords):
            continue
        if any(contract in desc for contract in excluded_contracts):
            continue
        if not any(good in desc_low for good in include_keywords):
            continue

        direction = "Входящий" if amount > 0 else "Исходящий"
        counterparty_key, counterparty_label = extract_counterparty(desc)

        candidates.append(
            {
                "date": dates[0],
                "time": times[0] if times else "",
                "direction": direction,
                "amount": amount,
                "counterparty_key": counterparty_key,
                "counterparty": counterparty_label,
                "description": desc,
                "card": card,
            }
        )

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

    return results, skipped_repeated


def build_excel(rows: list[dict]) -> bytes:
    """Формирует готовый .xlsx файл в оперативной памяти."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Переводы"

    headers = [
        "Дата",
        "Время",
        "Тип",
        "Сумма (₽)",
        "Номер / Договор",
        "Описание",
        "Карта",
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
                row["amount"],
                row["counterparty"],
                row["description"],
                row["card"],
            ]
        )
        ws.cell(row=ws.max_row, column=4).number_format = "#,##0.00"

    col_widths = [12, 8, 12, 14, 22, 46, 8]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    if not is_allowed(message):
        return
    await message.answer(
        "Кидай PDF-выписку из Т-Банка — верну Excel только с разовыми "
        "переводами и пополнениями на карту (повторяющиеся номера полностью "
        "отсекаются).\n\nКоманды:\n"
        "/excluded — список игнорируемых договоров\n"
        "/exclude 1234567890 — добавить договор в исключения\n"
        "/unexclude 1234567890 — убрать из исключений"
    )


@dp.message(Command("excluded"))
async def cmd_excluded(message: types.Message):
    if not is_allowed(message):
        return
    contracts = sorted(load_excluded())
    text = "\n".join(contracts) if contracts else "Список пуст."
    await message.answer("Игнорируемые договоры:\n" + text)


@dp.message(Command("exclude"))
async def cmd_exclude(message: types.Message):
    if not is_allowed(message):
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
    if not is_allowed(message):
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
    if not is_allowed(message):
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
        transfers, skipped_repeated = parse_tbank_pdf(pdf_bytes, excluded)

        if not transfers:
            await wait_msg.edit_text(
                f"Подходящих разовых переводов не найдено (отсечено повторяющихся: {skipped_repeated})."
            )
            return

        total_in = sum(t["amount"] for t in transfers if t["amount"] > 0)
        total_out = sum(t["amount"] for t in transfers if t["amount"] < 0)
        balance = total_in + total_out

        excel_bytes = build_excel(transfers)
        out_name = f"transfers_{Path(doc.file_name).stem}.xlsx"

        caption = (
            f"Найдено разовых переводов: {len(transfers)}\n"
            f"Отсечено повторяющихся операций: {skipped_repeated}\n"
            f"Входящие: +{total_in:,.2f} ₽\n"
            f"Исходящие: {total_out:,.2f} ₽\n"
            f"Разница: {balance:+,.2f} ₽"
        ).replace(",", " ")

        await message.answer_document(
            BufferedInputFile(excel_bytes, filename=out_name),
            caption=caption,
        )
        await wait_msg.delete()

    except Exception as e:
        logging.exception("Ошибка при разборе PDF")
        await wait_msg.edit_text(f"Не удалось разобрать файл: {e}")


async def main():
    if not BOT_TOKEN:
        raise ValueError("Не задан BOT_TOKEN в переменных окружения!")
    bot = Bot(token=BOT_TOKEN)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
