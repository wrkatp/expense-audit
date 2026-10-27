#!/usr/bin/env python3
"""Аудит расходов из CSV- или Excel-файла.

Использование:
    python3 audit.py <файл.csv|файл.xlsx> [-o report.txt]

Столбцы (дата, категория, описание, сумма) определяются автоматически:
сначала по названиям в заголовке (русским и английским), затем по содержимому.
Если автоопределение ошиблось, столбец можно указать явно:
    --amount-col "Сумма списания" --date-col 1   (название или номер с 1)

Работает без сторонних библиотек. Суммы считаются через Decimal,
поэтому итоги точные, без ошибок округления float.
"""

import argparse
import csv
import io
import re
import statistics
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation

NO_CATEGORY = "(без категории)"
SIGMA_THRESHOLD = 3
TOP_N = 5
HEADER_SCAN_ROWS = 30

# Названия столбцов: сначала точные совпадения, потом вхождение подстроки.
# Порядок важен: для суммы «расход» и «списание» важнее «прихода».
COLUMN_SYNONYMS = {
    "date": ["дата", "date", "дата операции", "дата платежа", "дата расхода", "день",
             "transaction date", "posting date", "период"],
    "category": ["категория", "category", "статья", "статья расходов", "вид расхода",
                 "группа", "тип расхода", "раздел", "category name", "expense type"],
    "description": ["описание", "description", "назначение", "назначение платежа",
                    "комментарий", "наименование", "контрагент", "получатель",
                    "поставщик", "memo", "details", "payee", "merchant", "примечание"],
    "amount": ["сумма", "amount", "сумма расхода", "расход", "списание", "сумма операции",
               "сумма платежа", "стоимость", "итого", "sum", "total", "cost", "debit",
               "price", "value"],
}
TOTAL_ROW_MARKERS = ("итого", "всего", "total", "subtotal")

DATE_FORMATS = ["%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%Y/%m/%d",
                "%d-%m-%Y", "%m/%d/%Y", "%Y.%m.%d"]


# ---------------------------------------------------------------- чтение файлов

def read_csv_rows(path):
    raw = open(path, "rb").read()
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")
    sample = text[:20000]
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        delimiter = max(",;\t|", key=sample.count)
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter)
    rows = []
    for cells in reader:
        # Номер строки как в файле (с учётом многострочных ячеек)
        rows.append((reader.line_num, cells))
    return rows


def _col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group()
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx_rows(path, sheet_name=None):
    """Минимальный читатель .xlsx на стандартной библиотеке (без openpyxl)."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", ns):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{ns['m']}}}t")))

        date_styles = set()
        if "xl/styles.xml" in names:
            styles = ET.fromstring(z.read("xl/styles.xml"))
            custom_date_fmts = {
                int(f.get("numFmtId")) for f in styles.findall("m:numFmts/m:numFmt", ns)
                if re.search(r"[dy]|mm.*yy|dd", re.sub(r'"[^"]*"|\[[^]]*]', "", f.get("formatCode", "").lower()))
            }
            for i, xf in enumerate(styles.findall("m:cellXfs/m:xf", ns)):
                fmt_id = int(xf.get("numFmtId", 0))
                if 14 <= fmt_id <= 22 or 45 <= fmt_id <= 47 or fmt_id in custom_date_fmts:
                    date_styles.add(i)

        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels}
        sheets = workbook.findall("m:sheets/m:sheet", ns)
        if sheet_name:
            sheets = [s for s in sheets if s.get("name") == sheet_name]
            if not sheets:
                sys.exit(f"Ошибка: в файле нет листа «{sheet_name}»")
        target = targets[sheets[0].get(f"{{{ns['r']}}}id")].lstrip("/")
        sheet_path = target if target.startswith("xl/") else "xl/" + target

        rows = []
        for row in ET.fromstring(z.read(sheet_path)).iter(f"{{{ns['m']}}}row"):
            values = {}
            for c in row.findall("m:c", ns):
                kind, style = c.get("t"), int(c.get("s", 0))
                v = c.find("m:v", ns)
                if kind == "s":
                    value = shared[int(v.text)]
                elif kind == "inlineStr":
                    value = "".join(t.text or "" for t in c.iter(f"{{{ns['m']}}}t"))
                elif v is None:
                    value = ""
                elif style in date_styles and kind not in ("str", "b"):
                    value = (datetime(1899, 12, 30) + timedelta(days=float(v.text))).date().isoformat()
                else:
                    value = v.text
                values[_col_index(c.get("r"))] = value
            cells = [values.get(i, "") for i in range(max(values) + 1)] if values else []
            rows.append((int(row.get("r")), cells))
        return rows


def read_table(path, sheet=None):
    lower = path.lower()
    if lower.endswith((".xlsx", ".xlsm")):
        return read_xlsx_rows(path, sheet)
    if lower.endswith(".xls"):
        sys.exit("Ошибка: старый формат .xls не поддерживается. "
                 "Пересохраните файл как .xlsx или .csv.")
    return read_csv_rows(path)


# ---------------------------------------------------------------- разбор значений

def parse_amount(text):
    """'1 234,56 ₽' / '-1,234.56' / '(500)' / '−300 руб.' -> Decimal или None."""
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    negative = s.startswith("(") and s.endswith(")")
    s = s.replace("−", "-").replace("–", "-")
    s = re.sub(r"(руб\.?|р\.|rub|usd|eur|₽|\$|€)", "", s, flags=re.I)
    s = re.sub(r"[\s  '()]", "", s)
    if not re.fullmatch(r"[-+]?[\d.,]+(?:[eE][-+]?\d+)?", s):
        return None
    if "," in s and "." in s:
        # Последний разделитель — десятичный
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(",") > 1:
        s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    if s.count(".") > 1:
        s = s.replace(".", "")
    try:
        value = Decimal(s)
    except InvalidOperation:
        return None
    return -value if negative else value


def parse_date(text):
    s = str(text or "").strip()
    if not s:
        return None
    s = re.split(r"[ T]", s)[0]
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def normalize_header(text):
    return re.sub(r"\s+", " ", str(text or "").strip().lower().replace("ё", "е"))


# ---------------------------------------------------------------- определение столбцов

def header_matches(header, field):
    h = normalize_header(header)
    if not h:
        return 0
    synonyms = COLUMN_SYNONYMS[field]
    if h in synonyms:
        return 2
    if any(re.search(rf"\b{re.escape(s)}", h) for s in synonyms):
        return 1
    return 0


def find_header_row(rows):
    """Ищет строку заголовка среди первых строк (в выписках бывает «шапка» сверху)."""
    best_index, best_score = 0, 0
    for i, (_, cells) in enumerate(rows[:HEADER_SCAN_ROWS]):
        score = sum(1 for c in cells for f in COLUMN_SYNONYMS if header_matches(c, f))
        if score > best_score:
            best_index, best_score = i, score
    return best_index if best_score >= 2 else 0


def resolve_override(spec, headers):
    if spec.isdigit():
        index = int(spec) - 1
        if 0 <= index < len(headers):
            return index
    else:
        for i, h in enumerate(headers):
            if normalize_header(h) == normalize_header(spec):
                return i
    sys.exit(f"Ошибка: столбец «{spec}» не найден. Есть: {', '.join(map(str, headers))}")


def detect_columns(headers, data, overrides):
    """Возвращает {поле: индекс столбца или None} и способ определения каждого поля."""
    columns, how = {}, {}
    for field, spec in overrides.items():
        if spec:
            columns[field] = resolve_override(spec, headers)
            how[field] = "указан вручную"

    # 1) По названиям в заголовке: сначала точные совпадения, затем частичные
    for strength in (2, 1):
        for field in ("amount", "date", "category", "description"):
            if field in columns:
                continue
            for i, h in enumerate(headers):
                if i not in columns.values() and header_matches(h, field) == strength:
                    columns[field] = i
                    how[field] = "по заголовку"
                    break

    # 2) По содержимому для того, что не нашлось
    width = max([len(headers)] + [len(c) for _, c in data[:200]])
    sample = [cells for _, cells in data[:500]]

    def values(i):
        return [c[i].strip() for c in sample if i < len(c) and str(c[i]).strip()]

    def share(i, parser):
        vals = values(i)
        return sum(parser(v) is not None for v in vals) / len(vals) if vals else 0

    free = lambda: [i for i in range(width) if i not in columns.values()]

    if "date" not in columns:
        candidates = [i for i in free() if share(i, parse_date) >= 0.8]
        if candidates:
            columns["date"], how["date"] = candidates[0], "по содержимому"
    if "amount" not in columns:
        candidates = [i for i in free()
                      if share(i, parse_amount) >= 0.8 and share(i, parse_date) < 0.5]
        if candidates:
            # Из нескольких числовых берём тот, где больше дробных значений (обычно это деньги)
            columns["amount"] = max(candidates, key=lambda i: sum(
                bool(re.search(r"[.,]\d{1,2}\b", v)) for v in values(i)))
            how["amount"] = "по содержимому"
    text_cols = [i for i in free() if values(i) and share(i, parse_amount) < 0.5]
    if "category" not in columns:
        # Категория — текст с небольшим набором повторяющихся значений
        candidates = [i for i in text_cols
                      if len(set(values(i))) <= max(50, len(values(i)) * 0.1)]
        if candidates:
            columns["category"] = min(candidates, key=lambda i: len(set(values(i))))
            how["category"] = "по содержимому"
            text_cols.remove(columns["category"])
    if "description" not in columns and text_cols:
        columns["description"] = max(text_cols, key=lambda i: len(set(values(i))))
        how["description"] = "по содержимому"

    for field in COLUMN_SYNONYMS:
        columns.setdefault(field, None)
    return columns, how


# ---------------------------------------------------------------- загрузка данных

def load(path, overrides, sheet=None):
    table = read_table(path, sheet)
    table = [(n, cells) for n, cells in table if any(str(c).strip() for c in cells)]
    if not table:
        sys.exit("Ошибка: файл пустой.")
    h = find_header_row(table)
    headers = [str(c).strip() for c in table[h][1]]
    data = table[h + 1:]
    columns, how = detect_columns(headers, data, overrides)
    if columns["amount"] is None:
        sys.exit("Ошибка: не удалось найти столбец с суммой. "
                 f"Столбцы в файле: {', '.join(headers)}. Укажите его через --amount-col.")

    def cell(cells, field):
        i = columns[field]
        return str(cells[i]).strip() if i is not None and i < len(cells) else ""

    rows, errors, skipped_totals = [], [], []
    for line_no, cells in data:
        raw_amount = cell(cells, "amount")
        joined = " ".join(str(c) for c in cells).strip().lower()
        if joined.startswith(TOTAL_ROW_MARKERS):
            skipped_totals.append(line_no)
            continue
        if not raw_amount:
            continue
        amount = parse_amount(raw_amount)
        if amount is None:
            errors.append((line_no, raw_amount))
            continue
        raw_date = cell(cells, "date")
        parsed = parse_date(raw_date)
        rows.append({
            "line": line_no,
            "date": parsed.isoformat() if parsed else raw_date,
            "category": cell(cells, "category") or NO_CATEGORY,
            "description": cell(cells, "description"),
            "amount": amount,
        })

    # В банковских выписках расходы часто идут со знаком минус.
    # Если отрицательных большинство, переворачиваем знак: расходы станут положительными,
    # а «минусами» окажутся поступления и возвраты.
    sign_flipped = False
    if rows and sum(r["amount"] < 0 for r in rows) > len(rows) / 2:
        sign_flipped = True
        for r in rows:
            r["amount"] = -r["amount"]

    meta = {"headers": headers, "columns": columns, "how": how, "header_line": table[h][0],
            "errors": errors, "skipped_totals": skipped_totals, "sign_flipped": sign_flipped}
    return rows, meta


# ---------------------------------------------------------------- отчёт

def fmt(amount):
    """Форматирует сумму с пробелами между разрядами, сохраняя все знаки после запятой."""
    # Дополняем до копеек (только нулями — если знаков больше двух, они сохраняются)
    if amount.as_tuple().exponent > -2:
        amount = amount.quantize(Decimal("0.01"))
    text = f"{amount:,f}".replace(",", " ")
    return text.replace(".", ",")


def row_text(row):
    parts = [f"стр. {row['line']:>5}", row["date"] or "—", row["category"]]
    if row["description"]:
        parts.append(row["description"])
    parts.append(fmt(row["amount"]))
    return " | ".join(parts)


def section(title):
    return ["", "=" * 70, title, "=" * 70]


FIELD_NAMES = {"date": "дата", "category": "категория",
               "description": "описание", "amount": "сумма"}


def build_report(path, rows, meta):
    columns, headers = meta["columns"], meta["headers"]
    out = [f"АУДИТ РАСХОДОВ: {path}", f"Строк с данными: {len(rows)}",
           f"Заголовок в строке {meta['header_line']}. Определённые столбцы:"]
    for field, name in FIELD_NAMES.items():
        i = columns[field]
        found = f"«{headers[i] if i < len(headers) and headers[i] else f'столбец {i + 1}'}» ({meta['how'][field]})" \
            if i is not None else "не найден"
        out.append(f"  {name:<10} -> {found}")
    if meta["sign_flipped"]:
        out.append("Большинство сумм было со знаком минус (похоже на банковскую выписку): "
                   "знак перевёрнут, расходы считаются положительными, "
                   "а в разделе 6 — поступления и возвраты.")
    if meta["skipped_totals"]:
        out.append(f"Пропущены строки «Итого/Всего»: {', '.join(map(str, meta['skipped_totals']))}")
    if meta["errors"]:
        out.append(f"Пропущено строк с нечитаемой суммой: {len(meta['errors'])}")
        out += [f"  стр. {n}: «{v}»" for n, v in meta["errors"][:20]]
    if not rows:
        out.append("\nНет строк с суммами — анализировать нечего.")
        return "\n".join(out) + "\n"

    # 1. Общая сумма
    total = sum((r["amount"] for r in rows), Decimal(0))
    out += section("1. ОБЩАЯ СУММА РАСХОДОВ")
    out.append(fmt(total))
    dates = sorted(r["date"] for r in rows if re.fullmatch(r"\d{4}-\d{2}-\d{2}", r["date"]))
    if dates:
        out.append(f"Период: {dates[0]} — {dates[-1]}")

    # 2. Суммы по категориям
    by_category = defaultdict(list)
    for r in rows:
        by_category[r["category"]].append(r)
    totals = {c: sum((r["amount"] for r in rs), Decimal(0)) for c, rs in by_category.items()}
    out += section("2. СУММЫ ПО КАТЕГОРИЯМ (по убыванию)")
    if columns["category"] is None:
        out.append("Столбец с категорией не найден — все траты в одной группе.")
    width = max(len(c) for c in totals)
    for cat, amount in sorted(totals.items(), key=lambda kv: kv[1], reverse=True):
        share = amount / total * 100 if total else Decimal(0)
        out.append(f"{cat:<{width}}  {fmt(amount):>20}  {share:5.1f}%  "
                   f"({len(by_category[cat])} шт.)")

    # 3. Топ крупнейших трат
    out += section(f"3. ТОП-{TOP_N} САМЫХ КРУПНЫХ ТРАТ")
    for i, r in enumerate(sorted(rows, key=lambda r: r["amount"], reverse=True)[:TOP_N], 1):
        out.append(f"{i}. {row_text(r)}")

    # 4. Дубликаты: одинаковые дата, описание и сумма
    #    (если нет описания — сравниваем по категории)
    label_field = "description" if columns["description"] is not None else "category"
    groups = defaultdict(list)
    for r in rows:
        groups[(r["date"], r[label_field], r["amount"])].append(r)
    duplicates = [g for g in groups.values() if len(g) > 1]
    extra = sum((g[0]["amount"] * (len(g) - 1) for g in duplicates), Decimal(0))
    key_name = "описание" if label_field == "description" else "категория"
    out += section(f"4. ДУБЛИКАТЫ (дата + {key_name} + сумма)")
    if columns["date"] is None:
        out.append("Столбец с датой не найден — совпадения ищутся без учёта даты, "
                   "часть из них может быть обычными повторяющимися платежами.")
    if duplicates:
        out.append(f"Групп: {len(duplicates)}, лишних строк: "
                   f"{sum(len(g) - 1 for g in duplicates)}, лишняя сумма: {fmt(extra)}")
        for g in sorted(duplicates, key=lambda g: g[0]["amount"], reverse=True):
            out.append("")
            out += [f"  {row_text(r)}" for r in g]
    else:
        out.append("Дубликатов не найдено.")

    # 5. Аномалии: отклонение от среднего по категории больше чем на 3σ
    out += section(f"5. АНОМАЛИИ (|сумма - среднее по категории| > {SIGMA_THRESHOLD}σ)")
    anomalies = []
    for cat, rs in by_category.items():
        if len(rs) < 2:
            continue
        amounts = [r["amount"] for r in rs]
        mean = statistics.mean(amounts)
        stdev = statistics.stdev(amounts)
        if stdev == 0:
            continue
        for r in rs:
            z = (r["amount"] - mean) / stdev
            if abs(z) > SIGMA_THRESHOLD:
                anomalies.append((r, mean, stdev, z))
    if anomalies:
        for r, mean, stdev, z in sorted(anomalies, key=lambda a: abs(a[3]), reverse=True):
            out.append(f"  {row_text(r)}")
            out.append(f"      среднее по категории {fmt(mean.quantize(Decimal('0.01')))}, "
                       f"σ = {fmt(stdev.quantize(Decimal('0.01')))}, отклонение {z:+.1f}σ")
    else:
        out.append("Аномалий не найдено.")

    # 6. Отрицательные суммы
    negative = [r for r in rows if r["amount"] < 0]
    out += section("6. СТРОКИ С ОТРИЦАТЕЛЬНЫМИ СУММАМИ")
    if negative:
        out += [f"  {row_text(r)}" for r in negative]
    else:
        out.append("Отрицательных сумм нет.")

    return "\n".join(out) + "\n"


def main():
    parser = argparse.ArgumentParser(description="Аудит расходов из CSV- или Excel-файла.")
    parser.add_argument("path", help="путь к файлу .csv или .xlsx")
    parser.add_argument("-o", "--output", default="report.txt",
                        help="куда сохранить отчёт (по умолчанию report.txt)")
    parser.add_argument("--sheet", help="лист Excel (по умолчанию первый)")
    for field, name in FIELD_NAMES.items():
        parser.add_argument(f"--{field}-col", help=f"столбец «{name}»: название или номер с 1")
    args = parser.parse_args()

    overrides = {f: getattr(args, f"{f}_col") for f in FIELD_NAMES}
    try:
        rows, meta = load(args.path, overrides, args.sheet)
    except FileNotFoundError:
        sys.exit(f"Ошибка: файл не найден: {args.path}")
    except zipfile.BadZipFile:
        sys.exit(f"Ошибка: {args.path} не похож на файл Excel (.xlsx).")

    report = build_report(args.path, rows, meta)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(report)
    print(report)
    print(f"Отчёт сохранён в {args.output}")


if __name__ == "__main__":
    main()
