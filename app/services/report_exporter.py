from __future__ import annotations

from pathlib import Path
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
          "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]


def _date(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    try:
        return datetime.fromisoformat(str(value)[:10])
    except (ValueError, TypeError):
        return str(value)


def _party(record):
    return f'{record.get("executor") or ""} - {record.get("month") or ""}'.strip(" -")


def _admin_amount(record, rates):
    if not record.get("admin_done_date"):
        return 0
    value = record.get("admin_amount")
    return int(value) if value not in (None, "") else int(rates["admin_rate"])


def _set_rows(ws, title, headers, rows, total=None, start_row=3):
    if title:
        title_cell = ws.cell(start_row - 2, 1, title)
        title_cell.font = Font(bold=True, size=12)
    header_row = start_row
    for col, value in enumerate(headers, 1):
        cell = ws.cell(header_row, col, value)
        cell.font = Font(bold=True)
    for row_index, row in enumerate(rows, header_row + 1):
        for col, value in enumerate(row, 1):
            ws.cell(row_index, col, value)
    last_row = header_row + len(rows)
    if total is not None:
        total_row = last_row + 1
        for col, value in enumerate(total, 1):
            cell = ws.cell(total_row, col, value)
            cell.font = Font(bold=True)
        last_row = total_row
    for col in range(1, len(headers) + 1):
        values = [str(ws.cell(row, col).value or "") for row in range(1, last_row + 1)]
        width = min(max(max((len(v) for v in values), default=10) + 2, 12), 42)
        letter = get_column_letter(col)
        current = ws.column_dimensions[letter].width or 0
        ws.column_dimensions[letter].width = max(current, width)
    ws.freeze_panes = ws.cell(header_row + 1, 1).coordinate
    ws.sheet_view.showGridLines = True
    return last_row


def _build_summary(ws, records, batches, year, rates):
    headers = [
        "Партия", "Кол-во контор", "Первичка отдана", "Первичка возвращена",
        "К получению, ₽", "Моя доля, ₽", "Отчёт по первичке сдан",
        "Первичка оплачена", "Админ. выполнено", "Админ. сумма, ₽",
        "Отчёт по админ. работе сдан", "Админ. работа оплачена", "Комментарий",
    ]
    batch_map = {
        (int(b.get("year") or year), str(b.get("executor") or "").casefold(),
         str(b.get("month") or "").casefold()): b for b in batches
    }
    groups = {}
    for record in records:
        key = (int(record.get("year") or year), str(record.get("executor") or "").casefold(),
               str(record.get("month") or "").casefold())
        groups.setdefault(key, []).append(record)
    keys = set(groups)
    keys.update(batch_map)
    month_order = {m.casefold(): i for i, m in enumerate(MONTHS)}
    rows = []
    for key in sorted(keys, key=lambda x: (x[0], x[1], month_order.get(x[2], 99))):
        group = groups.get(key, [])
        batch = batch_map.get(key, {})
        display_executor = next((str(r.get("executor") or "") for r in group if r.get("executor")), key[1])
        rows.append([
            f"{display_executor} - {key[2]}".strip(" -"),
            len(group),
            _date(batch.get("primary_given_date")),
            _date(batch.get("primary_returned_date")),
            sum(int(r.get("workplaces") or 0) * rates["primary_rate"] for r in group),
            sum(int(r.get("workplaces") or 0) * rates["primary_share_rate"] for r in group),
            _date(batch.get("primary_report_date")),
            _date(batch.get("primary_paid_date")),
            sum(1 for r in group if r.get("admin_done_date")),
            sum(_admin_amount(r, rates) for r in group),
            _date(batch.get("admin_report_date")),
            _date(batch.get("admin_paid_date")),
            batch.get("comment") or "",
        ])
    total = [
        "ИТОГО ЗА ГОД", sum(r[1] for r in rows), None, None,
        sum(r[4] for r in rows), sum(r[5] for r in rows), None, None,
        sum(r[8] for r in rows), sum(r[9] for r in rows), None, None, None,
    ]
    _set_rows(ws, f"Сводка по партиям — {year}", headers, rows, total)


def _build_annual(ws, records, batches, year, rates):
    ws["A1"] = f"ГОДОВОЙ ОТЧЁТ — {year}"
    month_headers = ["Месяц", "Кол-во контор", "Рабочих мест", "К получению, ₽",
                     "Моя доля, ₽", "Админ. работа, ₽", "Всего, ₽"]
    month_rows = []
    for month in MONTHS:
        group = [r for r in records if str(r.get("month") or "").casefold() == month.casefold()]
        fee = sum(int(r.get("workplaces") or 0) * rates["primary_rate"] for r in group)
        share = sum(int(r.get("workplaces") or 0) * rates["primary_share_rate"] for r in group)
        admin = sum(_admin_amount(r, rates) for r in group)
        month_rows.append([month, len(group), sum(int(r.get("workplaces") or 0) for r in group),
                           fee, share, admin, share + admin])
    month_total = ["ИТОГО ЗА ГОД"] + [sum(row[col] for row in month_rows) for col in range(1, 7)]
    _set_rows(ws, f"Годовой отчёт — {year}", month_headers, month_rows, month_total, start_row=3)

    executor_headers = ["Исполнитель", "Кол-во контор", "Рабочих мест", "К получению, ₽",
                        "Моя доля, ₽", "Админ. работа, ₽", "Всего, ₽"]
    executors = sorted({str(r.get("executor") or "") for r in records if r.get("executor")},
                       key=str.casefold)
    executor_rows = []
    for executor in executors:
        group = [r for r in records if str(r.get("executor") or "") == executor]
        fee = sum(int(r.get("workplaces") or 0) * rates["primary_rate"] for r in group)
        share = sum(int(r.get("workplaces") or 0) * rates["primary_share_rate"] for r in group)
        admin = sum(_admin_amount(r, rates) for r in group)
        executor_rows.append([executor, len(group), sum(int(r.get("workplaces") or 0) for r in group),
                              fee, share, admin, share + admin])
    executor_total = ["ИТОГО"] + [sum(row[col] for row in executor_rows) for col in range(1, 7)]
    start = 3 + len(month_rows) + 3
    _set_rows(ws, "По исполнителям", executor_headers, executor_rows, executor_total, start_row=start)

    primary_paid = primary_wait = admin_paid = admin_wait = 0
    for batch in batches:
        key = (str(batch.get("executor") or "").casefold(), str(batch.get("month") or "").casefold())
        group = [r for r in records if (str(r.get("executor") or "").casefold(),
                                        str(r.get("month") or "").casefold()) == key]
        primary_amount = sum(int(r.get("workplaces") or 0) * rates["primary_rate"] for r in group)
        admin_amount = sum(_admin_amount(r, rates) for r in group)
        if batch.get("primary_paid_date"):
            primary_paid += primary_amount
        else:
            primary_wait += primary_amount
        if batch.get("admin_paid_date"):
            admin_paid += admin_amount
        else:
            admin_wait += admin_amount
    payment_rows = [
        ["Первичка оплачена", primary_paid],
        ["Первичка ожидает оплаты", primary_wait],
        ["Административная работа оплачена", admin_paid],
        ["Административная работа ожидает оплаты", admin_wait],
    ]
    payment_start = start + len(executor_rows) + 5
    _set_rows(ws, "Статус оплаты", ["Показатель", "Сумма, ₽"], payment_rows,
              start_row=payment_start)


def _build_manager(ws, records, party, manager_kind, rates):
    group = [r for r in records if _party(r).casefold() == party.casefold()]
    ws["A1"] = "ОТЧЁТ РУКОВОДИТЕЛЮ"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = "Партия"
    ws["B2"] = party
    if manager_kind == "primary":
        headers = ["№ протокола", "Название организации", "Кол-во рабочих мест",
                   f"Сумма по ставке {rates['primary_rate']} ₽ за место"]
        rows = [[r.get("protocol"), r.get("organization"), int(r.get("workplaces") or 0),
                 int(r.get("workplaces") or 0) * rates["primary_rate"]]
                for r in sorted(group, key=lambda r: str(r.get("protocol") or ""))]
        total = ["ИТОГО", "", sum(row[2] for row in rows), sum(row[3] for row in rows)]
    else:
        headers = ["№ протокола", "Название организации", "Дата выполнения", "Сумма, ₽"]
        admin_records = [r for r in group if r.get("admin_done_date")]
        rows = [[r.get("protocol"), r.get("organization"), _date(r.get("admin_done_date")),
                 _admin_amount(r, rates)]
                for r in sorted(admin_records, key=lambda r: str(r.get("protocol") or ""))]
        total = ["ИТОГО", "", len(rows), sum(row[3] for row in rows)]
    _set_rows(ws, "Первичная документация" if manager_kind == "primary" else "Административная работа", headers, rows, total, start_row=5)


def export_report(template_path: str | Path | None, output_path: str | Path, records, batches,
                  year: int = 2026, report_type: str = "all", party: str = "",
                  manager_kind: str = "primary", rates: dict | None = None) -> Path:
    """Build a clean value-only workbook; no source-template objects, formulas or external links."""
    output_path = Path(output_path)
    rates = rates or {"primary_rate": 20, "primary_share_rate": 10, "admin_rate": 100}
    records = [dict(r) for r in records if int(r.get("year") or year) == int(year)]
    batches = [dict(b) for b in batches if int(b.get("year") or year) == int(year)]
    workbook = Workbook()
    workbook.remove(workbook.active)
    if report_type in ("all", "summary"):
        ws = workbook.create_sheet(f"{year} Сводка по партиям")
        _build_summary(ws, records, batches, year, rates)
    if report_type in ("all", "manager"):
        parties = sorted({_party(r) for r in records if _party(r)}, key=str.casefold)
        if not party and parties:
            party = parties[0]
        ws = workbook.create_sheet(f"{year} Отчёт руководителю")
        _build_manager(ws, records, party, manager_kind, rates)
    if report_type in ("all", "annual"):
        ws = workbook.create_sheet(f"{year} Годовой отчет")
        _build_annual(ws, records, batches, year, rates)
    if not workbook.sheetnames:
        raise ValueError("Неизвестный тип отчёта.")
    for ws in workbook.worksheets:
        ws.sheet_view.showGridLines = True
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, (date, datetime)):
                    cell.number_format = "DD.MM.YYYY"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    return output_path
