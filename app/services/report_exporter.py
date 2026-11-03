from __future__ import annotations

from copy import copy
from pathlib import Path
from datetime import date, datetime
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]


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
        return value


def _copy_row_style(ws, source_row, target_row, max_col):
    if source_row == target_row:
        return
    for col in range(1, max_col + 1):
        src, dst = ws.cell(source_row, col), ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def _party(record):
    return f'{record.get("executor") or ""} - {record.get("month") or ""}'.strip(" -")


def _sum(records, key):
    return sum(float(r.get(key) or 0) for r in records)


def _party_records(records, party):
    return [r for r in records if _party(r).casefold() == party.casefold()]


def _fill_summary(ws, records, batches, year):
    ws["A1"] = f"Сводка по партиям — {year}"
    rows = []
    batch_map = {(str(b["executor"]), str(b["month"]).casefold()): b for b in batches}
    keys = sorted({(str(b["executor"]), str(b["month"]).casefold()) for b in batches},
                  key=lambda x: (x[0].casefold(), MONTHS.index(x[1].capitalize()) if x[1].capitalize() in MONTHS else 99))
    record_groups = {}
    for rec in records:
        record_groups.setdefault((str(rec.get("executor") or ""), str(rec.get("month") or "").casefold()), []).append(rec)
    keys = sorted(set(keys) | set(record_groups), key=lambda x: (x[0].casefold(), MONTHS.index(x[1].capitalize()) if x[1].capitalize() in MONTHS else 99))
    for i, key in enumerate(keys, 5):
        batch = batch_map.get(key, {})
        group = record_groups.get(key, [])
        _copy_row_style(ws, 5, i, 13)
        values = [
            f"{key[0]} - {key[1]}",
            len(group),
            _date(batch.get("primary_given_date")),
            _date(batch.get("primary_returned_date")),
            sum(int(r.get("workplaces") or 0) * 20 for r in group),
            sum(int(r.get("workplaces") or 0) * 10 for r in group),
            _date(batch.get("primary_report_date")),
            _date(batch.get("primary_paid_date")),
            sum(1 for r in group if r.get("admin_done_date")),
            sum(100 for r in group if r.get("admin_done_date")),
            _date(batch.get("admin_report_date")),
            _date(batch.get("admin_paid_date")),
            batch.get("comment") or "",
        ]
        for c, value in enumerate(values, 1):
            ws.cell(i, c).value = value
        rows.append(values)
    total_row = 5 + len(rows)
    _copy_row_style(ws, max(5, ws.max_row), total_row, 13)
    totals = ["ИТОГО ЗА ГОД", sum(x[1] for x in rows), None, None,
              sum(x[4] for x in rows), sum(x[5] for x in rows), None, None,
              sum(x[8] for x in rows), sum(x[9] for x in rows), None, None, None]
    for c, value in enumerate(totals, 1):
        ws.cell(total_row, c).value = value
    for r in range(total_row + 1, ws.max_row + 1):
        for c in range(1, 14):
            ws.cell(r, c).value = None
    if "Свод2026" in ws.tables:
        ws.tables["Свод2026"].ref = f"A4:M{total_row-1}"


def _fill_yearly(ws, records, batches, year):
    ws["A1"] = f"ГОДОВОЙ ОТЧЁТ — {year}"
    for i, month in enumerate(MONTHS, 5):
        group = [r for r in records if str(r.get("month") or "").casefold() == month.casefold()]
        vals = [month, len(group), sum(int(r.get("workplaces") or 0) for r in group),
                sum(int(r.get("workplaces") or 0) * 20 for r in group),
                sum(int(r.get("workplaces") or 0) * 10 for r in group),
                sum(100 for r in group if r.get("admin_done_date"))]
        vals.append(vals[4] + vals[5])
        for c, value in enumerate(vals, 1):
            ws.cell(i, c).value = value
    totals = ["ИТОГО ЗА ГОД"] + [sum(float(ws.cell(r, c).value or 0) for r in range(5, 17)) for c in range(2, 8)]
    for c, value in enumerate(totals, 1):
        ws.cell(17, c).value = value

    executors = sorted({str(r.get("executor") or "") for r in records if r.get("executor")})
    for offset, executor in enumerate(executors):
        row = 21 + offset
        if row > 22:
            _copy_row_style(ws, 22, row, 7)
        group = [r for r in records if str(r.get("executor") or "") == executor]
        vals = [executor, len(group), sum(int(r.get("workplaces") or 0) for r in group),
                sum(int(r.get("workplaces") or 0) * 20 for r in group),
                sum(int(r.get("workplaces") or 0) * 10 for r in group),
                sum(100 for r in group if r.get("admin_done_date"))]
        vals.append(vals[4] + vals[5])
        for c, value in enumerate(vals, 1):
            ws.cell(row, c).value = value
    for row in range(21 + len(executors), max(ws.max_row, 23) + 1):
        for c in range(1, 8):
            ws.cell(row, c).value = None
    if "ГодПоИсполнителям" in ws.tables:
        ws.tables["ГодПоИсполнителям"].ref = f"A20:G{max(20,20+len(executors))}"

    # Payment status summary is based only on manually entered payment dates in batch control.
    primary_paid = primary_wait = admin_paid = admin_wait = 0
    for b in batches:
        group = _party_records(records, f'{b.get("executor")} - {b.get("month")}')
        primary_amount = sum(int(r.get("workplaces") or 0) * 20 for r in group)
        admin_amount = sum(100 for r in group if r.get("admin_done_date"))
        if b.get("primary_paid_date"):
            primary_paid += primary_amount
        else:
            primary_wait += primary_amount
        if b.get("admin_paid_date"):
            admin_paid += admin_amount
        else:
            admin_wait += admin_amount
    for row, value in zip(range(26, 30), [primary_paid, primary_wait, admin_paid, admin_wait]):
        ws.cell(row, 2).value = value
    for row in (4, 20, 25):
        for cell in ws[row]:
            if cell.value is not None:
                font = copy(cell.font)
                font.bold = True
                cell.font = font


def _clear_cells(ws, min_row, max_row, min_col=1, max_col=5):
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row, col)
            if not isinstance(cell, MergedCell):
                cell.value = None


def _fill_manager(ws, records, party, manager_kind):
    ws["B3"] = party
    chosen = _party_records(records, party)
    primary = sorted(chosen, key=lambda r: str(r.get("protocol") or ""))
    admin = [r for r in primary if r.get("admin_done_date")]
    if manager_kind == "primary":
        _clear_cells(ws, 139, 273)
        ws["A5"] = "ОТЧЁТ ПО ПАРТИИ КОНТОР — ПЕРВИЧНАЯ ДОКУМЕНТАЦИЯ"
        for i, rec in enumerate(primary[:130], 7):
            vals = [rec.get("protocol"), rec.get("organization"), int(rec.get("workplaces") or 0),
                    int(rec.get("workplaces") or 0) * 20]
            for c, value in enumerate(vals, 1):
                ws.cell(i, c).value = value
        ws["B137"] = "ИТОГО ПО ПЕРВИЧНОЙ ДОКУМЕНТАЦИИ:"
        ws["C137"] = len(primary)
        ws["D137"] = sum(int(r.get("workplaces") or 0) * 20 for r in primary)
        ws["A139"] = None
        ws["A273"] = None
    else:
        _clear_cells(ws, 5, 138)
        ws["A139"] = "ОТЧЁТ ПО ПАРТИИ КОНТОР — АДМИНИСТРАТИВНАЯ РАБОТА"
        for i, rec in enumerate(admin[:130], 141):
            vals = [rec.get("protocol"), rec.get("organization"), _date(rec.get("admin_done_date")), 100]
            for c, value in enumerate(vals, 1):
                ws.cell(i, c).value = value
        ws["B271"] = "ИТОГО ПО АДМИН. РАБОТЕ:"
        ws["C271"] = len(admin)
        ws["D271"] = len(admin) * 100
        ws["A273"] = None
    # Remove all leftover formula cells in both report blocks.
    if manager_kind == "primary":
        for row in range(7 + min(len(primary),130), 137):
            for col in range(1,5): ws.cell(row,col).value = None
    else:
        for row in range(141 + min(len(admin),130), 271):
            for col in range(1,5): ws.cell(row,col).value = None


def export_report(template_path: str | Path, output_path: str | Path, records, batches,
                  year: int = 2026, report_type: str = "all", party: str = "",
                  manager_kind: str = "primary") -> Path:
    """Populate the original workbook template with calculated values (never Excel formulas)."""
    template_path, output_path = Path(template_path), Path(output_path)
    if not template_path.exists():
        raise FileNotFoundError(f"Не найден шаблон отчёта: {template_path}")
    wb = load_workbook(template_path)
    required = {"2026 Сводка по партиям", "2026 Отчёт руководителю", "2026 Годовой отчет"}
    if not required.issubset(set(wb.sheetnames)):
        raise ValueError("Шаблон не похож на «2026_учет_контор.xlsx»: не найдены три листа отчётности.")
    records = [dict(r) for r in records if int(r.get("year") or year) == int(year)]
    batches = [dict(b) for b in batches if int(b.get("year") or year) == int(year)]
    sw = wb["2026 Сводка по партиям"]
    yw = wb["2026 Годовой отчет"]
    mw = wb["2026 Отчёт руководителю"]
    _fill_summary(sw, records, batches, year)
    _fill_yearly(yw, records, batches, year)
    parties = sorted({_party(r) for r in records if _party(r)})
    if not party and parties:
        party = parties[0]
    _fill_manager(mw, records, party, manager_kind)

    # Replace year labels and table references. Report tabs are retained in preview workbook.
    for suffix in ("Сводка по партиям", "Отчёт руководителю", "Годовой отчет"):
        old = f"2026 {suffix}"
        if year != 2026 and old in wb.sheetnames:
            wb[old].title = f"{year} {suffix}"
    if report_type != "all":
        selected = {"summary": f"{year} Сводка по партиям",
                    "annual": f"{year} Годовой отчет",
                    "manager": f"{year} Отчёт руководителю"}.get(report_type)
        if not selected or selected not in wb.sheetnames:
            raise ValueError("Неизвестный тип отчёта.")
        for name in list(wb.sheetnames):
            if name != selected:
                del wb[name]
    try:
        wb.calculation.fullCalcOnLoad = False
        wb.calculation.forceFullCalc = False
        wb.calculation.calcMode = "manual"
    except AttributeError:
        pass
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
