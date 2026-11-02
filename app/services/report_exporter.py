from __future__ import annotations

from copy import copy
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.datavalidation import DataValidation


MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]


def _copy_row_style(ws, source_row: int, target_row: int, max_col: int) -> None:
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
        if src.protection:
            dst.protection = copy(src.protection)
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def _date(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    try:
        return datetime.fromisoformat(str(value))
    except ValueError:
        return value


def export_report(template_path: str | Path, output_path: str | Path, records, batches, year: int = 2026) -> Path:
    """Fill a copy of the user's original workbook. Every original sheet is retained."""
    template_path, output_path = Path(template_path), Path(output_path)
    if not template_path.exists():
        raise FileNotFoundError(f"Не найден шаблон отчёта: {template_path}")
    wb = load_workbook(template_path)
    required = {"2026 Реестр", "2026 Сводка по партиям", "2026 Годовой отчет"}
    if not required.issubset(set(wb.sheetnames)):
        raise ValueError("Шаблон не похож на «2026_учет_контор.xlsx»: не найдены обязательные листы.")

    ws = wb["2026 Реестр"]
    # Keep the header, column widths, styles, formulas and Excel table.
    source_style_row = min(2, ws.max_row)
    last_data_row = max(1, len(records) + 1)
    for row in range(2, last_data_row + 1):
        if row > ws.max_row:
            _copy_row_style(ws, source_style_row, row, 12)
        rec = records[row - 2] if row - 2 < len(records) else None
        if rec:
            values = [
                f"=ROW()-1", rec["executor"], rec["month"],
                f'=IF(B{row}="","",B{row}&" - "&C{row})',
                rec["protocol"], rec["organization"], int(rec["workplaces"] or 0),
                f'=IF(G{row}="","",G{row}*20)', f'=IF(G{row}="","",G{row}*10)',
                _date(rec["admin_done_date"]), f'=IF(J{row}<>"",100,0)', None,
            ]
        else:
            values = [f"=ROW()-1", None, None, f'=IF(B{row}="","",B{row}&" - "&C{row})',
                      None, None, None, f'=IF(G{row}="","",G{row}*20)',
                      f'=IF(G{row}="","",G{row}*10)', None, f'=IF(J{row}<>"",100,0)', None]
        for col, value in enumerate(values, 1):
            ws.cell(row, col).value = value
    # Keep a subtotal row immediately after the current data, and remove stale rows.
    total_registry_row = last_data_row + 1
    _copy_row_style(ws, min(418, ws.max_row), total_registry_row, 12)
    if total_registry_row != 418 and 418 <= ws.max_row:
        for col in range(1, 13):
            ws.cell(418, col).value = None
    ws.cell(total_registry_row, 6).value = "ИТОГО (видимые строки):"
    for col, formula in {
        7: f"=SUBTOTAL(109,G2:G{last_data_row})",
        8: f"=SUBTOTAL(109,H2:H{last_data_row})",
        9: f"=SUBTOTAL(109,I2:I{last_data_row})",
        11: f"=SUBTOTAL(109,K2:K{last_data_row})",
    }.items():
        ws.cell(total_registry_row, col).value = formula
    for row in range(total_registry_row + 1, ws.max_row + 1):
        for col in range(1, 13):
            ws.cell(row, col).value = None
    table = ws.tables.get("Реестр2026")
    if table:
        table.ref = f"A1:L{last_data_row}"
    ws.auto_filter.ref = f"A1:L{last_data_row}"

    # Batch register (one row per executor + month).
    sw = wb["2026 Сводка по партиям"]
    batch_start, total_row = 5, 5 + len(batches)
    style_source = 5
    for i, batch in enumerate(batches):
        r = batch_start + i
        _copy_row_style(sw, style_source, r, 13)
        party = f'{batch["executor"]} - {batch["month"]}'
        vals = [
            party,
            f'=COUNTIF(Реестр2026[Партия],A{r})',
            _date(batch["primary_given_date"]),
            _date(batch["primary_returned_date"]),
            f'=SUMIF(Реестр2026[Партия],A{r},Реестр2026[Первичка: к получению (по 20 ₽)])',
            f'=SUMIF(Реестр2026[Партия],A{r},Реестр2026[Первичка: моя доля (по 10 ₽)])',
            _date(batch["primary_report_date"]),
            _date(batch["primary_paid_date"]),
            f'=COUNTIFS(Реестр2026[Партия],A{r},Реестр2026[Админ. работа выполнена (дата)],"<>")',
            f'=SUMIF(Реестр2026[Партия],A{r},Реестр2026[Админ. сумма (100 ₽)])',
            _date(batch["admin_report_date"]),
            _date(batch["admin_paid_date"]),
            batch["comment"] or "",
        ]
        for c, value in enumerate(vals, 1):
            sw.cell(r, c).value = value
    old_end = sw.max_row
    _copy_row_style(sw, old_end, total_row, 13)
    for r in range(total_row + 1, old_end + 1):
        for c in range(1, 14):
            sw.cell(r, c).value = None
    total_values = ["ИТОГО ЗА ГОД", f"=SUM(B5:B{total_row-1})", None, None,
                    f"=SUM(E5:E{total_row-1})", f"=SUM(F5:F{total_row-1})", None, None,
                    f"=SUM(I5:I{total_row-1})", f"=SUM(J5:J{total_row-1})", None, None, None]
    for c, value in enumerate(total_values, 1):
        sw.cell(total_row, c).value = value
    if "2026 Отчёт руководителю" in wb.sheetnames and batches:
        manager = wb["2026 Отчёт руководителю"]
        manager["B3"] = f'{batches[0]["executor"]} - {batches[0]["month"]}'
        batch_end = total_row - 1
        validation = DataValidation(
            type="list",
            formula1='=INDIRECT("' + "'" + str(year) + " Сводка по партиям'!$A$5:$A$" + str(batch_end) + '")',
            allow_blank=False,
        )
        manager.add_data_validation(validation)
        validation.add(manager["B3"])
    st = sw.tables.get("Свод2026")
    if st:
        st.ref = f"A4:M{total_row-1}"

    # Lists used by drop-downs and reports.
    executors = sorted({str(r["executor"]) for r in records if r["executor"]})
    if "Списки" in wb.sheetnames:
        lists = wb["Списки"]
        for r in range(2, max(lists.max_row, len(executors) + 1) + 1):
            lists.cell(r, 1).value = executors[r-2] if r-2 < len(executors) else None
        for i, month in enumerate(MONTHS, 2):
            lists.cell(i, 2).value = month

    # Yearly report is retained with its original layout and recalculates from the table.
    yr = wb["2026 Годовой отчет"]
    for i, month in enumerate(MONTHS, 5):
        yr.cell(i, 1).value = month
        for col in range(2, 8):
            letter = get_column_letter(col)
            formula = {
                2: f'=COUNTIF(Реестр2026[Месяц],A{i})',
                3: f'=SUMIF(Реестр2026[Месяц],A{i},Реестр2026[Кол-во рабочих мест])',
                4: f'=SUMIF(Реестр2026[Месяц],A{i},Реестр2026[Первичка: к получению (по 20 ₽)])',
                5: f'=SUMIF(Реестр2026[Месяц],A{i},Реестр2026[Первичка: моя доля (по 10 ₽)])',
                6: f'=SUMIF(Реестр2026[Месяц],A{i},Реестр2026[Админ. сумма (100 ₽)])',
                7: f'=E{i}+F{i}',
            }[col]
            yr.cell(i, col).value = formula
    # Update executor rows without overwriting the payment-status section below.
    header_row = 20
    extra_rows = max(0, len(executors) - 3)
    if extra_rows:
        yr.insert_rows(24, extra_rows)
    for offset, executor in enumerate(executors):
        row = header_row + 1 + offset
        if row >= 23:
            _copy_row_style(yr, 22, row, 7)
        yr.cell(row, 1).value = executor
        yr.cell(row, 2).value = f'=COUNTIF(Реестр2026[Исполнитель],A{row})'
        yr.cell(row, 3).value = f'=SUMIF(Реестр2026[Исполнитель],A{row},Реестр2026[Кол-во рабочих мест])'
        yr.cell(row, 4).value = f'=SUMIF(Реестр2026[Исполнитель],A{row},Реестр2026[Первичка: к получению (по 20 ₽)])'
        yr.cell(row, 5).value = f'=SUMIF(Реестр2026[Исполнитель],A{row},Реестр2026[Первичка: моя доля (по 10 ₽)])'
        yr.cell(row, 6).value = f'=SUMIF(Реестр2026[Исполнитель],A{row},Реестр2026[Админ. сумма (100 ₽)])'
        yr.cell(row, 7).value = f'=E{row}+F{row}'
    for row in range(header_row + 1 + len(executors), 23):
        for col in range(1, 8):
            yr.cell(row, col).value = None
    executor_table = yr.tables.get("ГодПоИсполнителям")
    if executor_table:
        executor_table.ref = f"A{header_row}:G{header_row + len(executors)}"
    if yr.print_area:
        yr.print_area = f"A1:G{29 + extra_rows}"

    # All report variants stay in the file and are recalculated by Excel on open.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass
    # The source workbook may be from 2026, but each exported workbook is year-specific.
    for suffix in ("Реестр", "Сводка по партиям", "Отчёт руководителю", "Годовой отчет"):
        old_name = f"2026 {suffix}"
        if old_name in wb.sheetnames and year != 2026:
            wb[old_name].title = f"{year} {suffix}"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
