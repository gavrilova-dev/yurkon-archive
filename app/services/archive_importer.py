from pathlib import Path
from openpyxl import load_workbook


def _header(value: object) -> str:
    return " ".join(str(value or "").replace("\n", " ").strip().casefold().split())


def read_archive_journal(path: str | Path) -> list[dict]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Вход Архив ИЛ-10"] if "Вход Архив ИЛ-10" in wb.sheetnames else wb.active
    header_row = None
    columns = {}
    for row_num in range(1, min(ws.max_row, 15) + 1):
        vals = [_header(c.value) for c in ws[row_num]]
        if any("индекс дела" in v for v in vals) and any("номер протокола" in v for v in vals):
            header_row = row_num
            columns = {v: i + 1 for i, v in enumerate(vals) if v}
            break
    if not header_row:
        raise ValueError("Не найдены столбцы «Индекс дела» и «Номер протокола».")

    def col(fragment: str):
        return next((n for name, n in columns.items() if fragment in name), None)

    index_col, org_col, protocol_col = col("индекс дела"), col("наименование организации"), col("номер протокола")
    if not all((index_col, org_col, protocol_col)):
        raise ValueError("В журнале не хватает столбцов индекса, организации или номера протокола.")

    result = []
    for row_num in range(header_row + 2, ws.max_row + 1):
        idx = ws.cell(row_num, index_col).value
        org = ws.cell(row_num, org_col).value
        protocol = ws.cell(row_num, protocol_col).value
        if idx and org and protocol:
            result.append({
                "archive_index": str(idx).strip(),
                "organization": str(org).strip(),
                "protocol": str(protocol).strip(),
                "source_row": row_num,
            })
    wb.close()
    return result
