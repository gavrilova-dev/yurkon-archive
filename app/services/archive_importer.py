from pathlib import Path
from openpyxl import load_workbook


def _header(value: object) -> str:
    return " ".join(str(value or "").replace("\n", " ").strip().casefold().split())


def _person_key(value: object) -> str:
    return "".join(char for char in str(value or "").casefold() if char.isalnum())


def read_archive_journal(path: str | Path) -> list[dict]:
    wb = load_workbook(path, data_only=True)
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
    submitter_col = col("сдавшего дело в архив")
    if not all((index_col, org_col, protocol_col, submitter_col)):
        raise ValueError(
            "В журнале не хватает столбцов индекса, организации, номера протокола "
            "или ФИО сдавшего дело в архив."
        )

    result = []
    max_col = max(index_col, org_col, protocol_col, submitter_col)
    for row_num, values in enumerate(
        ws.iter_rows(min_row=header_row + 1, min_col=1, max_col=max_col, values_only=True),
        start=header_row + 1,
    ):
        idx = values[index_col - 1]
        org = values[org_col - 1]
        protocol = values[protocol_col - 1]
        submitter = values[submitter_col - 1]
        # Some journal exports insert a numeric helper row directly below headers.
        # Skip only that row, not the first real organization if the helper row is absent.
        numeric_helper = all(
            isinstance(value, (int, float)) and not isinstance(value, bool)
            for value in (idx, org, protocol)
        )
        if numeric_helper:
            continue
        if (
            idx not in (None, "")
            and org not in (None, "")
            and protocol not in (None, "")
            and _person_key(submitter) == _person_key("Кормин М.В.")
        ):
            result.append({
                "archive_index": str(idx).strip(),
                "organization": str(org).strip(),
                "protocol": str(protocol).strip(),
                "source_row": row_num,
            })
    wb.close()
    return result
