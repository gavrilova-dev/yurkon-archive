from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Any


class Database:
    """Local SQLite storage for archive references, primary work and batch payments."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(self.path)
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.connection.execute("PRAGMA journal_mode = WAL")
        self.initialize()

    def initialize(self) -> None:
        self.connection.executescript("""
        CREATE TABLE IF NOT EXISTS primary_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            executor TEXT NOT NULL DEFAULT '',
            month TEXT NOT NULL DEFAULT '',
            protocol TEXT NOT NULL DEFAULT '',
            organization TEXT NOT NULL DEFAULT '',
            workplaces INTEGER NOT NULL DEFAULT 0,
            admin_done_date TEXT,
            assembled_date TEXT,
            returned_date TEXT,
            signed_date TEXT,
            scanned_date TEXT,
            yandex_uploaded_date TEXT,
            pdf_path TEXT,
            notes TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_primary_protocol ON primary_records(protocol);
        CREATE INDEX IF NOT EXISTS idx_primary_batch ON primary_records(executor, month);

        CREATE TABLE IF NOT EXISTS archive_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            archive_index TEXT NOT NULL,
            organization TEXT NOT NULL,
            protocol TEXT NOT NULL,
            protocol_norm TEXT NOT NULL,
            organization_norm TEXT NOT NULL,
            source_row INTEGER,
            imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(archive_index)
        );
        CREATE INDEX IF NOT EXISTS idx_archive_protocol ON archive_records(protocol_norm);
        CREATE INDEX IF NOT EXISTS idx_archive_index ON archive_records(archive_index);
        CREATE INDEX IF NOT EXISTS idx_archive_org ON archive_records(organization_norm);

        CREATE TABLE IF NOT EXISTS batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            executor TEXT NOT NULL,
            month TEXT NOT NULL,
            primary_given_date TEXT,
            primary_returned_date TEXT,
            primary_report_date TEXT,
            primary_paid_date TEXT,
            admin_report_date TEXT,
            admin_paid_date TEXT,
            comment TEXT NOT NULL DEFAULT '',
            UNIQUE(executor, month)
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """)
        columns = {row["name"] for row in self.connection.execute("PRAGMA table_info(primary_records)")}
        if "yandex_uploaded_date" not in columns:
            self.connection.execute("ALTER TABLE primary_records ADD COLUMN yandex_uploaded_date TEXT")
        self.connection.commit()

    @staticmethod
    def _norm(value: Any) -> str:
        return " ".join(str(value or "").strip().casefold().split())

    @staticmethod
    def _date(value: Any) -> str | None:
        if value in (None, ""):
            return None
        if isinstance(value, (datetime, date)):
            return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
        return str(value).strip()

    def backup_to(self, path: str | Path) -> Path:
        destination = Path(path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        target = sqlite3.connect(destination)
        try:
            self.connection.backup(target)
        finally:
            target.close()
        return destination

    def close(self) -> None:
        self.connection.close()

    def get_setting(self, key: str, default: str = "") -> str:
        row = self.connection.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key: str, value: str) -> None:
        self.connection.execute(
            "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        self.connection.commit()

    def add_record(self, data: dict[str, Any]) -> int:
        fields = (
            "executor", "month", "protocol", "organization", "workplaces",
            "admin_done_date", "assembled_date", "returned_date", "signed_date",
            "scanned_date", "yandex_uploaded_date", "pdf_path", "notes",
        )
        values = [data.get(k) for k in fields]
        values[5:11] = [self._date(v) for v in values[5:11]]
        values[4] = int(values[4] or 0)
        values[12] = values[12] or ""
        cur = self.connection.execute(
            f"INSERT INTO primary_records({','.join(fields)}) VALUES({','.join('?' for _ in fields)})",
            values,
        )
        self.connection.commit()
        return int(cur.lastrowid)

    def update_record(self, record_id: int, data: dict[str, Any]) -> None:
        fields = (
            "executor", "month", "protocol", "organization", "workplaces",
            "admin_done_date", "assembled_date", "returned_date", "signed_date",
            "scanned_date", "yandex_uploaded_date", "pdf_path", "notes",
        )
        values = [data.get(k) for k in fields]
        values[5:11] = [self._date(v) for v in values[5:11]]
        values[4] = int(values[4] or 0)
        values[12] = values[12] or ""
        assignments = ",".join(f"{k}=?" for k in fields)
        self.connection.execute(
            f"UPDATE primary_records SET {assignments},updated_at=CURRENT_TIMESTAMP WHERE id=?",
            values + [record_id],
        )
        self.connection.commit()

    def delete_record(self, record_id: int) -> None:
        self.connection.execute("DELETE FROM primary_records WHERE id=?", (record_id,))
        self.connection.commit()

    def get_record(self, record_id: int):
        return self.connection.execute("SELECT * FROM primary_records WHERE id=?", (record_id,)).fetchone()

    def list_records(self, search: str = "") -> list[sqlite3.Row]:
        where = ""
        params = ()
        if search.strip():
            q = f"%{search.strip()}%"
            where = """WHERE p.organization LIKE ? OR p.protocol LIKE ? OR p.executor LIKE ?
                       OR p.month LIKE ? OR EXISTS (
                           SELECT 1 FROM archive_records ax
                           WHERE ax.protocol_norm=replace(lower(trim(p.protocol)),' ','')
                             AND ax.organization_norm=lower(trim(p.organization))
                             AND ax.archive_index LIKE ?
                       )"""
            params = (q, q, q, q, q)
        return list(self.connection.execute(f"""
            SELECT p.*, (
                SELECT a.archive_index FROM archive_records a
                WHERE a.protocol_norm=replace(lower(trim(p.protocol)),' ','')
                  AND a.organization_norm=lower(trim(p.organization))
                ORDER BY a.id DESC LIMIT 1
            ) AS archive_index
            FROM primary_records p {where} ORDER BY p.id DESC
        """, params))

    def list_executors(self) -> list[str]:
        return [r[0] for r in self.connection.execute(
            "SELECT DISTINCT executor FROM primary_records WHERE executor<>'' ORDER BY executor"
        )]

    def list_months(self) -> list[str]:
        order = ["январь","февраль","март","апрель","май","июнь","июль","август","сентябрь","октябрь","ноябрь","декабрь"]
        found = [r[0] for r in self.connection.execute(
            "SELECT DISTINCT month FROM primary_records WHERE month<>''"
        )]
        return sorted(found, key=lambda x: order.index(x.casefold()) if x.casefold() in order else 99)

    def find_archive(self, protocol: str) -> list[sqlite3.Row]:
        norm = self._norm(protocol).replace(" ", "")
        if not norm:
            return []
        exact = list(self.connection.execute(
            "SELECT * FROM archive_records WHERE replace(protocol_norm,' ','')=? ORDER BY organization,archive_index",
            (norm,),
        ))
        if exact:
            return exact
        return list(self.connection.execute(
            "SELECT * FROM archive_records WHERE replace(protocol_norm,' ','') LIKE ? ORDER BY protocol_norm,organization",
            (norm + "%",),
        ))

    def import_archive_rows(self, rows: list[dict[str, Any]]) -> dict[str, int]:
        added = updated = skipped = 0
        with self.connection:
            for row in rows:
                idx = str(row.get("archive_index") or "").strip()
                org = str(row.get("organization") or "").strip()
                protocol = str(row.get("protocol") or "").strip()
                if not idx or not org or not protocol:
                    skipped += 1
                    continue
                org_norm = self._norm(org)
                protocol_norm = self._norm(protocol).replace(" ", "")
                exists = self.connection.execute(
                    "SELECT id FROM archive_records WHERE archive_index=?",
                    (idx,),
                ).fetchone()
                if exists:
                    self.connection.execute(
                        """UPDATE archive_records SET organization=?,protocol=?,protocol_norm=?,
                           organization_norm=?,source_row=?,imported_at=CURRENT_TIMESTAMP WHERE id=?""",
                        (org, protocol, protocol_norm, org_norm, row.get("source_row"), exists["id"]),
                    )
                    updated += 1
                else:
                    self.connection.execute(
                        """INSERT INTO archive_records
                           (archive_index,organization,protocol,protocol_norm,organization_norm,source_row)
                           VALUES(?,?,?,?,?,?)""",
                        (idx, org, protocol, protocol_norm, org_norm, row.get("source_row")),
                    )
                    added += 1
        return {"added": added, "updated": updated, "skipped": skipped}

    def archive_count(self) -> int:
        return int(self.connection.execute("SELECT COUNT(*) FROM archive_records").fetchone()[0])

    def save_batch(self, data: dict[str, Any]) -> None:
        fields = ("executor","month","primary_given_date","primary_returned_date",
                  "primary_report_date","primary_paid_date","admin_report_date","admin_paid_date","comment")
        values = [data.get(k) for k in fields]
        values[2:8] = [self._date(v) for v in values[2:8]]
        self.connection.execute(
            f"""INSERT INTO batches({','.join(fields)}) VALUES({','.join('?' for _ in fields)})
                ON CONFLICT(executor,month) DO UPDATE SET
                {','.join(f'{k}=excluded.{k}' for k in fields[2:])}""",
            values,
        )
        self.connection.commit()

    def list_batches(self) -> list[sqlite3.Row]:
        return list(self.connection.execute("SELECT * FROM batches ORDER BY executor,month"))

    def import_legacy_workbook(self, path: str | Path) -> dict[str, int]:
        from openpyxl import load_workbook

        wb = load_workbook(path, data_only=True)
        if "2026 Реестр" not in wb.sheetnames:
            raise ValueError("В книге нет листа «2026 Реестр».")
        ws = wb["2026 Реестр"]
        imported = 0
        with self.connection:
            for r in range(2, ws.max_row + 1):
                executor, month = ws.cell(r,2).value, ws.cell(r,3).value
                protocol, organization = ws.cell(r,5).value, ws.cell(r,6).value
                if not protocol and not organization:
                    continue
                if not protocol or not organization:
                    continue
                workplaces = ws.cell(r,7).value or 0
                try:
                    workplaces = int(workplaces)
                except (TypeError, ValueError):
                    workplaces = 0
                data = {
                    "executor": str(executor or ""),
                    "month": str(month or "").casefold(),
                    "protocol": str(protocol).strip(),
                    "organization": str(organization).strip(),
                    "workplaces": workplaces,
                    "admin_done_date": ws.cell(r,10).value,
                }
                self.connection.execute("""
                    INSERT INTO primary_records
                    (executor,month,protocol,organization,workplaces,admin_done_date)
                    SELECT ?,?,?,?,?,?
                    WHERE NOT EXISTS (
                      SELECT 1 FROM primary_records
                      WHERE protocol=? AND organization=? AND executor=? AND month=?
                    )
                """, (data["executor"],data["month"],data["protocol"],data["organization"],
                      workplaces,self._date(data["admin_done_date"]),
                      data["protocol"],data["organization"],data["executor"],data["month"]))
                imported += self.connection.execute("SELECT changes()").fetchone()[0]

        if "2026 Сводка по партиям" in wb.sheetnames:
            sw = wb["2026 Сводка по партиям"]
            for r in range(5, sw.max_row + 1):
                party = sw.cell(r,1).value
                if not party or str(party).startswith("ИТОГО"):
                    continue
                parts = str(party).split(" - ",1)
                if len(parts) != 2:
                    continue
                self.save_batch({
                    "executor": parts[0].strip(), "month": parts[1].strip().casefold(),
                    "primary_given_date": sw.cell(r,3).value,
                    "primary_returned_date": sw.cell(r,4).value,
                    "primary_report_date": sw.cell(r,7).value,
                    "primary_paid_date": sw.cell(r,8).value,
                    "admin_report_date": sw.cell(r,11).value,
                    "admin_paid_date": sw.cell(r,12).value,
                    "comment": sw.cell(r,13).value or "",
                })
        return {"imported": imported}
