from __future__ import annotations

import sqlite3
import re
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
        self.connection.execute("PRAGMA synchronous = FULL")
        self.initialize()

    def initialize(self) -> None:
        self.connection.executescript("""
        CREATE TABLE IF NOT EXISTS primary_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            executor TEXT NOT NULL DEFAULT '',
            month TEXT NOT NULL DEFAULT '',
            year INTEGER NOT NULL DEFAULT 2026,
            protocol TEXT NOT NULL DEFAULT '',
            organization TEXT NOT NULL DEFAULT '',
            workplaces INTEGER NOT NULL DEFAULT 0,
            admin_amount INTEGER,
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
            year INTEGER NOT NULL DEFAULT 2026,
            primary_given_date TEXT,
            primary_returned_date TEXT,
            primary_report_date TEXT,
            primary_paid_date TEXT,
            admin_report_date TEXT,
            admin_paid_date TEXT,
            comment TEXT NOT NULL DEFAULT '',
            UNIQUE(executor, month, year)
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """)
        columns = {row["name"] for row in self.connection.execute("PRAGMA table_info(primary_records)")}
        if "admin_amount" not in columns:
            self.connection.execute("ALTER TABLE primary_records ADD COLUMN admin_amount INTEGER")
        if "yandex_uploaded_date" not in columns:
            self.connection.execute("ALTER TABLE primary_records ADD COLUMN yandex_uploaded_date TEXT")
        if "year" not in columns:
            self.connection.execute("ALTER TABLE primary_records ADD COLUMN year INTEGER NOT NULL DEFAULT 2026")
        for column in ("primary_rate", "primary_share_rate", "admin_rate"):
            if column not in columns:
                self.connection.execute(f"ALTER TABLE primary_records ADD COLUMN {column} INTEGER")
        batch_columns = {row["name"] for row in self.connection.execute("PRAGMA table_info(batches)")}
        if "year" not in batch_columns:
            self.connection.execute("ALTER TABLE batches RENAME TO batches_old")
            self.connection.execute("""
                CREATE TABLE batches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    executor TEXT NOT NULL,
                    month TEXT NOT NULL,
                    year INTEGER NOT NULL DEFAULT 2026,
                    primary_given_date TEXT,
                    primary_returned_date TEXT,
                    primary_report_date TEXT,
                    primary_paid_date TEXT,
                    admin_report_date TEXT,
                    admin_paid_date TEXT,
                    comment TEXT NOT NULL DEFAULT '',
                    UNIQUE(executor, month, year)
                )
            """)
            self.connection.execute("""
                INSERT INTO batches(id,executor,month,year,primary_given_date,primary_returned_date,
                    primary_report_date,primary_paid_date,admin_report_date,admin_paid_date,comment)
                SELECT id,executor,month,2026,primary_given_date,primary_returned_date,
                    primary_report_date,primary_paid_date,admin_report_date,admin_paid_date,comment
                FROM batches_old
            """)
            self.connection.execute("DROP TABLE batches_old")
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_primary_year ON primary_records(year)")
        defaults = {"primary_rate": "20", "primary_share_rate": "10", "admin_rate": "100"}
        for key, value in defaults.items():
            self.connection.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (key, value))
        current_rates = {key: int(self.get_setting(key, value)) for key, value in defaults.items()}
        self.connection.execute("""
            UPDATE primary_records
            SET primary_rate=COALESCE(primary_rate, ?),
                primary_share_rate=COALESCE(primary_share_rate, ?),
                admin_rate=COALESCE(admin_rate, ?)
            WHERE primary_rate IS NULL OR primary_share_rate IS NULL OR admin_rate IS NULL
        """, (current_rates["primary_rate"], current_rates["primary_share_rate"],
              current_rates["admin_rate"]))
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_archive_protocol_norm ON archive_records(protocol_norm)")
        self.connection.commit()

    @staticmethod
    def _norm(value: Any) -> str:
        return " ".join(str(value or "").strip().casefold().split())

    @staticmethod
    def _org_norm(value: Any) -> str:
        text = str(value or "").casefold().replace("ё", "е")
        return re.sub(r"[^0-9a-zа-я]+", "", text)

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

    def get_rates(self) -> dict[str, int]:
        return {
            "primary_rate": int(self.get_setting("primary_rate", "20")),
            "primary_share_rate": int(self.get_setting("primary_share_rate", "10")),
            "admin_rate": int(self.get_setting("admin_rate", "100")),
        }

    def save_rates(self, rates: dict[str, int]) -> None:
        for key in ("primary_rate", "primary_share_rate", "admin_rate"):
            value = int(rates[key])
            if value < 0:
                raise ValueError("Суммы не могут быть отрицательными.")
            self.set_setting(key, str(value))

    def restore_from(self, path: str | Path) -> None:
        source_path = Path(path).resolve()
        if source_path == self.path.resolve():
            raise ValueError("Нельзя восстановить базу из текущего файла базы.")
        if not source_path.exists():
            raise FileNotFoundError(source_path)
        source = sqlite3.connect(source_path)
        try:
            source.backup(self.connection)
            self.connection.commit()
            self.connection.execute("PRAGMA foreign_keys = ON")
            self.connection.execute("PRAGMA synchronous = FULL")
            self.initialize()
        finally:
            source.close()

    def find_duplicate_protocol(self, protocol: str, exclude_id: int | None = None):
        normalized = self._norm(protocol).replace(" ", "").replace("ё", "е")
        if not normalized:
            return None
        rows = self.connection.execute(
            "SELECT id, protocol, organization, executor, month FROM primary_records ORDER BY id"
        ).fetchall()
        for row in rows:
            if exclude_id is not None and row["id"] == exclude_id:
                continue
            candidate = self._norm(row["protocol"]).replace(" ", "").replace("ё", "е")
            if candidate == normalized:
                return dict(row)
        return None

    def add_record(self, data: dict[str, Any]) -> int:
        duplicate = self.find_duplicate_protocol(data.get("protocol", ""))
        if duplicate:
            raise ValueError(
                f'Протокол {duplicate["protocol"]} уже есть в базе: '
                f'{duplicate["organization"]} (исполнитель: {duplicate["executor"]}, '
                f'партия: {duplicate["month"]}). Проверьте запись перед добавлением.'
            )
        fields = (
            "executor", "month", "year", "protocol", "organization", "workplaces", "admin_amount",
            "admin_done_date", "assembled_date", "returned_date", "signed_date",
            "scanned_date", "yandex_uploaded_date", "pdf_path", "notes",
            "primary_rate", "primary_share_rate", "admin_rate",
        )
        current_rates = self.get_rates()
        values = [data.get(k) for k in fields[:15]]
        values.extend(current_rates[key] for key in ("primary_rate", "primary_share_rate", "admin_rate"))
        values[2] = int(values[2] or date.today().year)
        values[5] = int(values[5] or 0)
        values[6] = int(values[6]) if values[6] not in (None, "") else None
        values[7:13] = [self._date(v) for v in values[7:13]]
        values[14] = values[14] or ""
        cur = self.connection.execute(
            f"INSERT INTO primary_records({','.join(fields)}) VALUES({','.join('?' for _ in fields)})",
            values,
        )
        self.connection.commit()
        self.ensure_batches_for_records()
        return int(cur.lastrowid)

    def update_record(self, record_id: int, data: dict[str, Any]) -> None:
        current = self.get_record(record_id)
        old_protocol = self._norm(current["protocol"]).replace(" ", "") if current else ""
        new_protocol = self._norm(data.get("protocol", "")).replace(" ", "")
        if old_protocol != new_protocol:
            duplicate = self.find_duplicate_protocol(data.get("protocol", ""), exclude_id=record_id)
            if duplicate:
                raise ValueError(
                    f'Протокол {duplicate["protocol"]} уже есть в базе: '
                    f'{duplicate["organization"]} (исполнитель: {duplicate["executor"]}, '
                    f'партия: {duplicate["month"]}).'
                )
        fields = (
            "executor", "month", "year", "protocol", "organization", "workplaces", "admin_amount",
            "admin_done_date", "assembled_date", "returned_date", "signed_date",
            "scanned_date", "yandex_uploaded_date", "pdf_path", "notes",
        )
        values = [data.get(k, current[k] if current else None) for k in fields]
        values[2] = int(values[2] or date.today().year)
        values[5] = int(values[5] or 0)
        values[6] = int(values[6]) if values[6] not in (None, "") else None
        values[7:13] = [self._date(v) for v in values[7:13]]
        values[14] = values[14] or ""
        assignments = ",".join(f"{k}=?" for k in fields)
        self.connection.execute(
            f"UPDATE primary_records SET {assignments},updated_at=CURRENT_TIMESTAMP WHERE id=?",
            values + [record_id],
        )
        self.connection.commit()
        self.ensure_batches_for_records()

    def ensure_batches_for_records(self) -> None:
        self.connection.execute("""
            INSERT OR IGNORE INTO batches(executor, month, year)
            SELECT DISTINCT executor, month, year FROM primary_records
            WHERE executor<>'' AND month<>''
        """)
        self.connection.commit()

    def delete_record(self, record_id: int) -> None:
        self.connection.execute("DELETE FROM primary_records WHERE id=?", (record_id,))
        self.connection.commit()

    def get_record(self, record_id: int):
        return self.connection.execute("SELECT * FROM primary_records WHERE id=?", (record_id,)).fetchone()

    def list_records(self, search: str = "") -> list[dict[str, Any]]:
        records = [
            dict(row) for row in self.connection.execute(
                "SELECT * FROM primary_records ORDER BY id DESC"
            )
        ]
        archive_rows = self.connection.execute("""
            SELECT protocol_norm, MIN(archive_index) AS archive_index
            FROM archive_records GROUP BY protocol_norm
        """).fetchall()
        by_protocol = {
            self._norm(row["protocol_norm"]).replace(" ", "").replace("ё", "е"): row["archive_index"]
            for row in archive_rows
        }
        for record in records:
            protocol_key = self._norm(record["protocol"]).replace(" ", "").replace("ё", "е")
            record["archive_index"] = by_protocol.get(protocol_key)

        query = self._norm(search)
        if query:
            searchable = ("organization", "protocol", "executor", "month", "archive_index", "notes")
            records = [
                record for record in records
                if any(query in self._norm(record.get(field)) for field in searchable)
            ]
        return records

    def mark_scanned(self, protocol: str, organization: str, pdf_path: str, scanned_date: str) -> None:
        protocol_norm = self._norm(protocol).replace(" ", "")
        organization_norm = self._norm(organization)
        matches = self.connection.execute(
            "SELECT id,protocol,organization FROM primary_records ORDER BY id DESC"
        ).fetchall()
        record_id = next(
            (row["id"] for row in matches
             if self._norm(row["protocol"]).replace(" ", "") == protocol_norm
             and self._norm(row["organization"]) == organization_norm),
            None,
        )
        if record_id is not None:
            self.connection.execute(
                "UPDATE primary_records SET scanned_date=?,admin_done_date=COALESCE(admin_done_date,?),pdf_path=?,updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (self._date(scanned_date), self._date(scanned_date), pdf_path, record_id),
            )
            self.connection.commit()

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
        fields = ("executor","month","year","primary_given_date","primary_returned_date",
                  "primary_report_date","primary_paid_date","admin_report_date","admin_paid_date","comment")
        values = [data.get(k) for k in fields]
        values[2] = int(values[2] or date.today().year)
        values[3:9] = [self._date(v) for v in values[3:9]]
        self.connection.execute(
            f"""INSERT INTO batches({','.join(fields)}) VALUES({','.join('?' for _ in fields)})
                ON CONFLICT(executor,month,year) DO UPDATE SET
                {','.join(f'{k}=excluded.{k}' for k in fields[3:])}""",
            values,
        )
        self.connection.commit()

    def list_batches(self) -> list[sqlite3.Row]:
        return list(self.connection.execute("SELECT * FROM batches ORDER BY year DESC, id"))

    def update_status(self, record_id: int, field: str, value: str | None) -> None:
        allowed = {"assembled_date", "scanned_date", "yandex_uploaded_date"}
        if field not in allowed:
            raise ValueError("Недопустимое поле статуса.")
        stamp = self._date(value)
        with self.connection:
            if field == "scanned_date":
                self.connection.execute("""
                    UPDATE primary_records
                    SET scanned_date=?, admin_done_date=?, updated_at=CURRENT_TIMESTAMP
                    WHERE id=?
                """, (stamp, stamp, record_id))
            else:
                self.connection.execute(
                    f"UPDATE primary_records SET {field}=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                    (stamp, record_id),
                )

    def import_legacy_workbook(self, path: str | Path) -> dict[str, int]:
        """Import only the five source fields. All statuses, fees and batch controls are app-owned."""
        from openpyxl import load_workbook
        import re

        wb = load_workbook(path, data_only=True, read_only=True)
        registry_name = next((name for name in wb.sheetnames if name.endswith(" Реестр") and name.split()[0].isdigit()), None)
        if not registry_name:
            wb.close()
            raise ValueError("В книге не найден лист вида «2026 Реестр».")
        year = int(registry_name.split()[0])
        ws = wb[registry_name]
        imported = skipped = 0
        with self.connection:
            for row in ws.iter_rows(min_row=2, values_only=True):
                # Source workbook columns: B executor, C month, E protocol, F organization, G workplaces.
                if len(row) < 7:
                    continue
                executor, month, protocol, organization, workplaces = row[1], row[2], row[4], row[5], row[6]
                if not protocol or not organization:
                    continue
                protocol = str(protocol).strip()
                organization = str(organization).strip()
                executor = str(executor or "").strip()
                month = str(month or "").strip().casefold()
                try:
                    workplaces = int(float(workplaces or 0))
                except (TypeError, ValueError):
                    workplaces = 0
                normalized = self._norm(protocol).replace(" ", "").replace("ё", "е")
                if self.find_duplicate_protocol(protocol):
                    skipped += 1
                    continue
                rates = self.get_rates()
                self.connection.execute("""
                    INSERT INTO primary_records
                    (executor,month,year,protocol,organization,workplaces,
                     primary_rate,primary_share_rate,admin_rate)
                    VALUES(?,?,?,?,?,?,?,?,?)
                """, (executor, month, year, protocol, organization, workplaces,
                      rates["primary_rate"], rates["primary_share_rate"], rates["admin_rate"]))
                imported += 1
        wb.close()
        self.ensure_batches_for_records()
        return {"imported": imported, "skipped": skipped, "statuses_updated": 0}
