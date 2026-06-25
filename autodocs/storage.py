from __future__ import annotations

import json
import sqlite3
import uuid
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass
class TemplateRecord:
    id: str
    name: str
    category: str
    original_filename: str
    stored_path: str
    placeholders: list[str]
    created_at: str
    updated_at: str


@dataclass
class ExportHistoryRecord:
    id: str
    template_id: str
    template_name: str
    export_format: str
    output_paths: list[str]
    values: dict[str, object]
    created_at: str


@dataclass
class CustomerRecord:
    cccd: str
    fields: dict[str, str]
    created_at: str
    updated_at: str

    @property
    def display_name(self) -> str:
        for key in ("ten_khach_hang", "ho_ten", "name", "customer_name", "client_name"):
            value = self.fields.get(key)
            if value:
                return value
        return self.cccd


class TemplateStore:
    def __init__(self, data_dir: str | Path):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.data_dir / "autodocs.sqlite3"
        self._init_db()

    @property
    def templates_dir(self) -> Path:
        path = self.data_dir / "templates"
        path.mkdir(parents=True, exist_ok=True)
        return path

    @contextmanager
    def _connect(self):
        connection = sqlite3.connect(self.db_path)
        try:
            yield connection
            connection.commit()
        finally:
            connection.close()

    def _init_db(self) -> None:
        with self._connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS templates (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    category TEXT NOT NULL DEFAULT '',
                    original_filename TEXT NOT NULL,
                    stored_path TEXT NOT NULL,
                    placeholders_json TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS export_history (
                    id TEXT PRIMARY KEY,
                    template_id TEXT NOT NULL,
                    template_name TEXT NOT NULL,
                    export_format TEXT NOT NULL,
                    output_paths_json TEXT NOT NULL,
                    values_json TEXT NOT NULL,
                    created_at TEXT NOT NULL
                )
                """
            )
            connection.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_export_history_template_created
                ON export_history(template_id, created_at DESC)
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS customers (
                    cccd TEXT PRIMARY KEY,
                    fields_json TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )
            connection.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_customers_updated
                ON customers(updated_at DESC)
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS field_aliases (
                    canonical TEXT NOT NULL,
                    alias TEXT NOT NULL,
                    PRIMARY KEY (canonical, alias)
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY,
                    value_json TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )

    def add_template(
        self,
        name: str,
        category: str,
        original_filename: str,
        stored_path: str | Path,
        placeholders: list[str],
        template_id: str | None = None,
    ) -> TemplateRecord:
        template_id = template_id or uuid.uuid4().hex
        now = datetime.now().isoformat(timespec="seconds")
        record = TemplateRecord(
            id=template_id,
            name=name.strip() or Path(original_filename).stem,
            category=category.strip(),
            original_filename=original_filename,
            stored_path=str(stored_path),
            placeholders=placeholders,
            created_at=now,
            updated_at=now,
        )
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO templates (
                    id, name, category, original_filename, stored_path,
                    placeholders_json, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.id,
                    record.name,
                    record.category,
                    record.original_filename,
                    record.stored_path,
                    json.dumps(record.placeholders, ensure_ascii=True),
                    record.created_at,
                    record.updated_at,
                ),
            )
        return record

    def list_templates(self) -> list[TemplateRecord]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT id, name, category, original_filename, stored_path,
                       placeholders_json, created_at, updated_at
                FROM templates
                ORDER BY updated_at DESC, name COLLATE NOCASE
                """
            ).fetchall()
        return [self._row_to_record(row) for row in rows]

    def get_template(self, template_id: str) -> TemplateRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT id, name, category, original_filename, stored_path,
                       placeholders_json, created_at, updated_at
                FROM templates
                WHERE id = ?
                """,
                (template_id,),
            ).fetchone()
        return self._row_to_record(row) if row else None

    def delete_template(self, template_id: str) -> None:
        record = self.get_template(template_id)
        with self._connect() as connection:
            connection.execute("DELETE FROM templates WHERE id = ?", (template_id,))
            connection.execute("DELETE FROM export_history WHERE template_id = ?", (template_id,))
        if record:
            path = Path(record.stored_path)
            if path.exists() and self.templates_dir in path.resolve().parents:
                path.unlink()

    def add_export_history(
        self,
        template_id: str,
        template_name: str,
        export_format: str,
        output_paths: list[str | Path],
        values: dict[str, object],
    ) -> ExportHistoryRecord:
        record = ExportHistoryRecord(
            id=uuid.uuid4().hex,
            template_id=template_id,
            template_name=template_name,
            export_format=export_format,
            output_paths=[str(path) for path in output_paths],
            values=values,
            created_at=datetime.now().isoformat(timespec="seconds"),
        )
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO export_history (
                    id, template_id, template_name, export_format,
                    output_paths_json, values_json, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.id,
                    record.template_id,
                    record.template_name,
                    record.export_format,
                    json.dumps(record.output_paths, ensure_ascii=False),
                    json.dumps(record.values, ensure_ascii=False),
                    record.created_at,
                ),
            )
        return record

    def list_export_history(self, template_id: str | None = None, limit: int = 300) -> list[ExportHistoryRecord]:
        params: list[object] = []
        where = ""
        if template_id:
            where = "WHERE template_id = ?"
            params.append(template_id)
        params.append(limit)
        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT id, template_id, template_name, export_format,
                       output_paths_json, values_json, created_at
                FROM export_history
                {where}
                ORDER BY created_at DESC
                LIMIT ?
                """,
                params,
            ).fetchall()
        return [self._history_row_to_record(row) for row in rows]

    def clear_export_history(self, template_id: str | None = None) -> None:
        with self._connect() as connection:
            if template_id:
                connection.execute("DELETE FROM export_history WHERE template_id = ?", (template_id,))
            else:
                connection.execute("DELETE FROM export_history")

    def upsert_customer(self, cccd: str, fields: dict[str, object]) -> CustomerRecord:
        cccd = cccd.strip()
        if not cccd:
            raise ValueError("CCCD khong duoc de trong.")

        clean_fields = {
            str(key).strip(): "" if value is None else str(value).strip()
            for key, value in fields.items()
            if str(key).strip()
        }
        clean_fields["cccd"] = cccd
        now = datetime.now().isoformat(timespec="seconds")
        existing = self.get_customer(cccd)
        created_at = existing.created_at if existing else now

        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO customers (cccd, fields_json, created_at, updated_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(cccd) DO UPDATE SET
                    fields_json = excluded.fields_json,
                    updated_at = excluded.updated_at
                """,
                (
                    cccd,
                    json.dumps(clean_fields, ensure_ascii=False),
                    created_at,
                    now,
                ),
            )
        return CustomerRecord(cccd=cccd, fields=clean_fields, created_at=created_at, updated_at=now)

    def get_customer(self, cccd: str) -> CustomerRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT cccd, fields_json, created_at, updated_at
                FROM customers
                WHERE cccd = ?
                """,
                (cccd,),
            ).fetchone()
        return self._customer_row_to_record(row) if row else None

    def list_customers(self, query: str = "") -> list[CustomerRecord]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT cccd, fields_json, created_at, updated_at
                FROM customers
                ORDER BY updated_at DESC
                """
            ).fetchall()
        records = [self._customer_row_to_record(row) for row in rows]
        query = query.strip().casefold()
        if not query:
            return records
        return [
            record
            for record in records
            if query in record.cccd.casefold()
            or query in record.display_name.casefold()
            or query in " ".join(record.fields.values()).casefold()
        ]

    def delete_customer(self, cccd: str) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM customers WHERE cccd = ?", (cccd,))

    def list_field_aliases(self) -> dict[str, list[str]]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT canonical, alias
                FROM field_aliases
                ORDER BY canonical COLLATE NOCASE, alias COLLATE NOCASE
                """
            ).fetchall()
        aliases: dict[str, list[str]] = {}
        for canonical, alias in rows:
            aliases.setdefault(canonical, []).append(alias)
        return aliases

    def add_field_alias(self, canonical: str, alias: str) -> None:
        canonical = canonical.strip()
        alias = alias.strip()
        if not canonical or not alias:
            raise ValueError("Truong chuan va alias khong duoc de trong.")
        with self._connect() as connection:
            connection.execute(
                """
                INSERT OR IGNORE INTO field_aliases (canonical, alias)
                VALUES (?, ?)
                """,
                (canonical, alias),
            )

    def delete_field_alias(self, canonical: str, alias: str) -> None:
        with self._connect() as connection:
            connection.execute(
                "DELETE FROM field_aliases WHERE canonical = ? AND alias = ?",
                (canonical, alias),
            )

    def set_app_setting(self, key: str, value: dict[str, object]) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO app_settings (key, value_json, updated_at)
                VALUES (?, ?, ?)
                ON CONFLICT(key) DO UPDATE SET
                    value_json = excluded.value_json,
                    updated_at = excluded.updated_at
                """,
                (key, json.dumps(value, ensure_ascii=False), now),
            )

    def get_app_setting(self, key: str) -> dict[str, object]:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT value_json FROM app_settings WHERE key = ?",
                (key,),
            ).fetchone()
        return json.loads(row[0]) if row else {}

    @staticmethod
    def _row_to_record(row) -> TemplateRecord:
        return TemplateRecord(
            id=row[0],
            name=row[1],
            category=row[2],
            original_filename=row[3],
            stored_path=row[4],
            placeholders=json.loads(row[5]),
            created_at=row[6],
            updated_at=row[7],
        )

    @staticmethod
    def _history_row_to_record(row) -> ExportHistoryRecord:
        return ExportHistoryRecord(
            id=row[0],
            template_id=row[1],
            template_name=row[2],
            export_format=row[3],
            output_paths=json.loads(row[4]),
            values=json.loads(row[5]),
            created_at=row[6],
        )

    @staticmethod
    def _customer_row_to_record(row) -> CustomerRecord:
        return CustomerRecord(
            cccd=row[0],
            fields=json.loads(row[1]),
            created_at=row[2],
            updated_at=row[3],
        )
