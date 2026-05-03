from __future__ import annotations

"""
abbreviation_database.py

Модуль ведения единой базы аббревиатур.

Что делает модуль:
1. Хранит единую базу сокращений в JSON.
2. Обновляет базу на основе existing_abbreviations.csv после этапа 2.
3. Выполняет самоочистку базы от некорректных записей.
4. Экспортирует базу в CSV/XLSX для ручной корректировки.
5. Импортирует ручные правки обратно в JSON-базу.

Основные поля записи:
- record_id
- abbreviation
- long_form
- normalized_long_form
- status
- source_documents
- source_detection_types
- comment
- created_at
- updated_at
"""

from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Any

import json
import pandas as pd


@dataclass
class AbbreviationRecord:
    record_id: str
    abbreviation: str
    long_form: str
    normalized_long_form: str
    status: str
    source_documents: list[str]
    source_detection_types: list[str]
    comment: str
    created_at: str
    updated_at: str


class AbbreviationDatabase:
    """
    Единая база аббревиатур.

    Поддерживает:
    - загрузку и сохранение JSON;
    - обновление из existing_abbreviations.csv;
    - очистку от некорректных записей;
    - экспорт для ручной корректировки;
    - импорт ручных исправлений обратно в JSON.
    """

    def __init__(self, database_path: str | Path) -> None:
        self.database_path = Path(database_path)
        self.records: list[dict[str, Any]] = []
        self.updated_at: str = ""

    @staticmethod
    def _now_str() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def _clean_text(value: Any) -> str:
        text = str(value).strip()
        if not text:
            return ""
        return " ".join(text.split())

    @classmethod
    def _normalize_long_form(cls, value: Any) -> str:
        text = cls._clean_text(value)
        if not text:
            return ""
        return text.lower()

    @staticmethod
    def _normalize_abbreviation_for_record_id(value: Any) -> str:
        text = str(value).strip()
        if not text:
            return ""
        return " ".join(text.split()).upper()

    @classmethod
    def _make_record_id(cls, abbreviation: str, normalized_long_form: str) -> str:
        abbr_part = cls._normalize_abbreviation_for_record_id(abbreviation)
        form_part = normalized_long_form.replace(" ", "_")
        return f"{abbr_part}__{form_part}"

    def _records_map(self) -> dict[str, dict[str, Any]]:
        return {record["record_id"]: record for record in self.records}

    def _sort_records(self) -> None:
        self.records.sort(
            key=lambda item: (
                str(item.get("abbreviation", "")).lower(),
                str(item.get("long_form", "")).lower(),
            )
        )

    def load(self) -> None:
        if not self.database_path.exists():
            self.records = []
            self.updated_at = ""
            return

        with open(self.database_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.records = data.get("records", [])
        self.updated_at = data.get("updated_at", "")

    load_database = load
    read = load

    def save(self) -> None:
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self.updated_at = self._now_str()
        self._sort_records()

        data = {
            "database_path": str(self.database_path).replace("/", "\\"),
            "updated_at": self.updated_at,
            "records_count": len(self.records),
            "records": self.records,
        }

        with open(self.database_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    save_database = save
    write = save

    def to_dataframe(self) -> pd.DataFrame:
        if not self.records:
            return pd.DataFrame(
                columns=[
                    "record_id",
                    "abbreviation",
                    "long_form",
                    "normalized_long_form",
                    "status",
                    "source_documents",
                    "source_detection_types",
                    "comment",
                    "created_at",
                    "updated_at",
                ]
            )

        df = pd.DataFrame(self.records).copy()

        if "source_documents" in df.columns:
            df["source_documents"] = df["source_documents"].apply(
                lambda value: "; ".join(value) if isinstance(value, list) else str(value)
            )
        if "source_detection_types" in df.columns:
            df["source_detection_types"] = df["source_detection_types"].apply(
                lambda value: "; ".join(value) if isinstance(value, list) else str(value)
            )

        return df

    def update_from_existing_abbreviations_csv(self, csv_path: str | Path) -> dict[str, int]:
        csv_path = Path(csv_path)
        if not csv_path.exists():
            raise FileNotFoundError(f"Файл не найден: {csv_path}")

        df = pd.read_csv(csv_path, encoding="utf-8-sig").copy()

        required_columns = {"abbreviation", "long_form"}
        missing = required_columns - set(df.columns)
        if missing:
            raise ValueError(
                "Во входном CSV отсутствуют обязательные столбцы: "
                + ", ".join(sorted(missing))
            )

        df["abbreviation"] = df["abbreviation"].fillna("").astype(str).map(self._clean_text)
        df["long_form"] = df["long_form"].fillna("").astype(str).map(self._clean_text)

        if "detection_type" not in df.columns:
            df["detection_type"] = ""
        df["detection_type"] = df["detection_type"].fillna("").astype(str).map(self._clean_text)

        df["normalized_long_form"] = df["long_form"].map(self._normalize_long_form)

        df = df[
            (df["abbreviation"] != "") &
            (df["long_form"] != "") &
            (df["long_form"].str.lower() != "nan") &
            (df["normalized_long_form"] != "") &
            (df["normalized_long_form"] != "nan")
        ].copy()

        records_map = self._records_map()
        source_document = str(csv_path).replace("/", "\\")
        now_str = self._now_str()

        added = 0
        updated = 0

        for _, row in df.iterrows():
            abbreviation = row["abbreviation"]
            long_form = row["long_form"]
            normalized_long_form = row["normalized_long_form"]
            detection_type = row["detection_type"]

            record_id = self._make_record_id(abbreviation, normalized_long_form)

            if record_id not in records_map:
                new_record = AbbreviationRecord(
                    record_id=record_id,
                    abbreviation=abbreviation,
                    long_form=long_form,
                    normalized_long_form=normalized_long_form,
                    status="active",
                    source_documents=[source_document],
                    source_detection_types=[detection_type] if detection_type else [],
                    comment="",
                    created_at=now_str,
                    updated_at=now_str,
                )
                records_map[record_id] = asdict(new_record)
                added += 1
            else:
                record = records_map[record_id]
                record["abbreviation"] = abbreviation
                record["long_form"] = long_form
                record["normalized_long_form"] = normalized_long_form

                if source_document not in record.get("source_documents", []):
                    record.setdefault("source_documents", []).append(source_document)

                if detection_type and detection_type not in record.get("source_detection_types", []):
                    record.setdefault("source_detection_types", []).append(detection_type)

                if not record.get("status"):
                    record["status"] = "active"
                if "comment" not in record:
                    record["comment"] = ""
                if "created_at" not in record or not record["created_at"]:
                    record["created_at"] = now_str

                record["updated_at"] = now_str
                updated += 1

        self.records = list(records_map.values())
        self._sort_records()

        return {
            "added": added,
            "updated": updated,
            "total_records": len(self.records),
        }

    update_from_existing_abbreviations = update_from_existing_abbreviations_csv
    update_from_csv = update_from_existing_abbreviations_csv
    sync_with_csv = update_from_existing_abbreviations_csv
    merge_from_csv = update_from_existing_abbreviations_csv
    add_from_csv = update_from_existing_abbreviations_csv
    update_database = update_from_existing_abbreviations_csv
    process_csv = update_from_existing_abbreviations_csv
    import_from_csv = update_from_existing_abbreviations_csv

    def cleanup_invalid_records(self) -> dict[str, int]:
        before_count = len(self.records)
        cleaned_records: list[dict[str, Any]] = []

        for record in self.records:
            abbreviation = self._clean_text(record.get("abbreviation", ""))
            long_form = self._clean_text(record.get("long_form", ""))
            normalized_long_form = self._normalize_long_form(
                record.get("normalized_long_form", "") or long_form
            )

            if not abbreviation:
                continue
            if not long_form:
                continue
            if long_form.lower() == "nan":
                continue
            if not normalized_long_form:
                continue
            if normalized_long_form == "nan":
                continue

            record["abbreviation"] = abbreviation
            record["long_form"] = long_form
            record["normalized_long_form"] = normalized_long_form

            if "status" not in record or not record["status"]:
                record["status"] = "active"
            if "comment" not in record or record["comment"] is None:
                record["comment"] = ""

            if not isinstance(record.get("source_documents", []), list):
                value = record.get("source_documents", "")
                record["source_documents"] = [str(value)] if value else []

            if not isinstance(record.get("source_detection_types", []), list):
                value = record.get("source_detection_types", "")
                record["source_detection_types"] = [str(value)] if value else []

            cleaned_records.append(record)

        self.records = cleaned_records
        self._sort_records()

        after_count = len(self.records)
        removed = before_count - after_count

        return {
            "before_cleanup": before_count,
            "after_cleanup": after_count,
            "removed": removed,
        }

    cleanup = cleanup_invalid_records
    clean_invalid_records = cleanup_invalid_records
    remove_invalid_records = cleanup_invalid_records
    self_clean = cleanup_invalid_records

    def export_to_csv_and_xlsx(self, export_dir: str | Path) -> dict[str, str]:
        export_dir = Path(export_dir)
        export_dir.mkdir(parents=True, exist_ok=True)

        df = self.to_dataframe()

        csv_path = export_dir / "abbreviation_database_export.csv"
        xlsx_path = export_dir / "abbreviation_database_export.xlsx"

        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        df.to_excel(xlsx_path, index=False)

        return {
            "database_export_csv": str(csv_path),
            "database_export_xlsx": str(xlsx_path),
        }

    export_for_manual_editing = export_to_csv_and_xlsx
    export_manual_edit_files = export_to_csv_and_xlsx

    def export_to_csv(self, csv_path: str | Path) -> None:
        csv_path = Path(csv_path)
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        self.to_dataframe().to_csv(csv_path, index=False, encoding="utf-8-sig")

    def export_to_xlsx(self, xlsx_path: str | Path) -> None:
        xlsx_path = Path(xlsx_path)
        xlsx_path.parent.mkdir(parents=True, exist_ok=True)
        self.to_dataframe().to_excel(xlsx_path, index=False)

    def import_manual_corrections(self, csv_path: str | Path) -> dict[str, int]:
        csv_path = Path(csv_path)
        if not csv_path.exists():
            raise FileNotFoundError(f"Файл не найден: {csv_path}")

        df = pd.read_csv(csv_path, encoding="utf-8-sig").copy()

        if "record_id" not in df.columns:
            raise ValueError("Для импорта ручных правок в CSV должен присутствовать столбец record_id.")

        records_map = self._records_map()
        updated = 0
        skipped = 0
        now_str = self._now_str()

        for _, row in df.iterrows():
            record_id = self._clean_text(row.get("record_id", ""))
            if not record_id or record_id not in records_map:
                skipped += 1
                continue

            record = records_map[record_id]

            abbreviation = self._clean_text(row.get("abbreviation", record.get("abbreviation", "")))
            long_form = self._clean_text(row.get("long_form", record.get("long_form", "")))
            status = self._clean_text(row.get("status", record.get("status", "active"))) or "active"
            comment = str(row.get("comment", record.get("comment", ""))).strip()

            normalized_long_form = self._normalize_long_form(long_form)
            new_record_id = self._make_record_id(abbreviation, normalized_long_form)

            record["abbreviation"] = abbreviation
            record["long_form"] = long_form
            record["normalized_long_form"] = normalized_long_form
            record["status"] = status
            record["comment"] = comment
            record["updated_at"] = now_str

            if new_record_id != record_id:
                record["record_id"] = new_record_id
                records_map[new_record_id] = record
                del records_map[record_id]

            updated += 1

        self.records = list(records_map.values())
        self.cleanup_invalid_records()
        self.save()

        return {
            "updated": updated,
            "skipped": skipped,
            "total_records": len(self.records),
        }

    def update_from_stage2_and_export(
        self,
        existing_abbreviations_csv: str | Path,
        export_dir: str | Path,
    ) -> dict[str, Any]:
        self.load()
        update_stats = self.update_from_existing_abbreviations_csv(existing_abbreviations_csv)
        cleanup_stats = self.cleanup_invalid_records()
        self.save()
        export_files = self.export_to_csv_and_xlsx(export_dir)

        return {
            "update_stats": update_stats,
            "cleanup_stats": cleanup_stats,
            "database_json": str(self.database_path),
            **export_files,
        }


if __name__ == "__main__":
    db = AbbreviationDatabase("abbreviation_database/abbreviation_database.json")

    result = db.update_from_stage2_and_export(
        existing_abbreviations_csv="result_all/stage2/existing_abbreviations.csv",
        export_dir="result_all/abbreviation_database_export",
    )

    print("=" * 72)
    print("ОБНОВЛЕНИЕ ЕДИНОЙ БАЗЫ АББРЕВИАТУР ЗАВЕРШЕНО")
    print("=" * 72)
    for key, value in result.items():
        print(f"{key}: {value}")
