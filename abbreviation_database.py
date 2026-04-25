from __future__ import annotations

"""
abbreviation_database.py

Задача 1:
единая база аббревиатур с накоплением и возможностью ручной корректировки.

Исправления этой версии:
1. Не добавляет в базу записи с пустой полной формой.
2. Не добавляет записи, где long_form превращается в "nan".
3. Корректно обрабатывает NaN из pandas при импорте CSV/XLSX.
4. Отбрасывает standalone-записи без полной формы.
"""

from dataclasses import dataclass, asdict, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
import json

import pandas as pd
import regex


@dataclass
class AbbreviationRecord:
    """
    Одна запись в единой базе аббревиатур.
    """
    record_id: str
    abbreviation: str
    long_form: str
    normalized_long_form: str
    status: str = "active"
    source_documents: List[str] = field(default_factory=list)
    source_detection_types: List[str] = field(default_factory=list)
    comment: str = ""
    created_at: str = ""
    updated_at: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


class AbbreviationDatabase:
    """
    Единая база аббревиатур.

    Основное хранилище — JSON-файл.
    Дополнительно база может экспортироваться в CSV/XLSX для ручной правки.
    """

    def __init__(self, db_path: str | Path = "abbreviation_database.json") -> None:
        self.db_path = Path(db_path)
        self.records: List[AbbreviationRecord] = []

    # -----------------------------------------------------------------
    # Вспомогательные методы
    # -----------------------------------------------------------------

    def _now(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _normalize_whitespace(self, text: str) -> str:
        return regex.sub(r"\s+", " ", str(text)).strip()

    def _is_empty_like(self, value) -> bool:
        """
        Проверяет, является ли значение пустым / NaN-подобным.
        """
        if value is None:
            return True

        try:
            if pd.isna(value):
                return True
        except Exception:
            pass

        text = str(value).strip()
        if not text:
            return True

        lowered = text.lower()
        return lowered in {"nan", "none", "null"}

    def _safe_text(self, value) -> str:
        """
        Аккуратно преобразует значение в строку.
        Пустые / NaN-подобные значения превращаются в пустую строку.
        """
        if self._is_empty_like(value):
            return ""
        return self._normalize_whitespace(value)

    def normalize_abbreviation(self, abbreviation: str) -> str:
        """
        Нормализует сокращение:
        - убирает лишние пробелы,
        - сохраняет регистр как есть,
        - для поиска чаще используется upper().
        """
        return self._normalize_whitespace(abbreviation)

    def normalize_long_form(self, long_form: str) -> str:
        """
        Нормализует полную форму для сравнения и поиска.
        """
        text = self._normalize_whitespace(long_form).lower()
        text = text.strip(" ,.;:()[]{}\"'«»")
        return text

    def build_record_id(self, abbreviation: str, long_form: str) -> str:
        """
        Формирует устойчивый идентификатор записи.
        """
        abbr = self.normalize_abbreviation(abbreviation).upper()
        norm_long = self.normalize_long_form(long_form)
        safe_long = regex.sub(r"[^a-zа-яё0-9]+", "_", norm_long, flags=regex.IGNORECASE).strip("_")
        return f"{abbr}__{safe_long}"

    def _find_record_index(self, abbreviation: str, long_form: str) -> Optional[int]:
        """
        Ищет точное совпадение по аббревиатуре и полной форме.
        """
        normalized_abbreviation = self.normalize_abbreviation(abbreviation).upper()
        normalized_long_form = self.normalize_long_form(long_form)

        for idx, record in enumerate(self.records):
            if (
                self.normalize_abbreviation(record.abbreviation).upper() == normalized_abbreviation
                and record.normalized_long_form == normalized_long_form
            ):
                return idx
        return None

    # -----------------------------------------------------------------
    # Загрузка / сохранение
    # -----------------------------------------------------------------

    def load(self) -> None:
        """
        Загружает JSON-базу.
        Если файла нет, создаётся пустая база в памяти.
        """
        if not self.db_path.exists():
            self.records = []
            return

        with self.db_path.open("r", encoding="utf-8") as f:
            raw = json.load(f)

        records_raw = raw.get("records", [])
        self.records = [AbbreviationRecord(**item) for item in records_raw]

    def save(self) -> None:
        """
        Сохраняет базу в JSON.
        """
        self.db_path.parent.mkdir(parents=True, exist_ok=True)

        payload = {
            "database_path": str(self.db_path),
            "updated_at": self._now(),
            "records_count": len(self.records),
            "records": [record.to_dict() for record in self.records],
        }

        with self.db_path.open("w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

    # -----------------------------------------------------------------
    # Просмотр базы
    # -----------------------------------------------------------------

    def to_dataframe(self) -> pd.DataFrame:
        """
        Преобразует базу в DataFrame.
        """
        if not self.records:
            return pd.DataFrame(columns=[
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
            ])

        rows = []
        for record in self.records:
            row = record.to_dict()
            row["source_documents"] = " || ".join(record.source_documents)
            row["source_detection_types"] = " || ".join(record.source_detection_types)
            rows.append(row)

        return pd.DataFrame(rows)

    def find_by_abbreviation(self, abbreviation: str) -> pd.DataFrame:
        """
        Ищет записи по аббревиатуре.
        """
        df = self.to_dataframe()
        if df.empty:
            return df

        target = self.normalize_abbreviation(abbreviation).upper()
        return df[df["abbreviation"].astype(str).str.upper() == target].reset_index(drop=True)

    def find_by_long_form(self, long_form: str) -> pd.DataFrame:
        """
        Ищет записи по полной форме.
        """
        df = self.to_dataframe()
        if df.empty:
            return df

        target = self.normalize_long_form(long_form)
        return df[df["normalized_long_form"].astype(str) == target].reset_index(drop=True)

    # -----------------------------------------------------------------
    # Добавление / обновление
    # -----------------------------------------------------------------

    def add_or_update_record(
        self,
        abbreviation: str,
        long_form: str,
        source_document: str = "",
        detection_type: str = "",
        status: str = "active",
        comment: str = "",
        force_update_comment: bool = False
    ) -> str:
        """
        Добавляет новую запись или обновляет существующую.

        Возвращает:
        - "added"   если запись добавлена,
        - "updated" если запись уже была и обновлена,
        - "skipped" если запись пропущена.
        """
        abbreviation = self._safe_text(abbreviation)
        long_form = self._safe_text(long_form)
        detection_type = self._safe_text(detection_type)
        source_document = self._safe_text(source_document)
        comment = self._safe_text(comment)

        if not abbreviation or not long_form:
            return "skipped"

        normalized_long_form = self.normalize_long_form(long_form)
        if not normalized_long_form:
            return "skipped"

        record_id = self.build_record_id(abbreviation, long_form)
        existing_index = self._find_record_index(abbreviation, long_form)

        if existing_index is None:
            created_at = self._now()
            updated_at = created_at

            record = AbbreviationRecord(
                record_id=record_id,
                abbreviation=abbreviation,
                long_form=long_form,
                normalized_long_form=normalized_long_form,
                status=status,
                source_documents=[source_document] if source_document else [],
                source_detection_types=[detection_type] if detection_type else [],
                comment=comment,
                created_at=created_at,
                updated_at=updated_at,
            )
            self.records.append(record)
            return "added"

        record = self.records[existing_index]

        if source_document and source_document not in record.source_documents:
            record.source_documents.append(source_document)

        if detection_type and detection_type not in record.source_detection_types:
            record.source_detection_types.append(detection_type)

        if status:
            record.status = status

        if force_update_comment:
            record.comment = comment
        elif comment and not record.comment:
            record.comment = comment

        record.updated_at = self._now()
        return "updated"

    def update_record_manually(
        self,
        record_id: str,
        abbreviation: Optional[str] = None,
        long_form: Optional[str] = None,
        status: Optional[str] = None,
        comment: Optional[str] = None
    ) -> bool:
        """
        Ручное обновление записи внутри программы.
        """
        for idx, record in enumerate(self.records):
            if record.record_id != record_id:
                continue

            if abbreviation is not None and not self._is_empty_like(abbreviation):
                record.abbreviation = self.normalize_abbreviation(abbreviation)

            if long_form is not None and not self._is_empty_like(long_form):
                record.long_form = self._safe_text(long_form)
                record.normalized_long_form = self.normalize_long_form(long_form)

            if status is not None and not self._is_empty_like(status):
                record.status = self._safe_text(status)

            if comment is not None:
                record.comment = self._safe_text(comment)

            record.record_id = self.build_record_id(record.abbreviation, record.long_form)
            record.updated_at = self._now()

            if record.status == "active":
                record.status = "edited"

            self.records[idx] = record
            return True

        return False

    # -----------------------------------------------------------------
    # Интеграция с результатами этапа 2
    # -----------------------------------------------------------------

    def import_from_existing_abbreviations_csv(
        self,
        csv_path: str | Path,
        source_document: str = ""
    ) -> Dict[str, int]:
        """
        Импортирует сокращения из existing_abbreviations.csv.

        Загружаются только строки, где есть:
        - abbreviation
        - long_form

        Пустые / NaN long_form пропускаются.
        Также пропускаются standalone-записи без полной формы.
        """
        csv_path = Path(csv_path)
        if not csv_path.exists():
            raise FileNotFoundError(f"Файл не найден: {csv_path}")

        df = pd.read_csv(csv_path, encoding="utf-8-sig")

        required_columns = {"abbreviation", "long_form", "detection_type"}
        missing = required_columns - set(df.columns)
        if missing:
            raise ValueError(
                "Во входном CSV отсутствуют обязательные столбцы: "
                + ", ".join(sorted(missing))
            )

        added = 0
        updated = 0
        skipped = 0

        for _, row in df.iterrows():
            abbreviation = self._safe_text(row.get("abbreviation", ""))
            long_form = self._safe_text(row.get("long_form", ""))
            detection_type = self._safe_text(row.get("detection_type", ""))

            # Пропускаем standalone без полной формы
            if detection_type == "standalone" and not long_form:
                skipped += 1
                continue

            # Пропускаем любые записи без нормальной полной формы
            if not abbreviation or not long_form:
                skipped += 1
                continue

            normalized_long_form = self.normalize_long_form(long_form)
            if not normalized_long_form:
                skipped += 1
                continue

            result = self.add_or_update_record(
                abbreviation=abbreviation,
                long_form=long_form,
                source_document=source_document,
                detection_type=detection_type,
                status="active",
            )

            if result == "added":
                added += 1
            elif result == "updated":
                updated += 1
            else:
                skipped += 1

        return {
            "added": added,
            "updated": updated,
            "skipped": skipped,
        }

    # -----------------------------------------------------------------
    # Экспорт / импорт для ручной корректировки
    # -----------------------------------------------------------------

    def export_for_manual_edit(self, output_dir: str | Path = "abbreviation_database_export") -> Dict[str, Path]:
        """
        Экспортирует базу в CSV/XLSX, чтобы пользователь мог вручную её править.
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        df = self.to_dataframe()

        csv_path = output_dir / "abbreviation_database_export.csv"
        xlsx_path = output_dir / "abbreviation_database_export.xlsx"

        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

        saved_files: Dict[str, Path] = {
            "csv": csv_path,
        }

        try:
            df.to_excel(xlsx_path, index=False)
            saved_files["xlsx"] = xlsx_path
        except ModuleNotFoundError:
            pass

        return saved_files

    def import_manual_corrections(self, edited_file_path: str | Path) -> Dict[str, int]:
        """
        Импортирует ручные исправления из CSV/XLSX,
        которые были отредактированы пользователем.

        Ожидаемые столбцы:
        - record_id
        - abbreviation
        - long_form
        - status
        - comment
        """
        edited_file_path = Path(edited_file_path)
        if not edited_file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {edited_file_path}")

        suffix = edited_file_path.suffix.lower()
        if suffix == ".csv":
            df = pd.read_csv(edited_file_path, encoding="utf-8-sig")
        elif suffix in {".xlsx", ".xls"}:
            df = pd.read_excel(edited_file_path)
        else:
            raise ValueError("Поддерживаются только CSV/XLSX файлы.")

        required_columns = {"record_id", "abbreviation", "long_form", "status", "comment"}
        missing = required_columns - set(df.columns)
        if missing:
            raise ValueError(
                "В файле ручной корректировки отсутствуют обязательные столбцы: "
                + ", ".join(sorted(missing))
            )

        updated = 0
        not_found = 0

        for _, row in df.iterrows():
            record_id = self._safe_text(row.get("record_id", ""))
            abbreviation = row.get("abbreviation", "")
            long_form = row.get("long_form", "")
            status = row.get("status", "")
            comment = row.get("comment", "")

            if not record_id:
                not_found += 1
                continue

            ok = self.update_record_manually(
                record_id=record_id,
                abbreviation=abbreviation,
                long_form=long_form,
                status=status if not self._is_empty_like(status) else None,
                comment=comment
            )
            if ok:
                updated += 1
            else:
                not_found += 1

        return {
            "updated": updated,
            "not_found": not_found,
        }

    # -----------------------------------------------------------------
    # Сводка
    # -----------------------------------------------------------------

    def build_summary(self) -> pd.DataFrame:
        """
        Формирует краткую сводку по базе.
        """
        df = self.to_dataframe()
        if df.empty:
            return pd.DataFrame([{
                "records_total": 0,
                "active_records": 0,
                "edited_records": 0,
                "deprecated_records": 0,
                "unique_abbreviations": 0,
            }])

        summary = {
            "records_total": len(df),
            "active_records": int((df["status"] == "active").sum()),
            "edited_records": int((df["status"] == "edited").sum()),
            "deprecated_records": int((df["status"] == "deprecated").sum()),
            "unique_abbreviations": int(df["abbreviation"].nunique()),
        }
        return pd.DataFrame([summary])


if __name__ == "__main__":
    """
    Пример автономного запуска:

    1. Если рядом есть result_stage2/existing_abbreviations.csv,
       база пополнится из него.
    2. JSON-файл базы будет сохранён как abbreviation_database.json
    3. Также будет экспортирована версия для ручной корректировки.
    """
    db = AbbreviationDatabase("abbreviation_database.json")
    db.load()

    input_csv = Path("result_stage2/existing_abbreviations.csv")
    if input_csv.exists():
        stats = db.import_from_existing_abbreviations_csv(
            csv_path=input_csv,
            source_document="result_stage2/existing_abbreviations.csv"
        )
        print("Импорт из existing_abbreviations.csv:")
        print(stats)

    db.save()
    print(f"JSON-база сохранена: {db.db_path}")

    exported = db.export_for_manual_edit("abbreviation_database_export")
    print("Файлы для ручной корректировки:")
    for name, path in exported.items():
        print(f"{name}: {path}")

    print("\nСводка по базе:")
    print(db.build_summary().to_string(index=False))
