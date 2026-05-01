from __future__ import annotations

"""
repeated_declaration_replacer.py

Модуль замены повторных объявлений сокращений на сами сокращения.

Что делает модуль:
1. Загружает existing_abbreviations.csv, сформированный на этапе 2.
2. Выделяет объявления сокращений:
   - полная форма (АББР)
   - полная форма (далее – АББР)
   - АББР (полная форма)
3. Сохраняет первое объявление без изменений.
4. Все последующие повторные объявления того же сокращения
   заменяет на само сокращение.
5. Формирует:
   - новый .docx-файл;
   - отчёт о произведённых заменах;
   - таблицу конфликтных сокращений, пропущенных без замены;
   - краткую сводку.

Важно:
- модуль заменяет повторные ОБЪЯВЛЕНИЯ сокращений,
  а не все повторные полные формы термина;
- обработка выполняется для основного текста документа и таблиц;
- содержимое сносок текущей версией не изменяется.
"""

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Set, Tuple

import pandas as pd
import regex
from docx import Document
from docx.document import Document as DocumentObject
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


@dataclass
class ReplacementRule:
    abbreviation: str
    long_form: str
    normalized_long_form: str


@dataclass
class ReplacementEvent:
    abbreviation: str
    long_form: str
    source_type: str
    source_index: int
    original_text: str
    new_text: str
    action: str
    comment: str


class RepeatedDeclarationReplacer:
    """
    Заменяет повторные объявления сокращений на сами сокращения.

    Логика:
    - первое объявление сокращения сохраняется;
    - все последующие повторные объявления того же сокращения
      заменяются на короткую форму;
    - конфликтные сокращения (одна аббревиатура -> несколько полных форм)
      автоматически не заменяются.
    """

    def __init__(self, logger=None) -> None:
        self.logger = logger
        self.declaration_types = {
            "declared_long_first",
            "declared_long_dalee",
            "declared_abbr_first",
        }
        self.section_titles = {
            "обозначения и сокращения",
            "перечень обозначений и сокращений",
        }

    # -----------------------------------------------------------------
    # Загрузка и подготовка данных
    # -----------------------------------------------------------------

    def load_existing_abbreviations(self, csv_path: str | Path) -> pd.DataFrame:
        csv_path = Path(csv_path)
        if not csv_path.exists():
            raise FileNotFoundError(f"Файл не найден: {csv_path}")

        df = pd.read_csv(csv_path, encoding="utf-8-sig")

        required_columns = {
            "abbreviation",
            "long_form",
            "detection_type",
        }
        missing = required_columns - set(df.columns)
        if missing:
            raise ValueError(
                "Во входном CSV отсутствуют обязательные столбцы: "
                + ", ".join(sorted(missing))
            )

        df = df.copy()
        df["abbreviation"] = df["abbreviation"].fillna("").astype(str).str.strip()
        df["long_form"] = df["long_form"].fillna("").astype(str).str.strip()
        df["detection_type"] = df["detection_type"].fillna("").astype(str).str.strip()
        df["normalized_long_form"] = df["long_form"].map(self._normalize_long_form)

        return df

    def build_rules(self, df: pd.DataFrame) -> Tuple[List[ReplacementRule], pd.DataFrame]:
        """
        Формирует правила замены и таблицу конфликтов.

        Конфликтный случай:
        одна и та же аббревиатура объявлена с разными полными формами.
        Для таких сокращений автоматическая замена не выполняется.
        """
        declarations_df = df[
            (df["detection_type"].isin(self.declaration_types)) &
            (df["abbreviation"] != "") &
            (df["long_form"] != "")
        ].copy()

        if declarations_df.empty:
            return [], pd.DataFrame(columns=[
                "abbreviation",
                "normalized_long_form",
                "long_form",
                "comment",
            ])

        unique_pairs = declarations_df[
            ["abbreviation", "long_form", "normalized_long_form"]
        ].copy()
        unique_pairs = unique_pairs.sort_values(
            by=["abbreviation", "normalized_long_form", "long_form"]
        ).drop_duplicates(
            subset=["abbreviation", "normalized_long_form"],
            keep="first"
        ).reset_index(drop=True)

        conflict_rows = []
        conflict_abbreviations: Set[str] = set()

        for abbreviation, group in unique_pairs.groupby("abbreviation"):
            unique_forms = group["normalized_long_form"].dropna().unique().tolist()
            if len(unique_forms) > 1:
                conflict_abbreviations.add(abbreviation)
                for _, row in group.iterrows():
                    conflict_rows.append({
                        "abbreviation": row["abbreviation"],
                        "normalized_long_form": row["normalized_long_form"],
                        "long_form": row["long_form"],
                        "comment": (
                            "Для сокращения обнаружено несколько разных полных форм. "
                            "Автоматическая замена пропущена."
                        ),
                    })

        rules: List[ReplacementRule] = []
        safe_df = unique_pairs[~unique_pairs["abbreviation"].isin(conflict_abbreviations)].copy()

        for _, row in safe_df.iterrows():
            rules.append(
                ReplacementRule(
                    abbreviation=row["abbreviation"],
                    long_form=row["long_form"],
                    normalized_long_form=row["normalized_long_form"],
                )
            )

        # Сначала более длинные полные формы, чтобы избежать частичных перекрытий
        rules.sort(key=lambda item: len(item.long_form), reverse=True)

        if conflict_rows:
            conflict_df = pd.DataFrame(conflict_rows)
        else:
            conflict_df = pd.DataFrame(columns=[
                "abbreviation",
                "normalized_long_form",
                "long_form",
                "comment",
            ])
        return rules, conflict_df

    # -----------------------------------------------------------------
    # Основная логика замены
    # -----------------------------------------------------------------

    def replace_in_document(
        self,
        source_docx_path: str | Path,
        existing_abbreviations_csv: str | Path,
        output_path: str | Path,
    ) -> Dict[str, Path]:
        source_docx_path = Path(source_docx_path)
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        if not source_docx_path.exists():
            raise FileNotFoundError(f"Файл не найден: {source_docx_path}")

        df = self.load_existing_abbreviations(existing_abbreviations_csv)
        rules, conflict_df = self.build_rules(df)

        doc = Document(source_docx_path)

        seen_keys: Set[Tuple[str, str]] = set()
        events: List[ReplacementEvent] = []
        total_replacements = 0
        total_first_kept = 0

        paragraph_counter = 0
        table_cell_counter = 0
        in_glossary_section = False

        for block in self._iter_block_items(doc):
            if isinstance(block, Paragraph):
                source_type = "paragraph"
                source_index = paragraph_counter
                paragraph_counter += 1

                paragraph_text = self._clean_spaces(block.text)
                if self._is_section_title(paragraph_text):
                    in_glossary_section = True
                elif in_glossary_section and self._looks_like_new_section(block, paragraph_text):
                    in_glossary_section = False

                if in_glossary_section:
                    continue

                new_text, block_events, kept_count, replaced_count = self._replace_in_text(
                    text=block.text,
                    rules=rules,
                    seen_keys=seen_keys,
                    source_type=source_type,
                    source_index=source_index,
                )

                if new_text != block.text:
                    block.text = new_text

                events.extend(block_events)
                total_first_kept += kept_count
                total_replacements += replaced_count

            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            source_type = "table_cell"
                            source_index = table_cell_counter
                            table_cell_counter += 1

                            if in_glossary_section:
                                continue

                            new_text, block_events, kept_count, replaced_count = self._replace_in_text(
                                text=paragraph.text,
                                rules=rules,
                                seen_keys=seen_keys,
                                source_type=source_type,
                                source_index=source_index,
                            )

                            if new_text != paragraph.text:
                                paragraph.text = new_text

                            events.extend(block_events)
                            total_first_kept += kept_count
                            total_replacements += replaced_count

        doc.save(output_path)

        report_dir = output_path.parent
        events_df = self.events_to_dataframe(events)
        summary_df = self.build_summary(
            rules=rules,
            events_df=events_df,
            conflict_df=conflict_df,
            output_docx=output_path,
        )

        replacements_csv = report_dir / f"{output_path.stem}_replacement_report.csv"
        summary_csv = report_dir / f"{output_path.stem}_replacement_summary.csv"
        conflict_csv = report_dir / f"{output_path.stem}_replacement_conflicts.csv"

        events_df.to_csv(replacements_csv, index=False, encoding="utf-8-sig")
        summary_df.to_csv(summary_csv, index=False, encoding="utf-8-sig")
        conflict_df.to_csv(conflict_csv, index=False, encoding="utf-8-sig")

        saved_files: Dict[str, Path] = {
            "output_docx": output_path,
            "replacement_report_csv": replacements_csv,
            "replacement_summary_csv": summary_csv,
            "replacement_conflicts_csv": conflict_csv,
        }

        try:
            replacements_xlsx = report_dir / f"{output_path.stem}_replacement_report.xlsx"
            summary_xlsx = report_dir / f"{output_path.stem}_replacement_summary.xlsx"
            conflict_xlsx = report_dir / f"{output_path.stem}_replacement_conflicts.xlsx"

            events_df.to_excel(replacements_xlsx, index=False)
            summary_df.to_excel(summary_xlsx, index=False)
            conflict_df.to_excel(conflict_xlsx, index=False)

            saved_files["replacement_report_xlsx"] = replacements_xlsx
            saved_files["replacement_summary_xlsx"] = summary_xlsx
            saved_files["replacement_conflicts_xlsx"] = conflict_xlsx
        except Exception:
            pass

        if self.logger is not None:
            for event in events:
                if event.action == "replaced_with_abbreviation":
                    try:
                        self.logger.log_replacement(
                            old_text=event.original_text,
                            new_text=event.new_text,
                        )
                    except Exception:
                        pass

            if not conflict_df.empty:
                try:
                    self.logger.warning(
                        "Обнаружены конфликтные сокращения. Они не были заменены автоматически: "
                        + ", ".join(sorted(conflict_df["abbreviation"].dropna().astype(str).unique().tolist()))
                    )
                except Exception:
                    pass

        return saved_files

    def _replace_in_text(
        self,
        text: str,
        rules: List[ReplacementRule],
        seen_keys: Set[Tuple[str, str]],
        source_type: str,
        source_index: int,
    ) -> Tuple[str, List[ReplacementEvent], int, int]:
        if not text or not rules:
            return text, [], 0, 0

        current_text = text
        events: List[ReplacementEvent] = []
        kept_count = 0
        replaced_count = 0

        for rule in rules:
            current_text, rule_events, rule_kept, rule_replaced = self._apply_rule(
                text=current_text,
                rule=rule,
                seen_keys=seen_keys,
                source_type=source_type,
                source_index=source_index,
            )
            events.extend(rule_events)
            kept_count += rule_kept
            replaced_count += rule_replaced

        return current_text, events, kept_count, replaced_count

    def _apply_rule(
        self,
        text: str,
        rule: ReplacementRule,
        seen_keys: Set[Tuple[str, str]],
        source_type: str,
        source_index: int,
    ) -> Tuple[str, List[ReplacementEvent], int, int]:
        key = (rule.abbreviation, rule.normalized_long_form)
        patterns = self._build_rule_patterns(rule)

        current_text = text
        events: List[ReplacementEvent] = []
        kept_count = 0
        replaced_count = 0

        for pattern in patterns:
            def _callback(match) -> str:
                nonlocal kept_count, replaced_count, events, seen_keys
                original = match.group(0)

                if key not in seen_keys:
                    seen_keys.add(key)
                    kept_count += 1
                    events.append(
                        ReplacementEvent(
                            abbreviation=rule.abbreviation,
                            long_form=rule.long_form,
                            source_type=source_type,
                            source_index=source_index,
                            original_text=original,
                            new_text=original,
                            action="kept_first_declaration",
                            comment="Первое объявление сокращения сохранено без изменений.",
                        )
                    )
                    return original

                replaced_count += 1
                events.append(
                    ReplacementEvent(
                        abbreviation=rule.abbreviation,
                        long_form=rule.long_form,
                        source_type=source_type,
                        source_index=source_index,
                        original_text=original,
                        new_text=rule.abbreviation,
                        action="replaced_with_abbreviation",
                        comment="Повторное объявление сокращения заменено на краткую форму.",
                    )
                )
                return rule.abbreviation

            current_text = pattern.sub(_callback, current_text)

        return current_text, events, kept_count, replaced_count

    def _build_rule_patterns(self, rule: ReplacementRule) -> List[regex.Pattern]:
        long_form_pattern = self._build_flexible_phrase_pattern(rule.long_form)
        abbr_pattern = regex.escape(rule.abbreviation)

        return [
            regex.compile(
                rf"(?<!\w){long_form_pattern}\s*\(\s*далее\s*[–—-]\s*{abbr_pattern}\s*\)",
                flags=regex.IGNORECASE,
            ),
            regex.compile(
                rf"(?<!\w){long_form_pattern}\s*\(\s*{abbr_pattern}\s*\)",
                flags=regex.IGNORECASE,
            ),
            regex.compile(
                rf"(?<!\w){abbr_pattern}\s*\(\s*{long_form_pattern}\s*\)",
                flags=regex.IGNORECASE,
            ),
        ]

    # -----------------------------------------------------------------
    # Сохранение и отчёты
    # -----------------------------------------------------------------

    def events_to_dataframe(self, events: List[ReplacementEvent]) -> pd.DataFrame:
        if not events:
            return pd.DataFrame(columns=[
                "abbreviation",
                "long_form",
                "source_type",
                "source_index",
                "original_text",
                "new_text",
                "action",
                "comment",
            ])
        return pd.DataFrame([asdict(item) for item in events])

    def build_summary(
        self,
        rules: List[ReplacementRule],
        events_df: pd.DataFrame,
        conflict_df: pd.DataFrame,
        output_docx: Path,
    ) -> pd.DataFrame:
        if events_df.empty:
            kept_first = 0
            replaced = 0
        else:
            kept_first = int((events_df["action"] == "kept_first_declaration").sum())
            replaced = int((events_df["action"] == "replaced_with_abbreviation").sum())

        summary = {
            "output_docx": str(output_docx),
            "safe_abbreviation_rules": len(rules),
            "conflict_abbreviations_count": int(
                conflict_df["abbreviation"].nunique() if not conflict_df.empty else 0
            ),
            "first_declarations_kept": kept_first,
            "repeated_declarations_replaced": replaced,
            "notes": (
                "Текущая версия заменяет повторные объявления сокращений "
                "в основном тексте и таблицах. Сноски не модифицируются."
            ),
        }
        return pd.DataFrame([summary])

    # -----------------------------------------------------------------
    # Вспомогательные методы
    # -----------------------------------------------------------------

    def _normalize_long_form(self, text: str) -> str:
        text = self._clean_spaces(text).strip(" ,.;:()[]{}\"'«»")
        return text.lower()

    def _clean_spaces(self, text: str) -> str:
        if not text:
            return ""
        return regex.sub(r"\s+", " ", str(text)).strip()

    def _build_flexible_phrase_pattern(self, text: str) -> str:
        parts = [regex.escape(part) for part in self._clean_spaces(text).split()]
        return r"\s+".join(parts)

    def _is_section_title(self, text: str) -> bool:
        return self._clean_spaces(text).lower() in self.section_titles

    def _looks_like_new_section(self, paragraph: Paragraph, text: str) -> bool:
        if not text:
            return False

        style_name = ""
        try:
            if paragraph.style is not None and paragraph.style.name:
                style_name = str(paragraph.style.name).lower()
        except Exception:
            style_name = ""

        if "heading" in style_name or "заголов" in style_name:
            return True

        if regex.match(r"^\d+(\.\d+)*\s+\S+", text):
            return True

        return False

    def _iter_block_items(self, parent) -> Iterator[Paragraph | Table]:
        """
        Итерирует абзацы и таблицы в реальном порядке следования в документе.
        """
        if isinstance(parent, DocumentObject):
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        else:
            raise ValueError("Неподдерживаемый тип контейнера для обхода документа.")

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    # -----------------------------------------------------------------
    # Полный запуск
    # -----------------------------------------------------------------

    def run(
        self,
        source_docx_path: str | Path,
        existing_abbreviations_csv: str | Path,
        output_dir: str | Path,
        output_filename: Optional[str] = None,
    ) -> Dict[str, Path]:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        source_docx_path = Path(source_docx_path)

        if output_filename is None:
            output_filename = f"{source_docx_path.stem}_replaced_declarations.docx"

        output_path = output_dir / output_filename

        return self.replace_in_document(
            source_docx_path=source_docx_path,
            existing_abbreviations_csv=existing_abbreviations_csv,
            output_path=output_path,
        )


if __name__ == "__main__":
    replacer = RepeatedDeclarationReplacer()

    source_docx = "test_reduction_input.docx"
    existing_csv = "result_all/stage2/existing_abbreviations.csv"
    output_dir = "result_all/replacement_stage"

    saved_files = replacer.run(
        source_docx_path=source_docx,
        existing_abbreviations_csv=existing_csv,
        output_dir=output_dir,
    )

    print("=" * 72)
    print("ЗАМЕНА ПОВТОРНЫХ ОБЪЯВЛЕНИЙ ЗАВЕРШЕНА")
    print("=" * 72)
    for name, path in saved_files.items():
        print(f"{name}: {path}")
