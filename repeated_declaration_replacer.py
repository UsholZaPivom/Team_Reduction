from __future__ import annotations

"""
repeated_declaration_replacer.py

Точечная доработка v3:
1. Добавлено распознавание строк глоссария:
   "АББР – полная форма"
   "полная форма – АББР"
2. Улучшен подбор длинной части перед "(далее – АББР)":
   при одинаковом покрытии выбирается наиболее короткий и точный хвост,
   чтобы не захватывать соседние объявления.
3. Сохраняется расчёт статистики по повторным употреблениям сокращения
   после первого объявления.
"""

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any
from collections import defaultdict
import re

import pandas as pd
from docx import Document


@dataclass
class SafeRule:
    abbreviation: str
    long_form: str
    normalized_long_form: str
    detection_types: list[str]
    match_count_in_csv: int


@dataclass
class ReplacementEvent:
    abbreviation: str
    long_form: str
    normalized_long_form: str
    source_type: str
    source_index: str
    matched_text: str
    replacement_text: str
    action: str
    occurrence_number_global: int
    comment: str


class RepeatedDeclarationReplacer:
    def __init__(self) -> None:
        pass

    # ------------------------------------------------------------------
    # Базовые утилиты
    # ------------------------------------------------------------------

    @staticmethod
    def _clean_text(value: Any) -> str:
        text = str(value).strip()
        if not text:
            return ""
        return " ".join(text.split())

    @classmethod
    def _normalize_long_form(cls, value: Any) -> str:
        return cls._clean_text(value).lower()

    @staticmethod
    def _make_flexible_whitespace_pattern(text: str) -> str:
        escaped_parts = [re.escape(part) for part in text.split()]
        return r"\s+".join(escaped_parts)

    @staticmethod
    def _spans_overlap(span_a: tuple[int, int], span_b: tuple[int, int]) -> bool:
        return max(span_a[0], span_b[0]) < min(span_a[1], span_b[1])

    @staticmethod
    def _word_tokens(text: str) -> list[str]:
        return re.findall(r"[A-Za-zА-Яа-яЁё0-9]+", text.lower())

    @staticmethod
    def _soft_stem(token: str) -> str:
        token = token.lower()
        if len(token) >= 10:
            return token[:7]
        if len(token) >= 8:
            return token[:6]
        if len(token) >= 6:
            return token[:5]
        if len(token) >= 4:
            return token[:4]
        return token

    @classmethod
    def _meaningful_stems(cls, text: str) -> list[str]:
        stopwords = {
            "и", "или", "по", "на", "в", "во", "с", "со", "к", "ко", "о", "об",
            "от", "до", "из", "за", "над", "под", "при", "для", "не", "как",
            "это", "тот", "та", "те", "данный", "данная", "данное", "данные",
            "далее", "the", "for", "and", "or", "of", "to", "in", "on"
        }
        result: list[str] = []
        for token in cls._word_tokens(text):
            if len(token) > 2 and token not in stopwords:
                result.append(cls._soft_stem(token))
        return result

    @classmethod
    def _token_overlap_score(cls, candidate_long_form: str, rule_long_form: str) -> tuple[float, int]:
        candidate_stems = set(cls._meaningful_stems(candidate_long_form))
        rule_stems = set(cls._meaningful_stems(rule_long_form))

        if not candidate_stems or not rule_stems:
            return 0.0, 0

        overlap = candidate_stems & rule_stems
        overlap_count = len(overlap)
        coverage = overlap_count / max(len(rule_stems), 1)
        return coverage, overlap_count

    @classmethod
    def _looks_like_same_long_form(cls, candidate_long_form: str, rule_long_form: str) -> bool:
        coverage, overlap_count = cls._token_overlap_score(candidate_long_form, rule_long_form)
        rule_token_count = len(set(cls._meaningful_stems(rule_long_form)))

        if rule_token_count <= 2:
            return overlap_count == rule_token_count and overlap_count >= 2

        if rule_token_count == 3:
            return overlap_count >= 2 and coverage >= 0.66

        return overlap_count >= 2 and coverage >= 0.6

    # ------------------------------------------------------------------
    # Шаблоны поиска
    # ------------------------------------------------------------------

    @classmethod
    def _build_strict_declaration_regex(cls, long_form: str, abbreviation: str) -> re.Pattern:
        long_form_pattern = cls._make_flexible_whitespace_pattern(long_form)
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)

        pattern = (
            rf"(?P<matched>"
            rf"(?P<long>{long_form_pattern})"
            rf"\s*"
            rf"\("
            rf"\s*"
            rf"(?:далее\s*[–—-]\s*)?"
            rf"(?P<abbr>{abbr_pattern})"
            rf"\s*"
            rf"\)"
            rf")"
        )
        return re.compile(pattern, flags=re.IGNORECASE)

    @classmethod
    def _build_parenthetical_abbreviation_regex(cls, abbreviation: str) -> re.Pattern:
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        pattern = (
            rf"\("
            rf"\s*"
            rf"(?:далее\s*[–—-]\s*)?"
            rf"(?P<abbr>{abbr_pattern})"
            rf"\s*"
            rf"\)"
        )
        return re.compile(pattern, flags=re.IGNORECASE)

    @classmethod
    def _build_abbreviation_regex(cls, abbreviation: str) -> re.Pattern:
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        pattern = rf"(?<!\w){abbr_pattern}(?!\w)"
        return re.compile(pattern, flags=re.IGNORECASE)

    @classmethod
    def _build_glossary_line_regex_abbr_first(cls, abbreviation: str) -> re.Pattern:
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        pattern = (
            rf"^\s*"
            rf"(?P<abbr>{abbr_pattern})"
            rf"\s*[–—-]\s*"
            rf"(?P<long>.+?)"
            rf"\s*$"
        )
        return re.compile(pattern, flags=re.IGNORECASE)

    @classmethod
    def _build_glossary_line_regex_long_first(cls, abbreviation: str) -> re.Pattern:
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        pattern = (
            rf"^\s*"
            rf"(?P<long>.+?)"
            rf"\s*[–—-]\s*"
            rf"(?P<abbr>{abbr_pattern})"
            rf"\s*$"
        )
        return re.compile(pattern, flags=re.IGNORECASE)

    # ------------------------------------------------------------------
    # Извлечение длинной части перед скобками
    # ------------------------------------------------------------------

    @staticmethod
    def _find_left_boundary(text: str, left_end: int) -> int:
        search_zone = text[max(0, left_end - 260):left_end]
        relative_boundary = 0
        for sep_match in re.finditer(r"[\n\r.;:!?]", search_zone):
            relative_boundary = sep_match.end()
        return max(0, left_end - len(search_zone) + relative_boundary)

    @classmethod
    def _extract_candidate_long_form_before_parentheses(
        cls,
        text: str,
        paren_start: int,
        rule_long_form: str,
    ) -> tuple[str, int]:
        left_boundary = cls._find_left_boundary(text, paren_start)
        fragment = text[left_boundary:paren_start].strip()
        fragment = re.sub(r"\s+", " ", fragment)

        if not fragment:
            return "", paren_start

        token_matches = list(re.finditer(r"[A-Za-zА-Яа-яЁё0-9]+", fragment))
        if not token_matches:
            return "", paren_start

        best_candidate = ""
        best_start_in_fragment = len(fragment)
        # Чем выше coverage/overlap, тем лучше.
        # При равенстве берём более короткий кандидат, чтобы не захватывать
        # лишний контекст и соседние объявления.
        best_score = (-1.0, -1, 10**9)  # coverage, overlap_count, token_count(min)

        max_tail_tokens = min(len(token_matches), 18)

        for tail_size in range(2, max_tail_tokens + 1):
            tail_tokens = token_matches[-tail_size:]
            start_in_fragment = tail_tokens[0].start()
            candidate = fragment[start_in_fragment:].strip(" ,;:«»\"'„“”’()-")
            if not candidate:
                continue

            coverage, overlap_count = cls._token_overlap_score(candidate, rule_long_form)
            if overlap_count == 0:
                continue

            token_count = len(cls._word_tokens(candidate))
            score = (coverage, overlap_count, -token_count)

            if score > best_score and cls._looks_like_same_long_form(candidate, rule_long_form):
                best_score = score
                best_candidate = candidate
                best_start_in_fragment = start_in_fragment

        if not best_candidate:
            return "", paren_start

        absolute_start = left_boundary + best_start_in_fragment
        return best_candidate, absolute_start

    # ------------------------------------------------------------------
    # Гибкий поиск объявлений
    # ------------------------------------------------------------------

    def _find_declaration_matches(
        self,
        text: str,
        rule: SafeRule,
    ) -> list[dict[str, Any]]:
        matches: list[dict[str, Any]] = []
        used_spans: list[tuple[int, int]] = []
        clean_text = self._clean_text(text)

        # 1. Строка глоссария: "АББР – полная форма"
        glossary_abbr_first = self._build_glossary_line_regex_abbr_first(rule.abbreviation)
        m = glossary_abbr_first.match(clean_text)
        if m:
            candidate_long = self._clean_text(m.group("long"))
            if self._looks_like_same_long_form(candidate_long, rule.long_form):
                matches.append(
                    {
                        "start": 0,
                        "end": len(text),
                        "matched_text": clean_text,
                        "long_text": candidate_long,
                        "match_type": "glossary_abbr_first",
                    }
                )
                used_spans.append((0, len(text)))

        # 2. Строка глоссария: "полная форма – АББР"
        glossary_long_first = self._build_glossary_line_regex_long_first(rule.abbreviation)
        m = glossary_long_first.match(clean_text)
        if m:
            candidate_long = self._clean_text(m.group("long"))
            if self._looks_like_same_long_form(candidate_long, rule.long_form):
                span = (0, len(text))
                if not any(self._spans_overlap(span, used_span) for used_span in used_spans):
                    matches.append(
                        {
                            "start": 0,
                            "end": len(text),
                            "matched_text": clean_text,
                            "long_text": candidate_long,
                            "match_type": "glossary_long_first",
                        }
                    )
                    used_spans.append(span)

        # 3. Строгое объявление "полная форма (далее – АББР)"
        strict_regex = self._build_strict_declaration_regex(rule.long_form, rule.abbreviation)
        for match in strict_regex.finditer(text):
            span = match.span("matched")
            if any(self._spans_overlap(span, used_span) for used_span in used_spans):
                continue
            matches.append(
                {
                    "start": span[0],
                    "end": span[1],
                    "matched_text": match.group("matched"),
                    "long_text": self._clean_text(match.group("long")),
                    "match_type": "strict",
                }
            )
            used_spans.append(span)

        # 4. Гибкое объявление по конструкции "(далее – АББР)"
        paren_regex = self._build_parenthetical_abbreviation_regex(rule.abbreviation)
        for paren_match in paren_regex.finditer(text):
            paren_span = paren_match.span()
            if any(self._spans_overlap(paren_span, used_span) for used_span in used_spans):
                continue

            candidate_long, absolute_start = self._extract_candidate_long_form_before_parentheses(
                text=text,
                paren_start=paren_span[0],
                rule_long_form=rule.long_form,
            )
            if not candidate_long:
                continue

            full_start = absolute_start
            full_end = paren_span[1]
            full_span = (full_start, full_end)

            if any(self._spans_overlap(full_span, used_span) for used_span in used_spans):
                continue

            matched_text = self._clean_text(text[full_start:full_end])

            matches.append(
                {
                    "start": full_start,
                    "end": full_end,
                    "matched_text": matched_text,
                    "long_text": candidate_long,
                    "match_type": "generic_token_match",
                }
            )
            used_spans.append(full_span)

        matches.sort(key=lambda item: item["start"])
        return matches

    # ------------------------------------------------------------------
    # Загрузка правил из CSV
    # ------------------------------------------------------------------

    def _load_rules_from_existing_abbreviations(
        self,
        existing_abbreviations_csv: str | Path,
    ) -> tuple[list[SafeRule], int]:
        csv_path = Path(existing_abbreviations_csv)
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
            (df["normalized_long_form"] != "")
        ].copy()

        safe_rules: list[SafeRule] = []
        conflict_count = 0

        grouped = df.groupby("abbreviation", dropna=False)
        for abbreviation, group in grouped:
            unique_forms = sorted(set(group["normalized_long_form"].tolist()))
            if len(unique_forms) != 1:
                conflict_count += 1
                continue

            long_form = group["long_form"].iloc[0]
            normalized_long_form = unique_forms[0]
            detection_types = sorted(set(x for x in group["detection_type"].tolist() if x))

            safe_rules.append(
                SafeRule(
                    abbreviation=abbreviation,
                    long_form=long_form,
                    normalized_long_form=normalized_long_form,
                    detection_types=detection_types,
                    match_count_in_csv=len(group),
                )
            )

        safe_rules.sort(key=lambda item: len(item.long_form), reverse=True)
        return safe_rules, conflict_count

    # ------------------------------------------------------------------
    # Обход документа
    # ------------------------------------------------------------------

    def _iter_document_containers(self, doc: Document):
        for index, paragraph in enumerate(doc.paragraphs):
            yield {
                "source_type": "paragraph",
                "source_index": str(index),
                "get_text": lambda p=paragraph: p.text,
                "set_text": lambda new_text, p=paragraph: self._set_paragraph_text(p, new_text),
            }

        seen_cells: set[int] = set()
        for table_index, table in enumerate(doc.tables):
            for row_index, row in enumerate(table.rows):
                for cell_index, cell in enumerate(row.cells):
                    cell_id = id(cell._tc)
                    if cell_id in seen_cells:
                        continue
                    seen_cells.add(cell_id)

                    yield {
                        "source_type": "table_cell",
                        "source_index": f"table_{table_index}_row_{row_index}_cell_{cell_index}",
                        "get_text": lambda c=cell: c.text,
                        "set_text": lambda new_text, c=cell: self._set_cell_text(c, new_text),
                    }

    def _extract_document_texts(self, doc: Document) -> list[dict[str, str]]:
        result: list[dict[str, str]] = []
        for container in self._iter_document_containers(doc):
            result.append(
                {
                    "source_type": container["source_type"],
                    "source_index": container["source_index"],
                    "text": container["get_text"](),
                }
            )
        return result

    @staticmethod
    def _set_paragraph_text(paragraph, new_text: str) -> None:
        paragraph.text = new_text

    @staticmethod
    def _set_cell_text(cell, new_text: str) -> None:
        cell.text = new_text

    # ------------------------------------------------------------------
    # Замена повторных объявлений
    # ------------------------------------------------------------------

    def _process_text_container(
        self,
        text: str,
        source_type: str,
        source_index: str,
        safe_rules: list[SafeRule],
        occurrence_counters: dict[tuple[str, str], int],
        events: list[ReplacementEvent],
    ) -> str:
        updated_text = text

        for rule in safe_rules:
            key = (rule.abbreviation, rule.normalized_long_form)
            matches = self._find_declaration_matches(updated_text, rule)
            if not matches:
                continue

            pieces: list[str] = []
            last_pos = 0

            for match_info in matches:
                start = match_info["start"]
                end = match_info["end"]
                matched_text = match_info["matched_text"]

                occurrence_counters[key] += 1
                occurrence_number = occurrence_counters[key]

                if occurrence_number == 1:
                    replacement_text = matched_text
                    action = "keep_first_declaration"
                    comment = "Первое объявление сокращения оставлено без изменений."
                else:
                    replacement_text = rule.abbreviation
                    action = "replace_repeated_declaration"
                    comment = "Повторное объявление заменено на сокращение."

                pieces.append(updated_text[last_pos:start])
                pieces.append(replacement_text)
                last_pos = end

                events.append(
                    ReplacementEvent(
                        abbreviation=rule.abbreviation,
                        long_form=rule.long_form,
                        normalized_long_form=rule.normalized_long_form,
                        source_type=source_type,
                        source_index=source_index,
                        matched_text=matched_text,
                        replacement_text=replacement_text,
                        action=action,
                        occurrence_number_global=occurrence_number,
                        comment=comment,
                    )
                )

            pieces.append(updated_text[last_pos:])
            updated_text = "".join(pieces)

        return updated_text

    # ------------------------------------------------------------------
    # Анализ статистики употреблений
    # ------------------------------------------------------------------

    def _analyze_usage_statistics(
        self,
        safe_rules: list[SafeRule],
        source_texts: list[dict[str, str]],
    ) -> dict[tuple[str, str], dict[str, Any]]:
        stats_map: dict[tuple[str, str], dict[str, Any]] = {}

        for rule in safe_rules:
            key = (rule.abbreviation, rule.normalized_long_form)
            abbreviation_regex = self._build_abbreviation_regex(rule.abbreviation)

            first_declaration_found = False
            first_declaration_source_type = ""
            first_declaration_source_index = ""

            total_declaration_occurrences_found = 0
            total_abbreviation_mentions_found = 0
            abbreviation_mentions_inside_declarations = 0
            plain_abbreviation_mentions_outside_declarations = 0
            plain_abbreviation_mentions_after_first_declaration = 0

            for container in source_texts:
                text = container["text"]
                if not text or not text.strip():
                    continue

                declaration_matches = self._find_declaration_matches(text, rule)
                declaration_spans = [(m["start"], m["end"]) for m in declaration_matches]

                total_declaration_occurrences_found += len(declaration_matches)

                if declaration_matches and not first_declaration_found:
                    first_declaration_found = True
                    first_declaration_source_type = container["source_type"]
                    first_declaration_source_index = container["source_index"]

                all_abbr_matches = list(abbreviation_regex.finditer(text))
                total_abbreviation_mentions_found += len(all_abbr_matches)

                inside_decl_matches = []
                outside_decl_matches = []

                for abbr_match in all_abbr_matches:
                    span = abbr_match.span()
                    if any(self._spans_overlap(span, decl_span) for decl_span in declaration_spans):
                        inside_decl_matches.append(abbr_match)
                    else:
                        outside_decl_matches.append(abbr_match)

                abbreviation_mentions_inside_declarations += len(inside_decl_matches)
                plain_abbreviation_mentions_outside_declarations += len(outside_decl_matches)

                if first_declaration_found:
                    if not declaration_matches or container["source_index"] != first_declaration_source_index:
                        plain_abbreviation_mentions_after_first_declaration += len(outside_decl_matches)
                    else:
                        first_decl_end = declaration_matches[0]["end"]
                        plain_abbreviation_mentions_after_first_declaration += sum(
                            1 for match in outside_decl_matches if match.start() > first_decl_end
                        )

            repeated_declarations_expected = max(total_declaration_occurrences_found - 1, 0)
            total_repeated_usages_after_first_declaration = (
                repeated_declarations_expected + plain_abbreviation_mentions_after_first_declaration
            )

            stats_map[key] = {
                "abbreviation": rule.abbreviation,
                "long_form": rule.long_form,
                "normalized_long_form": rule.normalized_long_form,
                "safe_rule": True,
                "match_count_in_existing_abbreviations_csv": rule.match_count_in_csv,
                "first_declaration_found": first_declaration_found,
                "first_declaration_source_type": first_declaration_source_type,
                "first_declaration_source_index": first_declaration_source_index,
                "total_declaration_occurrences_found": total_declaration_occurrences_found,
                "total_abbreviation_mentions_found": total_abbreviation_mentions_found,
                "abbreviation_mentions_inside_declarations": abbreviation_mentions_inside_declarations,
                "plain_abbreviation_mentions_outside_declarations": plain_abbreviation_mentions_outside_declarations,
                "plain_abbreviation_mentions_after_first_declaration": plain_abbreviation_mentions_after_first_declaration,
                "repeated_declarations_expected": repeated_declarations_expected,
                "total_repeated_usages_after_first_declaration": total_repeated_usages_after_first_declaration,
            }

        return stats_map

    # ------------------------------------------------------------------
    # Формирование DataFrame
    # ------------------------------------------------------------------

    def _build_report_dataframe(self, events: list[ReplacementEvent]) -> pd.DataFrame:
        if not events:
            return pd.DataFrame(
                columns=[
                    "abbreviation",
                    "long_form",
                    "normalized_long_form",
                    "source_type",
                    "source_index",
                    "matched_text",
                    "replacement_text",
                    "action",
                    "occurrence_number_global",
                    "comment",
                ]
            )
        return pd.DataFrame([asdict(event) for event in events])

    def _build_statistics_dataframe(
        self,
        safe_rules: list[SafeRule],
        events: list[ReplacementEvent],
        usage_stats_map: dict[tuple[str, str], dict[str, Any]],
    ) -> pd.DataFrame:
        stats_map: dict[tuple[str, str], dict[str, Any]] = {}

        for rule in safe_rules:
            key = (rule.abbreviation, rule.normalized_long_form)
            base = usage_stats_map.get(key, {}).copy()
            if not base:
                base = {
                    "abbreviation": rule.abbreviation,
                    "long_form": rule.long_form,
                    "normalized_long_form": rule.normalized_long_form,
                    "safe_rule": True,
                    "match_count_in_existing_abbreviations_csv": rule.match_count_in_csv,
                    "first_declaration_found": False,
                    "first_declaration_source_type": "",
                    "first_declaration_source_index": "",
                    "total_declaration_occurrences_found": 0,
                    "total_abbreviation_mentions_found": 0,
                    "abbreviation_mentions_inside_declarations": 0,
                    "plain_abbreviation_mentions_outside_declarations": 0,
                    "plain_abbreviation_mentions_after_first_declaration": 0,
                    "repeated_declarations_expected": 0,
                    "total_repeated_usages_after_first_declaration": 0,
                }
            base["first_declarations_kept"] = 0
            base["repeated_declarations_replaced"] = 0
            base["replacement_rate_percent"] = 0.0
            stats_map[key] = base

        for event in events:
            key = (event.abbreviation, event.normalized_long_form)
            item = stats_map[key]
            if event.action == "keep_first_declaration":
                item["first_declarations_kept"] += 1
            elif event.action == "replace_repeated_declaration":
                item["repeated_declarations_replaced"] += 1

        rows: list[dict[str, Any]] = []
        for item in stats_map.values():
            expected = int(item.get("repeated_declarations_expected", 0))
            replaced = int(item.get("repeated_declarations_replaced", 0))
            item["replacement_rate_percent"] = round(replaced / expected * 100, 2) if expected > 0 else 0.0
            rows.append(item)

        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values(by=["abbreviation", "long_form"], ascending=[True, True], kind="stable").reset_index(drop=True)
        return df

    def _build_summary_dataframe(
        self,
        output_docx_path: Path,
        safe_rules_count: int,
        conflict_count: int,
        events: list[ReplacementEvent],
        statistics_df: pd.DataFrame,
    ) -> pd.DataFrame:
        first_kept = sum(1 for event in events if event.action == "keep_first_declaration")
        repeated_replaced = sum(1 for event in events if event.action == "replace_repeated_declaration")
        total_occurrences = len(events)

        return pd.DataFrame(
            [
                {
                    "output_docx": str(output_docx_path).replace("/", "\\"),
                    "safe_abbreviation_rules": safe_rules_count,
                    "conflict_abbreviations_count": conflict_count,
                    "total_declaration_occurrences_found": total_occurrences,
                    "first_declarations_kept": first_kept,
                    "repeated_declarations_expected": int(statistics_df["repeated_declarations_expected"].sum()) if not statistics_df.empty else 0,
                    "repeated_declarations_replaced": repeated_replaced,
                    "plain_abbreviation_mentions_after_first_declaration": int(statistics_df["plain_abbreviation_mentions_after_first_declaration"].sum()) if not statistics_df.empty else 0,
                    "total_repeated_usages_after_first_declaration": int(statistics_df["total_repeated_usages_after_first_declaration"].sum()) if not statistics_df.empty else 0,
                    "statistics_by_abbreviation_created": True,
                    "notes": (
                        "Добавлено распознавание строк глоссария и улучшен выбор длинной части перед скобками. "
                        "Это повышает вероятность обнаружения первого объявления сокращения в реальном документе."
                    ),
                }
            ]
        )

    # ------------------------------------------------------------------
    # Сохранение результатов
    # ------------------------------------------------------------------

    def _save_dataframes(
        self,
        output_dir: Path,
        output_docx_path: Path,
        report_df: pd.DataFrame,
        summary_df: pd.DataFrame,
        statistics_df: pd.DataFrame,
    ) -> dict[str, Path]:
        output_dir.mkdir(parents=True, exist_ok=True)

        report_csv = output_dir / f"{output_docx_path.stem}_replacement_report.csv"
        report_xlsx = output_dir / f"{output_docx_path.stem}_replacement_report.xlsx"

        summary_csv = output_dir / f"{output_docx_path.stem}_replacement_summary.csv"
        summary_xlsx = output_dir / f"{output_docx_path.stem}_replacement_summary.xlsx"

        stats_csv = output_dir / f"{output_docx_path.stem}_replacement_statistics_by_abbreviation.csv"
        stats_xlsx = output_dir / f"{output_docx_path.stem}_replacement_statistics_by_abbreviation.xlsx"

        report_df.to_csv(report_csv, index=False, encoding="utf-8-sig")
        report_df.to_excel(report_xlsx, index=False)

        summary_df.to_csv(summary_csv, index=False, encoding="utf-8-sig")
        summary_df.to_excel(summary_xlsx, index=False)

        statistics_df.to_csv(stats_csv, index=False, encoding="utf-8-sig")
        statistics_df.to_excel(stats_xlsx, index=False)

        return {
            "output_docx": output_docx_path,
            "replacement_report_csv": report_csv,
            "replacement_report_xlsx": report_xlsx,
            "replacement_summary_csv": summary_csv,
            "replacement_summary_xlsx": summary_xlsx,
            "replacement_statistics_csv": stats_csv,
            "replacement_statistics_xlsx": stats_xlsx,
        }

    # ------------------------------------------------------------------
    # Основной сценарий
    # ------------------------------------------------------------------

    def run(
        self,
        source_docx_path: str | Path,
        existing_abbreviations_csv: str | Path,
        output_dir: str | Path,
    ) -> dict[str, Path]:
        source_docx_path = Path(source_docx_path)
        output_dir = Path(output_dir)

        if not source_docx_path.exists():
            raise FileNotFoundError(f"Файл не найден: {source_docx_path}")

        safe_rules, conflict_count = self._load_rules_from_existing_abbreviations(existing_abbreviations_csv)

        source_doc_for_analysis = Document(source_docx_path)
        source_texts = self._extract_document_texts(source_doc_for_analysis)
        usage_stats_map = self._analyze_usage_statistics(safe_rules, source_texts)

        doc = Document(source_docx_path)
        occurrence_counters: dict[tuple[str, str], int] = defaultdict(int)
        events: list[ReplacementEvent] = []

        for container in self._iter_document_containers(doc):
            current_text = container["get_text"]()
            if not current_text or not current_text.strip():
                continue

            new_text = self._process_text_container(
                text=current_text,
                source_type=container["source_type"],
                source_index=container["source_index"],
                safe_rules=safe_rules,
                occurrence_counters=occurrence_counters,
                events=events,
            )

            if new_text != current_text:
                container["set_text"](new_text)

        output_docx_path = output_dir / f"{source_docx_path.stem}_replaced_declarations.docx"
        output_docx_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_docx_path)

        report_df = self._build_report_dataframe(events)
        statistics_df = self._build_statistics_dataframe(
            safe_rules=safe_rules,
            events=events,
            usage_stats_map=usage_stats_map,
        )
        summary_df = self._build_summary_dataframe(
            output_docx_path=output_docx_path,
            safe_rules_count=len(safe_rules),
            conflict_count=conflict_count,
            events=events,
            statistics_df=statistics_df,
        )

        return self._save_dataframes(
            output_dir=output_dir,
            output_docx_path=output_docx_path,
            report_df=report_df,
            summary_df=summary_df,
            statistics_df=statistics_df,
        )


if __name__ == "__main__":
    replacer = RepeatedDeclarationReplacer()
    result = replacer.run(
        source_docx_path="test_reduction_input.docx",
        existing_abbreviations_csv="result_all/stage2/existing_abbreviations.csv",
        output_dir="result_all/replacement_stage",
    )
    print("=" * 72)
    print("ЗАМЕНА ПОВТОРНЫХ ОБЪЯВЛЕНИЙ ЗАВЕРШЕНА")
    print("=" * 72)
    for key, value in result.items():
        print(f"{key}: {value}")
