
from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any
import re

import pandas as pd
from docx import Document


@dataclass
class SafeRule:
    abbreviation: str
    long_form: str
    normalized_long_form: str
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

    @staticmethod
    def _clean_text(value: Any) -> str:
        text = str(value).strip()
        return " ".join(text.split()) if text else ""

    @classmethod
    def _normalize_long_form(cls, value: Any) -> str:
        return cls._clean_text(value).lower()

    @staticmethod
    def _extract_words(text: str) -> list[str]:
        return re.findall(r"[A-Za-zА-Яа-яЁё0-9-]+", str(text))

    @staticmethod
    def _make_flexible_whitespace_pattern(text: str) -> str:
        parts = [re.escape(part) for part in str(text).split()]
        return r"\s+".join(parts)

    @staticmethod
    def _spans_overlap(a: tuple[int, int], b: tuple[int, int]) -> bool:
        return max(a[0], b[0]) < min(a[1], b[1])

    @classmethod
    def _token_overlap_ratio(cls, left: str, right: str) -> float:
        left_tokens = {token.lower() for token in cls._extract_words(left) if len(token) > 2}
        right_tokens = {token.lower() for token in cls._extract_words(right) if len(token) > 2}
        if not left_tokens or not right_tokens:
            return 0.0
        return len(left_tokens & right_tokens) / max(len(right_tokens), 1)

    @classmethod
    def _looks_like_same_long_form(cls, candidate_long_form: str, rule_long_form: str) -> bool:
        candidate = cls._clean_text(candidate_long_form)
        rule = cls._clean_text(rule_long_form)
        if not candidate or not rule:
            return False
        if candidate.lower() == rule.lower():
            return True
        overlap = cls._token_overlap_ratio(candidate, rule)
        rule_word_count = len(cls._extract_words(rule))
        if rule_word_count <= 2:
            return overlap >= 1.0
        if rule_word_count == 3:
            return overlap >= 0.66
        return overlap >= 0.6

    @classmethod
    def _build_exact_declaration_regex(cls, long_form: str, abbreviation: str) -> re.Pattern:
        long_pattern = cls._make_flexible_whitespace_pattern(long_form)
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        pattern = (
            rf"(?P<matched>(?P<long>{long_pattern})\s*\(\s*(?:далее\s*[–—-]\s*)?(?P<abbr>{abbr_pattern})\s*\))"
        )
        return re.compile(pattern, flags=re.IGNORECASE)

    @classmethod
    def _build_parenthetical_abbreviation_regex(cls, abbreviation: str) -> re.Pattern:
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        return re.compile(rf"\(\s*(?:далее\s*[–—-]\s*)?(?P<abbr>{abbr_pattern})\s*\)", flags=re.IGNORECASE)

    @classmethod
    def _build_glossary_abbr_first_regex(cls, abbreviation: str) -> re.Pattern:
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        return re.compile(rf"^\s*(?P<abbr>{abbr_pattern})\s*[–—-]\s*(?P<long>.+?)\s*$", flags=re.IGNORECASE)

    @classmethod
    def _build_glossary_long_first_regex(cls, abbreviation: str) -> re.Pattern:
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        return re.compile(rf"^\s*(?P<long>.+?)\s*[–—-]\s*(?P<abbr>{abbr_pattern})\s*$", flags=re.IGNORECASE)

    @classmethod
    def _build_abbreviation_regex(cls, abbreviation: str) -> re.Pattern:
        abbr_pattern = cls._make_flexible_whitespace_pattern(abbreviation)
        return re.compile(rf"(?<!\w){abbr_pattern}(?!\w)", flags=re.IGNORECASE)

    def _load_safe_rules(self, existing_abbreviations_csv: Path) -> list[SafeRule]:
        df = pd.read_csv(existing_abbreviations_csv, encoding="utf-8-sig")
        if df.empty:
            return []
        if "abbreviation" not in df.columns:
            raise ValueError("Во входном CSV отсутствует столбец abbreviation.")
        possible_long_columns = [col for col in ["long_form", "matched_term"] if col in df.columns]
        if not possible_long_columns:
            raise ValueError("Во входном CSV отсутствуют long_form и matched_term.")

        rules_map: dict[tuple[str, str], SafeRule] = {}
        for _, row in df.iterrows():
            abbreviation = self._clean_text(row.get("abbreviation", "")).upper()
            if not abbreviation:
                continue
            long_form = ""
            for col in possible_long_columns:
                value = self._clean_text(row.get(col, ""))
                if value:
                    long_form = value
                    break
            if not long_form or len(self._extract_words(long_form)) < 2:
                continue

            normalized = self._normalize_long_form(long_form)
            key = (abbreviation, normalized)
            if key not in rules_map:
                rules_map[key] = SafeRule(abbreviation=abbreviation, long_form=long_form, normalized_long_form=normalized, match_count_in_csv=1)
            else:
                rules_map[key].match_count_in_csv += 1

        return sorted(rules_map.values(), key=lambda x: (x.abbreviation, x.long_form))

    def _iter_text_containers(self, document: Document) -> list[dict[str, Any]]:
        containers = []
        for i, paragraph in enumerate(document.paragraphs):
            containers.append({"source_type": "paragraph", "source_index": str(i), "text": paragraph.text, "object": paragraph})
        table_cell_index = 0
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        containers.append({"source_type": "table_cell", "source_index": str(table_cell_index), "text": paragraph.text, "object": paragraph})
                        table_cell_index += 1
        return containers

    def _extract_candidate_long_form_before_parentheses(self, text: str, paren_start: int, rule_long_form: str) -> tuple[str, int]:
        left_boundary = 0
        search_zone = text[max(0, paren_start - 220):paren_start]
        for match in re.finditer(r"[\n\r.;:!?]", search_zone):
            left_boundary = match.end()
        fragment = re.sub(r"\s+", " ", search_zone[left_boundary:]).strip()
        if not fragment:
            return "", paren_start

        token_matches = list(re.finditer(r"[A-Za-zА-Яа-яЁё0-9-]+", fragment))
        if len(token_matches) < 2:
            return "", paren_start

        best_candidate = ""
        best_start = 0
        best_score = (-1.0, 0)

        max_tail = min(len(token_matches), 14)
        for tail_size in range(2, max_tail + 1):
            tail = token_matches[-tail_size:]
            start_in_fragment = tail[0].start()
            candidate = fragment[start_in_fragment:].strip(" ,;:«»\"'()[]{}")
            if not candidate:
                continue
            if not self._looks_like_same_long_form(candidate, rule_long_form):
                continue
            overlap = self._token_overlap_ratio(candidate, rule_long_form)
            token_count = len(self._extract_words(candidate))
            score = (overlap, -token_count)
            if score > best_score:
                best_score = score
                best_candidate = candidate
                best_start = start_in_fragment

        absolute_start = paren_start - len(fragment) + best_start
        return best_candidate, absolute_start

    def _find_declaration_matches(self, text: str, rule: SafeRule) -> list[dict[str, Any]]:
        matches = []
        used_spans: list[tuple[int, int]] = []
        clean_line = self._clean_text(text)

        m = self._build_glossary_abbr_first_regex(rule.abbreviation).match(clean_line)
        if m:
            candidate_long = self._clean_text(m.group("long"))
            if self._looks_like_same_long_form(candidate_long, rule.long_form):
                matches.append({"start": 0, "end": len(text), "matched_text": text, "long_text": candidate_long, "match_type": "glossary_abbr_first"})
                used_spans.append((0, len(text)))

        m = self._build_glossary_long_first_regex(rule.abbreviation).match(clean_line)
        if m:
            candidate_long = self._clean_text(m.group("long"))
            if self._looks_like_same_long_form(candidate_long, rule.long_form):
                span = (0, len(text))
                if span not in used_spans:
                    matches.append({"start": 0, "end": len(text), "matched_text": text, "long_text": candidate_long, "match_type": "glossary_long_first"})
                    used_spans.append(span)

        strict_regex = self._build_exact_declaration_regex(rule.long_form, rule.abbreviation)
        for match in strict_regex.finditer(text):
            span = match.span("matched")
            if any(self._spans_overlap(span, used) for used in used_spans):
                continue
            matches.append({"start": span[0], "end": span[1], "matched_text": match.group("matched"), "long_text": self._clean_text(match.group("long")), "match_type": "strict"})
            used_spans.append(span)

        paren_regex = self._build_parenthetical_abbreviation_regex(rule.abbreviation)
        for paren_match in paren_regex.finditer(text):
            paren_span = paren_match.span()
            if any(self._spans_overlap(paren_span, used) for used in used_spans):
                continue
            candidate_long, absolute_start = self._extract_candidate_long_form_before_parentheses(text, paren_span[0], rule.long_form)
            if not candidate_long or not self._looks_like_same_long_form(candidate_long, rule.long_form):
                continue
            span = (absolute_start, paren_span[1])
            if any(self._spans_overlap(span, used) for used in used_spans):
                continue
            matches.append({"start": span[0], "end": span[1], "matched_text": text[span[0]:span[1]], "long_text": candidate_long, "match_type": "flexible_parenthetical"})
            used_spans.append(span)

        matches.sort(key=lambda x: (x["start"], x["end"]))
        return matches

    def _process_text_container(self, text: str, source_type: str, source_index: str, safe_rules: list[SafeRule], occurrence_counters: dict[tuple[str, str], int], events: list[ReplacementEvent]) -> str:
        updated_text = text
        for rule in safe_rules:
            key = (rule.abbreviation, rule.normalized_long_form)
            matches = self._find_declaration_matches(updated_text, rule)
            if not matches:
                continue

            pieces = []
            last_pos = 0
            for info in matches:
                start, end = info["start"], info["end"]
                matched_text = info["matched_text"]
                occurrence_counters[key] = occurrence_counters.get(key, 0) + 1
                occurrence_number = occurrence_counters[key]

                if occurrence_number == 1:
                    replacement_text = matched_text
                    action = "keep_first_declaration"
                    comment = "Первое объявление сокращения оставлено без изменений."
                else:
                    replacement_text = rule.abbreviation
                    action = "replace_repeated_declaration"
                    comment = "Повторное объявление заменено целиком на сокращение."

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

    def _build_report_dataframe(self, events: list[ReplacementEvent]) -> pd.DataFrame:
        if not events:
            return pd.DataFrame(columns=["abbreviation", "long_form", "normalized_long_form", "source_type", "source_index", "matched_text", "replacement_text", "action", "occurrence_number_global", "comment"])
        return pd.DataFrame([asdict(item) for item in events])

    def _analyze_usage_statistics(self, safe_rules: list[SafeRule], source_texts: list[dict[str, str]], events: list[ReplacementEvent]) -> pd.DataFrame:
        rows = []
        for rule in safe_rules:
            declaration_occurrences = 0
            first_declaration_found = False
            first_declaration_source_type = ""
            first_declaration_source_index = ""
            total_abbreviation_mentions_found = 0
            abbreviation_mentions_inside_declarations = 0
            plain_abbreviation_mentions_outside_declarations = 0
            plain_abbreviation_mentions_after_first_declaration = 0

            abbr_regex = self._build_abbreviation_regex(rule.abbreviation)
            first_decl_end_map: dict[tuple[str, str], int] = {}

            for container in source_texts:
                text = container["text"]
                if not self._clean_text(text):
                    continue
                declaration_matches = self._find_declaration_matches(text, rule)
                declaration_occurrences += len(declaration_matches)
                if declaration_matches and not first_declaration_found:
                    first_declaration_found = True
                    first_declaration_source_type = container["source_type"]
                    first_declaration_source_index = container["source_index"]
                    first_decl_end_map[(container["source_type"], container["source_index"])] = declaration_matches[0]["end"]

                declaration_spans = [(m["start"], m["end"]) for m in declaration_matches]
                abbr_matches = list(abbr_regex.finditer(text))
                total_abbreviation_mentions_found += len(abbr_matches)

                first_decl_end = first_decl_end_map.get((container["source_type"], container["source_index"]), -1)
                for match in abbr_matches:
                    span = match.span()
                    inside = any(self._spans_overlap(span, decl_span) for decl_span in declaration_spans)
                    if inside:
                        abbreviation_mentions_inside_declarations += 1
                    else:
                        plain_abbreviation_mentions_outside_declarations += 1
                        if first_declaration_found:
                            if container["source_index"] != first_declaration_source_index:
                                plain_abbreviation_mentions_after_first_declaration += 1
                            elif span[0] > first_decl_end:
                                plain_abbreviation_mentions_after_first_declaration += 1

            repeated_expected = max(declaration_occurrences - 1, 0)
            repeated_replaced = sum(1 for e in events if e.abbreviation == rule.abbreviation and e.normalized_long_form == rule.normalized_long_form and e.action == "replace_repeated_declaration")

            rows.append({
                "abbreviation": rule.abbreviation,
                "long_form": rule.long_form,
                "normalized_long_form": rule.normalized_long_form,
                "safe_rule": True,
                "match_count_in_existing_abbreviations_csv": rule.match_count_in_csv,
                "first_declaration_found": first_declaration_found,
                "first_declaration_source_type": first_declaration_source_type,
                "first_declaration_source_index": first_declaration_source_index,
                "total_declaration_occurrences_found": declaration_occurrences,
                "total_abbreviation_mentions_found": total_abbreviation_mentions_found,
                "abbreviation_mentions_inside_declarations": abbreviation_mentions_inside_declarations,
                "plain_abbreviation_mentions_outside_declarations": plain_abbreviation_mentions_outside_declarations,
                "plain_abbreviation_mentions_after_first_declaration": plain_abbreviation_mentions_after_first_declaration,
                "repeated_declarations_expected": repeated_expected,
                "repeated_declarations_replaced": repeated_replaced,
                "total_repeated_usages_after_first_declaration": repeated_expected + plain_abbreviation_mentions_after_first_declaration,
                "replacement_rate_percent": round(repeated_replaced / repeated_expected * 100, 2) if repeated_expected > 0 else 0.0,
            })

        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values(by=["abbreviation", "long_form"], ascending=[True, True], kind="stable").reset_index(drop=True)
        return df

    def _build_summary_dataframe(self, output_docx_path: Path, safe_rules_count: int, conflict_count: int, events: list[ReplacementEvent], statistics_df: pd.DataFrame) -> pd.DataFrame:
        first_kept = sum(1 for e in events if e.action == "keep_first_declaration")
        repeated_replaced = sum(1 for e in events if e.action == "replace_repeated_declaration")
        total_occurrences = len(events)
        return pd.DataFrame([{
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
            "notes": "Заменяются только повторные объявления целиком. Это уменьшает риск поломки конструкции '(далее – АББР)' и склейки слов.",
        }])

    def _save_dataframes(self, output_dir: Path, output_docx_path: Path, report_df: pd.DataFrame, summary_df: pd.DataFrame, statistics_df: pd.DataFrame) -> dict[str, Path]:
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

    def run(self, source_docx_path: str | Path, existing_abbreviations_csv: str | Path, output_dir: str | Path) -> dict[str, Path]:
        source_docx_path = Path(source_docx_path)
        existing_abbreviations_csv = Path(existing_abbreviations_csv)
        output_dir = Path(output_dir)

        if not source_docx_path.exists():
            raise FileNotFoundError(f"Файл не найден: {source_docx_path}")
        if not existing_abbreviations_csv.exists():
            raise FileNotFoundError(f"Файл не найден: {existing_abbreviations_csv}")

        safe_rules = self._load_safe_rules(existing_abbreviations_csv)
        document = Document(source_docx_path)
        containers = self._iter_text_containers(document)

        occurrence_counters: dict[tuple[str, str], int] = {}
        events: list[ReplacementEvent] = []

        for container in containers:
            original_text = container["text"]
            if not self._clean_text(original_text):
                continue
            updated_text = self._process_text_container(
                text=original_text,
                source_type=container["source_type"],
                source_index=container["source_index"],
                safe_rules=safe_rules,
                occurrence_counters=occurrence_counters,
                events=events,
            )
            if updated_text != original_text:
                container["object"].text = updated_text
                container["text"] = updated_text

        output_dir.mkdir(parents=True, exist_ok=True)
        output_docx_path = output_dir / f"{source_docx_path.stem}_replaced_declarations.docx"
        document.save(output_docx_path)

        source_texts = [{"source_type": c["source_type"], "source_index": c["source_index"], "text": c["text"]} for c in containers]
        report_df = self._build_report_dataframe(events)
        statistics_df = self._analyze_usage_statistics(safe_rules, source_texts, events)
        summary_df = self._build_summary_dataframe(output_docx_path, len(safe_rules), 0, events, statistics_df)
        return self._save_dataframes(output_dir, output_docx_path, report_df, summary_df, statistics_df)
