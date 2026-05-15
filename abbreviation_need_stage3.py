
from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Any

import pandas as pd
import regex

from abbreviation_extraction_stage2 import Stage2ReductionAnalyzer


@dataclass
class AbbreviationDecision:
    term: str
    suggested_abbreviation: str
    abbreviation_found_in_text: bool
    found_abbreviation: str
    frequency: int
    word_count: int
    char_length: int
    decision_score: int
    need_to_introduce: bool
    priority: str
    reason: str


class AbbreviationNeedAnalyzer:
    def __init__(self) -> None:
        self.stage2_analyzer = Stage2ReductionAnalyzer()
        self.bad_tokens = {
            "в", "во", "на", "по", "при", "для", "из", "с", "со", "к", "ко",
            "и", "или", "а", "но", "как", "что", "чтобы", "который", "которые",
            "которых", "которому", "данный", "данных", "данного", "данной",
            "этот", "эта", "эти", "того", "таких", "например", "включая",
            "целью", "рамках", "выбраны", "следующие", "представлен", "представлены",
            "осуществляется", "используются", "используется", "необходимо",
            "выполняется", "реализованы", "реализации", "разработке",
            "адрес", "целью", "внедрения", "состоящий", "состоящими",
        }
        self.bad_term_starts = {
            "в", "во", "на", "по", "при", "для", "из", "с", "со", "к",
            "данных", "выбраны", "следующие", "целью", "адрес", "включая",
        }
        self.bad_term_ends = {
            "которые", "которых", "используются", "используется", "представлен",
            "представлены", "реализованы", "реализации", "данных", "включая",
        }

    def run(self, docx_path: str | Path, output_dir: str | Path) -> Dict[str, Path]:
        docx_path = Path(docx_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        stage2_output_dir = output_dir / "stage2_intermediate"
        stage2_saved = self.stage2_analyzer.run(docx_path, stage2_output_dir)

        merged_csv = stage2_saved["merged_csv"]
        existing_csv = stage2_saved["existing_abbreviations_csv"]

        merged_df = pd.read_csv(merged_csv, encoding="utf-8-sig")
        existing_abbreviations_df = pd.read_csv(existing_csv, encoding="utf-8-sig")

        decisions = self._build_decisions(merged_df, existing_abbreviations_df)
        decisions_df = pd.DataFrame([asdict(item) for item in decisions])
        recommendations_df = self._build_recommendations_table(decisions_df)

        return self._save_results(
            merged_df=merged_df,
            existing_abbreviations_df=existing_abbreviations_df,
            decisions_df=decisions_df,
            recommendations_df=recommendations_df,
            output_dir=output_dir,
        )

    def _build_decisions(self, merged_df: pd.DataFrame, existing_abbreviations_df: pd.DataFrame) -> List[AbbreviationDecision]:
        decisions: List[AbbreviationDecision] = []
        if merged_df.empty:
            return decisions

        existing_abbrs = {
            str(value).strip().upper()
            for value in existing_abbreviations_df.get("abbreviation", pd.Series(dtype=str)).fillna("").tolist()
            if str(value).strip()
        }

        long_forms_in_document = set()
        for column in ["long_form", "matched_term"]:
            if column in existing_abbreviations_df.columns:
                for value in existing_abbreviations_df[column].fillna("").tolist():
                    cleaned = self._clean_text(value).lower()
                    if cleaned:
                        long_forms_in_document.add(cleaned)

        temp: List[AbbreviationDecision] = []

        for row in merged_df.to_dict("records"):
            term = self._clean_text(row.get("term", ""))
            suggested_abbreviation = self._clean_text(row.get("suggested_abbreviation", "")).upper()
            abbreviation_found_in_text = bool(row.get("abbreviation_found_in_text", False))
            found_abbreviation = self._clean_text(row.get("found_abbreviation", "")).upper()
            frequency = int(self._safe_int(row.get("frequency", 0)))
            word_count = int(self._safe_int(row.get("word_count", len(self._extract_words(term)))))
            char_length = len(term)

            score, reason, priority, need_to_introduce = self._evaluate_term(
                term=term,
                suggested_abbreviation=suggested_abbreviation,
                abbreviation_found_in_text=abbreviation_found_in_text,
                frequency=frequency,
                word_count=word_count,
                char_length=char_length,
                existing_abbrs=existing_abbrs,
                long_forms_in_document=long_forms_in_document,
            )

            temp.append(
                AbbreviationDecision(
                    term=term,
                    suggested_abbreviation=suggested_abbreviation,
                    abbreviation_found_in_text=abbreviation_found_in_text,
                    found_abbreviation=found_abbreviation,
                    frequency=frequency,
                    word_count=word_count,
                    char_length=char_length,
                    decision_score=score,
                    need_to_introduce=need_to_introduce,
                    priority=priority,
                    reason=reason,
                )
            )

        best_by_abbr: dict[str, AbbreviationDecision] = {}
        for item in temp:
            if not item.suggested_abbreviation:
                continue
            prev = best_by_abbr.get(item.suggested_abbreviation)
            if prev is None or (
                item.decision_score > prev.decision_score
                or (item.decision_score == prev.decision_score and item.frequency > prev.frequency)
                or (item.decision_score == prev.decision_score and item.frequency == prev.frequency and len(item.term) < len(prev.term))
            ):
                best_by_abbr[item.suggested_abbreviation] = item

        decisions = list(best_by_abbr.values())
        decisions.sort(
            key=lambda item: (
                item.need_to_introduce,
                item.decision_score,
                item.frequency,
                item.word_count,
                -item.char_length,
                item.term,
            ),
            reverse=True,
        )
        return decisions

    def _evaluate_term(
        self,
        term: str,
        suggested_abbreviation: str,
        abbreviation_found_in_text: bool,
        frequency: int,
        word_count: int,
        char_length: int,
        existing_abbrs: set[str],
        long_forms_in_document: set[str],
    ) -> tuple[int, str, str, bool]:
        if abbreviation_found_in_text:
            return (0, "Аббревиатура уже присутствует в тексте, дополнительный ввод не требуется.", "none",
                    False)

        if not term or not suggested_abbreviation:
            return (0, "Недостаточно данных для рекомендации.", "none", False)



        hard_fail_reason = self._hard_filter_reason(term, suggested_abbreviation, word_count, existing_abbrs, long_forms_in_document)
        if hard_fail_reason:
            return (0, " " + hard_fail_reason, "none", False)

        reasons: list[str] = []
        score = 0

        if frequency >= 5:
            score += 30
            reasons.append("термин встречается часто")
        elif frequency >= 3:
            score += 20
            reasons.append("термин встречается несколько раз")
        elif frequency >= 2:
            score += 10
            reasons.append("термин встречается более одного раза")
        else:
            reasons.append("термин встречается редко")

        if 3 <= word_count <= 4:
            score += 20
            reasons.append(f"термин состоит из {word_count} слов")
        elif word_count == 2:
            score += 10
            reasons.append("термин состоит из 2 слов")
        elif word_count == 5:
            score += 8
            reasons.append("термин довольно длинный по количеству слов")
        elif word_count >= 6:
            score -= 20
            reasons.append("термин слишком длинный для автоматического ввода аббревиатуры")

        if char_length >= 40:
            score += 12
            reasons.append("полная форма очень длинная")
        elif char_length >= 25:
            score += 8
            reasons.append("полная форма достаточно длинная")
        elif char_length >= 18:
            score += 4
            reasons.append("полная форма средней длины")

        abbr_quality_score, abbr_reason = self._evaluate_suggested_abbreviation(suggested_abbreviation)
        score += abbr_quality_score
        if abbr_reason:
            reasons.append(abbr_reason)

        soft_penalty, penalty_reasons = self._soft_penalties(term)
        score += soft_penalty
        reasons.extend(penalty_reasons)

        if score >= 75:
            priority = "high"
            need_to_introduce = True
            reasons.insert(0, "рекомендуется ввести аббревиатуру")
        elif score >= 60:
            priority = "medium"
            need_to_introduce = True
            reasons.insert(0, "ввод аббревиатуры целесообразен")
        elif score >= 45:
            priority = "low"
            need_to_introduce = False
            reasons.insert(0, "кандидат спорный, требуется ручная проверка")
        else:
            priority = "none"
            need_to_introduce = False
            reasons.insert(0, "ввод аббревиатуры не требуется")

        reason_text = "" + "; ".join(dict.fromkeys(reasons))
        return score, reason_text, priority, need_to_introduce

    def _hard_filter_reason(self, term: str, suggested_abbreviation: str, word_count: int, existing_abbrs: set[str], long_forms_in_document: set[str]) -> str:
        words = self._extract_words(term)
        words_lower = [word.lower() for word in words]
        cleaned_term = self._clean_text(term).lower()

        if word_count < 2:
            return "Термин слишком короткий для ввода аббревиатуры."
        if word_count > 5:
            return "Термин слишком длинный и больше похож на фрагмент предложения, чем на устойчивый термин."
        if any(ch in term for ch in ";:!?"):
            return "Термин содержит знаки препинания предложения и не рассматривается как устойчивое словосочетание."
        if words_lower and words_lower[0] in self.bad_term_starts:
            return "Термин начинается с контекстного слова и выглядит как фрагмент предложения."
        if words_lower and words_lower[-1] in self.bad_term_ends:
            return "Термин заканчивается контекстным словом и выглядит как фрагмент предложения."
        if sum(1 for token in words_lower if token in self.bad_tokens) >= 2:
            return "В термине слишком много контекстных слов, поэтому автоматический ввод аббревиатуры запрещён."
        if suggested_abbreviation in existing_abbrs:
            return "Такое сокращение уже есть в документе."
        for existing_abbr in existing_abbrs:
            if len(existing_abbr) >= 3 and existing_abbr != suggested_abbreviation:
                if existing_abbr in suggested_abbreviation and len(suggested_abbreviation) > len(existing_abbr):
                    return "Предлагаемая аббревиатура выглядит как искусственное расширение уже существующего сокращения."
        if cleaned_term in long_forms_in_document:
            return "Для этого термина уже есть полная форма в документе, новое сокращение вводить не нужно."
        if len([w for w in words_lower if w not in self.bad_tokens]) < 2:
            return "Термин не содержит достаточного числа значимых слов."
        return ""

    def _soft_penalties(self, term: str) -> tuple[int, list[str]]:
        penalties = 0
        reasons: list[str] = []
        words = [word.lower() for word in self._extract_words(term)]
        if any(word in {"выбраны", "представлены", "осуществляется", "реализованы"} for word in words):
            penalties -= 25
            reasons.append("термин содержит глагольный контекст")
        if len(words) >= 5 and any(word in self.bad_tokens for word in words[:2]):
            penalties -= 10
            reasons.append("начало словосочетания похоже на контекстный хвост")
        return penalties, reasons

    def _evaluate_suggested_abbreviation(self, abbreviation: str) -> tuple[int, str]:
        abbreviation = self._clean_text(abbreviation).upper()
        if not abbreviation:
            return -40, "аббревиатура не сформирована"
        pure_len = len(abbreviation.replace(" ", ""))
        if pure_len < 2:
            return -30, "аббревиатура слишком короткая"
        if pure_len > 8:
            return -25, "аббревиатура слишком длинная"
        if 3 <= pure_len <= 6:
            return 12, "аббревиатура имеет удобную длину"
        if pure_len == 2:
            return 2, "аббревиатура короткая, но допустима"
        return 6, "аббревиатура допустима по длине"

    def _build_recommendations_table(self, decisions_df: pd.DataFrame) -> pd.DataFrame:
        if decisions_df.empty:
            return pd.DataFrame(columns=["term", "suggested_abbreviation", "abbreviation_found_in_text", "need_to_introduce", "priority", "decision_score", "reason"])
        result = decisions_df[["term", "suggested_abbreviation", "abbreviation_found_in_text", "need_to_introduce", "priority", "decision_score", "reason"]].copy()
        return result.sort_values(by=["need_to_introduce", "decision_score", "term"], ascending=[False, False, True], kind="stable").reset_index(drop=True)

    def _save_results(self, merged_df: pd.DataFrame, existing_abbreviations_df: pd.DataFrame, decisions_df: pd.DataFrame, recommendations_df: pd.DataFrame, output_dir: Path) -> Dict[str, Path]:
        saved_files: Dict[str, Path] = {}
        merged_csv = output_dir / "merged_terms_and_abbreviations.csv"
        merged_xlsx = output_dir / "merged_terms_and_abbreviations.xlsx"
        existing_csv = output_dir / "existing_abbreviations.csv"
        existing_xlsx = output_dir / "existing_abbreviations.xlsx"
        decisions_csv = output_dir / "abbreviation_decisions.csv"
        decisions_xlsx = output_dir / "abbreviation_decisions.xlsx"
        recommendations_csv = output_dir / "abbreviation_recommendations.csv"
        recommendations_xlsx = output_dir / "abbreviation_recommendations.xlsx"

        merged_df.to_csv(merged_csv, index=False, encoding="utf-8-sig")
        existing_abbreviations_df.to_csv(existing_csv, index=False, encoding="utf-8-sig")
        decisions_df.to_csv(decisions_csv, index=False, encoding="utf-8-sig")
        recommendations_df.to_csv(recommendations_csv, index=False, encoding="utf-8-sig")

        saved_files["merged_csv"] = merged_csv
        saved_files["existing_abbreviations_csv"] = existing_csv
        saved_files["abbreviation_decisions_csv"] = decisions_csv
        saved_files["abbreviation_recommendations_csv"] = recommendations_csv

        try:
            merged_df.to_excel(merged_xlsx, index=False)
            existing_abbreviations_df.to_excel(existing_xlsx, index=False)
            decisions_df.to_excel(decisions_xlsx, index=False)
            recommendations_df.to_excel(recommendations_xlsx, index=False)
            saved_files["merged_xlsx"] = merged_xlsx
            saved_files["existing_abbreviations_xlsx"] = existing_xlsx
            saved_files["abbreviation_decisions_xlsx"] = decisions_xlsx
            saved_files["abbreviation_recommendations_xlsx"] = recommendations_xlsx
        except ModuleNotFoundError as exc:
            print("Внимание: не удалось сохранить XLSX-файлы.")
            print("Причина:", exc)
            print("CSV-файлы при этом успешно сохранены.")

        return saved_files

    @staticmethod
    def _extract_words(text: str) -> list[str]:
        return regex.findall(r"[A-Za-zА-Яа-яЁё0-9-]+", str(text))

    @staticmethod
    def _clean_text(value: Any) -> str:
        text = str(value).strip()
        return " ".join(text.split()) if text else ""

    @staticmethod
    def _safe_int(value: Any) -> int:
        try:
            return int(value)
        except Exception:
            try:
                return int(float(value))
            except Exception:
                return 0
