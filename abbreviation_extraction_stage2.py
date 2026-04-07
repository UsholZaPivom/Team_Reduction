from __future__ import annotations

"""
abbreviation_extraction_stage2.py

Этап 2 проекта:
вычленение всех доступных к сокращению словоформ
и уже имеющихся в документе аббревиатур.

Что улучшено в этой версии:
1. Полная форма около скобок очищается точнее:
   - вместо захвата лишнего левого контекста
     берётся ближайшая терминная группа перед / после аббревиатуры.
2. Дубли standalone уменьшаются:
   - если в том же предложении уже найдено объявление вида
     "полная форма (АББР)" или "АББР (полная форма)",
     отдельный standalone для той же аббревиатуры не добавляется.
3. Если аббревиатура объявлена с полной формой, но этой полной формы
   нет в списке reducible_terms, она автоматически добавляется.

Что НЕ делает модуль:
- не решает, нужно ли вводить новую аббревиатуру;
- не изменяет текст документа;
- не выполняет замену расшифровок на аббревиатуры.
"""

# ============================================================
# 1. Импорт библиотек
# ============================================================

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import regex
from rapidfuzz import fuzz
from natasha import Doc, Segmenter
from pymorphy2 import MorphAnalyzer

# Модуль первого этапа должен лежать рядом с этим файлом.
from text_recognition_candidates_v3 import ReducibleWordformRecognizerV3


# ============================================================
# 2. Структуры данных
# ============================================================

@dataclass
class TextFragment:
    """
    Один текстовый фрагмент, извлечённый из документа.
    """
    source_type: str
    source_index: int
    text: str


@dataclass
class FoundAbbreviation:
    """
    Одно найденное сокращение / аббревиатура.
    """
    abbreviation: str
    long_form: str
    detection_type: str
    source_type: str
    source_index: int
    sentence: str
    matched_term: str
    match_score: float


# ============================================================
# 3. Извлечение текста из документа
# ============================================================

class DocumentTextExtractor:
    """
    Извлекает текст из .docx-документа:
    - заголовки,
    - абзацы,
    - таблицы,
    - сноски (если есть в word/footnotes.xml).
    """

    def extract_fragments(self, docx_path: str | Path) -> List[TextFragment]:
        from docx import Document

        docx_path = Path(docx_path)
        doc = Document(docx_path)

        fragments: List[TextFragment] = []

        paragraph_index = 0
        heading_index = 0
        table_cell_index = 0
        footnote_index = 0

        # -------- Заголовки и абзацы --------
        for paragraph in doc.paragraphs:
            text = self._clean_text(paragraph.text)
            if not text:
                continue

            style_name = ""
            if paragraph.style is not None and paragraph.style.name:
                style_name = paragraph.style.name.lower()

            if "heading" in style_name or "заголов" in style_name:
                fragments.append(
                    TextFragment("heading", heading_index, text)
                )
                heading_index += 1
            else:
                fragments.append(
                    TextFragment("paragraph", paragraph_index, text)
                )
                paragraph_index += 1

        # -------- Таблицы --------
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = self._clean_text(cell.text)
                    if not text:
                        continue

                    fragments.append(
                        TextFragment("table_cell", table_cell_index, text)
                    )
                    table_cell_index += 1

        # -------- Сноски --------
        footnotes = self._extract_footnotes_from_docx(docx_path)
        for text in footnotes:
            cleaned = self._clean_text(text)
            if not cleaned:
                continue

            fragments.append(
                TextFragment("footnote", footnote_index, cleaned)
            )
            footnote_index += 1

        return fragments

    def _clean_text(self, text: str) -> str:
        """
        Убирает лишние пробелы и переводы строк.
        """
        if not text:
            return ""

        text = text.replace("\n", " ")
        text = regex.sub(r"\s+", " ", text)
        return text.strip()

    def _extract_footnotes_from_docx(self, docx_path: Path) -> List[str]:
        """
        Пытается прочитать word/footnotes.xml из .docx.
        """
        result: List[str] = []

        try:
            with zipfile.ZipFile(docx_path, "r") as archive:
                if "word/footnotes.xml" not in archive.namelist():
                    return result

                xml_bytes = archive.read("word/footnotes.xml")
                root = ET.fromstring(xml_bytes)

                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

                for footnote in root.findall(".//w:footnote", ns):
                    texts = []
                    for node in footnote.findall(".//w:t", ns):
                        if node.text:
                            texts.append(node.text)

                    footnote_text = " ".join(texts).strip()
                    if footnote_text:
                        result.append(footnote_text)

        except Exception:
            return []

        return result


# ============================================================
# 4. Поиск имеющихся аббревиатур
# ============================================================

class ExistingAbbreviationExtractor:
    """
    Ищет уже существующие аббревиатуры в тексте.

    Поддерживаемые случаи:
    1. полная форма (АББР)
    2. АББР (полная форма)
    3. standalone-аббревиатуры в верхнем регистре

    В этой версии улучшено:
    - long_form очищается до ближайшей терминной группы;
    - standalone-дубли по тому же предложению подавляются.
    """

    def __init__(self) -> None:
        self.segmenter = Segmenter()
        self.morph = MorphAnalyzer()

        # Шаблон токена-аббревиатуры
        self.abbr_pattern = r"[A-ZА-ЯЁ][A-ZА-ЯЁ0-9-]{1,9}"

        # Более широкий шаблон длинной левой/правой части;
        # потом он будет дополнительно очищен эвристикой.
        self.long_form_word = r"[A-Za-zА-Яа-яЁё-]{3,}"
        self.long_form_chunk = rf"(?:{self.long_form_word}(?:\s+{self.long_form_word}){{1,11}})"

        self.pattern_long_first = regex.compile(
            rf"(?P<long>{self.long_form_chunk})\s*\((?P<abbr>{self.abbr_pattern})\)",
            flags=regex.IGNORECASE
        )

        self.pattern_abbr_first = regex.compile(
            rf"(?P<abbr>{self.abbr_pattern})\s*\((?P<long>{self.long_form_chunk})\)",
            flags=regex.IGNORECASE
        )

        self.pattern_standalone_abbr = regex.compile(
            rf"\b{self.abbr_pattern}\b"
        )

        self.false_positive_abbr = {
            "РФ", "РТФ", "MS", "WORD", "PDF", "DOCX"
        }

        # Части речи, которые чаще входят в терминные группы
        self.allowed_pos = {"NOUN", "ADJF", "ADJS", "PRTF", "PRTS"}

        # Слова-паразиты контекста, которые часто "прилипают" к long_form
        self.context_lemmas = {
            "в", "во", "на", "при", "по", "внутри", "внутренний",
            "сноска", "дополнительно", "указанный", "указать",
            "подсистема", "контроль", "использоваться",
            "отчёт", "документ", "пример", "тестирование", "проверка",
            "данный", "этот", "тот"
        }

        self.word_pattern = regex.compile(r"[A-Za-zА-Яа-яЁё-]+")

    def split_into_sentences(self, text: str) -> List[str]:
        if not text.strip():
            return []

        doc = Doc(text)
        doc.segment(self.segmenter)

        result: List[str] = []
        for sent in doc.sents:
            sentence = sent.text.strip()
            if sentence:
                result.append(sentence)

        return result

    def extract_from_fragments(self, fragments: List[TextFragment]) -> List[FoundAbbreviation]:
        """
        Ищет все аббревиатуры во всех фрагментах.
        """
        found: List[FoundAbbreviation] = []

        for fragment in fragments:
            sentences = self.split_into_sentences(fragment.text)

            for sentence in sentences:
                declared_items = self._extract_declared_abbreviations(
                    sentence=sentence,
                    source_type=fragment.source_type,
                    source_index=fragment.source_index
                )

                found.extend(declared_items)

                # Собираем набор уже объявленных в этом же предложении сокращений,
                # чтобы не дублировать их как standalone.
                declared_abbreviations_in_sentence = {
                    item.abbreviation for item in declared_items
                }

                found.extend(
                    self._extract_standalone_abbreviations(
                        sentence=sentence,
                        source_type=fragment.source_type,
                        source_index=fragment.source_index,
                        declared_in_same_sentence=declared_abbreviations_in_sentence
                    )
                )

        return self._deduplicate(found)

    def _extract_declared_abbreviations(
        self,
        sentence: str,
        source_type: str,
        source_index: int
    ) -> List[FoundAbbreviation]:
        """
        Ищет:
        - полная форма (АББР)
        - АББР (полная форма)
        """
        result: List[FoundAbbreviation] = []

        # -------- Вариант 1: полная форма (АББР) --------
        for match in self.pattern_long_first.finditer(sentence):
            raw_long_form = match.group("long")
            abbreviation = match.group("abbr").upper()

            if not self._is_valid_abbreviation(abbreviation):
                continue

            long_form = self._shrink_long_form(raw_long_form, direction="right")

            result.append(
                FoundAbbreviation(
                    abbreviation=abbreviation,
                    long_form=long_form,
                    detection_type="declared_long_first",
                    source_type=source_type,
                    source_index=source_index,
                    sentence=sentence,
                    matched_term="",
                    match_score=0.0
                )
            )

        # -------- Вариант 2: АББР (полная форма) --------
        for match in self.pattern_abbr_first.finditer(sentence):
            raw_long_form = match.group("long")
            abbreviation = match.group("abbr").upper()

            if not self._is_valid_abbreviation(abbreviation):
                continue

            long_form = self._shrink_long_form(raw_long_form, direction="left")

            result.append(
                FoundAbbreviation(
                    abbreviation=abbreviation,
                    long_form=long_form,
                    detection_type="declared_abbr_first",
                    source_type=source_type,
                    source_index=source_index,
                    sentence=sentence,
                    matched_term="",
                    match_score=0.0
                )
            )

        return result

    def _extract_standalone_abbreviations(
        self,
        sentence: str,
        source_type: str,
        source_index: int,
        declared_in_same_sentence: Set[str]
    ) -> List[FoundAbbreviation]:
        """
        Ищет отдельные standalone-аббревиатуры.
        Если та же аббревиатура уже была объявлена в этом предложении,
        standalone не добавляем.
        """
        result: List[FoundAbbreviation] = []

        for match in self.pattern_standalone_abbr.finditer(sentence):
            abbreviation = match.group(0).upper()

            if not self._is_valid_abbreviation(abbreviation):
                continue

            if abbreviation in declared_in_same_sentence:
                continue

            result.append(
                FoundAbbreviation(
                    abbreviation=abbreviation,
                    long_form="",
                    detection_type="standalone",
                    source_type=source_type,
                    source_index=source_index,
                    sentence=sentence,
                    matched_term="",
                    match_score=0.0
                )
            )

        return result

    def _clean_long_form(self, text: str) -> str:
        """
        Базовая очистка полной формы.
        """
        text = regex.sub(r"\s+", " ", text).strip()
        text = text.strip(" ,.;:()[]{}\"'«»")
        return text

    def _extract_words(self, text: str) -> List[str]:
        """
        Выделяет слова из текста.
        """
        return self.word_pattern.findall(text)

    def _parse_word(self, word: str):
        return self.morph.parse(word)[0]

    def _get_pos(self, word: str) -> str:
        pos = self._parse_word(word).tag.POS
        return pos if pos else ""

    def _get_lemma(self, word: str) -> str:
        return self._parse_word(word).normal_form

    def _is_content_word(self, word: str) -> bool:
        """
        Содержательное слово для терминной группы.
        """
        if len(word) < 3:
            return False

        lemma = self._get_lemma(word)
        pos = self._get_pos(word)

        if lemma in self.context_lemmas:
            return False

        if pos not in self.allowed_pos:
            return False

        return True

    def _is_term_like_words(self, words: List[str]) -> bool:
        """
        Проверяет, похожа ли последовательность слов на терминную группу.
        """
        if len(words) < 2 or len(words) > 6:
            return False

        pos_list = [self._get_pos(word) for word in words]
        noun_count = sum(1 for pos in pos_list if pos == "NOUN")
        modifier_count = sum(1 for pos in pos_list if pos in {"ADJF", "ADJS", "PRTF", "PRTS"})

        # Последнее слово желательно существительное
        if pos_list[-1] != "NOUN":
            return False

        if not (noun_count >= 2 or (noun_count >= 1 and modifier_count >= 1)):
            return False

        return True

    def _shrink_long_form(self, raw_text: str, direction: str) -> str:
        """
        Обрезает long_form до ближайшей терминной группы.

        direction:
        - "right"  -> берём правый конец (для случая "полная форма (АББР)")
        - "left"   -> берём левый конец  (для случая "АББР (полная форма)")

        Пример:
        "подсистеме контроля используется система управления доступом"
        -> "система управления доступом"

        "сноске дополнительно указан центр оперативной обработки сообщений"
        -> "центр оперативной обработки сообщений"
        """
        text = self._clean_long_form(raw_text)
        words = self._extract_words(text)

        if len(words) < 2:
            return text

        best_candidate = None

        # Перебираем окна 2..6 слов.
        # Для direction="right" предпочитаем окна ближе к правому краю.
        # Для direction="left"  предпочитаем окна ближе к левому краю.
        max_window = min(6, len(words))

        for window_size in range(max_window, 1, -1):
            for start in range(0, len(words) - window_size + 1):
                end = start + window_size
                chunk = words[start:end]

                # Все слова должны быть содержательными
                if not all(self._is_content_word(w) for w in chunk):
                    continue

                if not self._is_term_like_words(chunk):
                    continue

                candidate_text = " ".join(chunk)

                if direction == "right":
                    # Чем ближе окно к правому краю, тем лучше
                    score = end
                else:
                    # Чем ближе окно к левому краю, тем лучше
                    score = -start

                candidate = (window_size, score, candidate_text)

                if best_candidate is None or candidate > best_candidate:
                    best_candidate = candidate

        if best_candidate:
            return best_candidate[2]

        # Если терминная группа не нашлась, пытаемся хотя бы отрезать
        # длинный контекст с краёв по содержательным словам.
        filtered_words = [w for w in words if self._is_content_word(w)]
        if len(filtered_words) >= 2:
            return " ".join(filtered_words[:6])

        return text

    def _is_valid_abbreviation(self, abbreviation: str) -> bool:
        """
        Проверяет, похоже ли сокращение на настоящую аббревиатуру.
        """
        if len(abbreviation) < 2 or len(abbreviation) > 10:
            return False

        if abbreviation in self.false_positive_abbr:
            return False

        letters_only = regex.sub(r"[^A-ZА-ЯЁ]", "", abbreviation)
        if len(letters_only) < 2:
            return False

        return True

    def _deduplicate(self, items: List[FoundAbbreviation]) -> List[FoundAbbreviation]:
        """
        Убирает повторы найденных аббревиатур.
        """
        seen = set()
        result: List[FoundAbbreviation] = []

        for item in items:
            key = (
                item.abbreviation,
                item.long_form,
                item.detection_type,
                item.source_type,
                item.source_index,
                item.sentence
            )
            if key in seen:
                continue

            seen.add(key)
            result.append(item)

        return result


# ============================================================
# 5. Основной модуль второго этапа
# ============================================================

class Stage2ReductionAnalyzer:
    """
    Реализует вторую задачу:
    - выделяет словоформы, доступные к сокращению;
    - выделяет уже существующие аббревиатуры;
    - сопоставляет их;
    - при необходимости расширяет список терминов за счёт
      объявленных полных форм.
    """

    def __init__(self) -> None:
        self.stage1_recognizer = ReducibleWordformRecognizerV3()
        self.text_extractor = DocumentTextExtractor()
        self.abbreviation_extractor = ExistingAbbreviationExtractor()

    def run(self, docx_path: str | Path, output_dir: str | Path) -> Dict[str, Path]:
        docx_path = Path(docx_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        # ----------------------------------------------------
        # Шаг 1. Термины из первого этапа
        # ----------------------------------------------------
        mentions = self.stage1_recognizer.analyze_document(docx_path)
        reducible_terms_df = self.stage1_recognizer.aggregate_mentions(mentions).copy()

        reducible_terms_df = reducible_terms_df.rename(columns={
            "phrase_example": "term",
            "normalized_phrase": "normalized_term",
            "proposed_abbreviation": "suggested_abbreviation"
        })

        # ----------------------------------------------------
        # Шаг 2. Все текстовые фрагменты документа
        # ----------------------------------------------------
        fragments = self.text_extractor.extract_fragments(docx_path)
        fragments_df = pd.DataFrame([asdict(f) for f in fragments])

        # ----------------------------------------------------
        # Шаг 3. Поиск уже существующих аббревиатур
        # ----------------------------------------------------
        found_abbreviations = self.abbreviation_extractor.extract_from_fragments(fragments)

        # ----------------------------------------------------
        # Шаг 4. Расширяем список reducible_terms объявленными формами
        # ----------------------------------------------------
        reducible_terms_df = self._extend_reducible_terms_with_declared_forms(
            reducible_terms_df=reducible_terms_df,
            found_abbreviations=found_abbreviations
        )

        # ----------------------------------------------------
        # Шаг 5. Сопоставляем аббревиатуры и термины
        # ----------------------------------------------------
        matched_abbreviations = self._match_abbreviations_to_terms(
            found_abbreviations=found_abbreviations,
            reducible_terms_df=reducible_terms_df
        )

        existing_abbreviations_df = pd.DataFrame([asdict(item) for item in matched_abbreviations])

        # ----------------------------------------------------
        # Шаг 6. Формируем сводную таблицу
        # ----------------------------------------------------
        merged_df = self._build_merged_table(
            reducible_terms_df=reducible_terms_df,
            existing_abbreviations_df=existing_abbreviations_df
        )

        # ----------------------------------------------------
        # Шаг 7. Сохраняем результаты
        # ----------------------------------------------------
        saved_files = self._save_results(
            fragments_df=fragments_df,
            reducible_terms_df=reducible_terms_df,
            existing_abbreviations_df=existing_abbreviations_df,
            merged_df=merged_df,
            output_dir=output_dir
        )

        return saved_files

    def _normalize_term_for_compare(self, term: str) -> str:
        """
        Нормализует термин для сравнения:
        lower + схлопывание пробелов.
        """
        term = regex.sub(r"\s+", " ", str(term)).strip().lower()
        return term

    def _build_abbreviation_from_phrase(self, phrase: str) -> str:
        """
        Строит аббревиатуру по первым буквам слов.
        """
        words = regex.findall(r"[A-Za-zА-Яа-яЁё-]+", phrase)
        letters = []
        for word in words:
            if word:
                letters.append(word[0].upper())
        return "".join(letters)

    def _extend_reducible_terms_with_declared_forms(
        self,
        reducible_terms_df: pd.DataFrame,
        found_abbreviations: List[FoundAbbreviation]
    ) -> pd.DataFrame:
        """
        Если в документе объявлена аббревиатура с полной формой,
        а такой полной формы нет среди reducible_terms,
        добавляем её в список терминов.

        Это нужно, например, для случаев вроде:
        ЦООС из сноски, когда аббревиатура и полная форма уже есть в документе,
        но первый этап термин не выделил.
        """
        rows = reducible_terms_df.to_dict("records") if not reducible_terms_df.empty else []

        existing_terms_normalized = {
            self._normalize_term_for_compare(str(row.get("term", "")))
            for row in rows
        }

        additions = []

        for item in found_abbreviations:
            if not item.long_form:
                continue

            normalized_long_form = self._normalize_term_for_compare(item.long_form)

            if normalized_long_form in existing_terms_normalized:
                continue

            suggested_abbreviation = self._build_abbreviation_from_phrase(item.long_form)
            word_count = len(regex.findall(r"[A-Za-zА-Яа-яЁё-]+", item.long_form))

            additions.append({
                "term": item.long_form,
                "normalized_term": normalized_long_form,
                "suggested_abbreviation": suggested_abbreviation,
                "word_count": word_count,
                "frequency": 1,
                "examples": item.sentence
            })

            existing_terms_normalized.add(normalized_long_form)

        if additions:
            reducible_terms_df = pd.concat(
                [reducible_terms_df, pd.DataFrame(additions)],
                ignore_index=True
            )

        reducible_terms_df = reducible_terms_df.sort_values(
            by=["frequency", "word_count", "term"],
            ascending=[False, False, True]
        ).reset_index(drop=True)

        return reducible_terms_df

    def _match_abbreviations_to_terms(
        self,
        found_abbreviations: List[FoundAbbreviation],
        reducible_terms_df: pd.DataFrame
    ) -> List[FoundAbbreviation]:
        """
        Сопоставляет найденные аббревиатуры с найденными словоформами.

        Логика:
        1. Точное совпадение abbreviation == suggested_abbreviation.
        2. Если есть long_form, сравнение с term и normalized_term через rapidfuzz.
        3. Лучшее совпадение с score >= 70 считается валидным.
        """
        if reducible_terms_df.empty:
            return found_abbreviations

        terms = reducible_terms_df.to_dict("records")
        matched: List[FoundAbbreviation] = []

        for item in found_abbreviations:
            best_term = ""
            best_score = 0.0

            abbreviation = item.abbreviation.upper()
            long_form = item.long_form.strip().lower()

            for term_row in terms:
                term = str(term_row.get("term", ""))
                normalized_term = str(term_row.get("normalized_term", ""))
                suggested_abbreviation = str(term_row.get("suggested_abbreviation", "")).upper()

                score = 0.0

                # 1. Совпадение по самой аббревиатуре
                if abbreviation and abbreviation == suggested_abbreviation:
                    score = max(score, 100.0)

                # 2. Совпадение по полной форме
                if long_form:
                    score = max(score, float(fuzz.ratio(long_form, term.lower())))
                    score = max(score, float(fuzz.ratio(long_form, normalized_term.lower())))

                if score > best_score:
                    best_score = score
                    best_term = term

            matched.append(
                FoundAbbreviation(
                    abbreviation=item.abbreviation,
                    long_form=item.long_form,
                    detection_type=item.detection_type,
                    source_type=item.source_type,
                    source_index=item.source_index,
                    sentence=item.sentence,
                    matched_term=best_term if best_score >= 70 else "",
                    match_score=best_score
                )
            )

        return matched

    def _build_merged_table(
        self,
        reducible_terms_df: pd.DataFrame,
        existing_abbreviations_df: pd.DataFrame
    ) -> pd.DataFrame:
        """
        Формирует сводную таблицу:
        - найденная словоформа,
        - предложенная аббревиатура,
        - есть ли такая аббревиатура в тексте,
        - какая именно аббревиатура найдена,
        - score сопоставления.
        """
        if reducible_terms_df.empty:
            return pd.DataFrame(columns=[
                "term",
                "normalized_term",
                "suggested_abbreviation",
                "frequency",
                "abbreviation_found_in_text",
                "found_abbreviation",
                "match_score"
            ])

        existing_records = existing_abbreviations_df.to_dict("records") if not existing_abbreviations_df.empty else []

        rows = []
        for term_row in reducible_terms_df.to_dict("records"):
            term = str(term_row["term"])
            suggested_abbreviation = str(term_row["suggested_abbreviation"]).upper()

            matched_items = []
            for abbr_row in existing_records:
                found_abbreviation = str(abbr_row.get("abbreviation", "")).upper()
                matched_term = str(abbr_row.get("matched_term", ""))

                if found_abbreviation == suggested_abbreviation or matched_term == term:
                    matched_items.append(abbr_row)

            if matched_items:
                best_match = max(matched_items, key=lambda x: float(x.get("match_score", 0)))
                rows.append({
                    "term": term,
                    "normalized_term": term_row["normalized_term"],
                    "suggested_abbreviation": suggested_abbreviation,
                    "word_count": term_row["word_count"],
                    "frequency": term_row["frequency"],
                    "abbreviation_found_in_text": True,
                    "found_abbreviation": best_match.get("abbreviation", ""),
                    "match_score": best_match.get("match_score", 0),
                    "examples": term_row["examples"]
                })
            else:
                rows.append({
                    "term": term,
                    "normalized_term": term_row["normalized_term"],
                    "suggested_abbreviation": suggested_abbreviation,
                    "word_count": term_row["word_count"],
                    "frequency": term_row["frequency"],
                    "abbreviation_found_in_text": False,
                    "found_abbreviation": "",
                    "match_score": 0,
                    "examples": term_row["examples"]
                })

        df = pd.DataFrame(rows)
        df = df.sort_values(
            by=["abbreviation_found_in_text", "frequency", "word_count", "term"],
            ascending=[False, False, False, True]
        ).reset_index(drop=True)

        return df

    def _save_results(
        self,
        fragments_df: pd.DataFrame,
        reducible_terms_df: pd.DataFrame,
        existing_abbreviations_df: pd.DataFrame,
        merged_df: pd.DataFrame,
        output_dir: Path
    ) -> Dict[str, Path]:
        """
        Сохраняет таблицы в CSV и XLSX.
        """
        saved_files: Dict[str, Path] = {}

        fragments_csv = output_dir / "document_fragments.csv"
        fragments_xlsx = output_dir / "document_fragments.xlsx"

        reducible_csv = output_dir / "reducible_terms.csv"
        reducible_xlsx = output_dir / "reducible_terms.xlsx"

        existing_csv = output_dir / "existing_abbreviations.csv"
        existing_xlsx = output_dir / "existing_abbreviations.xlsx"

        merged_csv = output_dir / "merged_terms_and_abbreviations.csv"
        merged_xlsx = output_dir / "merged_terms_and_abbreviations.xlsx"

        fragments_df.to_csv(fragments_csv, index=False, encoding="utf-8-sig")
        reducible_terms_df.to_csv(reducible_csv, index=False, encoding="utf-8-sig")
        existing_abbreviations_df.to_csv(existing_csv, index=False, encoding="utf-8-sig")
        merged_df.to_csv(merged_csv, index=False, encoding="utf-8-sig")

        saved_files["document_fragments_csv"] = fragments_csv
        saved_files["reducible_terms_csv"] = reducible_csv
        saved_files["existing_abbreviations_csv"] = existing_csv
        saved_files["merged_csv"] = merged_csv

        try:
            fragments_df.to_excel(fragments_xlsx, index=False)
            reducible_terms_df.to_excel(reducible_xlsx, index=False)
            existing_abbreviations_df.to_excel(existing_xlsx, index=False)
            merged_df.to_excel(merged_xlsx, index=False)

            saved_files["document_fragments_xlsx"] = fragments_xlsx
            saved_files["reducible_terms_xlsx"] = reducible_xlsx
            saved_files["existing_abbreviations_xlsx"] = existing_xlsx
            saved_files["merged_xlsx"] = merged_xlsx

        except ModuleNotFoundError as exc:
            print("Внимание: не удалось сохранить XLSX-файлы.")
            print("Причина:", exc)
            print("CSV-файлы при этом успешно сохранены.")

        return saved_files


# ============================================================
# 6. Локальный запуск
# ============================================================

if __name__ == "__main__":
    """
    Пример запуска.

    Перед запуском:
    1. Убедитесь, что рядом лежит text_recognition_candidates_v3.py
    2. Укажите имя входного .docx в input_docx
    3. Запустите файл
    """

    input_docx = "test_reduction_input.docx"
    output_dir = "result_stage2"

    analyzer = Stage2ReductionAnalyzer()
    saved_files = analyzer.run(input_docx, output_dir)

    print("=" * 60)
    print("ЭТАП 2 ЗАВЕРШЁН: выделение словоформ и аббревиатур")
    print("=" * 60)
    print("Сформированы файлы:")

    for name, path in saved_files.items():
        print(f"{name}: {path}")
