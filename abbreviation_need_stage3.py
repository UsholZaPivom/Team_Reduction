from __future__ import annotations

"""
abbreviation_need_stage3.py

Этап 3 проекта:
определение необходимости ввода аббревиатуры.

Назначение модуля:
1. Использовать результаты второго этапа:
   - список словоформ, доступных к сокращению,
   - список уже имеющихся аббревиатур,
   - сводную таблицу "термин -> аббревиатура найдена / не найдена".
2. Для каждого термина определить:
   - нужно ли вводить аббревиатуру;
   - насколько это целесообразно;
   - почему принято такое решение.
3. Сформировать итоговую таблицу рекомендаций.

Что понимается под "необходимостью ввода аббревиатуры":
- если термин уже имеет объявленную аббревиатуру в документе,
  то вводить новую не нужно;
- если аббревиатуры ещё нет, но термин длинный и/или повторяется,
  то её введение рекомендуется;
- если термин короткий и встречается редко, аббревиатура обычно не нужна.

Этот модуль НЕ меняет сам документ.
Он только выносит аналитическое решение для следующего шага.
"""

# ============================================================
# 1. Импорт библиотек
# ============================================================

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List

import pandas as pd
import regex

# Используем модуль второго этапа.
# Важно: файл abbreviation_extraction_stage2.py должен лежать рядом.
from abbreviation_extraction_stage2 import Stage2ReductionAnalyzer


# ============================================================
# 2. Структура итогового решения
# ============================================================

@dataclass
class AbbreviationDecision:
    """
    Итоговое решение по одному термину.

    Поля:
    - term: полная форма термина;
    - suggested_abbreviation: предлагаемая аббревиатура;
    - abbreviation_found_in_text: уже есть ли эта аббревиатура в документе;
    - found_abbreviation: какая именно аббревиатура найдена в тексте;
    - frequency: сколько раз термин был замечен на предыдущих этапах;
    - word_count: количество слов в термине;
    - char_length: длина термина в символах;
    - decision_score: числовая оценка целесообразности ввода аббревиатуры;
    - need_to_introduce: итоговое решение True/False;
    - priority: приоритет рекомендации (high/medium/low/none);
    - reason: текстовое объяснение решения.
    """
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


# ============================================================
# 3. Основной класс 3-го этапа
# ============================================================

class AbbreviationNeedAnalyzer:
    """
    Класс реализует 3-ю задачу куратора:
    определить необходимость ввода аббревиатуры.

    Источник данных:
    - результаты второго этапа (Stage2ReductionAnalyzer).

    Общая логика принятия решения:
    1. Если аббревиатура уже есть в тексте, новую вводить не нужно.
    2. Если аббревиатуры нет, рассчитывается score:
       - чем чаще встречается термин, тем выше score;
       - чем длиннее термин и чем больше в нём слов, тем выше score;
       - если предлагаемая аббревиатура выглядит удачной, score растёт.
    3. По score выносится итог:
       - high / medium / low priority;
       - need_to_introduce = True/False.
    """

    def __init__(self) -> None:
        """
        Инициализируем модуль второго этапа.
        """
        self.stage2_analyzer = Stage2ReductionAnalyzer()

    # ========================================================
    # 4. Публичный метод запуска
    # ========================================================

    def run(self, docx_path: str | Path, output_dir: str | Path) -> Dict[str, Path]:
        """
        Полный запуск третьего этапа.

        Что делает:
        1. Запускает второй этап и получает таблицы:
           - reducible_terms,
           - existing_abbreviations,
           - merged_terms_and_abbreviations.
        2. На основе merged-таблицы определяет необходимость ввода аббревиатуры.
        3. Сохраняет таблицы с решениями.

        Возвращает словарь с путями к созданным файлам.
        """
        docx_path = Path(docx_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        # ----------------------------------------------------
        # Шаг 1. Запускаем второй этап и получаем его файлы
        # ----------------------------------------------------
        stage2_output_dir = output_dir / "stage2_intermediate"
        stage2_saved = self.stage2_analyzer.run(docx_path, stage2_output_dir)

        # Читаем сводную таблицу второго этапа.
        merged_csv = stage2_saved["merged_csv"]
        merged_df = pd.read_csv(merged_csv)

        # Также читаем таблицу уже найденных аббревиатур —
        # это полезно для расширенного анализа и отладки.
        existing_csv = stage2_saved["existing_abbreviations_csv"]
        existing_abbreviations_df = pd.read_csv(existing_csv)

        # ----------------------------------------------------
        # Шаг 2. Формируем решения по каждому термину
        # ----------------------------------------------------
        decisions = self._build_decisions(merged_df)
        decisions_df = pd.DataFrame([asdict(item) for item in decisions])

        # ----------------------------------------------------
        # Шаг 3. Дополнительная краткая таблица рекомендаций
        # ----------------------------------------------------
        recommendations_df = self._build_recommendations_table(decisions_df)

        # ----------------------------------------------------
        # Шаг 4. Сохраняем результат
        # ----------------------------------------------------
        saved_files = self._save_results(
            merged_df=merged_df,
            existing_abbreviations_df=existing_abbreviations_df,
            decisions_df=decisions_df,
            recommendations_df=recommendations_df,
            output_dir=output_dir
        )

        return saved_files

    # ========================================================
    # 5. Построение решений
    # ========================================================

    def _build_decisions(self, merged_df: pd.DataFrame) -> List[AbbreviationDecision]:
        """
        Преобразует строки merged-таблицы в набор решений.

        Для каждого термина:
        - считаем score;
        - определяем, нужно ли вводить аббревиатуру;
        - формируем объяснение.
        """
        decisions: List[AbbreviationDecision] = []

        if merged_df.empty:
            return decisions

        for row in merged_df.to_dict("records"):
            term = str(row.get("term", "")).strip()
            suggested_abbreviation = str(row.get("suggested_abbreviation", "")).strip().upper()
            abbreviation_found_in_text = bool(row.get("abbreviation_found_in_text", False))
            found_abbreviation = str(row.get("found_abbreviation", "")).strip().upper()
            frequency = int(row.get("frequency", 0))
            word_count = int(row.get("word_count", 0))
            char_length = len(term)

            score, reason, priority, need_to_introduce = self._evaluate_term(
                term=term,
                suggested_abbreviation=suggested_abbreviation,
                abbreviation_found_in_text=abbreviation_found_in_text,
                frequency=frequency,
                word_count=word_count,
                char_length=char_length
            )

            decisions.append(
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
                    reason=reason
                )
            )

        # Сортируем так, чтобы сначала шли термины,
        # для которых точно рекомендуется вводить аббревиатуру.
        decisions.sort(
            key=lambda item: (
                item.need_to_introduce,
                item.decision_score,
                item.frequency,
                item.word_count,
                item.term
            ),
            reverse=True
        )

        return decisions

    def _evaluate_term(
        self,
        term: str,
        suggested_abbreviation: str,
        abbreviation_found_in_text: bool,
        frequency: int,
        word_count: int,
        char_length: int
    ) -> tuple[int, str, str, bool]:
        """
        Главная эвристика принятия решения.

        Логика:
        1. Если аббревиатура уже есть в тексте, новую вводить не нужно.
        2. Иначе считаем decision_score по нескольким критериям:
           - частота использования термина;
           - длина термина в словах;
           - длина термина в символах;
           - качество предлагаемой аббревиатуры.
        3. По итоговому score определяем приоритет и итоговое решение.
        """
        # ----------------------------------------------------
        # Случай 1. Аббревиатура уже есть в документе
        # ----------------------------------------------------
        if abbreviation_found_in_text:
            return (
                0,
                "Аббревиатура уже присутствует в тексте, дополнительный ввод не требуется.",
                "none",
                False
            )

        # ----------------------------------------------------
        # Случай 2. Аббревиатуры нет -> оцениваем целесообразность
        # ----------------------------------------------------
        score = 0
        reasons: List[str] = []

        # --- Критерий 1. Частота ---
        # Чем чаще встречается термин, тем полезнее его сократить.
        if frequency >= 4:
            score += 40
            reasons.append("термин часто встречается в документе")
        elif frequency == 3:
            score += 30
            reasons.append("термин повторяется несколько раз")
        elif frequency == 2:
            score += 20
            reasons.append("термин встречается более одного раза")
        else:
            score += 5
            reasons.append("термин встречается редко")

        # --- Критерий 2. Количество слов ---
        # Длинные многословные конструкции сильнее выигрывают от сокращения.
        if word_count >= 5:
            score += 35
            reasons.append("термин состоит из 5 и более слов")
        elif word_count == 4:
            score += 25
            reasons.append("термин состоит из 4 слов")
        elif word_count == 3:
            score += 15
            reasons.append("термин состоит из 3 слов")
        elif word_count == 2:
            score += 5
            reasons.append("термин состоит из 2 слов")

        # --- Критерий 3. Длина термина в символах ---
        # Чем длиннее полная форма, тем заметнее польза аббревиатуры.
        if char_length >= 40:
            score += 20
            reasons.append("полная форма очень длинная")
        elif char_length >= 25:
            score += 12
            reasons.append("полная форма достаточно длинная")
        elif char_length >= 18:
            score += 6
            reasons.append("полная форма средней длины")

        # --- Критерий 4. Качество предлагаемой аббревиатуры ---
        # Если аббревиатура получается компактной и читаемой,
        # её ввод обычно удобнее.
        abbr_quality_score, abbr_reason = self._evaluate_suggested_abbreviation(suggested_abbreviation)
        score += abbr_quality_score
        if abbr_reason:
            reasons.append(abbr_reason)

        # ----------------------------------------------------
        # Итоговое решение по score
        # ----------------------------------------------------
        if score >= 65:
            priority = "high"
            need_to_introduce = True
            reasons.insert(0, "рекомендуется ввести аббревиатуру")
        elif score >= 45:
            priority = "medium"
            need_to_introduce = True
            reasons.insert(0, "ввод аббревиатуры целесообразен")
        elif score >= 30:
            priority = "low"
            need_to_introduce = False
            reasons.insert(0, "ввод аббревиатуры возможен, но не обязателен")
        else:
            priority = "low"
            need_to_introduce = False
            reasons.insert(0, "ввод аббревиатуры не требуется")

        reason_text = "; ".join(reasons)
        return score, reason_text, priority, need_to_introduce

    def _evaluate_suggested_abbreviation(self, abbreviation: str) -> tuple[int, str]:
        """
        Оценивает качество автоматически предложенной аббревиатуры.

        Идея:
        - слишком короткая аббревиатура малоинформативна;
        - слишком длинная аббревиатура неудобна;
        - оптимальна длина 3..6 символов.
        """
        abbreviation = abbreviation.strip().upper()

        if not abbreviation:
            return 0, "не удалось построить надёжную аббревиатуру"

        # Оставляем только буквы / цифры для оценки длины
        normalized = regex.sub(r"[^A-ZА-ЯЁ0-9]", "", abbreviation)
        length = len(normalized)

        if length < 2:
            return -5, "предлагаемая аббревиатура слишком короткая"

        if 3 <= length <= 6:
            return 10, "предлагаемая аббревиатура компактная и удобная"
        elif length in (2, 7):
            return 5, "предлагаемая аббревиатура допустима по длине"
        else:
            return -5, "предлагаемая аббревиатура слишком длинная"

    # ========================================================
    # 6. Подготовка компактной таблицы рекомендаций
    # ========================================================

    def _build_recommendations_table(self, decisions_df: pd.DataFrame) -> pd.DataFrame:
        """
        Формирует компактную таблицу рекомендаций для пользователя / отчёта.

        Оставляем только самые полезные поля:
        - термин,
        - предлагаемая аббревиатура,
        - уже ли есть аббревиатура в тексте,
        - нужно ли вводить,
        - приоритет,
        - краткая причина.
        """
        if decisions_df.empty:
            return pd.DataFrame(columns=[
                "term",
                "suggested_abbreviation",
                "abbreviation_found_in_text",
                "need_to_introduce",
                "priority",
                "decision_score",
                "reason"
            ])

        columns = [
            "term",
            "suggested_abbreviation",
            "abbreviation_found_in_text",
            "need_to_introduce",
            "priority",
            "decision_score",
            "reason"
        ]

        result = decisions_df[columns].copy()

        # Сортируем: сначала те, для кого надо вводить аббревиатуру.
        priority_order = {"high": 3, "medium": 2, "low": 1, "none": 0}
        result["_priority_sort"] = result["priority"].map(priority_order).fillna(0)

        result = result.sort_values(
            by=["need_to_introduce", "_priority_sort", "decision_score", "term"],
            ascending=[False, False, False, True]
        ).drop(columns=["_priority_sort"]).reset_index(drop=True)

        return result

    # ========================================================
    # 7. Сохранение результатов
    # ========================================================

    def _save_results(
        self,
        merged_df: pd.DataFrame,
        existing_abbreviations_df: pd.DataFrame,
        decisions_df: pd.DataFrame,
        recommendations_df: pd.DataFrame,
        output_dir: Path
    ) -> Dict[str, Path]:
        """
        Сохраняет таблицы в CSV и XLSX.

        Сохраняемые файлы:
        - merged_terms_and_abbreviations.csv/xlsx
        - existing_abbreviations.csv/xlsx
        - abbreviation_decisions.csv/xlsx
        - abbreviation_recommendations.csv/xlsx
        """
        saved_files: Dict[str, Path] = {}

        merged_csv = output_dir / "merged_terms_and_abbreviations.csv"
        merged_xlsx = output_dir / "merged_terms_and_abbreviations.xlsx"

        existing_csv = output_dir / "existing_abbreviations.csv"
        existing_xlsx = output_dir / "existing_abbreviations.xlsx"

        decisions_csv = output_dir / "abbreviation_decisions.csv"
        decisions_xlsx = output_dir / "abbreviation_decisions.xlsx"

        recommendations_csv = output_dir / "abbreviation_recommendations.csv"
        recommendations_xlsx = output_dir / "abbreviation_recommendations.xlsx"

        # CSV сохраняем всегда
        merged_df.to_csv(merged_csv, index=False, encoding="utf-8-sig")
        existing_abbreviations_df.to_csv(existing_csv, index=False, encoding="utf-8-sig")
        decisions_df.to_csv(decisions_csv, index=False, encoding="utf-8-sig")
        recommendations_df.to_csv(recommendations_csv, index=False, encoding="utf-8-sig")

        saved_files["merged_csv"] = merged_csv
        saved_files["existing_abbreviations_csv"] = existing_csv
        saved_files["abbreviation_decisions_csv"] = decisions_csv
        saved_files["abbreviation_recommendations_csv"] = recommendations_csv

        # XLSX — если установлен openpyxl
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


# ============================================================
# 8. Локальный запуск
# ============================================================

if __name__ == "__main__":
    """
    Пример запуска.

    Перед запуском:
    1. Убедитесь, что рядом лежит abbreviation_extraction_stage2.py
    2. Укажите имя входного документа Word
    3. Запустите файл
    """

    input_docx = "test_reduction_input.docx"
    output_dir = "result_stage3"

    analyzer = AbbreviationNeedAnalyzer()
    saved_files = analyzer.run(input_docx, output_dir)

    print("=" * 60)
    print("ЭТАП 3 ЗАВЕРШЁН: определение необходимости ввода аббревиатуры")
    print("=" * 60)
    print("Сформированы файлы:")

    for name, path in saved_files.items():
        print(f"{name}: {path}")
