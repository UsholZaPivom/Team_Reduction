from __future__ import annotations

"""
declaration_validator.py

Модуль проверки ошибочных объявлений сокращений.

Логика задачи:
ошибочным объявлением считается такое объявление сокращения,
которое встретилось в документе один раз в момент объявления
и далее по тексту больше не понадобилось.

Примеры объявлений, которые считаются "объявлениями":
- полная форма (АББР)
- полная форма (далее – АББР)
- АББР (полная форма)

На вход модулю удобно подавать файл existing_abbreviations.csv,
который формируется на этапе 2.

Что делает модуль:
1. Загружает existing_abbreviations.csv
2. Находит все объявления сокращений
3. Проверяет, встречается ли сокращение после объявления
4. Формирует:
   - таблицу всех проверенных объявлений
   - таблицу только ошибочных объявлений
   - краткую сводку

Основная идея проверки:
если после позиции объявления сокращение больше не встречается
как самостоятельное употребление или повторное объявление,
то объявление считается ошибочным.

Важно:
модуль работает без изменения документа.
Он только анализирует данные и формирует отчёт.
"""

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List

import pandas as pd


@dataclass
class DeclarationCheckResult:
    """
    Результат проверки одного объявления сокращения.
    """
    abbreviation: str
    long_form: str
    detection_type: str
    source_type: str
    source_index: int
    sentence: str
    later_mentions_count: int
    later_standalone_count: int
    later_declared_count: int
    is_erroneous_declaration: bool
    comment: str


class DeclarationValidator:
    """
    Валидатор ошибочных объявлений сокращений.

    Объявления ищутся среди типов:
    - declared_long_first
    - declared_long_dalee
    - declared_abbr_first

    Проверка:
    - если после объявления сокращение ни разу больше не встретилось,
      объявление считается ошибочным;
    - если после объявления были только технически эквивалентные записи
      самого объявления, но не было использования дальше по тексту,
      объявление также считается ошибочным.
    """

    def __init__(self) -> None:
        self.declaration_types = {
            "declared_long_first",
            "declared_long_dalee",
            "declared_abbr_first",
        }

    # -----------------------------------------------------------------
    # Загрузка и подготовка данных
    # -----------------------------------------------------------------

    def load_existing_abbreviations(self, csv_path: str | Path) -> pd.DataFrame:
        """
        Загружает existing_abbreviations.csv.
        """
        csv_path = Path(csv_path)
        if not csv_path.exists():
            raise FileNotFoundError(f"Файл не найден: {csv_path}")

        df = pd.read_csv(csv_path, encoding="utf-8-sig")

        required_columns = {
            "abbreviation",
            "long_form",
            "detection_type",
            "source_type",
            "source_index",
            "sentence",
        }

        missing = required_columns - set(df.columns)
        if missing:
            raise ValueError(
                "Во входном CSV отсутствуют обязательные столбцы: "
                + ", ".join(sorted(missing))
            )

        # Нормализация типов и пустых значений
        df = df.copy()
        df["abbreviation"] = df["abbreviation"].fillna("").astype(str).str.strip()
        df["long_form"] = df["long_form"].fillna("").astype(str).str.strip()
        df["detection_type"] = df["detection_type"].fillna("").astype(str).str.strip()
        df["source_type"] = df["source_type"].fillna("").astype(str).str.strip()
        df["source_index"] = pd.to_numeric(df["source_index"], errors="coerce").fillna(-1).astype(int)
        df["sentence"] = df["sentence"].fillna("").astype(str).str.strip()

        return df

    # -----------------------------------------------------------------
    # Основная логика проверки
    # -----------------------------------------------------------------

    def validate_declarations(self, df: pd.DataFrame) -> List[DeclarationCheckResult]:
        """
        Проверяет все объявления сокращений в таблице existing_abbreviations.
        """
        if df.empty:
            return []

        results: List[DeclarationCheckResult] = []

        # Выделяем только строки-объявления
        declarations_df = df[df["detection_type"].isin(self.declaration_types)].copy()

        # Чтобы корректно искать употребления "после объявления",
        # сортируем по позиции в документе.
        declarations_df = declarations_df.sort_values(
            by=["source_index", "abbreviation", "detection_type"],
            ascending=[True, True, True]
        ).reset_index(drop=True)

        for _, row in declarations_df.iterrows():
            abbreviation = row["abbreviation"]
            source_index = int(row["source_index"])
            sentence = row["sentence"]

            # Ищем все более поздние упоминания того же сокращения
            later_df = df[
                (df["abbreviation"] == abbreviation) &
                (df["source_index"] > source_index)
            ].copy()

            later_mentions_count = len(later_df)
            later_standalone_count = len(
                later_df[later_df["detection_type"] == "standalone"]
            )
            later_declared_count = len(
                later_df[later_df["detection_type"].isin(self.declaration_types)]
            )

            # Критерий ошибочного объявления:
            # после объявления сокращение больше не встречается.
            is_erroneous = later_mentions_count == 0

            if is_erroneous:
                comment = (
                    "Сокращение объявлено, но далее по тексту не используется."
                )
            else:
                if later_standalone_count > 0:
                    comment = (
                        "Сокращение объявлено и далее используется в тексте."
                    )
                else:
                    comment = (
                        "После объявления найдены поздние упоминания, "
                        "но требуется ручная проверка их контекста."
                    )

            results.append(
                DeclarationCheckResult(
                    abbreviation=abbreviation,
                    long_form=row["long_form"],
                    detection_type=row["detection_type"],
                    source_type=row["source_type"],
                    source_index=source_index,
                    sentence=sentence,
                    later_mentions_count=later_mentions_count,
                    later_standalone_count=later_standalone_count,
                    later_declared_count=later_declared_count,
                    is_erroneous_declaration=is_erroneous,
                    comment=comment,
                )
            )

        return results

    # -----------------------------------------------------------------
    # Формирование таблиц
    # -----------------------------------------------------------------

    def results_to_dataframe(self, results: List[DeclarationCheckResult]) -> pd.DataFrame:
        """
        Преобразует список результатов в DataFrame.
        """
        if not results:
            return pd.DataFrame(columns=[
                "abbreviation",
                "long_form",
                "detection_type",
                "source_type",
                "source_index",
                "sentence",
                "later_mentions_count",
                "later_standalone_count",
                "later_declared_count",
                "is_erroneous_declaration",
                "comment",
            ])

        return pd.DataFrame([asdict(item) for item in results])

    def build_summary(self, results_df: pd.DataFrame) -> pd.DataFrame:
        """
        Формирует короткую сводку по проверке.
        """
        total_declarations = len(results_df)

        if total_declarations == 0:
            summary = {
                "total_declarations": 0,
                "erroneous_declarations": 0,
                "valid_or_reused_declarations": 0,
                "error_share_percent": 0.0,
            }
            return pd.DataFrame([summary])

        erroneous_count = int(results_df["is_erroneous_declaration"].sum())
        valid_count = total_declarations - erroneous_count
        error_share = round((erroneous_count / total_declarations) * 100, 2)

        summary = {
            "total_declarations": total_declarations,
            "erroneous_declarations": erroneous_count,
            "valid_or_reused_declarations": valid_count,
            "error_share_percent": error_share,
        }
        return pd.DataFrame([summary])

    # -----------------------------------------------------------------
    # Сохранение
    # -----------------------------------------------------------------

    def save_results(
        self,
        results_df: pd.DataFrame,
        output_dir: str | Path
    ) -> Dict[str, Path]:
        """
        Сохраняет результаты проверки в CSV/XLSX.
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        erroneous_df = results_df[results_df["is_erroneous_declaration"] == True].copy()
        summary_df = self.build_summary(results_df)

        all_csv = output_dir / "declaration_validation_all.csv"
        erroneous_csv = output_dir / "erroneous_declarations.csv"
        summary_csv = output_dir / "declaration_validation_summary.csv"

        results_df.to_csv(all_csv, index=False, encoding="utf-8-sig")
        erroneous_df.to_csv(erroneous_csv, index=False, encoding="utf-8-sig")
        summary_df.to_csv(summary_csv, index=False, encoding="utf-8-sig")

        saved_files: Dict[str, Path] = {
            "all_declarations_csv": all_csv,
            "erroneous_declarations_csv": erroneous_csv,
            "summary_csv": summary_csv,
        }

        try:
            all_xlsx = output_dir / "declaration_validation_all.xlsx"
            erroneous_xlsx = output_dir / "erroneous_declarations.xlsx"
            summary_xlsx = output_dir / "declaration_validation_summary.xlsx"

            results_df.to_excel(all_xlsx, index=False)
            erroneous_df.to_excel(erroneous_xlsx, index=False)
            summary_df.to_excel(summary_xlsx, index=False)

            saved_files["all_declarations_xlsx"] = all_xlsx
            saved_files["erroneous_declarations_xlsx"] = erroneous_xlsx
            saved_files["summary_xlsx"] = summary_xlsx

        except ModuleNotFoundError as exc:
            print("Внимание: не удалось сохранить XLSX-файлы.")
            print("Причина:", exc)
            print("CSV-файлы при этом успешно сохранены.")

        return saved_files

    # -----------------------------------------------------------------
    # Полный запуск
    # -----------------------------------------------------------------

    def run(
        self,
        existing_abbreviations_csv: str | Path,
        output_dir: str | Path
    ) -> Dict[str, Path]:
        """
        Полный запуск проверки ошибочных объявлений.
        """
        df = self.load_existing_abbreviations(existing_abbreviations_csv)
        results = self.validate_declarations(df)
        results_df = self.results_to_dataframe(results)
        return self.save_results(results_df, output_dir)


if __name__ == "__main__":
    """
    Пример автономного запуска:
    python declaration_validator.py

    По умолчанию ожидается, что рядом есть файл:
    result_stage2/existing_abbreviations.csv
    """
    input_csv = "result_stage2/existing_abbreviations.csv"
    output_dir = "result_declaration_validation"

    validator = DeclarationValidator()
    saved_files = validator.run(input_csv, output_dir)

    print("=" * 72)
    print("ПРОВЕРКА ОШИБОЧНЫХ ОБЪЯВЛЕНИЙ ЗАВЕРШЕНА")
    print("=" * 72)
    print("Сформированы файлы:")
    for name, path in saved_files.items():
        print(f"{name}: {path}")
