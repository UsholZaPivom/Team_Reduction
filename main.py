from __future__ import annotations

"""
main.py

Главный файл проекта для последовательного запуска:
1. Этап 1 — распознавание текста и выделение словоформ-кандидатов.
2. Этап 2 — извлечение существующих аббревиатур и полных форм.
3. Проверка ошибочных объявлений сокращений.
4. Этап 3 — определение необходимости ввода аббревиатуры.

Поддерживается логирование через product_logger.py.

Важно:
рядом с этим файлом должны находиться:
- text_recognition_candidates_v3.py
- abbreviation_extraction_stage2.py
- abbreviation_need_stage3.py
- product_logger.py
- declaration_validator.py
"""

from pathlib import Path
import traceback
import pandas as pd

from text_recognition_candidates_v3 import ReducibleWordformRecognizerV3
from abbreviation_extraction_stage2 import Stage2ReductionAnalyzer
from abbreviation_need_stage3 import AbbreviationNeedAnalyzer
from declaration_validator import DeclarationValidator
from product_logger import ProductLogger


def print_header(title: str) -> None:
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


def print_saved_files(saved_files: dict) -> None:
    if not saved_files:
        print("Файлы не были сформированы.")
        return

    print("Сформированы файлы:")
    for name, path in saved_files.items():
        print(f"  {name}: {path}")


def ensure_input_exists(docx_path: Path) -> None:
    if not docx_path.exists():
        raise FileNotFoundError(
            f"Входной документ не найден: {docx_path}\n"
            f"Проверьте имя файла и его расположение."
        )


def run_stage_1(docx_path: Path, output_dir: Path, logger: ProductLogger) -> dict:
    stage_name = "Этап 1: распознавание текста и выделение словоформ"
    logger.stage_started(stage_name)

    print_header("ЭТАП 1. Распознавание текста и выделение словоформ")

    recognizer = ReducibleWordformRecognizerV3()
    mentions = recognizer.analyze_document(docx_path)

    print(f"Найдено сырых вхождений-кандидатов: {len(mentions)}")
    logger.log_candidates_found(len(mentions))

    saved_files = recognizer.save_results(mentions, output_dir)

    print_saved_files(saved_files)
    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )

    return saved_files


def run_stage_2(docx_path: Path, output_dir: Path, logger: ProductLogger) -> dict:
    stage_name = "Этап 2: вычленение словоформ и имеющихся аббревиатур"
    logger.stage_started(stage_name)

    print_header("ЭТАП 2. Вычленение словоформ и имеющихся аббревиатур")

    analyzer = Stage2ReductionAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)

    existing_csv = saved_files.get("existing_abbreviations_csv")
    if existing_csv and Path(existing_csv).exists():
        try:
            existing_df = pd.read_csv(existing_csv, encoding="utf-8-sig")
            logger.log_abbreviations_found(len(existing_df))
        except Exception as exc:
            logger.warning(f"Не удалось прочитать existing_abbreviations.csv для логирования: {exc}")

    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )

    return saved_files


def run_declaration_validation(stage2_files: dict, output_dir: Path, logger: ProductLogger) -> dict:
    stage_name = "Проверка ошибочных объявлений"
    logger.stage_started(stage_name)

    print_header("ПРОВЕРКА ОШИБОЧНЫХ ОБЪЯВЛЕНИЙ")

    existing_csv = stage2_files.get("existing_abbreviations_csv")
    if not existing_csv:
        raise FileNotFoundError(
            "Не найден путь к existing_abbreviations.csv в результатах этапа 2."
        )

    validator = DeclarationValidator()
    saved_files = validator.run(existing_csv, output_dir)

    print_saved_files(saved_files)

    summary_csv = saved_files.get("summary_csv")
    erroneous_csv = saved_files.get("erroneous_declarations_csv")

    if summary_csv and Path(summary_csv).exists():
        try:
            summary_df = pd.read_csv(summary_csv, encoding="utf-8-sig")
            if not summary_df.empty:
                row = summary_df.iloc[0]
                total = int(row.get("total_declarations", 0))
                erroneous = int(row.get("erroneous_declarations", 0))
                valid = int(row.get("valid_or_reused_declarations", 0))
                share = float(row.get("error_share_percent", 0))

                logger.info(
                    "Сводка по объявлениям: "
                    f"всего={total}, ошибочных={erroneous}, корректных/использованных={valid}, "
                    f"доля ошибочных={share}%"
                )
        except Exception as exc:
            logger.warning(f"Не удалось прочитать summary_csv для логирования: {exc}")

    if erroneous_csv and Path(erroneous_csv).exists():
        try:
            erroneous_df = pd.read_csv(erroneous_csv, encoding="utf-8-sig")
            for _, row in erroneous_df.iterrows():
                logger.log_declaration_error(
                    abbreviation=str(row.get("abbreviation", "")).strip(),
                    long_form=str(row.get("long_form", "")).strip(),
                    page=None,
                    line=None
                )
        except Exception as exc:
            logger.warning(f"Не удалось прочитать erroneous_declarations.csv для логирования: {exc}")

    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )

    return saved_files


def run_stage_3(docx_path: Path, output_dir: Path, logger: ProductLogger) -> dict:
    stage_name = "Этап 3: определение необходимости ввода аббревиатуры"
    logger.stage_started(stage_name)

    print_header("ЭТАП 3. Определение необходимости ввода аббревиатуры")

    analyzer = AbbreviationNeedAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)
    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )

    return saved_files


def run_all_stages(
    docx_path: str | Path,
    root_output_dir: str | Path,
    logger: ProductLogger
) -> dict:
    docx_path = Path(docx_path)
    root_output_dir = Path(root_output_dir)

    ensure_input_exists(docx_path)
    root_output_dir.mkdir(parents=True, exist_ok=True)

    stage1_dir = root_output_dir / "stage1"
    stage2_dir = root_output_dir / "stage2"
    validation_dir = root_output_dir / "declaration_validation"
    stage3_dir = root_output_dir / "stage3"

    logger.log_document_loaded(docx_path)
    logger.info(f"Корневая папка результатов: {root_output_dir.resolve()}")

    results = {}
    results["stage1"] = run_stage_1(docx_path, stage1_dir, logger)
    results["stage2"] = run_stage_2(docx_path, stage2_dir, logger)
    results["declaration_validation"] = run_declaration_validation(results["stage2"], validation_dir, logger)
    results["stage3"] = run_stage_3(docx_path, stage3_dir, logger)

    return results


if __name__ == "__main__":
    INPUT_DOCX = "test_reduction_input.docx"
    OUTPUT_ROOT = "result_all"

    logger = ProductLogger(log_dir="logs", console_output=True)

    try:
        logger.start_session(INPUT_DOCX)

        print_header("ПОСЛЕДОВАТЕЛЬНЫЙ ЗАПУСК ВСЕХ ЭТАПОВ ПРОЕКТА")

        all_results = run_all_stages(INPUT_DOCX, OUTPUT_ROOT, logger)

        print_header("ВСЕ ЭТАПЫ УСПЕШНО ЗАВЕРШЕНЫ")
        print(f"Итоговая папка результатов: {Path(OUTPUT_ROOT).resolve()}")

        logger.info(f"Итоговая папка результатов: {Path(OUTPUT_ROOT).resolve()}")

        print("\nКраткая структура результатов:")
        for stage_name, files in all_results.items():
            print(f"\n[{stage_name}]")
            for key, value in files.items():
                print(f"  {key}: {value}")

        logger.finish_session(success=True)

    except Exception as exc:
        print_header("ВО ВРЕМЯ ЗАПУСКА ПРОИЗОШЛА ОШИБКА")
        print(exc)
        print("\nПодробная трассировка:")
        print(traceback.format_exc())

        logger.log_exception("main", exc, with_traceback=True)
        logger.finish_session(success=False)

        raise
