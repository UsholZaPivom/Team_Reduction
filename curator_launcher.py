from __future__ import annotations

"""
curator_launcher.py

Упрощённый консольный запускатор для куратора.

Ключевой принцип:
куратору НЕ нужно вручную указывать пути к документу, базе или результатам.
Он просто запускает .exe, выбирает пункт меню и получает результат.

Как это работает:
1. Программа автоматически берёт входной документ:
   test_reduction_input.docx
   из папки рядом с .exe

2. Все результаты автоматически складываются рядом с .exe:
   - result_all/
   - logs/
   - abbreviation_database/

3. Если нужные промежуточные файлы ещё не созданы,
   программа сама создаёт их в нужной последовательности.

Что можно проверить:
1. Полный прогон всей программы
2. Создание отдельного файла со списком сокращений
3. Вставка списка сокращений в конец документа
4. Вставка списка сокращений перед маркером
5. Обновление существующего раздела сокращений
6. Обновление единой базы аббревиатур
7. Показать, где лежат результаты и логи
"""

import sys
import traceback
from pathlib import Path
from typing import Dict, Optional

import pandas as pd

from text_recognition_candidates_v3 import ReducibleWordformRecognizerV3
from abbreviation_extraction_stage2 import Stage2ReductionAnalyzer
from abbreviation_need_stage3 import AbbreviationNeedAnalyzer
from declaration_validator import DeclarationValidator
from product_logger import ProductLogger
from abbreviation_database import AbbreviationDatabase
from abbreviation_list_inserter import AbbreviationListInserter


# ---------------------------------------------------------------------
# Базовые пути приложения
# ---------------------------------------------------------------------

def get_app_dir() -> Path:
    """
    Возвращает рабочую папку приложения.
    Для .exe — папка рядом с exe.
    Для .py — папка рядом со скриптом.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
INPUT_DOCX = APP_DIR / "test_reduction_input.docx"
OUTPUT_ROOT = APP_DIR / "result_all"
DATABASE_PATH = APP_DIR / "abbreviation_database" / "abbreviation_database.json"
LOG_DIR = APP_DIR / "logs"


# ---------------------------------------------------------------------
# Печать и проверки
# ---------------------------------------------------------------------

def print_header(title: str) -> None:
    print("\n" + "=" * 78)
    print(title)
    print("=" * 78)


def print_saved_files(saved_files: Dict) -> None:
    if not saved_files:
        print("Файлы не были сформированы.")
        return

    print("Сформированы файлы:")
    for name, path in saved_files.items():
        print(f"  {name}: {path}")


def ensure_input_exists() -> None:
    if not INPUT_DOCX.exists():
        raise FileNotFoundError(
            "Не найден входной файл test_reduction_input.docx.\n"
            f"Положите его рядом с exe в папку:\n{APP_DIR}"
        )


def ensure_app_dirs() -> None:
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------
# Этапы пайплайна
# ---------------------------------------------------------------------

def run_stage_1(output_dir: Path, logger: ProductLogger) -> Dict:
    stage_name = "Этап 1: распознавание текста и выделение словоформ"
    logger.stage_started(stage_name)
    print_header("ЭТАП 1. Распознавание текста и выделение словоформ")

    recognizer = ReducibleWordformRecognizerV3()
    mentions = recognizer.analyze_document(INPUT_DOCX)
    logger.log_candidates_found(len(mentions))

    saved_files = recognizer.save_results(mentions, output_dir)
    print_saved_files(saved_files)

    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )
    return saved_files


def run_stage_2(output_dir: Path, logger: ProductLogger) -> Dict:
    stage_name = "Этап 2: извлечение аббревиатур"
    logger.stage_started(stage_name)
    print_header("ЭТАП 2. Извлечение аббревиатур")

    analyzer = Stage2ReductionAnalyzer()
    saved_files = analyzer.run(INPUT_DOCX, output_dir)
    print_saved_files(saved_files)

    existing_csv = saved_files.get("existing_abbreviations_csv")
    if existing_csv and Path(existing_csv).exists():
        try:
            existing_df = pd.read_csv(existing_csv, encoding="utf-8-sig")
            logger.log_abbreviations_found(len(existing_df))
        except Exception as exc:
            logger.warning(f"Не удалось прочитать existing_abbreviations.csv: {exc}")

    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )
    return saved_files


def ensure_stage2_results(logger: ProductLogger) -> Dict:
    """
    Гарантирует наличие результатов этапа 2.
    Если их нет — запускает этап 2 автоматически.
    """
    stage2_dir = OUTPUT_ROOT / "stage2"
    existing_csv = stage2_dir / "existing_abbreviations.csv"

    if existing_csv.exists():
        logger.info(f"Используются готовые результаты этапа 2: {existing_csv}")
        return {"existing_abbreviations_csv": existing_csv}

    logger.info("Результаты этапа 2 не найдены. Выполняется автоматический запуск этапа 2.")
    return run_stage_2(stage2_dir, logger)


def run_declaration_validation(stage2_files: Dict, output_dir: Path, logger: ProductLogger) -> Dict:
    stage_name = "Проверка ошибочных объявлений"
    logger.stage_started(stage_name)
    print_header("ПРОВЕРКА ОШИБОЧНЫХ ОБЪЯВЛЕНИЙ")

    existing_csv = stage2_files.get("existing_abbreviations_csv")
    if not existing_csv:
        raise FileNotFoundError("Не найден existing_abbreviations.csv после этапа 2.")

    validator = DeclarationValidator()
    saved_files = validator.run(existing_csv, output_dir)
    print_saved_files(saved_files)

    summary_csv = saved_files.get("summary_csv")
    if summary_csv and Path(summary_csv).exists():
        try:
            summary_df = pd.read_csv(summary_csv, encoding="utf-8-sig")
            if not summary_df.empty:
                row = summary_df.iloc[0]
                logger.info(
                    "Сводка по ошибочным объявлениям: "
                    f"всего={int(row.get('total_declarations', 0))}, "
                    f"ошибочных={int(row.get('erroneous_declarations', 0))}, "
                    f"корректных={int(row.get('valid_or_reused_declarations', 0))}"
                )
        except Exception as exc:
            logger.warning(f"Не удалось прочитать summary_csv: {exc}")

    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )
    return saved_files


def update_abbreviation_database(stage2_files: Dict, export_dir: Path, logger: ProductLogger) -> Dict:
    stage_name = "Обновление единой базы аббревиатур"
    logger.stage_started(stage_name)
    print_header("ОБНОВЛЕНИЕ ЕДИНОЙ БАЗЫ АББРЕВИАТУР")

    existing_csv = stage2_files.get("existing_abbreviations_csv")
    if not existing_csv:
        raise FileNotFoundError("Не найден existing_abbreviations.csv после этапа 2.")

    db = AbbreviationDatabase(DATABASE_PATH)
    db.load()
    logger.log_database_loaded(DATABASE_PATH, len(db.records))

    import_stats = db.import_from_existing_abbreviations_csv(
        csv_path=existing_csv,
        source_document=str(existing_csv)
    )
    db.save()

    export_files = db.export_for_manual_edit(export_dir)
    summary_df = db.build_summary()
    summary_csv = export_dir / "abbreviation_database_summary.csv"
    summary_df.to_csv(summary_csv, index=False, encoding="utf-8-sig")

    saved_files = {
        "database_json": db.db_path,
        "manual_export_csv": export_files.get("csv"),
        "summary_csv": summary_csv,
    }
    if "xlsx" in export_files:
        saved_files["manual_export_xlsx"] = export_files["xlsx"]

    print(f"Статистика обновления базы: {import_stats}")
    print_saved_files(saved_files)

    logger.log_database_updated(
        added_count=int(import_stats.get("added", 0)),
        updated_count=int(import_stats.get("updated", 0))
    )
    logger.stage_finished(
        stage_name,
        details=(
            f"добавлено {import_stats.get('added', 0)}, "
            f"обновлено {import_stats.get('updated', 0)}, "
            f"пропущено {import_stats.get('skipped', 0)}"
        )
    )
    return saved_files


def run_stage_3(output_dir: Path, logger: ProductLogger) -> Dict:
    stage_name = "Этап 3: определение необходимости ввода аббревиатуры"
    logger.stage_started(stage_name)
    print_header("ЭТАП 3. Определение необходимости ввода аббревиатуры")

    analyzer = AbbreviationNeedAnalyzer()
    saved_files = analyzer.run(INPUT_DOCX, output_dir)
    print_saved_files(saved_files)

    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )
    return saved_files


def run_full_pipeline(logger: ProductLogger) -> Dict:
    ensure_input_exists()
    ensure_app_dirs()

    stage1_dir = OUTPUT_ROOT / "stage1"
    stage2_dir = OUTPUT_ROOT / "stage2"
    validation_dir = OUTPUT_ROOT / "declaration_validation"
    db_export_dir = OUTPUT_ROOT / "abbreviation_database_export"
    stage3_dir = OUTPUT_ROOT / "stage3"

    logger.log_document_loaded(INPUT_DOCX)
    logger.info(f"Корневая папка результатов: {OUTPUT_ROOT}")
    logger.info(f"Путь к базе аббревиатур: {DATABASE_PATH}")

    results = {}
    results["stage1"] = run_stage_1(stage1_dir, logger)
    results["stage2"] = run_stage_2(stage2_dir, logger)
    results["declaration_validation"] = run_declaration_validation(results["stage2"], validation_dir, logger)
    results["abbreviation_database"] = update_abbreviation_database(results["stage2"], db_export_dir, logger)
    results["stage3"] = run_stage_3(stage3_dir, logger)
    return results


# ---------------------------------------------------------------------
# Проверка задачи 2
# ---------------------------------------------------------------------

def get_default_abbreviation_source() -> Path:
    return OUTPUT_ROOT / "stage2" / "existing_abbreviations.csv"


def ensure_abbreviation_source(logger: ProductLogger) -> Path:
    """
    Гарантирует наличие файла existing_abbreviations.csv для задачи 2.
    Если его нет — автоматически запускает этап 2.
    """
    path = get_default_abbreviation_source()
    if path.exists():
        logger.info(f"Используется файл сокращений: {path}")
        return path

    logger.info("Файл existing_abbreviations.csv не найден. Автоматически запускается этап 2.")
    ensure_input_exists()
    ensure_app_dirs()
    run_stage_2(OUTPUT_ROOT / "stage2", logger)
    return path


def run_task2_mode(
    mode: str,
    logger: ProductLogger,
    marker_text: Optional[str] = None
) -> Dict:
    ensure_input_exists()
    ensure_app_dirs()

    input_data_path = ensure_abbreviation_source(logger)
    if not input_data_path.exists():
        raise FileNotFoundError(f"Файл с сокращениями не найден: {input_data_path}")

    inserter = AbbreviationListInserter()
    saved = inserter.run(
        input_data_path=input_data_path,
        source_docx_path=INPUT_DOCX,
        mode=mode,
        marker_text=marker_text
    )
    print_saved_files(saved)

    output_docx = saved.get("output_docx")
    if output_docx:
        if mode == "separate_file":
            logger.log_list_created(output_docx)
        elif mode == "insert_end":
            logger.log_list_inserted(output_docx, "в конец документа")
        elif mode == "insert_before_marker":
            logger.log_list_inserted(output_docx, f'перед маркером "{marker_text}"')
        elif mode == "append_existing_list":
            logger.log_existing_list_updated(output_docx)

    return saved


# ---------------------------------------------------------------------
# Меню
# ---------------------------------------------------------------------

def print_menu() -> None:
    print_header("МЕНЮ ПРОВЕРКИ ФУНКЦИОНАЛА")
    print(f"Входной файл: {INPUT_DOCX.name}")
    print("1. Полный прогон всей программы")
    print("2. Создать отдельный файл со списком сокращений")
    print("3. Вставить список сокращений в конец документа")
    print('4. Вставить список сокращений перед маркером "Общие положения"')
    print("5. Обновить уже существующий раздел сокращений")
    print("6. Обновить единую базу аббревиатур из результатов этапа 2")
    print("7. Показать, где лежат результаты и логи")
    print("0. Выход")


def print_paths_info() -> None:
    print_header("РАСПОЛОЖЕНИЕ ФАЙЛОВ")
    print(f"Папка приложения: {APP_DIR}")
    print(f"Входной документ: {INPUT_DOCX}")
    print(f"Папка результатов: {OUTPUT_ROOT}")
    print(f"Папка логов: {LOG_DIR}")
    print(f"База аббревиатур: {DATABASE_PATH}")


def run_with_session(action_name: str, callback):
    logger = ProductLogger(log_dir=LOG_DIR, console_output=True)
    logger.start_session(action_name)
    try:
        result = callback(logger)
        logger.finish_session(success=True)
        print("\nГотово.")
        print(f"Результаты находятся в: {OUTPUT_ROOT.resolve()}")
        print(f"Логи находятся в: {LOG_DIR.resolve()}")
        return result
    except Exception as exc:
        print("\nВо время выполнения произошла ошибка:")
        print(exc)
        print("\nПодробности:")
        print(traceback.format_exc())
        logger.log_exception(action_name, exc, with_traceback=True)
        logger.finish_session(success=False)
        return None


def main():
    ensure_app_dirs()

    while True:
        print_menu()
        choice = input("Выберите действие: ").strip()

        if choice == "0":
            print("Завершение работы.")
            break

        elif choice == "1":
            def action(logger: ProductLogger):
                results = run_full_pipeline(logger)
                print_header("ПОЛНЫЙ ПРОГОН ЗАВЕРШЁН")
                for stage_name, files in results.items():
                    print(f"\n[{stage_name}]")
                    for key, value in files.items():
                        print(f"  {key}: {value}")
                return results

            run_with_session("full_pipeline", action)

        elif choice == "2":
            run_with_session(
                "task2_separate_file",
                lambda logger: run_task2_mode("separate_file", logger)
            )

        elif choice == "3":
            run_with_session(
                "task2_insert_end",
                lambda logger: run_task2_mode("insert_end", logger)
            )

        elif choice == "4":
            run_with_session(
                "task2_insert_before_marker",
                lambda logger: run_task2_mode("insert_before_marker", logger, marker_text="Общие положения")
            )

        elif choice == "5":
            run_with_session(
                "task2_append_existing_list",
                lambda logger: run_task2_mode("append_existing_list", logger)
            )

        elif choice == "6":
            def action(logger: ProductLogger):
                stage2_files = ensure_stage2_results(logger)
                return update_abbreviation_database(
                    stage2_files=stage2_files,
                    export_dir=OUTPUT_ROOT / "abbreviation_database_export",
                    logger=logger
                )

            run_with_session("database_update", action)

        elif choice == "7":
            print_paths_info()

        else:
            print("Неизвестная команда. Повторите ввод.")


if __name__ == "__main__":
    main()
