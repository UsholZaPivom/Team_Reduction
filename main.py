from __future__ import annotations

"""
main.py

Главный файл для последовательного запуска этапов проекта
с логированием, замером времени выполнения и обновлением
единой базы аббревиатур после этапа 2.
"""

from pathlib import Path
import json
import logging
import time
import traceback
from datetime import datetime
from typing import Any

import pandas as pd

from text_recognition_candidates_v3 import ReducibleWordformRecognizerV3
from abbreviation_extraction_stage2 import Stage2ReductionAnalyzer
from abbreviation_need_stage3 import AbbreviationNeedAnalyzer
from repeated_declaration_replacer import RepeatedDeclarationReplacer

try:
    from abbreviation_database import AbbreviationDatabase
except Exception:
    AbbreviationDatabase = None


# ----------------------------------------------------------------------
# Логирование
# ----------------------------------------------------------------------

def get_logger(log_dir: Path) -> tuple[logging.Logger, Path]:
    log_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_dir / f"run_{timestamp}.log"

    logger = logging.getLogger("team_reduction_main")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        fmt="[%(asctime)s] [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger, log_path


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


def log_saved_files(logger: logging.Logger, saved_files: dict) -> None:
    if not saved_files:
        logger.warning("Файлы не были сформированы.")
        return

    for name, path in saved_files.items():
        logger.info(f"Сформирован файл [{name}]: {path}")


# ----------------------------------------------------------------------
# Базовые проверки
# ----------------------------------------------------------------------

def ensure_input_exists(docx_path: Path) -> None:
    if not docx_path.exists():
        raise FileNotFoundError(
            f"Входной документ не найден: {docx_path}\n"
            f"Проверьте имя файла и его расположение."
        )


def normalize_long_form(value: Any) -> str:
    text = str(value).strip()
    if not text:
        return ""
    text = " ".join(text.split())
    return text.lower()


def make_record_id(abbreviation: str, normalized_long_form: str) -> str:
    safe_form = normalized_long_form.replace(" ", "_")
    return f"{abbreviation}__{safe_form}"


# ----------------------------------------------------------------------
# Этап 1
# ----------------------------------------------------------------------

def run_stage_1(docx_path: Path, output_dir: Path, logger: logging.Logger) -> dict:
    stage_name = "ЭТАП 1. Распознавание текста и выделение словоформ"
    print_header(stage_name)
    logger.info(f"Запуск: {stage_name}")

    stage_start = time.perf_counter()

    recognizer = ReducibleWordformRecognizerV3()
    mentions = recognizer.analyze_document(docx_path)

    logger.info(f"Найдено сырых вхождений-кандидатов: {len(mentions)}")
    print(f"Найдено сырых вхождений-кандидатов: {len(mentions)}")

    saved_files = recognizer.save_results(mentions, output_dir)

    print_saved_files(saved_files)
    log_saved_files(logger, saved_files)

    duration = time.perf_counter() - stage_start
    logger.info(f"Завершён {stage_name}. Время выполнения: {duration:.3f} с")

    return saved_files


# ----------------------------------------------------------------------
# Этап 2
# ----------------------------------------------------------------------

def run_stage_2(docx_path: Path, output_dir: Path, logger: logging.Logger) -> dict:
    stage_name = "ЭТАП 2. Вычленение словоформ и имеющихся аббревиатур"
    print_header(stage_name)
    logger.info(f"Запуск: {stage_name}")

    stage_start = time.perf_counter()

    analyzer = Stage2ReductionAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)
    log_saved_files(logger, saved_files)

    duration = time.perf_counter() - stage_start
    logger.info(f"Завершён {stage_name}. Время выполнения: {duration:.3f} с")

    return saved_files


# ----------------------------------------------------------------------
# Этап обновления базы аббревиатур
# ----------------------------------------------------------------------

def _call_first_existing_method(obj: Any, method_names: list[str], *args, **kwargs):
    for method_name in method_names:
        if hasattr(obj, method_name):
            method = getattr(obj, method_name)
            return method(*args, **kwargs)
    raise AttributeError(
        f"Не найден ни один из методов: {', '.join(method_names)}"
    )


def _update_database_via_module(
    existing_csv: Path,
    db_json_path: Path,
    export_dir: Path,
    logger: logging.Logger,
) -> dict | None:
    if AbbreviationDatabase is None:
        return None

    logger.info("Попытка обновления базы через модуль abbreviation_database.py")

    db_json_path.parent.mkdir(parents=True, exist_ok=True)
    export_dir.mkdir(parents=True, exist_ok=True)

    db = AbbreviationDatabase(str(db_json_path))

    # Загрузка существующей базы
    for method_name in ["load", "load_database", "read"]:
        if hasattr(db, method_name):
            getattr(db, method_name)()
            break

    # Обновление по existing_abbreviations.csv
    _call_first_existing_method(
        db,
        [
            "update_from_existing_abbreviations_csv",
            "update_from_existing_abbreviations",
            "update_from_csv",
            "sync_with_csv",
            "merge_from_csv",
            "add_from_csv",
            "update_database",
            "process_csv",
            "import_from_csv",
        ],
        str(existing_csv),
    )

    # Самоочистка базы
    for method_name in [
        "cleanup",
        "cleanup_invalid_records",
        "clean_invalid_records",
        "remove_invalid_records",
        "self_clean",
    ]:
        if hasattr(db, method_name):
            getattr(db, method_name)()
            break

    # Сохранение JSON
    for method_name in ["save", "save_database", "write"]:
        if hasattr(db, method_name):
            getattr(db, method_name)()
            break

    saved_files: dict[str, Path] = {
        "database_json": db_json_path,
    }

    # Экспорт для ручной корректировки
    export_done = False

    if hasattr(db, "export_to_csv_and_xlsx"):
        result = db.export_to_csv_and_xlsx(str(export_dir))
        export_done = True
        if isinstance(result, dict):
            for key, value in result.items():
                saved_files[key] = Path(value)
    elif hasattr(db, "export_for_manual_editing"):
        result = db.export_for_manual_editing(str(export_dir))
        export_done = True
        if isinstance(result, dict):
            for key, value in result.items():
                saved_files[key] = Path(value)
    elif hasattr(db, "export_manual_edit_files"):
        result = db.export_manual_edit_files(str(export_dir))
        export_done = True
        if isinstance(result, dict):
            for key, value in result.items():
                saved_files[key] = Path(value)
    else:
        csv_path = export_dir / "abbreviation_database_export.csv"
        xlsx_path = export_dir / "abbreviation_database_export.xlsx"

        if hasattr(db, "export_to_csv"):
            db.export_to_csv(str(csv_path))
            saved_files["database_export_csv"] = csv_path
            export_done = True

        if hasattr(db, "export_to_xlsx"):
            db.export_to_xlsx(str(xlsx_path))
            saved_files["database_export_xlsx"] = xlsx_path
            export_done = True

    if not export_done:
        raise RuntimeError(
            "Модуль abbreviation_database.py найден, но в нём не удалось определить методы экспорта."
        )

    return saved_files


def _update_database_fallback(
    existing_csv: Path,
    db_json_path: Path,
    export_dir: Path,
    logger: logging.Logger,
) -> dict:
    """
    Резервная реализация обновления базы прямо из main.py.
    Используется, если не удалось автоматически вызвать методы
    модуля abbreviation_database.py.
    """
    logger.warning(
        "Переход к резервному сценарию обновления базы из main.py"
    )

    db_json_path.parent.mkdir(parents=True, exist_ok=True)
    export_dir.mkdir(parents=True, exist_ok=True)

    if not existing_csv.exists():
        raise FileNotFoundError(f"Файл не найден: {existing_csv}")

    df = pd.read_csv(existing_csv, encoding="utf-8-sig").copy()

    required_columns = {"abbreviation", "long_form"}
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(
            "Во входном CSV отсутствуют обязательные столбцы: "
            + ", ".join(sorted(missing))
        )

    df["abbreviation"] = df["abbreviation"].fillna("").astype(str).str.strip()
    df["long_form"] = df["long_form"].fillna("").astype(str).str.strip()

    if "detection_type" not in df.columns:
        df["detection_type"] = ""
    df["detection_type"] = df["detection_type"].fillna("").astype(str).str.strip()

    df["normalized_long_form"] = df["long_form"].map(normalize_long_form)

    # Самоочистка
    df = df[
        (df["abbreviation"] != "") &
        (df["long_form"] != "") &
        (df["long_form"].str.lower() != "nan") &
        (df["normalized_long_form"] != "") &
        (df["normalized_long_form"] != "nan")
    ].copy()

    records_map: dict[str, dict] = {}

    if db_json_path.exists():
        with open(db_json_path, "r", encoding="utf-8") as f:
            existing_data = json.load(f)

        for record in existing_data.get("records", []):
            records_map[record["record_id"]] = record

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    source_document = str(existing_csv)

    for _, row in df.iterrows():
        abbreviation = row["abbreviation"]
        long_form = row["long_form"]
        normalized = row["normalized_long_form"]
        detection_type = row["detection_type"]

        record_id = make_record_id(abbreviation, normalized)

        if record_id not in records_map:
            records_map[record_id] = {
                "record_id": record_id,
                "abbreviation": abbreviation,
                "long_form": long_form,
                "normalized_long_form": normalized,
                "status": "active",
                "source_documents": [source_document],
                "source_detection_types": [detection_type] if detection_type else [],
                "comment": "",
                "created_at": now_str,
                "updated_at": now_str,
            }
        else:
            record = records_map[record_id]
            if source_document not in record.get("source_documents", []):
                record.setdefault("source_documents", []).append(source_document)
            if detection_type and detection_type not in record.get("source_detection_types", []):
                record.setdefault("source_detection_types", []).append(detection_type)
            record["updated_at"] = now_str

    records = sorted(
        records_map.values(),
        key=lambda item: (str(item.get("abbreviation", "")).lower(), str(item.get("long_form", "")).lower())
    )

    data = {
        "database_path": str(db_json_path),
        "updated_at": now_str,
        "records_count": len(records),
        "records": records,
    }

    with open(db_json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    export_dir.mkdir(parents=True, exist_ok=True)

    export_df = pd.DataFrame(records)
    csv_path = export_dir / "abbreviation_database_export.csv"
    xlsx_path = export_dir / "abbreviation_database_export.xlsx"

    export_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    export_df.to_excel(xlsx_path, index=False)

    return {
        "database_json": db_json_path,
        "database_export_csv": csv_path,
        "database_export_xlsx": xlsx_path,
    }


def run_database_stage(
    stage2_files: dict,
    logger: logging.Logger,
    db_json_path: Path,
    export_dir: Path,
) -> dict:
    stage_name = "ЭТАП 2.1. Обновление единой базы аббревиатур"
    print_header(stage_name)
    logger.info(f"Запуск: {stage_name}")

    stage_start = time.perf_counter()

    existing_csv = stage2_files.get("existing_abbreviations_csv")
    if not existing_csv:
        raise FileNotFoundError(
            "Не найден existing_abbreviations.csv. "
            "Обновление базы требует результатов этапа 2."
        )

    existing_csv = Path(existing_csv)

    saved_files = None

    try:
        saved_files = _update_database_via_module(
            existing_csv=existing_csv,
            db_json_path=db_json_path,
            export_dir=export_dir,
            logger=logger,
        )
    except Exception as module_exc:
        logger.warning(
            "Не удалось обновить базу через abbreviation_database.py: "
            f"{module_exc}"
        )
        logger.warning("Будет использован резервный сценарий обновления базы.")

    if saved_files is None:
        saved_files = _update_database_fallback(
            existing_csv=existing_csv,
            db_json_path=db_json_path,
            export_dir=export_dir,
            logger=logger,
        )

    print_saved_files(saved_files)
    log_saved_files(logger, saved_files)

    duration = time.perf_counter() - stage_start
    logger.info(f"Завершён {stage_name}. Время выполнения: {duration:.3f} с")

    return saved_files


# ----------------------------------------------------------------------
# Этап 3
# ----------------------------------------------------------------------

def run_stage_3(docx_path: Path, output_dir: Path, logger: logging.Logger) -> dict:
    stage_name = "ЭТАП 3. Определение необходимости ввода аббревиатуры"
    print_header(stage_name)
    logger.info(f"Запуск: {stage_name}")

    stage_start = time.perf_counter()

    analyzer = AbbreviationNeedAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)
    log_saved_files(logger, saved_files)

    duration = time.perf_counter() - stage_start
    logger.info(f"Завершён {stage_name}. Время выполнения: {duration:.3f} с")

    return saved_files


# ----------------------------------------------------------------------
# Этап 4
# ----------------------------------------------------------------------

def run_replacement_stage(docx_path: Path, stage2_files: dict, output_dir: Path, logger: logging.Logger) -> dict:
    stage_name = "ЭТАП 4. Замена повторных объявлений на сокращения"
    print_header(stage_name)
    logger.info(f"Запуск: {stage_name}")

    stage_start = time.perf_counter()

    existing_csv = stage2_files.get("existing_abbreviations_csv")
    if not existing_csv:
        raise FileNotFoundError(
            "Не найден existing_abbreviations.csv. "
            "Этап замены требует результатов этапа 2."
        )

    replacer = RepeatedDeclarationReplacer()
    saved_files = replacer.run(
        source_docx_path=docx_path,
        existing_abbreviations_csv=existing_csv,
        output_dir=output_dir,
    )

    print_saved_files(saved_files)
    log_saved_files(logger, saved_files)

    duration = time.perf_counter() - stage_start
    logger.info(f"Завершён {stage_name}. Время выполнения: {duration:.3f} с")

    return saved_files


# ----------------------------------------------------------------------
# Полный запуск
# ----------------------------------------------------------------------

def run_all_stages(docx_path: str | Path, root_output_dir: str | Path, logger: logging.Logger) -> dict:
    docx_path = Path(docx_path)
    root_output_dir = Path(root_output_dir)

    ensure_input_exists(docx_path)
    root_output_dir.mkdir(parents=True, exist_ok=True)

    stage1_dir = root_output_dir / "stage1"
    stage2_dir = root_output_dir / "stage2"
    stage3_dir = root_output_dir / "stage3"
    replacement_dir = root_output_dir / "replacement_stage"

    db_json_path = Path("abbreviation_database") / "abbreviation_database.json"
    db_export_dir = root_output_dir / "abbreviation_database_export"

    results = {}
    results["stage1"] = run_stage_1(docx_path, stage1_dir, logger)
    results["stage2"] = run_stage_2(docx_path, stage2_dir, logger)
    results["abbreviation_database"] = run_database_stage(
        results["stage2"],
        logger=logger,
        db_json_path=db_json_path,
        export_dir=db_export_dir,
    )
    results["stage3"] = run_stage_3(docx_path, stage3_dir, logger)
    results["replacement_stage"] = run_replacement_stage(docx_path, results["stage2"], replacement_dir, logger)

    return results


if __name__ == "__main__":
    INPUT_DOCX = "test_reduction_input.docx"
    OUTPUT_ROOT = "result_all"
    LOG_DIR = Path("logs")

    logger, log_path = get_logger(LOG_DIR)
    total_start = time.perf_counter()

    try:
        print_header("ПОСЛЕДОВАТЕЛЬНЫЙ ЗАПУСК ВСЕХ ЭТАПОВ ПРОЕКТА")
        logger.info("Начало сессии обработки")
        logger.info(f"Входной документ: {Path(INPUT_DOCX).resolve()}")
        logger.info(f"Корневая папка результатов: {Path(OUTPUT_ROOT).resolve()}")

        all_results = run_all_stages(INPUT_DOCX, OUTPUT_ROOT, logger)

        total_duration = time.perf_counter() - total_start

        print_header("ВСЕ ЭТАПЫ УСПЕШНО ЗАВЕРШЕНЫ")
        print(f"Итоговая папка результатов: {Path(OUTPUT_ROOT).resolve()}")
        print(f"Лог-файл: {log_path.resolve()}")

        logger.info("Все этапы успешно завершены")
        logger.info(f"Общее время выполнения программы: {total_duration:.3f} с")
        logger.info(f"Лог-файл сохранён: {log_path.resolve()}")

        print("\nКраткая структура результатов:")
        for stage_name, files in all_results.items():
            print(f"\n[{stage_name}]")
            logger.info(f"Результаты блока [{stage_name}]")
            for key, value in files.items():
                print(f"  {key}: {value}")
                logger.info(f"  {key}: {value}")

    except Exception as exc:
        total_duration = time.perf_counter() - total_start

        print_header("ВО ВРЕМЯ ЗАПУСКА ПРОИЗОШЛА ОШИБКА")
        print(exc)
        print("\nПодробная трассировка:")
        print(traceback.format_exc())

        logger.error(f"Во время выполнения произошла ошибка: {exc}")
        logger.error(traceback.format_exc())
        logger.info(f"Общее время до ошибки: {total_duration:.3f} с")
        logger.info(f"Лог-файл сохранён: {log_path.resolve()}")
