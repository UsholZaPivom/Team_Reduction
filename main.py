from __future__ import annotations

"""
main.py

Главный файл для поочерёдного запуска всех трёх этапов проекта
с поддержкой логирования.

Что делает данный файл:
1. Запускает этап 1:
   - распознавание текста;
   - выделение словоформ, поддающихся сокращению.
2. Запускает этап 2:
   - вычленение доступных к сокращению словоформ;
   - поиск уже имеющихся аббревиатур;
   - сопоставление полных форм и сокращений.
3. Запускает этап 3:
   - определение необходимости ввода аббревиатуры.
4. Записывает ход работы в лог:
   - начало и завершение этапов;
   - количество найденных кандидатов;
   - ошибки и traceback.
"""

from pathlib import Path
import traceback

from text_recognition_candidates_v3 import ReducibleWordformRecognizerV3
from abbreviation_extraction_stage2 import Stage2ReductionAnalyzer
from abbreviation_need_stage3 import AbbreviationNeedAnalyzer
from product_logger import ProductLogger


def print_header(title: str) -> None:
    """
    Печатает заголовок этапа в консоль.
    """
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


def print_saved_files(saved_files: dict) -> None:
    """
    Печатает список файлов, созданных модулем.
    """
    if not saved_files:
        print("Файлы не были сформированы.")
        return

    print("Сформированы файлы:")
    for name, path in saved_files.items():
        print(f"  {name}: {path}")


def ensure_input_exists(docx_path: Path) -> None:
    """
    Проверяет наличие входного документа.
    """
    if not docx_path.exists():
        raise FileNotFoundError(
            f"Входной документ не найден: {docx_path}\n"
            f"Проверьте имя файла и его расположение."
        )


def run_stage_1(docx_path: Path, output_dir: Path, logger: ProductLogger) -> dict:
    """
    Запускает этап 1:
    распознавание текста и выделение словоформ,
    поддающихся сокращению.
    """
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
    """
    Запускает этап 2:
    вычленение доступных к сокращению словоформ
    и уже имеющихся аббревиатур.
    """
    stage_name = "Этап 2: вычленение словоформ и имеющихся аббревиатур"
    logger.stage_started(stage_name)

    print_header("ЭТАП 2. Вычленение словоформ и имеющихся аббревиатур")

    analyzer = Stage2ReductionAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)
    logger.stage_finished(
        stage_name,
        details=f"Сохранены файлы: {', '.join(saved_files.keys())}"
    )

    return saved_files


def run_stage_3(docx_path: Path, output_dir: Path, logger: ProductLogger) -> dict:
    """
    Запускает этап 3:
    определение необходимости ввода аббревиатуры.
    """
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
    """
    Поочерёдно запускает все три этапа проекта.
    """
    docx_path = Path(docx_path)
    root_output_dir = Path(root_output_dir)

    ensure_input_exists(docx_path)
    root_output_dir.mkdir(parents=True, exist_ok=True)

    stage1_dir = root_output_dir / "stage1"
    stage2_dir = root_output_dir / "stage2"
    stage3_dir = root_output_dir / "stage3"

    logger.log_document_loaded(docx_path)
    logger.info(f"Корневая папка результатов: {root_output_dir.resolve()}")

    results = {}
    results["stage1"] = run_stage_1(docx_path, stage1_dir, logger)
    results["stage2"] = run_stage_2(docx_path, stage2_dir, logger)
    results["stage3"] = run_stage_3(docx_path, stage3_dir, logger)

    return results


if __name__ == "__main__":
    """
    Самый удобный способ тестирования в PyCharm:

    1. Положите рядом с этим файлом:
       - test_reduction_input.docx
       - text_recognition_candidates_v3.py
       - abbreviation_extraction_stage2.py
       - abbreviation_need_stage3.py
       - product_logger.py

    2. При необходимости измените имя файла ниже:
       INPUT_DOCX = "test_reduction_input.docx"

    3. Нажмите Run.

    Что получится на выходе:
    - папка result_all
      ├── stage1
      ├── stage2
      └── stage3
    - папка logs с .log-файлом
    """

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
