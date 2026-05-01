from __future__ import annotations

"""
main.py

Главный файл для последовательного запуска этапов проекта.

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
4. Запускает этап замены повторных объявлений:
   - сохраняет первое объявление сокращения;
   - заменяет повторные объявления на краткую форму.

Результаты автоматически раскладываются по папкам:
- result_all/stage1
- result_all/stage2
- result_all/stage3
- result_all/replacement_stage
"""

from pathlib import Path
import traceback

from text_recognition_candidates_v3 import ReducibleWordformRecognizerV3
from abbreviation_extraction_stage2 import Stage2ReductionAnalyzer
from abbreviation_need_stage3 import AbbreviationNeedAnalyzer
from repeated_declaration_replacer import RepeatedDeclarationReplacer


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


def run_stage_1(docx_path: Path, output_dir: Path) -> dict:
    print_header("ЭТАП 1. Распознавание текста и выделение словоформ")

    recognizer = ReducibleWordformRecognizerV3()
    mentions = recognizer.analyze_document(docx_path)

    print(f"Найдено сырых вхождений-кандидатов: {len(mentions)}")

    saved_files = recognizer.save_results(mentions, output_dir)
    print_saved_files(saved_files)
    return saved_files


def run_stage_2(docx_path: Path, output_dir: Path) -> dict:
    print_header("ЭТАП 2. Вычленение словоформ и имеющихся аббревиатур")

    analyzer = Stage2ReductionAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)
    return saved_files


def run_stage_3(docx_path: Path, output_dir: Path) -> dict:
    print_header("ЭТАП 3. Определение необходимости ввода аббревиатуры")

    analyzer = AbbreviationNeedAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)
    return saved_files


def run_replacement_stage(docx_path: Path, stage2_files: dict, output_dir: Path) -> dict:
    print_header("ЭТАП 4. Замена повторных объявлений на сокращения")

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
    return saved_files


def run_all_stages(docx_path: str | Path, root_output_dir: str | Path) -> dict:
    docx_path = Path(docx_path)
    root_output_dir = Path(root_output_dir)

    ensure_input_exists(docx_path)
    root_output_dir.mkdir(parents=True, exist_ok=True)

    stage1_dir = root_output_dir / "stage1"
    stage2_dir = root_output_dir / "stage2"
    stage3_dir = root_output_dir / "stage3"
    replacement_dir = root_output_dir / "replacement_stage"

    results = {}
    results["stage1"] = run_stage_1(docx_path, stage1_dir)
    results["stage2"] = run_stage_2(docx_path, stage2_dir)
    results["stage3"] = run_stage_3(docx_path, stage3_dir)
    results["replacement_stage"] = run_replacement_stage(docx_path, results["stage2"], replacement_dir)

    return results


if __name__ == "__main__":
    INPUT_DOCX = "test_reduction_input.docx"
    OUTPUT_ROOT = "result_all"

    try:
        print_header("ПОСЛЕДОВАТЕЛЬНЫЙ ЗАПУСК ВСЕХ ЭТАПОВ ПРОЕКТА")

        all_results = run_all_stages(INPUT_DOCX, OUTPUT_ROOT)

        print_header("ВСЕ ЭТАПЫ УСПЕШНО ЗАВЕРШЕНЫ")
        print(f"Итоговая папка результатов: {Path(OUTPUT_ROOT).resolve()}")

        print("\nКраткая структура результатов:")
        for stage_name, files in all_results.items():
            print(f"\n[{stage_name}]")
            for key, value in files.items():
                print(f"  {key}: {value}")

    except Exception as exc:
        print_header("ВО ВРЕМЯ ЗАПУСКА ПРОИЗОШЛА ОШИБКА")
        print(exc)
        print("\nПодробная трассировка:")
        print(traceback.format_exc())
