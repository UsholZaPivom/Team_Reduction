from __future__ import annotations

"""
main.py

Главный файл для поочерёдного запуска всех трёх этапов проекта.

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

Для удобства тестирования все результаты автоматически раскладываются
по отдельным папкам:
- result_all/stage1
- result_all/stage2
- result_all/stage3

Важно:
рядом с этим файлом должны находиться:
- text_recognition_candidates_v3.py
- abbreviation_extraction_stage2.py
- abbreviation_need_stage3.py

При необходимости можно изменить имя входного документа
и корневую папку результатов внизу файла.
"""

from pathlib import Path
import traceback

from text_recognition_candidates_v3 import ReducibleWordformRecognizerV3
from abbreviation_extraction_stage2 import Stage2ReductionAnalyzer
from abbreviation_need_stage3 import AbbreviationNeedAnalyzer


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


def run_stage_1(docx_path: Path, output_dir: Path) -> dict:
    """
    Запускает этап 1:
    распознавание текста и выделение словоформ,
    поддающихся сокращению.
    """
    print_header("ЭТАП 1. Распознавание текста и выделение словоформ")

    recognizer = ReducibleWordformRecognizerV3()
    mentions = recognizer.analyze_document(docx_path)

    print(f"Найдено сырых вхождений-кандидатов: {len(mentions)}")

    saved_files = recognizer.save_results(mentions, output_dir)

    print_saved_files(saved_files)
    return saved_files


def run_stage_2(docx_path: Path, output_dir: Path) -> dict:
    """
    Запускает этап 2:
    вычленение доступных к сокращению словоформ
    и уже имеющихся аббревиатур.
    """
    print_header("ЭТАП 2. Вычленение словоформ и имеющихся аббревиатур")

    analyzer = Stage2ReductionAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)
    return saved_files


def run_stage_3(docx_path: Path, output_dir: Path) -> dict:
    """
    Запускает этап 3:
    определение необходимости ввода аббревиатуры.
    """
    print_header("ЭТАП 3. Определение необходимости ввода аббревиатуры")

    analyzer = AbbreviationNeedAnalyzer()
    saved_files = analyzer.run(docx_path, output_dir)

    print_saved_files(saved_files)
    return saved_files


def run_all_stages(docx_path: str | Path, root_output_dir: str | Path) -> dict:
    """
    Поочерёдно запускает все три этапа проекта.

    Параметры:
    - docx_path: путь к входному .docx-документу;
    - root_output_dir: корневая папка, куда будут сложены результаты.

    Возвращает словарь:
    {
        "stage1": {...},
        "stage2": {...},
        "stage3": {...}
    }
    """
    docx_path = Path(docx_path)
    root_output_dir = Path(root_output_dir)

    ensure_input_exists(docx_path)
    root_output_dir.mkdir(parents=True, exist_ok=True)

    stage1_dir = root_output_dir / "stage1"
    stage2_dir = root_output_dir / "stage2"
    stage3_dir = root_output_dir / "stage3"

    results = {}
    results["stage1"] = run_stage_1(docx_path, stage1_dir)
    results["stage2"] = run_stage_2(docx_path, stage2_dir)
    results["stage3"] = run_stage_3(docx_path, stage3_dir)

    return results


if __name__ == "__main__":
    """
    Самый удобный способ тестирования в PyCharm:

    1. Положите рядом с этим файлом:
       - input.docx
       - text_recognition_candidates_v3.py
       - abbreviation_extraction_stage2.py
       - abbreviation_need_stage3.py

    2. При необходимости измените имя файла ниже:
       INPUT_DOCX = "input.docx"

    3. Нажмите Run.

    Что получится на выходе:
    - папка result_all
      ├── stage1
      ├── stage2
      └── stage3
    """

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
