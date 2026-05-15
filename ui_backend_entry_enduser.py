from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any

import pandas as pd


def _patch_pymorphy2_dictionary_path() -> None:
    """
    Fix for PyInstaller one-dir build:
    pymorphy2 sometimes cannot find bundled Russian dictionaries because
    entry points are unavailable in the frozen environment.

    We patch pymorphy2.analyzer.lang_dict_path so that lang='ru' resolves
    directly to the bundled pymorphy2_dicts_ru/data folder.
    """
    try:
        import pymorphy2.analyzer as _pymorphy2_analyzer

        original_lang_dict_path = _pymorphy2_analyzer.lang_dict_path

        def _patched_lang_dict_path(lang: str):
            if lang == "ru":
                candidate_roots = []

                meipass = getattr(sys, "_MEIPASS", None)
                if meipass:
                    candidate_roots.append(Path(meipass))

                exe_dir = Path(sys.executable).resolve().parent
                candidate_roots.append(exe_dir)
                candidate_roots.append(exe_dir / "_internal")

                try:
                    candidate_roots.append(Path(__file__).resolve().parent)
                except Exception:
                    pass

                checked = []
                for root in candidate_roots:
                    data_path = root / "pymorphy2_dicts_ru" / "data"
                    checked.append(str(data_path))
                    if data_path.exists():
                        return str(data_path)

            return original_lang_dict_path(lang)

        _pymorphy2_analyzer.lang_dict_path = _patched_lang_dict_path
    except Exception:
        # If patching fails, application will still raise the original error later,
        # which is better than hiding the failure here.
        pass


_patch_pymorphy2_dictionary_path()

from abbreviation_need_stage3 import AbbreviationNeedAnalyzer
from abbreviation_list_inserter import AbbreviationListInserter
from repeated_declaration_replacer import RepeatedDeclarationReplacer


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def workspace_root() -> Path:
    local = os.environ.get("LOCALAPPDATA")
    if local:
        root = Path(local) / "ReductionApp"
    else:
        root = Path.home() / ".reduction_app"
    root.mkdir(parents=True, exist_ok=True)
    return root


def stage3_dir() -> Path:
    path = workspace_root() / "stage3"
    path.mkdir(parents=True, exist_ok=True)
    return path


def generated_dir() -> Path:
    path = workspace_root() / "generated"
    path.mkdir(parents=True, exist_ok=True)
    return path


def replacement_dir() -> Path:
    path = workspace_root() / "replacement_stage"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _ensure_docx_suffix(name: str) -> str:
    name = str(name).strip()
    if not name:
        name = "processed_document"
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name


def _safe_text(value: Any) -> str:
    text = "" if value is None else str(value)
    return " ".join(text.split()).strip()


def _load_stage3_outputs_or_build(source_docx: Path) -> dict[str, Path]:
    decisions_csv = stage3_dir() / "abbreviation_decisions.csv"
    existing_csv = stage3_dir() / "existing_abbreviations.csv"

    if decisions_csv.exists() and existing_csv.exists():
        return {
            "abbreviation_decisions_csv": decisions_csv,
            "existing_abbreviations_csv": existing_csv,
            "stage3_dir": stage3_dir(),
        }

    analyzer = AbbreviationNeedAnalyzer()
    saved_files = analyzer.run(source_docx, stage3_dir())
    saved_files["stage3_dir"] = stage3_dir()
    return saved_files


def _build_selected_new_entries_dataframe(selected_abbreviations: list[dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for item in selected_abbreviations:
        abbreviation = _safe_text(item.get("abbreviation", ""))
        long_form = _safe_text(item.get("long_form", ""))
        if not abbreviation or not long_form:
            continue
        rows.append({
            "abbreviation": abbreviation,
            "long_form": long_form,
        })
    return pd.DataFrame(rows)


def _build_existing_entries_dataframe(existing_abbreviations_csv: Path) -> pd.DataFrame:
    df = pd.read_csv(existing_abbreviations_csv, encoding="utf-8-sig").copy()

    required = {"abbreviation", "long_form"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            "Во входном existing_abbreviations.csv отсутствуют обязательные столбцы: "
            + ", ".join(sorted(missing))
        )

    df["abbreviation"] = df["abbreviation"].fillna("").astype(str).map(_safe_text)
    df["long_form"] = df["long_form"].fillna("").astype(str).map(_safe_text)

    df = df[
        (df["abbreviation"] != "") &
        (df["long_form"] != "") &
        (df["abbreviation"].str.lower() != "nan") &
        (df["long_form"].str.lower() != "nan")
    ][["abbreviation", "long_form"]].drop_duplicates().reset_index(drop=True)

    return df


def _build_combined_entries_file(
    existing_abbreviations_csv: Path,
    selected_abbreviations: list[dict[str, Any]],
) -> Path:
    existing_df = _build_existing_entries_dataframe(existing_abbreviations_csv)
    selected_df = _build_selected_new_entries_dataframe(selected_abbreviations)

    combined_df = pd.concat([existing_df, selected_df], ignore_index=True)
    if combined_df.empty:
        raise ValueError(
            "После объединения существующих и выбранных новых сокращений список оказался пуст."
        )

    combined_df["abbr_upper"] = combined_df["abbreviation"].astype(str).str.upper()
    combined_df["long_lower"] = combined_df["long_form"].astype(str).str.lower()
    combined_df = combined_df.drop_duplicates(subset=["abbr_upper", "long_lower"]).copy()
    combined_df = combined_df.drop(columns=["abbr_upper", "long_lower"])
    combined_df = combined_df.sort_values(by=["abbreviation", "long_form"], ascending=[True, True]).reset_index(drop=True)

    output_path = generated_dir() / "entries_for_insertion.csv"
    combined_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    return output_path


def _run_repeated_declaration_replacement_if_needed(
    source_docx: Path,
    existing_abbreviations_csv: Path,
    enabled: bool,
) -> Path:
    if not enabled:
        return source_docx

    replacer = RepeatedDeclarationReplacer()
    saved_files = replacer.run(
        source_docx_path=source_docx,
        existing_abbreviations_csv=existing_abbreviations_csv,
        output_dir=replacement_dir(),
    )

    output_docx = saved_files.get("output_docx")
    if not output_docx:
        raise RuntimeError("Этап замены повторных объявлений не вернул итоговый DOCX-файл.")
    return Path(output_docx)


def run_analyze_mode(source_docx: Path) -> None:
    source_docx = Path(source_docx)
    if not source_docx.exists():
        raise FileNotFoundError(f"Файл не найден: {source_docx}")

    _load_stage3_outputs_or_build(source_docx)


def run_process_mode(config_path: Path) -> Path:
    config = json.loads(Path(config_path).read_text(encoding="utf-8"))

    source_docx = Path(config["source_docx"]).resolve()
    output_dir = Path(config["output_dir"]).resolve()
    output_name = _ensure_docx_suffix(config["output_name"])

    insertion_mode = str(config.get("insertion_mode", "append_to_end")).strip()
    marker_text = _safe_text(config.get("marker_text", ""))
    existing_section_title = _safe_text(config.get("existing_section_title", "")) or "Перечень обозначений и сокращений"
    create_separate_file = bool(config.get("create_separate_file", False))
    run_repeated_replacement = bool(config.get("run_repeated_replacement", True))
    selected_abbreviations = list(config.get("selected_abbreviations", []))

    if not source_docx.exists():
        raise FileNotFoundError(f"Файл не найден: {source_docx}")

    output_dir.mkdir(parents=True, exist_ok=True)

    stage3_outputs = _load_stage3_outputs_or_build(source_docx)
    existing_abbreviations_csv = Path(stage3_outputs["existing_abbreviations_csv"])

    entries_csv = _build_combined_entries_file(
        existing_abbreviations_csv=existing_abbreviations_csv,
        selected_abbreviations=selected_abbreviations,
    )

    base_docx = _run_repeated_declaration_replacement_if_needed(
        source_docx=source_docx,
        existing_abbreviations_csv=existing_abbreviations_csv,
        enabled=run_repeated_replacement,
    )

    final_docx_path = output_dir / output_name
    inserter = AbbreviationListInserter()

    if insertion_mode == "append_to_end":
        created = inserter.run(
            input_data_path=entries_csv,
            source_docx_path=base_docx,
            mode="insert_end",
            output_path=final_docx_path,
        )
        processed_docx = Path(created["output_docx"])

    elif insertion_mode == "by_marker":
        if not marker_text:
            raise ValueError("Для режима by_marker нужно указать текст маркера.")
        created = inserter.run(
            input_data_path=entries_csv,
            source_docx_path=base_docx,
            mode="insert_before_marker",
            output_path=final_docx_path,
            marker_text=marker_text,
        )
        processed_docx = Path(created["output_docx"])

    elif insertion_mode == "existing_section":
        created = inserter.run(
            input_data_path=entries_csv,
            source_docx_path=base_docx,
            mode="append_existing_list",
            output_path=final_docx_path,
            section_title=existing_section_title,
        )
        processed_docx = Path(created["output_docx"])

    elif insertion_mode == "separate_file":
        shutil.copy2(base_docx, final_docx_path)
        processed_docx = final_docx_path
    else:
        raise ValueError(f"Неизвестный режим вставки: {insertion_mode}")

    if create_separate_file or insertion_mode == "separate_file":
        separate_docx = output_dir / f"{Path(output_name).stem}_abbreviations_list.docx"
        inserter.run(
            input_data_path=entries_csv,
            source_docx_path=base_docx,
            mode="separate_file",
            output_path=separate_docx,
        )

    return processed_docx
