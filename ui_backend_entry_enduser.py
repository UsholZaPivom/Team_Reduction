from __future__ import annotations

import json
import os
import shutil
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from typing import Callable, Dict, List, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from abbreviation_need_stage3 import AbbreviationNeedAnalyzer
from abbreviation_list_inserter import AbbreviationListInserter
from repeated_declaration_replacer import RepeatedDeclarationReplacer


# ==========================================================
# Утилиты путей / метаданных
# ==========================================================


def _get_app_root() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        root = Path(local_app_data) / "ReductionApp"
    else:
        root = Path.home() / ".reduction_app"
    root.mkdir(parents=True, exist_ok=True)
    return root


APP_ROOT = _get_app_root()
CONFIG_PATH = APP_ROOT / "ui_user_run_config.json"
RUNS_ROOT = APP_ROOT / "runs"
RUNS_ROOT.mkdir(parents=True, exist_ok=True)


@dataclass
class RecommendationRow:
    selected: bool
    abbreviation: str
    long_form: str
    recommended: bool
    already_in_document: bool
    score: int
    comment: str
    source: str


class EndUserBackend:
    """
    Backend пользовательского приложения.

    Основная доработка: результаты анализа stage 3 жёстко привязаны
    к конкретному входному документу. Старые рекомендации не подмешиваются
    к новому документу.
    """

    def __init__(self, logger: Optional[Callable[[str], None]] = None) -> None:
        self._log_callback = logger or (lambda msg: None)
        self.stage3_analyzer = AbbreviationNeedAnalyzer()
        self.list_inserter = AbbreviationListInserter()
        self.replacer = RepeatedDeclarationReplacer()

        self.current_source_docx: Optional[Path] = None
        self.current_run_dir: Optional[Path] = None
        self.current_metadata: Optional[dict] = None
        self.recommendation_rows: List[RecommendationRow] = []
        self.stage3_saved_files: Dict[str, Path] = {}

    # ---------------------------
    # Логирование
    # ---------------------------

    def log(self, message: str) -> None:
        self._log_callback(message)

    # ---------------------------
    # Работа с конфигом
    # ---------------------------

    def load_ui_config(self) -> dict:
        if not CONFIG_PATH.exists():
            return {}
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def save_ui_config(self, data: dict) -> None:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.log(f"Сохранён конфиг запуска: {CONFIG_PATH}")

    # ---------------------------
    # Идентификация документа / сессии
    # ---------------------------

    def _build_doc_metadata(self, source_docx: Path) -> dict:
        source_docx = source_docx.resolve()
        stat = source_docx.stat()
        identity = f"{source_docx}|{stat.st_size}|{stat.st_mtime_ns}"
        run_id = sha256(identity.encode("utf-8")).hexdigest()[:24]
        return {
            "source_docx": str(source_docx),
            "source_docx_name": source_docx.name,
            "source_docx_size": int(stat.st_size),
            "source_docx_mtime_ns": int(stat.st_mtime_ns),
            "run_id": run_id,
        }

    def _get_run_dir_for_doc(self, source_docx: Path) -> Path:
        metadata = self._build_doc_metadata(source_docx)
        return RUNS_ROOT / metadata["run_id"]

    def _read_run_metadata(self, run_dir: Path) -> Optional[dict]:
        path = run_dir / "run_metadata.json"
        if not path.exists():
            return None
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return None

    def _write_run_metadata(self, run_dir: Path, metadata: dict) -> None:
        run_dir.mkdir(parents=True, exist_ok=True)
        (run_dir / "run_metadata.json").write_text(
            json.dumps(metadata, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _is_metadata_compatible(self, metadata: dict, source_docx: Path) -> bool:
        try:
            current = self._build_doc_metadata(source_docx)
        except Exception:
            return False
        keys = ["source_docx", "source_docx_size", "source_docx_mtime_ns"]
        return all(metadata.get(key) == current.get(key) for key in keys)

    # ---------------------------
    # Смена входного файла / сброс состояния
    # ---------------------------

    def set_source_docx(self, source_docx: str | Path) -> bool:
        path = Path(source_docx).expanduser()
        if not path.exists():
            raise FileNotFoundError(f"Файл не найден: {path}")
        path = path.resolve()

        changed = self.current_source_docx is None or self.current_source_docx != path
        self.current_source_docx = path
        self.current_run_dir = self._get_run_dir_for_doc(path)
        self.current_metadata = self._build_doc_metadata(path)

        if changed:
            self.reset_analysis_state()
            self.log(f"Выбран новый исходный файл: {path}")
        else:
            self.log(f"Исходный файл подтверждён: {path}")
        return changed

    def reset_analysis_state(self) -> None:
        self.recommendation_rows = []
        self.stage3_saved_files = {}

    # ---------------------------
    # Очистка результатов текущей сессии
    # ---------------------------

    def clear_current_run_outputs(self) -> None:
        if not self.current_run_dir:
            return
        for name in ["stage3", "generated", "replacement_stage"]:
            target = self.current_run_dir / name
            if target.exists():
                shutil.rmtree(target, ignore_errors=True)
                self.log(f"Удалена папка старых результатов: {target}")

    # ---------------------------
    # Анализ stage 3
    # ---------------------------

    def analyze_current_document(self) -> List[RecommendationRow]:
        if not self.current_source_docx:
            raise ValueError("Не выбран входной Word-файл.")
        if not self.current_run_dir or not self.current_metadata:
            raise RuntimeError("Не удалось подготовить сессию запуска.")

        self.reset_analysis_state()
        self.clear_current_run_outputs()

        stage3_dir = self.current_run_dir / "stage3"
        self.log("Запущен анализ документа.")
        self.log("Запускаю анализ документа (этапы 3 и подготовка данных).")

        saved = self.stage3_analyzer.run(self.current_source_docx, stage3_dir)
        self.stage3_saved_files = {key: Path(value) for key, value in saved.items()}
        self._write_run_metadata(self.current_run_dir, self.current_metadata)

        rec_path = self.stage3_saved_files.get("abbreviation_recommendations_csv")
        self.log(f"Найден файл кандидатов: {rec_path}")
        rows = self._load_recommendations_from_current_stage3()
        self.log(f"Загружено вариантов: {len(rows)}")
        return rows

    def refresh_from_stage3(self) -> List[RecommendationRow]:
        if not self.current_source_docx:
            raise ValueError("Не выбран входной Word-файл.")
        if not self.current_run_dir:
            raise RuntimeError("Не определена текущая сессия запуска.")

        metadata = self._read_run_metadata(self.current_run_dir)
        if metadata is None:
            raise FileNotFoundError("Для текущего документа не найден результат stage 3. Сначала выполните анализ.")
        if not self._is_metadata_compatible(metadata, self.current_source_docx):
            raise RuntimeError(
                "Результаты stage 3 относятся к другому документу. "
                "Нужно заново выполнить анализ для текущего файла."
            )

        stage3_dir = self.current_run_dir / "stage3"
        rec_csv = stage3_dir / "abbreviation_recommendations.csv"
        existing_csv = stage3_dir / "existing_abbreviations.csv"
        decisions_csv = stage3_dir / "abbreviation_decisions.csv"

        if not rec_csv.exists():
            raise FileNotFoundError("Не найден abbreviation_recommendations.csv для текущего документа.")

        self.stage3_saved_files = {
            "abbreviation_recommendations_csv": rec_csv,
            "existing_abbreviations_csv": existing_csv,
            "abbreviation_decisions_csv": decisions_csv,
        }
        rows = self._load_recommendations_from_current_stage3()
        self.log(f"Обновлён список из stage 3: {len(rows)}")
        return rows

    def _load_recommendations_from_current_stage3(self) -> List[RecommendationRow]:
        rec_csv = self.stage3_saved_files.get("abbreviation_recommendations_csv")
        if not rec_csv or not Path(rec_csv).exists():
            raise FileNotFoundError("Файл рекомендаций stage 3 не найден.")

        df = pd.read_csv(rec_csv, encoding="utf-8-sig")
        if df.empty:
            self.recommendation_rows = []
            return []

        possible_score_columns = ["decision_score", "score"]
        score_col = next((c for c in possible_score_columns if c in df.columns), None)
        if score_col is None:
            score_col = "decision_score"
            df[score_col] = 0

        possible_term_columns = ["term", "long_form"]
        term_col = next((c for c in possible_term_columns if c in df.columns), None)
        if term_col is None:
            raise ValueError("В файле рекомендаций не найдена колонка term/long_form.")

        possible_abbr_columns = ["suggested_abbreviation", "abbreviation"]
        abbr_col = next((c for c in possible_abbr_columns if c in df.columns), None)
        if abbr_col is None:
            raise ValueError("В файле рекомендаций не найдена колонка suggested_abbreviation/abbreviation.")

        rows: List[RecommendationRow] = []
        for row in df.to_dict("records"):
            abbreviation = self._clean_text(row.get(abbr_col, "")).upper()
            long_form = self._clean_text(row.get(term_col, ""))
            if not abbreviation or not long_form:
                continue

            already = bool(row.get("abbreviation_found_in_text", False))
            recommended = bool(row.get("need_to_introduce", False)) and not already
            score = self._safe_int(row.get(score_col, 0))
            comment = self._clean_text(row.get("reason", ""))
            source = "abbreviation_recommendations"

            rows.append(
                RecommendationRow(
                    selected=recommended,
                    abbreviation=abbreviation,
                    long_form=long_form,
                    recommended=recommended,
                    already_in_document=already,
                    score=score,
                    comment=comment,
                    source=source,
                )
            )

        self.recommendation_rows = rows
        return rows

    # ---------------------------
    # Формирование итоговых данных
    # ---------------------------

    def set_row_selected(self, index: int, selected: bool) -> None:
        if 0 <= index < len(self.recommendation_rows):
            row = self.recommendation_rows[index]
            if row.already_in_document:
                row.selected = False
            else:
                row.selected = selected

    def select_all(self) -> None:
        for row in self.recommendation_rows:
            if not row.already_in_document:
                row.selected = True

    def unselect_all(self) -> None:
        for row in self.recommendation_rows:
            row.selected = False

    def invert_selection(self) -> None:
        for row in self.recommendation_rows:
            if not row.already_in_document:
                row.selected = not row.selected

    def _clean_text(self, value) -> str:
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        return " ".join(str(value).split()).strip()

    def _safe_int(self, value) -> int:
        try:
            if pd.isna(value):
                return 0
        except Exception:
            pass
        try:
            return int(value)
        except Exception:
            try:
                return int(float(value))
            except Exception:
                return 0

    def _existing_entries_df(self) -> pd.DataFrame:
        path = self.stage3_saved_files.get("existing_abbreviations_csv")
        if not path or not Path(path).exists():
            return pd.DataFrame(columns=["abbreviation", "long_form"])

        df = pd.read_csv(path, encoding="utf-8-sig")
        rows: List[dict] = []
        for record in df.to_dict("records"):
            abbreviation = self._clean_text(
                record.get("abbreviation")
                or record.get("found_abbreviation")
                or record.get("abbr")
            ).upper()
            long_form = self._clean_text(
                record.get("long_form")
                or record.get("term")
                or record.get("matched_term")
            )
            if abbreviation and long_form:
                rows.append({"abbreviation": abbreviation, "long_form": long_form})
        return pd.DataFrame(rows)

    def _selected_new_entries_df(self) -> pd.DataFrame:
        rows = [
            {"abbreviation": row.abbreviation, "long_form": row.long_form}
            for row in self.recommendation_rows
            if row.selected and not row.already_in_document
        ]
        return pd.DataFrame(rows)

    def _build_combined_entries_csv(self) -> Path:
        if not self.current_run_dir:
            raise RuntimeError("Не определена папка текущего запуска.")

        existing_df = self._existing_entries_df()
        selected_df = self._selected_new_entries_df()
        combined_df = pd.concat([existing_df, selected_df], ignore_index=True)

        if combined_df.empty:
            raise ValueError(
                "Не найдено ни одного сокращения для формирования перечня. "
                "Либо документ не содержит существующих сокращений, либо пользователь не выбрал новые варианты."
            )

        combined_df["abbreviation"] = combined_df["abbreviation"].map(lambda x: self._clean_text(x).upper())
        combined_df["long_form"] = combined_df["long_form"].map(self._clean_text)
        combined_df = combined_df[(combined_df["abbreviation"] != "") & (combined_df["long_form"] != "")]
        combined_df = combined_df.drop_duplicates(subset=["abbreviation", "long_form"], keep="first")
        combined_df = combined_df.sort_values(by=["abbreviation", "long_form"], ascending=[True, True], kind="stable")

        generated_dir = self.current_run_dir / "generated"
        generated_dir.mkdir(parents=True, exist_ok=True)
        csv_path = generated_dir / "combined_abbreviation_entries.csv"
        combined_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        self.log(f"Сформирован объединённый CSV со списком сокращений: {csv_path}")
        return csv_path

    def _apply_repeated_declaration_replacement_if_needed(self, source_docx: Path, enabled: bool) -> Path:
        if not enabled:
            return source_docx

        existing_csv = self.stage3_saved_files.get("existing_abbreviations_csv")
        if not existing_csv or not Path(existing_csv).exists():
            raise FileNotFoundError("Не найден existing_abbreviations.csv для этапа замены повторных объявлений.")

        if not self.current_run_dir:
            raise RuntimeError("Не определена текущая сессия запуска.")

        replacement_dir = self.current_run_dir / "replacement_stage"
        replacement_dir.mkdir(parents=True, exist_ok=True)
        self.log("Запущена замена повторных объявлений на сокращения.")
        saved = self.replacer.run(
            source_docx_path=source_docx,
            existing_abbreviations_csv=existing_csv,
            output_dir=replacement_dir,
        )
        output_docx = saved.get("output_docx")
        if not output_docx:
            raise RuntimeError("Этап замены повторных объявлений завершился без output_docx.")
        output_docx = Path(output_docx)
        self.log(f"Сформирован документ после замены повторных объявлений: {output_docx}")
        return output_docx

    def process_document(
        self,
        save_dir: str | Path,
        output_name: str,
        insert_mode_ui: str,
        marker_text: str,
        section_title: str,
        replace_repeated_declarations: bool,
        create_separate_list_docx: bool,
    ) -> Dict[str, Path]:
        if not self.current_source_docx:
            raise ValueError("Не выбран входной Word-файл.")
        if not self.current_run_dir:
            raise RuntimeError("Не определена текущая сессия запуска.")
        if not self.recommendation_rows:
            raise RuntimeError("Список сокращений не загружен. Сначала выполните анализ документа.")

        metadata = self._read_run_metadata(self.current_run_dir)
        if metadata is None or not self._is_metadata_compatible(metadata, self.current_source_docx):
            raise RuntimeError(
                "Результаты stage 3 не соответствуют текущему документу. "
                "Нужно заново выполнить анализ после выбора входного файла."
            )

        save_dir = Path(save_dir).expanduser().resolve()
        save_dir.mkdir(parents=True, exist_ok=True)
        output_name = self._clean_text(output_name)
        if not output_name:
            raise ValueError("Не указано имя итогового файла.")

        marker_text = self._clean_text(marker_text)
        section_title = self._clean_text(section_title) or "Перечень обозначений и сокращений"

        ui_to_run_mode = {
            "В конец документа": "insert_end",
            "Перед маркером": "insert_before_marker",
            "В существующий раздел": "append_existing_list",
        }
        run_mode = ui_to_run_mode.get(insert_mode_ui)
        if not run_mode:
            raise ValueError(f"Неизвестный режим вставки: {insert_mode_ui}")
        if run_mode == "insert_before_marker" and not marker_text:
            raise ValueError("Для режима 'Перед маркером' нужно заполнить поле 'Текст маркера'.")

        combined_csv = self._build_combined_entries_csv()
        working_docx = self._apply_repeated_declaration_replacement_if_needed(
            self.current_source_docx,
            replace_repeated_declarations,
        )

        output_docx = save_dir / f"{output_name}.docx"
        self.log("Запущена финальная обработка документа.")

        saved_paths: Dict[str, Path] = {}
        insert_result = self.list_inserter.run(
            input_data_path=combined_csv,
            source_docx_path=working_docx,
            mode=run_mode,
            output_path=output_docx,
            marker_text=marker_text if run_mode == "insert_before_marker" else None,
            section_title=section_title,
        )
        saved_paths.update({key: Path(value) for key, value in insert_result.items()})
        self.log(f"Сформирован итоговый документ: {output_docx}")

        if create_separate_list_docx:
            separate_path = save_dir / f"{output_name}_abbreviations_list.docx"
            separate_result = self.list_inserter.run(
                input_data_path=combined_csv,
                source_docx_path=working_docx,
                mode="separate_file",
                output_path=separate_path,
                section_title=section_title,
            )
            saved_paths["separate_abbreviation_list_docx"] = Path(separate_result["output_docx"])
            self.log(f"Сформирован отдельный Word-файл со списком сокращений: {separate_path}")

        self.save_ui_config(
            {
                "last_source_docx": str(self.current_source_docx),
                "last_save_dir": str(save_dir),
                "last_output_name": output_name,
                "last_insert_mode": insert_mode_ui,
                "last_marker_text": marker_text,
                "last_section_title": section_title,
                "last_replace_repeated_declarations": bool(replace_repeated_declarations),
                "last_create_separate_list_docx": bool(create_separate_list_docx),
            }
        )
        return saved_paths


class EndUserApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Интерфейс обработки сокращений")
        self.geometry("1600x900")
        self.minsize(1200, 700)

        self.backend = EndUserBackend(logger=self._append_log)
        self._build_ui()
        self._restore_config()

    # ---------------------------
    # UI
    # ---------------------------

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        outer = ttk.Frame(self, padding=8)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)
        outer.rowconfigure(1, weight=0)

        canvas = tk.Canvas(outer, highlightthickness=0)
        v_scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(outer, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        self.content = ttk.Frame(canvas, padding=4)
        self.content.columnconfigure(0, weight=1)
        self.content_window = canvas.create_window((0, 0), window=self.content, anchor="nw")

        def _on_content_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfigure(self.content_window, width=max(event.width, 1200))

        self.content.bind("<Configure>", _on_content_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # 1. Основные параметры
        frm1 = ttk.LabelFrame(self.content, text="1. Основные параметры", padding=8)
        frm1.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        frm1.columnconfigure(1, weight=1)

        self.source_docx_var = tk.StringVar()
        ttk.Label(frm1, text="Word-файл пользователя:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=2)
        source_entry = ttk.Entry(frm1, textvariable=self.source_docx_var)
        source_entry.grid(row=0, column=1, sticky="ew", pady=2)
        source_entry.bind("<FocusOut>", lambda _e: self._on_source_entry_focus_out())
        ttk.Button(frm1, text="Выбрать", command=self._choose_source_docx).grid(row=0, column=2, padx=(8, 0), pady=2)

        btn_row = ttk.Frame(frm1)
        btn_row.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Button(btn_row, text="Запустить анализ и загрузить варианты", command=self._run_analysis).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(btn_row, text="Обновить список из stage 3", command=self._refresh_from_stage3).grid(row=0, column=1)

        # 2. Таблица сокращений
        frm2 = ttk.LabelFrame(self.content, text="2. Выбор новых сокращений для ввода", padding=8)
        frm2.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        frm2.columnconfigure(0, weight=1)
        frm2.rowconfigure(2, weight=1)
        ttk.Label(
            frm2,
            text=(
                "В таблице выбираются только НОВЫЕ сокращения, которые пользователь хочет ввести. "
                "Сокращения, уже найденные в документе, будут добавлены в итоговый список автоматически."
            ),
            wraplength=1300,
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        tool_row = ttk.Frame(frm2)
        tool_row.grid(row=1, column=0, sticky="w", pady=(0, 8))
        ttk.Button(tool_row, text="Выбрать все", command=self._select_all).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(tool_row, text="Снять всё", command=self._unselect_all).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(tool_row, text="Инвертировать", command=self._invert_selection).grid(row=0, column=2, padx=(0, 8))
        ttk.Label(tool_row, text="Двойной щелчок по строке переключает выбор пользователя.").grid(row=0, column=3, padx=(8, 0))

        table_frame = ttk.Frame(frm2)
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = (
            "selected",
            "abbreviation",
            "long_form",
            "recommended",
            "already_in_document",
            "score",
            "comment",
            "source",
        )
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=18)
        headings = {
            "selected": "Выбор",
            "abbreviation": "Сокращение",
            "long_form": "Полная форма",
            "recommended": "Рекомендация",
            "already_in_document": "Уже есть в документе",
            "score": "Оценка",
            "comment": "Комментарий",
            "source": "Источник",
        }
        widths = {
            "selected": 80,
            "abbreviation": 140,
            "long_form": 360,
            "recommended": 120,
            "already_in_document": 150,
            "score": 80,
            "comment": 420,
            "source": 180,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w", stretch=True)

        self.tree.grid(row=0, column=0, sticky="nsew")
        t_v = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        t_h = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=t_v.set, xscrollcommand=t_h.set)
        t_v.grid(row=0, column=1, sticky="ns")
        t_h.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # 3. Параметры обработки и сохранения
        frm3 = ttk.LabelFrame(self.content, text="3. Параметры обработки и сохранения", padding=8)
        frm3.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        frm3.columnconfigure(1, weight=1)

        self.save_dir_var = tk.StringVar()
        self.output_name_var = tk.StringVar(value="processed_document")
        self.insert_mode_var = tk.StringVar(value="В конец документа")
        self.marker_text_var = tk.StringVar()
        self.section_title_var = tk.StringVar(value="Перечень обозначений и сокращений")
        self.replace_repeated_var = tk.BooleanVar(value=True)
        self.create_separate_list_var = tk.BooleanVar(value=False)

        ttk.Label(frm3, text="Куда сохранить результат:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Entry(frm3, textvariable=self.save_dir_var).grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Button(frm3, text="Выбрать", command=self._choose_save_dir).grid(row=0, column=2, padx=(8, 0), pady=2)

        ttk.Label(frm3, text="Имя итогового файла:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Entry(frm3, textvariable=self.output_name_var).grid(row=1, column=1, sticky="ew", pady=2)

        ttk.Label(frm3, text="Место вставки списка сокращений:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=2)
        insert_mode_cb = ttk.Combobox(
            frm3,
            textvariable=self.insert_mode_var,
            values=["В конец документа", "Перед маркером", "В существующий раздел"],
            state="readonly",
        )
        insert_mode_cb.grid(row=2, column=1, sticky="ew", pady=2)
        insert_mode_cb.bind("<<ComboboxSelected>>", lambda _e: self._update_insert_mode_state())

        ttk.Label(frm3, text="Текст маркера:").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=2)
        self.marker_entry = ttk.Entry(frm3, textvariable=self.marker_text_var)
        self.marker_entry.grid(row=3, column=1, sticky="ew", pady=2)

        ttk.Label(frm3, text="Название существующего раздела:").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Entry(frm3, textvariable=self.section_title_var).grid(row=4, column=1, sticky="ew", pady=2)

        ttk.Button(frm3, text="Запустить финальную обработку", command=self._run_final_processing).grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 8))
        ttk.Checkbutton(
            frm3,
            text="Выполнять замену повторных объявлений на сокращения",
            variable=self.replace_repeated_var,
        ).grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(
            frm3,
            text="Дополнительно сформировать отдельный Word-файл со списком сокращений",
            variable=self.create_separate_list_var,
        ).grid(row=7, column=0, columnspan=2, sticky="w")

        # 4. Логи
        frm4 = ttk.LabelFrame(self.content, text="4. Логи", padding=8)
        frm4.grid(row=3, column=0, sticky="nsew")
        frm4.columnconfigure(0, weight=1)
        frm4.rowconfigure(0, weight=1)
        self.log_text = tk.Text(frm4, height=10, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(frm4, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.grid(row=0, column=1, sticky="ns")

        self._update_insert_mode_state()

    # ---------------------------
    # Вспомогательные методы UI
    # ---------------------------

    def _append_log(self, message: str) -> None:
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.update_idletasks()

    def _restore_config(self) -> None:
        config = self.backend.load_ui_config()
        if not config:
            return
        self.source_docx_var.set(config.get("last_source_docx", ""))
        self.save_dir_var.set(config.get("last_save_dir", ""))
        self.output_name_var.set(config.get("last_output_name", "processed_document"))
        self.insert_mode_var.set(config.get("last_insert_mode", "В конец документа"))
        self.marker_text_var.set(config.get("last_marker_text", ""))
        self.section_title_var.set(config.get("last_section_title", "Перечень обозначений и сокращений"))
        self.replace_repeated_var.set(bool(config.get("last_replace_repeated_declarations", True)))
        self.create_separate_list_var.set(bool(config.get("last_create_separate_list_docx", False)))
        self._update_insert_mode_state()

        source = self.source_docx_var.get().strip()
        if source:
            try:
                self.backend.set_source_docx(source)
            except Exception:
                # старый путь мог стать невалидным; таблицу автоматически не восстанавливаем
                pass

    def _on_source_entry_focus_out(self) -> None:
        value = self.source_docx_var.get().strip()
        if not value:
            return
        try:
            changed = self.backend.set_source_docx(value)
            if changed:
                self._clear_tree()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def _choose_source_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите Word-файл",
            filetypes=[("Word documents", "*.docx")],
        )
        if not path:
            return
        self.source_docx_var.set(path)
        try:
            changed = self.backend.set_source_docx(path)
            if changed:
                self._clear_tree()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def _choose_save_dir(self) -> None:
        path = filedialog.askdirectory(title="Выберите папку для сохранения результата")
        if path:
            self.save_dir_var.set(path)

    def _update_insert_mode_state(self) -> None:
        mode = self.insert_mode_var.get()
        if mode == "Перед маркером":
            self.marker_entry.configure(state="normal")
        else:
            self.marker_entry.configure(state="disabled")

    def _clear_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _render_rows(self, rows: List[RecommendationRow]) -> None:
        self._clear_tree()
        for index, row in enumerate(rows):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    "Да" if row.selected else "Нет",
                    row.abbreviation,
                    row.long_form,
                    "Да" if row.recommended else "Нет",
                    "Да" if row.already_in_document else "Нет",
                    row.score,
                    row.comment,
                    row.source,
                ),
            )

    def _run_analysis(self) -> None:
        try:
            source = self.source_docx_var.get().strip()
            if not source:
                raise ValueError("Не выбран входной Word-файл.")
            changed = self.backend.set_source_docx(source)
            if changed:
                self._clear_tree()
            rows = self.backend.analyze_current_document()
            self._render_rows(rows)
        except Exception as exc:
            self._append_log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))

    def _refresh_from_stage3(self) -> None:
        try:
            source = self.source_docx_var.get().strip()
            if not source:
                raise ValueError("Не выбран входной Word-файл.")
            changed = self.backend.set_source_docx(source)
            if changed:
                self._clear_tree()
            rows = self.backend.refresh_from_stage3()
            self._render_rows(rows)
        except Exception as exc:
            self._append_log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))

    def _select_all(self) -> None:
        self.backend.select_all()
        self._render_rows(self.backend.recommendation_rows)

    def _unselect_all(self) -> None:
        self.backend.unselect_all()
        self._render_rows(self.backend.recommendation_rows)

    def _invert_selection(self) -> None:
        self.backend.invert_selection()
        self._render_rows(self.backend.recommendation_rows)

    def _on_tree_double_click(self, _event) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        iid = selection[0]
        try:
            index = int(iid)
        except Exception:
            return
        if 0 <= index < len(self.backend.recommendation_rows):
            row = self.backend.recommendation_rows[index]
            self.backend.set_row_selected(index, not row.selected)
            self._render_rows(self.backend.recommendation_rows)

    def _run_final_processing(self) -> None:
        try:
            source = self.source_docx_var.get().strip()
            if not source:
                raise ValueError("Не выбран входной Word-файл.")
            self.backend.set_source_docx(source)
            result = self.backend.process_document(
                save_dir=self.save_dir_var.get().strip(),
                output_name=self.output_name_var.get().strip(),
                insert_mode_ui=self.insert_mode_var.get().strip(),
                marker_text=self.marker_text_var.get().strip(),
                section_title=self.section_title_var.get().strip(),
                replace_repeated_declarations=self.replace_repeated_var.get(),
                create_separate_list_docx=self.create_separate_list_var.get(),
            )
            result_lines = [f"{key}: {value}" for key, value in result.items()]
            messagebox.showinfo("Успешно", "Финальная обработка завершена.\n\n" + "\n".join(result_lines))
        except Exception as exc:
            self._append_log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))
