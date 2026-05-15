from __future__ import annotations

import json
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ui_backend_entry_enduser import (
    run_analyze_mode,
    run_process_mode,
    workspace_root,
    stage3_dir,
)


@dataclass
class AbbreviationDecision:
    abbreviation: str
    long_form: str
    recommended: bool = True
    selected: bool = True
    source_file: str = ""
    note: str = ""
    score: str = ""
    already_exists: bool = False


@dataclass
class UiRunConfig:
    source_docx: str
    output_dir: str
    output_name: str
    insertion_mode: str
    marker_text: str
    existing_section_title: str
    create_separate_file: bool
    run_repeated_replacement: bool
    selected_abbreviations: list[dict[str, Any]]


class ProjectAdapter:
    CANDIDATE_FILE_PATTERNS = (
        "*abbreviation_recommendations*.csv",
        "*abbreviation_recommendations*.xlsx",
        "*abbreviation_decisions*.csv",
        "*abbreviation_decisions*.xlsx",
        "*stage3*.csv",
        "*stage3*.xlsx",
        "*abbreviation_need*.csv",
        "*abbreviation_need*.xlsx",
        "*decision*.csv",
        "*decision*.xlsx",
        "*candidate*.csv",
        "*candidate*.xlsx",
    )

    POSSIBLE_ABBR_COLUMNS = [
        "abbreviation", "abbr", "short_form", "сокращение", "аббревиатура",
        "suggested_abbreviation", "found_abbreviation"
    ]
    POSSIBLE_LONG_COLUMNS = [
        "long_form", "full_form", "term", "расшифровка", "полная_форма",
        "полная форма", "термин", "matched_term"
    ]
    POSSIBLE_RECOMMEND_COLUMNS = [
        "need_abbreviation", "recommended", "introduce", "decision",
        "нужно_вводить", "рекомендовано", "need_to_introduce"
    ]
    POSSIBLE_SCORE_COLUMNS = [
        "score", "confidence", "probability", "вес", "оценка", "уверенность",
        "decision_score", "match_score"
    ]
    POSSIBLE_NOTE_COLUMNS = [
        "note", "comment", "reason", "пояснение", "комментарий", "priority"
    ]
    POSSIBLE_ALREADY_EXISTS_COLUMNS = [
        "already_exists", "already_in_document", "exists_in_document",
        "уже_есть_в_документе", "уже есть в документе",
        "abbreviation_found_in_text"
    ]

    def __init__(self, logger) -> None:
        self.logger = logger

    def run_analysis(self, source_docx: Path) -> list[AbbreviationDecision]:
        self.logger("Запускаю анализ документа (этапы 3 и подготовка данных).")
        run_analyze_mode(source_docx)
        return self.load_stage3_candidates()

    def load_stage3_candidates(self) -> list[AbbreviationDecision]:
        candidate_file = self._find_candidate_file()
        if candidate_file is None:
            raise FileNotFoundError(
                "Не удалось найти выходной файл stage 3. "
                "Проверьте, что анализ завершился корректно."
            )

        self.logger(f"Найден файл кандидатов: {candidate_file}")
        df = self._read_table(candidate_file)
        decisions = self._convert_dataframe_to_decisions(df, candidate_file)
        if not decisions:
            raise ValueError(
                f"Файл найден, но из него не удалось извлечь сокращения: {candidate_file}"
            )
        return decisions

    def run_final_processing(self, config: UiRunConfig) -> None:
        config_path = workspace_root() / "ui_user_run_config.json"
        config_path.write_text(
            json.dumps(asdict(config), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self.logger(f"Сохранён конфиг запуска: {config_path}")
        run_process_mode(config_path)

    def _find_candidate_file(self) -> Optional[Path]:
        root = stage3_dir()
        if not root.exists():
            return None

        for pattern in self.CANDIDATE_FILE_PATTERNS:
            for path in sorted(root.rglob(pattern)):
                if path.is_file():
                    return path
        return None

    @staticmethod
    def _read_table(path: Path) -> pd.DataFrame:
        if path.suffix.lower() == ".csv":
            return pd.read_csv(path, encoding="utf-8-sig")
        if path.suffix.lower() in {".xlsx", ".xls"}:
            return pd.read_excel(path)
        raise ValueError(f"Неподдерживаемый формат файла: {path}")

    def _first_existing_column(self, df: pd.DataFrame, variants: list[str]) -> Optional[str]:
        lower_map = {str(col).strip().lower(): col for col in df.columns}
        for variant in variants:
            if variant.lower() in lower_map:
                return lower_map[variant.lower()]
        return None

    def _convert_dataframe_to_decisions(self, df: pd.DataFrame, source_file: Path) -> list[AbbreviationDecision]:
        # Поддерживаем как компактную таблицу рекомендаций из stage 3:
        # term, suggested_abbreviation, abbreviation_found_in_text, need_to_introduce, priority, decision_score, reason
        # так и более общие форматы с abbreviation/long_form.
        abbr_col = self._first_existing_column(df, self.POSSIBLE_ABBR_COLUMNS)
        long_col = self._first_existing_column(df, self.POSSIBLE_LONG_COLUMNS)
        recommend_col = self._first_existing_column(df, self.POSSIBLE_RECOMMEND_COLUMNS)
        score_col = self._first_existing_column(df, self.POSSIBLE_SCORE_COLUMNS)
        note_col = self._first_existing_column(df, self.POSSIBLE_NOTE_COLUMNS)
        exists_col = self._first_existing_column(df, self.POSSIBLE_ALREADY_EXISTS_COLUMNS)

        if abbr_col is None or long_col is None:
            return []

        def _bool_from_value(value: Any) -> bool:
            raw = str(value).strip().lower()
            return raw in {"1", "true", "yes", "да", "нужно", "recommended"}

        results: list[AbbreviationDecision] = []

        for _, row in df.iterrows():
            abbreviation = str(row.get(abbr_col, "")).strip()
            long_form = str(row.get(long_col, "")).strip()

            if not abbreviation or not long_form or abbreviation.lower() == "nan" or long_form.lower() == "nan":
                continue

            recommended = True
            if recommend_col is not None:
                recommended = _bool_from_value(row.get(recommend_col, ""))

            already_exists = False
            if exists_col is not None:
                already_exists = _bool_from_value(row.get(exists_col, ""))

            score = ""
            if score_col is not None:
                score = str(row.get(score_col, "")).strip()

            note_parts = []
            if note_col is not None:
                note_val = str(row.get(note_col, "")).strip()
                if note_val and note_val.lower() != "nan":
                    note_parts.append(note_val)

            # Если есть и priority, и reason — объединим их в один комментарий
            priority_col = self._first_existing_column(df, ["priority"])
            reason_col = self._first_existing_column(df, ["reason"])
            if priority_col is not None and priority_col != note_col:
                pr = str(row.get(priority_col, "")).strip()
                if pr and pr.lower() != "nan":
                    note_parts.append(f"priority={pr}")
            if reason_col is not None and reason_col != note_col:
                rs = str(row.get(reason_col, "")).strip()
                if rs and rs.lower() != "nan":
                    note_parts.append(rs)

            note = " | ".join(dict.fromkeys(note_parts))

            selected = recommended and not already_exists

            results.append(
                AbbreviationDecision(
                    abbreviation=abbreviation,
                    long_form=long_form,
                    recommended=recommended,
                    selected=selected,
                    source_file=str(source_file.name),
                    note=note,
                    score=score,
                    already_exists=already_exists,
                )
            )

        return results


class EndUserApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Интерфейс обработки сокращений")
        self.geometry("1060x760")
        self.minsize(900, 620)

        self.source_docx_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.output_name_var = tk.StringVar(value="processed_document")
        self.insertion_mode_var = tk.StringVar(value="append_to_end")
        self.mode_display_var = tk.StringVar(value="В конец документа")
        self.marker_text_var = tk.StringVar()
        self.existing_section_title_var = tk.StringVar(value="Перечень обозначений и сокращений")
        self.create_separate_file_var = tk.BooleanVar(value=False)
        self.run_repeated_replacement_var = tk.BooleanVar(value=True)

        self.decisions: list[AbbreviationDecision] = []
        self.adapter = ProjectAdapter(self.log)

        self._build_ui()
        self._refresh_mode_fields()
        self.after(100, self._bind_mousewheel)

    def _build_ui(self) -> None:
        # Внешний контейнер
        outer = ttk.Frame(self, padding=6)
        outer.pack(fill="both", expand=True)

        # Холст со скроллами
        self.canvas = tk.Canvas(outer, highlightthickness=0)
        self.v_scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.h_scrollbar = ttk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)

        self.canvas.configure(
            yscrollcommand=self.v_scrollbar.set,
            xscrollcommand=self.h_scrollbar.set,
        )

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")

        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        # Внутренний прокручиваемый фрейм
        self.scrollable_frame = ttk.Frame(self.canvas, padding=6)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        top_frame = ttk.LabelFrame(self.scrollable_frame, text="1. Основные параметры", padding=12)
        top_frame.pack(fill="x", pady=(0, 10))
        self._build_path_section(top_frame)

        middle_frame = ttk.LabelFrame(self.scrollable_frame, text="2. Выбор новых сокращений для ввода", padding=12)
        middle_frame.pack(fill="both", expand=True, pady=(0, 10))
        self._build_decisions_section(middle_frame)

        bottom_frame = ttk.LabelFrame(self.scrollable_frame, text="3. Параметры обработки и сохранения", padding=12)
        bottom_frame.pack(fill="x", pady=(0, 10))
        self._build_output_section(bottom_frame)

        log_frame = ttk.LabelFrame(self.scrollable_frame, text="4. Логи", padding=12)
        log_frame.pack(fill="both", expand=True)
        self._build_log_section(log_frame)

    def _build_path_section(self, parent: ttk.LabelFrame) -> None:
        parent.columnconfigure(1, weight=1)

        ttk.Label(parent, text="Word-файл пользователя:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=self.source_docx_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Выбрать", command=self._choose_source_docx).grid(row=0, column=2, padx=(8, 0), pady=4)

        buttons_frame = ttk.Frame(parent)
        buttons_frame.grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 0))

        ttk.Button(
            buttons_frame,
            text="Запустить анализ и загрузить варианты",
            command=self._run_analysis,
        ).pack(side="left")

        ttk.Button(
            buttons_frame,
            text="Обновить список из stage 3",
            command=self._reload_candidates_only,
        ).pack(side="left", padx=(8, 0))

    def _build_decisions_section(self, parent: ttk.LabelFrame) -> None:
        ttk.Label(
            parent,
            text=(
                "В таблице выбираются только НОВЫЕ сокращения, которые пользователь хочет ввести. "
                "Сокращения, уже найденные в документе, будут добавлены в итоговый список автоматически."
            ),
        ).pack(anchor="w", pady=(0, 8))

        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", pady=(0, 8))

        ttk.Button(toolbar, text="Выбрать всё", command=self._select_all).pack(side="left")
        ttk.Button(toolbar, text="Снять всё", command=self._unselect_all).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Инвертировать", command=self._invert_selection).pack(side="left", padx=(8, 0))
        ttk.Label(toolbar, text="Двойной щелчок по строке переключает выбор пользователя.").pack(side="left", padx=(16, 0))

        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True)

        columns = (
            "selected",
            "abbreviation",
            "long_form",
            "recommended",
            "already_exists",
            "score",
            "note",
            "source_file",
        )
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)

        self.tree.heading("selected", text="Выбор")
        self.tree.heading("abbreviation", text="Сокращение")
        self.tree.heading("long_form", text="Полная форма")
        self.tree.heading("recommended", text="Рекомендация")
        self.tree.heading("already_exists", text="Уже есть в документе")
        self.tree.heading("score", text="Оценка")
        self.tree.heading("note", text="Комментарий")
        self.tree.heading("source_file", text="Источник")

        self.tree.column("selected", width=70, anchor="center", stretch=False)
        self.tree.column("abbreviation", width=140, anchor="center", stretch=False)
        self.tree.column("long_form", width=360, anchor="w", stretch=False)
        self.tree.column("recommended", width=120, anchor="center", stretch=False)
        self.tree.column("already_exists", width=140, anchor="center", stretch=False)
        self.tree.column("score", width=90, anchor="center", stretch=False)
        self.tree.column("note", width=420, anchor="w", stretch=False)
        self.tree.column("source_file", width=180, anchor="center", stretch=False)

        self.tree_v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree_h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.tree_v_scroll.set, xscrollcommand=self.tree_h_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree_v_scroll.grid(row=0, column=1, sticky="ns")
        self.tree_h_scroll.grid(row=1, column=0, sticky="ew")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._toggle_selected_row)

    def _build_output_section(self, parent: ttk.LabelFrame) -> None:
        parent.columnconfigure(1, weight=1)

        ttk.Label(parent, text="Куда сохранить результат:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=self.output_dir_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Выбрать", command=self._choose_output_dir).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(parent, text="Имя итогового файла:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=self.output_name_var).grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Label(parent, text="Место вставки списка сокращений:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        mode_combo = ttk.Combobox(
            parent,
            textvariable=self.mode_display_var,
            state="readonly",
            values=["В конец документа", "По маркеру", "В существующий раздел", "В отдельный файл"],
        )
        mode_combo.current(0)
        mode_combo.grid(row=2, column=1, sticky="ew", pady=4)
        mode_combo.bind("<<ComboboxSelected>>", self._on_mode_changed)

        self.marker_label = ttk.Label(parent, text="Текст маркера:")
        self.marker_label.grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        self.marker_entry = ttk.Entry(parent, textvariable=self.marker_text_var)
        self.marker_entry.grid(row=3, column=1, sticky="ew", pady=4)

        self.section_label = ttk.Label(parent, text="Название существующего раздела:")
        self.section_label.grid(row=4, column=0, sticky="w", padx=(0, 8), pady=4)
        self.section_entry = ttk.Entry(parent, textvariable=self.existing_section_title_var)
        self.section_entry.grid(row=4, column=1, sticky="ew", pady=4)

        ttk.Button(
            parent,
            text="Запустить финальную обработку",
            command=self._run_final_processing,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 6))

        ttk.Checkbutton(
            parent,
            text="Выполнять замену повторных объявлений на сокращения",
            variable=self.run_repeated_replacement_var,
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 2))

        ttk.Checkbutton(
            parent,
            text="Дополнительно сформировать отдельный Word-файл со списком сокращений",
            variable=self.create_separate_file_var,
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 8))

    def _build_log_section(self, parent: ttk.LabelFrame) -> None:
        self.log_text = tk.Text(parent, wrap="word", height=10)
        self.log_text.pack(fill="both", expand=True)

    def _on_frame_configure(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event) -> None:
        # Подстраиваем ширину внутреннего фрейма минимум под ширину холста,
        # но не сжимаем его, если контент шире окна.
        req_width = self.scrollable_frame.winfo_reqwidth()
        new_width = max(req_width, event.width)
        self.canvas.itemconfigure(self.canvas_window, width=new_width)

    def _bind_mousewheel(self) -> None:
        self.bind_all("<MouseWheel>", self._on_mousewheel_windows)
        self.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel_windows)
        self.bind_all("<Button-4>", self._on_mousewheel_linux_up)
        self.bind_all("<Button-5>", self._on_mousewheel_linux_down)

    def _on_mousewheel_windows(self, event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel_windows(self, event) -> None:
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux_up(self, _event) -> None:
        self.canvas.yview_scroll(-1, "units")

    def _on_mousewheel_linux_down(self, _event) -> None:
        self.canvas.yview_scroll(1, "units")

    def log(self, message: str) -> None:
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.update_idletasks()

    def _choose_source_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите Word-файл",
            filetypes=[("Word files", "*.docx")],
        )
        if path:
            self.source_docx_var.set(path)
            if not self.output_dir_var.get().strip():
                self.output_dir_var.set(str(Path(path).resolve().parent))
            self.log(f"Выбран исходный файл: {path}")

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Выберите папку для сохранения результата")
        if path:
            self.output_dir_var.set(path)
            self.log(f"Выбрана папка результата: {path}")

    def _reload_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for index, decision in enumerate(self.decisions):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    "Да" if decision.selected else "Нет",
                    decision.abbreviation,
                    decision.long_form,
                    "Да" if decision.recommended else "Нет",
                    "Да" if decision.already_exists else "Нет",
                    decision.score,
                    decision.note,
                    decision.source_file,
                ),
            )

    def _toggle_selected_row(self, _event=None) -> None:
        selected_item = self.tree.focus()
        if not selected_item:
            return
        index = int(selected_item)
        if self.decisions[index].already_exists:
            return
        self.decisions[index].selected = not self.decisions[index].selected
        self._reload_tree()

    def _select_all(self) -> None:
        for decision in self.decisions:
            if not decision.already_exists:
                decision.selected = True
        self._reload_tree()

    def _unselect_all(self) -> None:
        for decision in self.decisions:
            if not decision.already_exists:
                decision.selected = False
        self._reload_tree()

    def _invert_selection(self) -> None:
        for decision in self.decisions:
            if not decision.already_exists:
                decision.selected = not decision.selected
        self._reload_tree()

    def _on_mode_changed(self, _event=None) -> None:
        mapping = {
            "В конец документа": "append_to_end",
            "По маркеру": "by_marker",
            "В существующий раздел": "existing_section",
            "В отдельный файл": "separate_file",
        }
        self.insertion_mode_var.set(mapping.get(self.mode_display_var.get(), "append_to_end"))
        self._refresh_mode_fields()

    def _refresh_mode_fields(self) -> None:
        mode = self.insertion_mode_var.get()
        self._set_widget_state(self.marker_entry, mode == "by_marker")
        self._set_widget_state(self.section_entry, mode == "existing_section")

    @staticmethod
    def _set_widget_state(widget, enabled: bool) -> None:
        widget.configure(state="normal" if enabled else "disabled")

    def _validate_basic_paths(self) -> bool:
        if not self.source_docx_var.get().strip():
            messagebox.showerror("Ошибка", "Сначала выберите Word-файл пользователя.")
            return False

        source_docx = Path(self.source_docx_var.get().strip())
        if not source_docx.exists():
            messagebox.showerror("Ошибка", "Указанный Word-файл не существует.")
            return False

        return True

    def _run_analysis(self) -> None:
        try:
            if not self._validate_basic_paths():
                return

            source_docx = Path(self.source_docx_var.get().strip())

            self.log("Запущен анализ документа.")
            self.decisions = self.adapter.run_analysis(source_docx)
            self._reload_tree()

            self.log(f"Загружено вариантов: {len(self.decisions)}")
            messagebox.showinfo("Готово", "Анализ выполнен. Список вариантов загружен.")
        except Exception as exc:
            self.log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))

    def _reload_candidates_only(self) -> None:
        try:
            if not self._validate_basic_paths():
                return
            self.decisions = self.adapter.load_stage3_candidates()
            self._reload_tree()
            self.log(f"Обновлён список вариантов: {len(self.decisions)}")
            messagebox.showinfo("Готово", "Список вариантов обновлён.")
        except Exception as exc:
            self.log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))

    def _build_run_config(self) -> UiRunConfig:
        if not self.output_dir_var.get().strip():
            raise ValueError("Укажите папку для сохранения результата.")

        if not self.output_name_var.get().strip():
            raise ValueError("Укажите имя итогового файла.")

        if self.insertion_mode_var.get() == "by_marker" and not self.marker_text_var.get().strip():
            raise ValueError("Для режима вставки по маркеру укажите текст маркера.")

        if self.insertion_mode_var.get() == "existing_section" and not self.existing_section_title_var.get().strip():
            raise ValueError("Для режима existing_section укажите название существующего раздела.")

        selected = [
            asdict(item)
            for item in self.decisions
            if item.selected and not item.already_exists
        ]

        return UiRunConfig(
            source_docx=self.source_docx_var.get().strip(),
            output_dir=self.output_dir_var.get().strip(),
            output_name=self.output_name_var.get().strip(),
            insertion_mode=self.insertion_mode_var.get().strip(),
            marker_text=self.marker_text_var.get().strip(),
            existing_section_title=self.existing_section_title_var.get().strip(),
            create_separate_file=bool(self.create_separate_file_var.get()),
            run_repeated_replacement=bool(self.run_repeated_replacement_var.get()),
            selected_abbreviations=selected,
        )

    def _run_final_processing(self) -> None:
        try:
            if not self._validate_basic_paths():
                return

            config = self._build_run_config()

            self.log("Запущена финальная обработка документа.")
            self.adapter.run_final_processing(config)

            self.log("Финальная обработка завершена.")
            messagebox.showinfo("Готово", "Итоговый Word-файл сформирован.")
        except Exception as exc:
            self.log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))


if __name__ == "__main__":
    app = EndUserApp()
    app.mainloop()
