from __future__ import annotations

import json
import subprocess
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


@dataclass
class AbbreviationDecision:
    abbreviation: str
    long_form: str
    recommended: bool = True
    selected: bool = True
    source_file: str = ""
    note: str = ""
    score: str = ""
    already_exists_in_document: bool = False


@dataclass
class UiRunConfig:
    project_root: str
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
    def __init__(self, project_root: Path, logger) -> None:
        self.project_root = project_root
        self.logger = logger

    def run_analysis(self, source_docx: Path) -> list[AbbreviationDecision]:
        backend_script = self.project_root / "ui_backend_entry.py"
        if not backend_script.exists():
            raise FileNotFoundError(
                "В папке проекта не найден ui_backend_entry.py. "
                "Скопируйте backend-файл рядом с main.py."
            )

        self.logger("Запускаю анализ документа через ui_backend_entry.py")
        command = [
            sys.executable,
            str(backend_script),
            "--mode", "analyze",
            "--source-docx", str(source_docx),
        ]
        self._run_subprocess(command, cwd=self.project_root)
        return self.load_stage3_candidates()

    def load_stage3_candidates(self) -> list[AbbreviationDecision]:
        stage3_dir = self.project_root / "result_ui_session" / "stage3"
        csv_path = stage3_dir / "abbreviation_decisions.csv"
        xlsx_path = stage3_dir / "abbreviation_decisions.xlsx"

        candidate_path: Optional[Path] = None
        if xlsx_path.exists():
            candidate_path = xlsx_path
        elif csv_path.exists():
            candidate_path = csv_path

        if candidate_path is None:
            raise FileNotFoundError(
                "Не найден файл abbreviation_decisions.csv/.xlsx в result_ui_session/stage3. "
                "Сначала запустите анализ."
            )

        self.logger(f"Загружаю решения из файла: {candidate_path}")
        if candidate_path.suffix.lower() == ".csv":
            df = pd.read_csv(candidate_path, encoding="utf-8-sig")
        else:
            df = pd.read_excel(candidate_path)

        return self._convert_stage3_dataframe_to_decisions(df, candidate_path)

    def run_final_processing(self, config: UiRunConfig) -> None:
        backend_script = self.project_root / "ui_backend_entry.py"
        if not backend_script.exists():
            raise FileNotFoundError(
                "В папке проекта не найден ui_backend_entry.py. "
                "Скопируйте backend-файл рядом с main.py."
            )

        config_path = self.project_root / "ui_user_run_config.json"
        config_path.write_text(
            json.dumps(asdict(config), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self.logger(f"Сохранён конфиг запуска: {config_path}")

        command = [
            sys.executable,
            str(backend_script),
            "--mode", "process",
            "--config", str(config_path),
        ]
        self._run_subprocess(command, cwd=self.project_root)

    def _run_subprocess(self, command: list[str], cwd: Path) -> None:
        self.logger("Выполняется команда:")
        self.logger(" ".join(command))

        completed = subprocess.run(
            command,
            cwd=str(cwd),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        if completed.stdout:
            self.logger("STDOUT:")
            self.logger(completed.stdout.strip())

        if completed.stderr:
            self.logger("STDERR:")
            self.logger(completed.stderr.strip())

        if completed.returncode != 0:
            raise RuntimeError(
                f"Команда завершилась с кодом {completed.returncode}. "
                "Подробности смотрите в окне логов."
            )

    def _convert_stage3_dataframe_to_decisions(
        self,
        df: pd.DataFrame,
        source_file: Path,
    ) -> list[AbbreviationDecision]:
        required = {"term", "suggested_abbreviation"}
        if not required.issubset(df.columns):
            raise ValueError(
                "Файл stage 3 не содержит ожидаемых колонок term и suggested_abbreviation."
            )

        decisions: list[AbbreviationDecision] = []

        for _, row in df.iterrows():
            term = str(row.get("term", "")).strip()
            abbreviation = str(row.get("suggested_abbreviation", "")).strip()
            if not term or not abbreviation or term.lower() == "nan" or abbreviation.lower() == "nan":
                continue

            need_to_introduce = bool(row.get("need_to_introduce", False))
            already_exists = bool(row.get("abbreviation_found_in_text", False))

            score = str(row.get("decision_score", "")).strip()
            priority = str(row.get("priority", "")).strip()
            reason = str(row.get("reason", "")).strip()

            note_parts: list[str] = []
            if already_exists:
                note_parts.append("Аббревиатура уже присутствует в документе и будет добавлена в список автоматически.")
            if priority:
                note_parts.append(f"Приоритет: {priority}")
            if reason:
                note_parts.append(reason)

            decisions.append(
                AbbreviationDecision(
                    abbreviation=abbreviation,
                    long_form=term,
                    recommended=need_to_introduce,
                    selected=need_to_introduce,
                    source_file=source_file.name,
                    note=" ".join(note_parts).strip(),
                    score=score,
                    already_exists_in_document=already_exists,
                )
            )

        decisions.sort(
            key=lambda item: (
                item.already_exists_in_document,
                item.recommended,
                len(item.long_form),
                item.abbreviation,
            ),
            reverse=True,
        )
        return decisions


class UserInterfaceApp(tk.Tk):
    MODE_TO_CODE = {
        "В конец документа": "append_to_end",
        "По маркеру": "by_marker",
        "В существующий раздел сокращений": "existing_section",
        "Отдельным файлом": "separate_file",
    }

    def __init__(self) -> None:
        super().__init__()

        self.title("Интерфейс обработки сокращений")
        self.geometry("1320x860")
        self.minsize(1180, 780)

        self.project_root_var = tk.StringVar()
        self.source_docx_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.output_name_var = tk.StringVar(value="processed_document")
        self.insertion_mode_label_var = tk.StringVar(value="В конец документа")
        self.marker_text_var = tk.StringVar()
        self.existing_section_title_var = tk.StringVar(value="Перечень обозначений и сокращений")
        self.create_separate_file_var = tk.BooleanVar(value=False)
        self.run_repeated_replacement_var = tk.BooleanVar(value=True)

        self.decisions: list[AbbreviationDecision] = []
        self.adapter: Optional[ProjectAdapter] = None

        self._build_ui()
        self._refresh_mode_fields()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        top_frame = ttk.LabelFrame(root, text="1. Пути и основные параметры", padding=12)
        top_frame.pack(fill="x", pady=(0, 10))
        self._build_path_section(top_frame)

        middle_frame = ttk.LabelFrame(root, text="2. Выбор новых сокращений для ввода", padding=12)
        middle_frame.pack(fill="both", expand=True, pady=(0, 10))
        self._build_decisions_section(middle_frame)

        bottom_frame = ttk.LabelFrame(root, text="3. Параметры обработки и сохранения", padding=12)
        bottom_frame.pack(fill="x", pady=(0, 10))
        self._build_output_section(bottom_frame)

        log_frame = ttk.LabelFrame(root, text="4. Логи", padding=12)
        log_frame.pack(fill="both", expand=True)
        self._build_log_section(log_frame)

    def _build_path_section(self, parent: ttk.LabelFrame) -> None:
        parent.columnconfigure(1, weight=1)

        ttk.Label(parent, text="Папка проекта:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=self.project_root_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Выбрать", command=self._choose_project_root).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(parent, text="Word-файл пользователя:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=self.source_docx_var).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Выбрать", command=self._choose_source_docx).grid(row=1, column=2, padx=(8, 0), pady=4)

        buttons_frame = ttk.Frame(parent)
        buttons_frame.grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))

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
        info = ttk.Label(
            parent,
            text=(
                "В таблице выбираются только НОВЫЕ сокращения, которые пользователь хочет ввести. "
                "Сокращения, уже найденные в документе, будут добавлены в итоговый список автоматически."
            ),
            wraplength=1120,
            justify="left",
        )
        info.pack(anchor="w", pady=(0, 8))

        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", pady=(0, 8))

        ttk.Button(toolbar, text="Выбрать всё", command=self._select_all).pack(side="left")
        ttk.Button(toolbar, text="Снять всё", command=self._unselect_all).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Инвертировать", command=self._invert_selection).pack(side="left", padx=(8, 0))
        ttk.Label(toolbar, text="Двойной щелчок по строке переключает выбор пользователя.").pack(side="left", padx=(16, 0))

        columns = ("selected", "abbreviation", "long_form", "recommended", "already_exists", "score", "note", "source_file")
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", height=15)

        self.tree.heading("selected", text="Выбор")
        self.tree.heading("abbreviation", text="Сокращение")
        self.tree.heading("long_form", text="Полная форма")
        self.tree.heading("recommended", text="Рекомендация")
        self.tree.heading("already_exists", text="Уже есть в документе")
        self.tree.heading("score", text="Оценка")
        self.tree.heading("note", text="Комментарий")
        self.tree.heading("source_file", text="Источник")

        self.tree.column("selected", width=70, anchor="center")
        self.tree.column("abbreviation", width=140, anchor="center")
        self.tree.column("long_form", width=360, anchor="w")
        self.tree.column("recommended", width=110, anchor="center")
        self.tree.column("already_exists", width=150, anchor="center")
        self.tree.column("score", width=90, anchor="center")
        self.tree.column("note", width=310, anchor="w")
        self.tree.column("source_file", width=120, anchor="center")

        self.tree.pack(fill="both", expand=True)
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
            textvariable=self.insertion_mode_label_var,
            state="readonly",
            values=list(self.MODE_TO_CODE.keys()),
        )
        mode_combo.grid(row=2, column=1, sticky="ew", pady=4)
        mode_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_mode_fields())

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
        self.log_text = tk.Text(parent, wrap="word", height=12)
        self.log_text.pack(fill="both", expand=True)

    def log(self, message: str) -> None:
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.update_idletasks()

    def _choose_project_root(self) -> None:
        path = filedialog.askdirectory(title="Выберите папку проекта")
        if path:
            self.project_root_var.set(path)
            self.log(f"Выбрана папка проекта: {path}")

    def _choose_source_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите Word-файл",
            filetypes=[("Word files", "*.docx")],
        )
        if path:
            self.source_docx_var.set(path)
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
                    "Да" if decision.already_exists_in_document else "Нет",
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
        if self.decisions[index].already_exists_in_document:
            messagebox.showinfo(
                "Информация",
                "Это сокращение уже найдено в документе. "
                "Оно будет добавлено в итоговый список автоматически."
            )
            return

        self.decisions[index].selected = not self.decisions[index].selected
        self._reload_tree()

    def _select_all(self) -> None:
        for decision in self.decisions:
            if not decision.already_exists_in_document:
                decision.selected = True
        self._reload_tree()

    def _unselect_all(self) -> None:
        for decision in self.decisions:
            if not decision.already_exists_in_document:
                decision.selected = False
        self._reload_tree()

    def _invert_selection(self) -> None:
        for decision in self.decisions:
            if not decision.already_exists_in_document:
                decision.selected = not decision.selected
        self._reload_tree()

    def _refresh_mode_fields(self) -> None:
        mode_code = self.MODE_TO_CODE[self.insertion_mode_label_var.get()]
        self._set_widget_state(self.marker_entry, mode_code == "by_marker")
        self._set_widget_state(self.section_entry, mode_code == "existing_section")

    @staticmethod
    def _set_widget_state(widget, enabled: bool) -> None:
        widget.configure(state="normal" if enabled else "disabled")

    def _validate_basic_paths(self) -> bool:
        if not self.project_root_var.get().strip():
            messagebox.showerror("Ошибка", "Сначала выберите папку проекта.")
            return False

        if not self.source_docx_var.get().strip():
            messagebox.showerror("Ошибка", "Сначала выберите Word-файл пользователя.")
            return False

        project_root = Path(self.project_root_var.get().strip())
        if not project_root.exists():
            messagebox.showerror("Ошибка", "Указанная папка проекта не существует.")
            return False

        source_docx = Path(self.source_docx_var.get().strip())
        if not source_docx.exists():
            messagebox.showerror("Ошибка", "Указанный Word-файл не существует.")
            return False

        return True

    def _get_adapter(self) -> ProjectAdapter:
        current_root = Path(self.project_root_var.get().strip())
        if self.adapter is None or self.adapter.project_root != current_root:
            self.adapter = ProjectAdapter(current_root, self.log)
        return self.adapter

    def _run_analysis(self) -> None:
        try:
            if not self._validate_basic_paths():
                return

            adapter = self._get_adapter()
            source_docx = Path(self.source_docx_var.get().strip())

            self.log("Запущен анализ документа.")
            self.decisions = adapter.run_analysis(source_docx)
            self._reload_tree()

            self.log(f"Загружено строк для выбора: {len(self.decisions)}")
            messagebox.showinfo("Готово", "Анализ выполнен. Варианты сокращений загружены.")
        except Exception as exc:
            self.log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))

    def _reload_candidates_only(self) -> None:
        try:
            if not self._validate_basic_paths():
                return

            adapter = self._get_adapter()
            self.decisions = adapter.load_stage3_candidates()
            self._reload_tree()

            self.log(f"Список обновлён. Всего строк: {len(self.decisions)}")
            messagebox.showinfo("Готово", "Список сокращений обновлён.")
        except Exception as exc:
            self.log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))

    def _build_run_config(self) -> UiRunConfig:
        if not self.output_dir_var.get().strip():
            raise ValueError("Укажите папку для сохранения результата.")

        if not self.output_name_var.get().strip():
            raise ValueError("Укажите имя итогового файла.")

        mode_code = self.MODE_TO_CODE[self.insertion_mode_label_var.get()]

        if mode_code == "by_marker" and not self.marker_text_var.get().strip():
            raise ValueError("Для режима вставки по маркеру укажите текст маркера.")

        if mode_code == "existing_section" and not self.existing_section_title_var.get().strip():
            raise ValueError("Для режима вставки в существующий раздел укажите название раздела.")

        selected = [
            asdict(item)
            for item in self.decisions
            if item.selected and not item.already_exists_in_document
        ]

        return UiRunConfig(
            project_root=self.project_root_var.get().strip(),
            source_docx=self.source_docx_var.get().strip(),
            output_dir=self.output_dir_var.get().strip(),
            output_name=self.output_name_var.get().strip(),
            insertion_mode=mode_code,
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
            adapter = self._get_adapter()

            self.log("Запущена финальная обработка документа.")
            adapter.run_final_processing(config)

            self.log("Финальная обработка завершена.")
            messagebox.showinfo("Готово", "Документ успешно обработан.")
        except Exception as exc:
            self.log(f"ОШИБКА: {exc}")
            messagebox.showerror("Ошибка", str(exc))


if __name__ == "__main__":
    app = UserInterfaceApp()
    app.mainloop()
