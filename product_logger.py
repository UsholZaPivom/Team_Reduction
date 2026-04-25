from __future__ import annotations

"""
product_logger.py

Модуль логирования действий продукта при обработке документов Word.

Назначение:
- выводить понятные пользователю логи;
- сохранять технические сообщения для отладки;
- фиксировать изменения в документе;
- фиксировать ошибки по этапам обработки.

Пример использования:

    from product_logger import ProductLogger

    logger = ProductLogger(log_dir="logs", console_output=True)
    logger.start_session("input.docx")
    logger.stage_started("Этап 1")
    logger.info("Документ успешно загружен")
    logger.log_replacement(
        old_text="Федеральная служба безопасности",
        new_text="ФСБ",
        page=2,
        line=13
    )
    logger.stage_finished("Этап 1", details="Найдено 24 кандидата")
    logger.close()
"""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional
import traceback


@dataclass
class LogEvent:
    timestamp: str
    level: str
    message: str


class ProductLogger:
    """
    Универсальный логгер для проекта.

    Параметры:
    - log_dir: папка, где будут храниться .log файлы;
    - console_output: печатать ли сообщения ещё и в консоль;
    - file_prefix: префикс имени лог-файла.
    """

    def __init__(
        self,
        log_dir: str | Path = "logs",
        console_output: bool = True,
        file_prefix: str = "product_log"
    ) -> None:
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

        self.console_output = console_output
        self.session_started_at = datetime.now()

        timestamp = self.session_started_at.strftime("%Y-%m-%d_%H-%M-%S")
        self.log_path = self.log_dir / f"{file_prefix}_{timestamp}.log"

        self._write_raw_line("=" * 80)
        self._write_raw_line("ЛОГ РАБОТЫ ПРОДУКТА")
        self._write_raw_line(f"Начало сессии: {self.session_started_at.strftime('%Y-%m-%d %H:%M:%S')}")
        self._write_raw_line("=" * 80)

    def _current_timestamp(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _write_raw_line(self, text: str) -> None:
        with self.log_path.open("a", encoding="utf-8") as f:
            f.write(text + "\n")

        if self.console_output:
            print(text)

    def _write_event(self, level: str, message: str) -> None:
        event = LogEvent(
            timestamp=self._current_timestamp(),
            level=level.upper(),
            message=message
        )
        line = f"[{event.timestamp}] [{event.level}] {event.message}"
        self._write_raw_line(line)

    def info(self, message: str) -> None:
        self._write_event("INFO", message)

    def warning(self, message: str) -> None:
        self._write_event("WARNING", message)

    def error(self, message: str) -> None:
        self._write_event("ERROR", message)

    def debug(self, message: str) -> None:
        self._write_event("DEBUG", message)

    def action(self, message: str) -> None:
        self._write_event("ACTION", message)

    def start_session(self, document_path: str | Path) -> None:
        self.info(f"Начата обработка документа: {Path(document_path)}")

    def finish_session(self, success: bool = True) -> None:
        finished_at = datetime.now()
        duration = finished_at - self.session_started_at

        status_text = "обработка завершена успешно" if success else "обработка завершена с ошибками"
        self.info(f"Сессия завершена: {status_text}")
        self.info(f"Длительность сессии: {duration}")
        self._write_raw_line("=" * 80)

    def close(self) -> None:
        self.finish_session(success=True)

    def stage_started(self, stage_name: str) -> None:
        self.info(f"Запущен этап: {stage_name}")

    def stage_finished(self, stage_name: str, details: Optional[str] = None) -> None:
        if details:
            self.info(f"Этап завершён: {stage_name}. {details}")
        else:
            self.info(f"Этап завершён: {stage_name}")

    def stage_failed(self, stage_name: str, error_text: str) -> None:
        self.error(f"Ошибка на этапе '{stage_name}': {error_text}")

    def log_document_loaded(self, document_path: str | Path) -> None:
        self.info(f"Документ загружен: {Path(document_path)}")

    def log_candidates_found(self, count: int) -> None:
        self.info(f"Найдено кандидатов на сокращение: {count}")

    def log_abbreviations_found(self, count: int) -> None:
        self.info(f"Найдено существующих аббревиатур: {count}")

    def log_database_loaded(self, db_path: str | Path, entries_count: int) -> None:
        self.info(f"База аббревиатур загружена: {Path(db_path)}; записей: {entries_count}")

    def log_database_updated(self, added_count: int, updated_count: int) -> None:
        self.info(
            f"База аббревиатур обновлена: добавлено {added_count}, "
            f"обновлено {updated_count}"
        )

    def log_declaration_error(
        self,
        abbreviation: str,
        long_form: str,
        page: Optional[int] = None,
        line: Optional[int] = None
    ) -> None:
        location = self._format_location(page=page, line=line)
        if location:
            self.warning(
                f'Ошибочное объявление сокращения {location}: "{long_form}" -> "{abbreviation}" '
                f"(сокращение далее не используется)"
            )
        else:
            self.warning(
                f'Ошибочное объявление сокращения: "{long_form}" -> "{abbreviation}" '
                f"(сокращение далее не используется)"
            )

    def log_replacement(
        self,
        old_text: str,
        new_text: str,
        page: Optional[int] = None,
        line: Optional[int] = None
    ) -> None:
        location = self._format_location(page=page, line=line)

        if location:
            message = f'{location} заменена фраза "{old_text}" на "{new_text}"'
        else:
            message = f'Заменена фраза "{old_text}" на "{new_text}"'

        self.action(message)

    def log_list_created(self, output_path: str | Path) -> None:
        self.action(f"Создан отдельный файл со списком сокращений: {Path(output_path)}")

    def log_list_inserted(self, destination_path: str | Path, position_description: str) -> None:
        self.action(
            f"Перечень обозначений и сокращений вставлен в документ "
            f"{Path(destination_path)}; позиция вставки: {position_description}"
        )

    def log_existing_list_updated(self, document_path: str | Path) -> None:
        self.action(
            f"Обновлён существующий перечень обозначений и сокращений в документе: "
            f"{Path(document_path)}"
        )

    def log_exception(
        self,
        stage_name: str,
        exc: Exception,
        with_traceback: bool = True
    ) -> None:
        self.error(f"Исключение на этапе '{stage_name}': {exc}")

        if with_traceback:
            self._write_raw_line("TRACEBACK:")
            self._write_raw_line(traceback.format_exc().rstrip())

    def _format_location(
        self,
        page: Optional[int] = None,
        line: Optional[int] = None
    ) -> str:
        if page is not None and line is not None:
            return f"На странице {page}, строке {line}"
        if page is not None:
            return f"На странице {page}"
        if line is not None:
            return f"В строке {line}"
        return ""

    def get_log_path(self) -> Path:
        return self.log_path


if __name__ == "__main__":
    logger = ProductLogger(log_dir="logs", console_output=True)
    logger.start_session("example_document.docx")

    logger.stage_started("Этап 1")
    logger.log_document_loaded("example_document.docx")
    logger.log_candidates_found(24)
    logger.stage_finished("Этап 1", details="Подготовлены кандидаты на сокращение")

    logger.stage_started("Этап 2")
    logger.log_abbreviations_found(11)
    logger.log_declaration_error(
        abbreviation="ФСБ",
        long_form="Федеральная служба безопасности",
        page=2,
        line=13
    )
    logger.stage_finished("Этап 2", details="Проверены существующие аббревиатуры")

    logger.stage_started("Этап 3")
    logger.log_replacement(
        old_text="Федеральная служба безопасности",
        new_text="ФСБ",
        page=2,
        line=13
    )
    logger.log_list_created("Перечень_сокращений.docx")
    logger.stage_finished("Этап 3", details="Сформирован перечень обозначений и сокращений")

    logger.close()
