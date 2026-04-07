from __future__ import annotations

"""
text_recognition_candidates_v3.py

Этап 1 проекта:
распознавание текста документа и выделение словоформ / словосочетаний,
которые потенциально поддаются сокращению.

Что делает эта версия:
1. Читает .docx-документ Microsoft Word.
2. Извлекает текст из заголовков, обычных абзацев и таблиц.
3. Делит текст на предложения.
4. Выполняет морфологический анализ слов.
5. Находит терминоподобные словосочетания, которые могут сокращаться.
6. Уменьшает шум за счёт более строгих правил фильтрации.
7. Сохраняет "сырые" и агрегированные результаты в CSV и, при наличии openpyxl, в XLSX.

Что НЕ делает эта версия:
- не ищет уже существующие аббревиатуры;
- не сопоставляет полную форму с аббревиатурой;
- не решает, нужно ли вводить аббревиатуру;
- не изменяет документ Word.

Это именно первый изолированный шаг конвейера.
"""

# ============================================================
# 1. Импорт библиотек
# ============================================================

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Tuple

# Библиотека для чтения .docx-файлов Word
from docx import Document

# Более мощные регулярные выражения
import regex

# Табличное представление и сохранение результата
import pandas as pd

# Natasha используется для сегментации текста на предложения
from natasha import Doc, Segmenter

# Морфологический анализатор русского языка
from pymorphy2 import MorphAnalyzer


# ============================================================
# 2. Структуры данных
# ============================================================

@dataclass
class TextFragment:
    """
    Описывает один текстовый фрагмент, извлечённый из документа.

    Пояснение:
    документ состоит из разных типов текстовых областей.
    Для отладки и последующего анализа важно понимать,
    откуда именно был получен текст:
    - из заголовка,
    - из обычного абзаца,
    - из ячейки таблицы.
    """
    source_type: str      # heading / paragraph / table_cell
    source_index: int     # порядковый номер внутри своего типа
    text: str             # сам текст фрагмента


@dataclass
class CandidateMention:
    """
    Описывает одно найденное вхождение словосочетания,
    которое потенциально можно сократить.

    Это НЕ "готовая аббревиатура", а только кандидат.
    """
    source_type: str
    source_index: int
    sentence: str
    phrase: str
    normalized_phrase: str
    proposed_abbreviation: str
    word_count: int


# ============================================================
# 3. Основной класс
# ============================================================

class ReducibleWordformRecognizerV3:
    """
    Основной класс, реализующий первый этап проекта:
    поиск словоформ и словосочетаний, поддающихся сокращению.

    В версии V3 дополнительно усилена фильтрация шума:
    - отбрасываются слишком общие / контекстные конструкции;
    - отбрасываются вложенные подфразы, если есть более полная;
    - отбрасываются явно "метатекстовые" словосочетания
      про сам отчёт, тестирование, проверку и т.д.
    """

    def __init__(self) -> None:
        """
        Инициализация NLP-инструментов и правил фильтрации.
        """
        # Сегментатор Natasha для деления текста на предложения
        self.segmenter = Segmenter()

        # Морфологический анализатор русского языка
        self.morph = MorphAnalyzer()

        # Регулярное выражение для выделения слов.
        # Разрешаем русские и латинские буквы, а также дефис.
        self.word_pattern = regex.compile(r"[A-Za-zА-Яа-яЁё-]+")

        # Части речи, которые чаще всего формируют термины.
        # Например:
        # - автоматизированная (ADJF)
        # - система (NOUN)
        # - мониторинга (NOUN)
        # - событий (NOUN)
        self.allowed_pos = {"NOUN", "ADJF", "ADJS", "PRTF", "PRTS"}

        # Базовые стоп-леммы: слишком служебные слова,
        # которые не должны становиться основой термина.
        self.stop_lemmas = {
            "и", "или", "в", "во", "на", "по", "с", "со", "к", "ко",
            "у", "о", "об", "от", "до", "за", "из", "под", "над",
            "при", "для", "не", "это", "тот", "такой", "данный",
            "этот", "который", "как", "а", "но", "же"
        }

        # Общие леммы, которые часто встречаются в отчётах и описаниях,
        # но редко являются предметными терминами для сокращения.
        # Такие слова дают "метатекстовый шум".
        self.generic_lemmas = {
            "задача", "результат", "пример", "этап", "область",
            "требование", "метод", "процесс", "алгоритм",
            "описание", "проект", "работа", "часть", "случай",
            "средство", "способ", "документ", "файл", "система"
        }

        # Специальный список контекстных / шумовых слов,
        # по которым можно понять, что найдено не предметное понятие,
        # а служебная конструкция текста отчёта.
        #
        # Примеры того, что хотим убрать:
        # - "Существующая аббревиатура"
        # - "отчёте модуль анализа ..."
        # - "тестировании автоматизированная система ..."
        self.context_lemmas = {
            "отчёт", "документ", "тестирование", "проверка", "задача",
            "пример", "этап", "раздел", "предложение", "сочетание",
            "термин", "аббревиатура", "введение", "вывод", "анализ",
            "необходимость", "определение", "последующий", "существующий",
            "поздний", "короткий", "новый", "первый", "данный",
            "следующий", "описание", "проект", "работа"
        }

        # Дополнительный список лемм, которые характерны
        # не для предметных терминов, а для формулировок задания,
        # процессов и организационных описаний.
        self.extra_noise_lemmas = {
            "рабочий", "группа", "ввод", "сокращение",
            "распознавание", "выделение"
        }

        # Слишком общие "головы" коротких словосочетаний.
        # Если двухсловная фраза начинается с такого слова,
        # она часто оказывается неполным куском более длинного термина.
        self.weak_head_lemmas = {
            "модуль", "блок", "часть", "уровень", "элемент", "подсистема"
        }

    # ========================================================
    # 4. Работа с документом
    # ========================================================

    def load_docx_fragments(self, docx_path: str | Path) -> List[TextFragment]:
        """
        Извлекает текстовые фрагменты из Word-документа.

        На текущем этапе извлекаются:
        - заголовки;
        - обычные абзацы;
        - текст из таблиц.

        Почему так:
        для первого этапа достаточно покрыть основной поток текста.
        Сноски можно добавить отдельным расширением позже.
        """
        docx_path = Path(docx_path)
        doc = Document(docx_path)

        fragments: List[TextFragment] = []

        paragraph_index = 0
        heading_index = 0
        table_cell_index = 0

        # -------- Извлечение абзацев и заголовков --------
        for paragraph in doc.paragraphs:
            text = self._clean_text(paragraph.text)
            if not text:
                continue

            style_name = ""
            if paragraph.style is not None and paragraph.style.name:
                style_name = paragraph.style.name.lower()

            # Если стиль содержит "heading" или "заголов",
            # считаем фрагмент заголовком.
            if "heading" in style_name or "заголов" in style_name:
                fragments.append(
                    TextFragment(
                        source_type="heading",
                        source_index=heading_index,
                        text=text
                    )
                )
                heading_index += 1
            else:
                fragments.append(
                    TextFragment(
                        source_type="paragraph",
                        source_index=paragraph_index,
                        text=text
                    )
                )
                paragraph_index += 1

        # -------- Извлечение текста из таблиц --------
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = self._clean_text(cell.text)
                    if not text:
                        continue

                    fragments.append(
                        TextFragment(
                            source_type="table_cell",
                            source_index=table_cell_index,
                            text=text
                        )
                    )
                    table_cell_index += 1

        return fragments

    # ========================================================
    # 5. Предобработка текста
    # ========================================================

    def _clean_text(self, text: str) -> str:
        """
        Базовая очистка текста:
        - замена переводов строк пробелами;
        - схлопывание повторяющихся пробелов;
        - удаление пробелов по краям.

        Это нужно, чтобы фразы сравнивались в более стабильном виде.
        """
        if not text:
            return ""

        text = text.replace("\n", " ")
        text = regex.sub(r"\s+", " ", text)
        return text.strip()

    def split_into_sentences(self, text: str) -> List[str]:
        """
        Делит текст на предложения с помощью Natasha.

        Natasha лучше подходит для русского текста, чем простое
        деление по точке, особенно когда в тексте есть сокращения,
        кавычки, числа и составные конструкции.
        """
        if not text.strip():
            return []

        doc = Doc(text)
        doc.segment(self.segmenter)

        sentences: List[str] = []
        for sent in doc.sents:
            sentence_text = sent.text.strip()
            if sentence_text:
                sentences.append(sentence_text)

        return sentences

    def extract_words(self, text: str) -> List[str]:
        """
        Извлекает слова из текста.

        Пример:
        "Автоматизированная система мониторинга событий"
        ->
        ["Автоматизированная", "система", "мониторинга", "событий"]
        """
        return self.word_pattern.findall(text)

    # ========================================================
    # 6. Морфологический анализ
    # ========================================================

    def parse_word(self, word: str):
        """
        Возвращает наиболее вероятный морфологический разбор слова.
        """
        return self.morph.parse(word)[0]

    def get_normal_form(self, word: str) -> str:
        """
        Возвращает нормальную форму слова (лемму).
        """
        return self.parse_word(word).normal_form

    def get_pos(self, word: str) -> str:
        """
        Возвращает часть речи слова.
        Если часть речи не определилась, возвращает пустую строку.
        """
        pos = self.parse_word(word).tag.POS
        return pos if pos else ""

    # ========================================================
    # 7. Фильтры уровня отдельных слов
    # ========================================================

    def is_existing_abbreviation_token(self, word: str) -> bool:
        """
        Проверяет, является ли слово уже готовой аббревиатурой.

        Примеры:
        - "АСМС"
        - "ЦОД"
        - "ООН"

        Такие слова не должны участвовать в построении "полной формы",
        потому что задача первого этапа — искать именно словоформы,
        которые можно сократить, а не уже существующие сокращения.
        """
        clean_word = word.strip("-")
        if len(clean_word) < 2:
            return False

        # Только буквы, без цифр и прочих символов
        if not regex.fullmatch(r"[A-Za-zА-Яа-яЁё-]+", clean_word):
            return False

        # Считаем аббревиатурой короткие слова в верхнем регистре
        if clean_word.isupper() and len(clean_word) <= 8:
            return True

        return False

    def is_content_word(self, word: str) -> bool:
        """
        Проверяет, можно ли считать слово содержательным
        для формирования термина.

        Условия:
        - слово не слишком короткое;
        - слово не является уже существующей аббревиатурой;
        - часть речи подходит;
        - лемма не является служебным словом.
        """
        # Очень короткие слова чаще всего не образуют термин
        if len(word) < 3:
            return False

        # Уже готовые аббревиатуры не рассматриваем как полную форму
        if self.is_existing_abbreviation_token(word):
            return False

        parse = self.parse_word(word)
        lemma = parse.normal_form
        pos = parse.tag.POS

        if lemma in self.stop_lemmas:
            return False

        if pos not in self.allowed_pos:
            return False

        return True

    # ========================================================
    # 8. Фильтры уровня словосочетания
    # ========================================================

    def is_context_noise_phrase(self, lemmas: List[str]) -> bool:
        """
        Проверяет, является ли фраза контекстным шумом,
        а не термином предметной области.

        Фраза считается шумовой, если:
        - первое слово уже задаёт контекст отчёта / тестирования;
        - внутри слишком много контекстных лемм.
        """
        if not lemmas:
            return True

        # Если первая лемма уже явно контекстная,
        # то фраза почти всегда шумовая.
        if lemmas[0] in self.context_lemmas:
            return True

        context_count = sum(1 for lemma in lemmas if lemma in self.context_lemmas)

        # Если контекстных лемм слишком много,
        # это уже описание текста, а не предметный термин.
        if context_count >= 2:
            return True

        return False

    def is_generic_noise_phrase(self, lemmas: List[str]) -> bool:
        """
        Отсекает слишком общие фразы.

        Принцип:
        если большинство слов во фразе относятся к слишком общим леммам,
        то это, скорее всего, не предметный термин.
        """
        if not lemmas:
            return True

        generic_count = sum(1 for lemma in lemmas if lemma in self.generic_lemmas)

        # Если общих слов слишком много,
        # фраза не подходит как хороший термин-кандидат.
        if generic_count >= max(2, len(lemmas) - 1):
            return True

        return False

    def is_process_noise_phrase(self, lemmas: List[str]) -> bool:
        """
        Отсекает процессные / метатекстовые словосочетания.

        Примеры того, что хотим убрать:
        - "ввода новых сокращений"
        - "распознавания словоформ выделения"
        - "рабочая группа"
        """
        if not lemmas:
            return True

        noise_count = sum(1 for lemma in lemmas if lemma in self.extra_noise_lemmas)
        if noise_count >= 2:
            return True

        return False

    def is_incomplete_phrase(self, lemmas: List[str]) -> bool:
        """
        Проверяет, является ли фраза неполной.

        Неполной считаем фразу, если:
        - она состоит всего из 2 слов;
        - первое слово слишком общее;
        - второе слово похоже на зависимое слово без главной части слева.

        Это помогает убрать случаи вроде:
        - "модуль анализа"
        - "текстовых документов"
        """
        if len(lemmas) != 2:
            return False

        # Общее первое слово часто означает, что перед нами только начало
        # более длинного термина, а не законченная полная форма.
        if lemmas[0] in self.weak_head_lemmas:
            return True

        # Такие вторые слова часто появляются как хвост более длинного термина.
        dependent_second_words = {"документ", "данные", "событие", "профиль"}
        if lemmas[1] in dependent_second_words:
            return True

        return False

    def is_term_like_chunk(self, chunk: List[Tuple[str, str, str]]) -> bool:
        """
        Проверяет, похож ли фрагмент на термин.

        chunk содержит кортежи:
        (слово, лемма, часть_речи)

        Правила в V3 более строгие:
        - длина от 2 до 5 слов;
        - последнее слово должно быть существительным;
        - должен быть терминный "каркас":
          либо 2+ существительных,
          либо существительное с определением;
        - фраза должна быть не слишком короткой по суммарной длине.
        """
        if len(chunk) < 2 or len(chunk) > 5:
            return False

        words = [word for word, _, _ in chunk]
        pos_list = [pos for _, _, pos in chunk]

        # Последнее слово в термине чаще всего является существительным:
        # "система", "документов", "событий", "данных"
        if pos_list[-1] != "NOUN":
            return False

        noun_count = sum(1 for pos in pos_list if pos == "NOUN")
        modifier_count = sum(1 for pos in pos_list if pos in {"ADJF", "ADJS", "PRTF", "PRTS"})

        # Терминный каркас:
        # 1) либо хотя бы два существительных
        # 2) либо существительное + определение
        if not (noun_count >= 2 or (noun_count >= 1 and modifier_count >= 1)):
            return False

        # Если фраза слишком короткая по общему числу символов,
        # это часто случайный шум.
        total_letters = sum(len(w) for w in words)
        if total_letters < 10:
            return False

        return True

    def build_abbreviation(self, words: List[str]) -> str:
        """
        Строит предлагаемую аббревиатуру по первым буквам слов.

        Пример:
        "Автоматизированная система мониторинга событий"
        -> "АСМС"
        """
        letters: List[str] = []
        for word in words:
            clean_word = word.strip("-")
            if clean_word:
                letters.append(clean_word[0].upper())

        return "".join(letters)

    # ========================================================
    # 9. Выделение кандидатов из предложения
    # ========================================================

    def extract_candidates_from_sentence(
        self,
        sentence: str,
        source_type: str,
        source_index: int
    ) -> List[CandidateMention]:
        """
        Извлекает из одного предложения все потенциальные
        словосочетания, поддающиеся сокращению.

        Общая схема:
        1. Достаём слова из предложения.
        2. Разбираем каждое слово морфологически.
        3. Собираем цепочки подряд идущих содержательных слов.
        4. Из каждой цепочки строим кандидатов.
        5. Оставляем только лучшие / максимальные варианты.
        """
        words = self.extract_words(sentence)
        if len(words) < 2:
            return []

        analyzed_words: List[Tuple[str, str, str]] = []

        # Морфологический разбор каждого слова
        for word in words:
            lemma = self.get_normal_form(word)
            pos = self.get_pos(word)
            analyzed_words.append((word, lemma, pos))

        # Собираем цепочки подряд идущих содержательных слов.
        # Если встречается слово, которое не проходит фильтр,
        # текущая цепочка завершается.
        candidate_chains: List[List[Tuple[str, str, str]]] = []
        current_chain: List[Tuple[str, str, str]] = []

        for word, lemma, pos in analyzed_words:
            if self.is_content_word(word):
                current_chain.append((word, lemma, pos))
            else:
                if current_chain:
                    candidate_chains.append(current_chain)
                    current_chain = []

        if current_chain:
            candidate_chains.append(current_chain)

        all_mentions: List[CandidateMention] = []

        # Из каждой цепочки извлекаем кандидатов
        for chain in candidate_chains:
            chain_mentions = self._extract_candidates_from_chain(
                chain=chain,
                sentence=sentence,
                source_type=source_type,
                source_index=source_index
            )
            all_mentions.extend(chain_mentions)

        return all_mentions

    def _extract_candidates_from_chain(
        self,
        chain: List[Tuple[str, str, str]],
        sentence: str,
        source_type: str,
        source_index: int
    ) -> List[CandidateMention]:
        """
        Из одной цепочки содержательных слов генерирует
        терминоподобные словосочетания.

        В версии V3 после генерации всех допустимых окон
        дополнительно оставляются только "максимальные" фразы,
        чтобы уменьшить количество вложенных подфраз.
        """
        candidates: List[Tuple[List[str], List[str], int, int]] = []

        max_window = min(5, len(chain))

        # Шаг 1. Генерация всех допустимых окон
        for start in range(len(chain)):
            for window_size in range(2, max_window + 1):
                end = start + window_size
                if end > len(chain):
                    break

                chunk = chain[start:end]

                if not self.is_term_like_chunk(chunk):
                    continue

                words = [word for word, _, _ in chunk]
                lemmas = [lemma for _, lemma, _ in chunk]

                # Дополнительная фильтрация контекстного шума
                if self.is_context_noise_phrase(lemmas):
                    continue

                # Дополнительная фильтрация слишком общих фраз
                if self.is_generic_noise_phrase(lemmas):
                    continue

                # Убираем процессные и организационные конструкции.
                if self.is_process_noise_phrase(lemmas):
                    continue

                # Убираем короткие неполные словосочетания,
                # которые являются лишь куском более длинного термина.
                if self.is_incomplete_phrase(lemmas):
                    continue

                candidates.append((words, lemmas, start, end))

        if not candidates:
            return []

        # Шаг 2. Оставляем только максимальные фразы,
        # а вложенные подфразы убираем.
        #
        # Например, если есть:
        # - "система мониторинга"
        # - "мониторинга событий"
        # - "автоматизированная система мониторинга событий"
        #
        # то хотим оставить прежде всего более полную конструкцию.
        maximal_candidates = self._keep_maximal_candidates(candidates)

        mentions: List[CandidateMention] = []

        # Шаг 3. Превращаем оставшиеся кандидаты в CandidateMention
        for words, lemmas, start, end in maximal_candidates:
            phrase = " ".join(words)
            normalized_phrase = " ".join(lemmas)
            proposed_abbreviation = self.build_abbreviation(words)

            mentions.append(
                CandidateMention(
                    source_type=source_type,
                    source_index=source_index,
                    sentence=sentence,
                    phrase=phrase,
                    normalized_phrase=normalized_phrase,
                    proposed_abbreviation=proposed_abbreviation,
                    word_count=len(words)
                )
            )

        return mentions

    def _keep_maximal_candidates(
        self,
        candidates: List[Tuple[List[str], List[str], int, int]]
    ) -> List[Tuple[List[str], List[str], int, int]]:
        """
        Убирает вложенные подфразы, если есть более полная фраза.

        Идея:
        если один кандидат целиком содержится внутри другого,
        и второй длиннее, то предпочитаем второй.
        """
        # Сортируем по длине окна: сначала самые длинные
        sorted_candidates = sorted(
            candidates,
            key=lambda item: (item[3] - item[2], len(" ".join(item[0]))),
            reverse=True
        )

        selected: List[Tuple[List[str], List[str], int, int]] = []

        for candidate in sorted_candidates:
            _, _, start, end = candidate

            is_nested = False
            for _, _, selected_start, selected_end in selected:
                # Если текущий кандидат полностью вложен
                # в уже выбранный более длинный кандидат,
                # то отбрасываем его.
                if selected_start <= start and end <= selected_end:
                    is_nested = True
                    break

            if not is_nested:
                selected.append(candidate)

        # Возвращаем кандидаты в естественном порядке следования
        # по исходной цепочке
        selected = sorted(selected, key=lambda item: item[2])
        return selected

    # ========================================================
    # 10. Обработка всего документа
    # ========================================================

    def analyze_document(self, docx_path: str | Path) -> List[CandidateMention]:
        """
        Полный конвейер первого этапа:
        - загрузка документа;
        - разбиение фрагментов на предложения;
        - извлечение кандидатов из каждого предложения.
        """
        fragments = self.load_docx_fragments(docx_path)

        all_mentions: List[CandidateMention] = []

        for fragment in fragments:
            sentences = self.split_into_sentences(fragment.text)

            for sentence in sentences:
                mentions = self.extract_candidates_from_sentence(
                    sentence=sentence,
                    source_type=fragment.source_type,
                    source_index=fragment.source_index
                )
                all_mentions.extend(mentions)

        return all_mentions

    # ========================================================
    # 11. Агрегация и финальная очистка
    # ========================================================

    def mentions_to_dataframe(self, mentions: List[CandidateMention]) -> pd.DataFrame:
        """
        Преобразует список найденных вхождений в DataFrame.

        Это полезно для детальной проверки каждого отдельного случая.
        """
        return pd.DataFrame([asdict(item) for item in mentions])

    def aggregate_mentions(self, mentions: List[CandidateMention]) -> pd.DataFrame:
        """
        Группирует найденные вхождения по нормализованной форме.

        После группировки получается более удобная таблица:
        - пример словосочетания;
        - его нормализованная форма;
        - предлагаемая аббревиатура;
        - количество слов;
        - частота встречаемости;
        - несколько примеров предложений.
        """
        if not mentions:
            return pd.DataFrame(columns=[
                "phrase_example",
                "normalized_phrase",
                "proposed_abbreviation",
                "word_count",
                "frequency",
                "examples"
            ])

        grouped: Dict[str, Dict] = {}

        for mention in mentions:
            key = mention.normalized_phrase

            if key not in grouped:
                grouped[key] = {
                    "phrase_example": mention.phrase,
                    "normalized_phrase": mention.normalized_phrase,
                    "proposed_abbreviation": mention.proposed_abbreviation,
                    "word_count": mention.word_count,
                    "frequency": 0,
                    "examples": []
                }

            grouped[key]["frequency"] += 1

            # Храним не более 3 примеров,
            # чтобы таблица не становилась слишком большой
            if len(grouped[key]["examples"]) < 3:
                grouped[key]["examples"].append(mention.sentence)

        data = []
        for item in grouped.values():
            data.append({
                "phrase_example": item["phrase_example"],
                "normalized_phrase": item["normalized_phrase"],
                "proposed_abbreviation": item["proposed_abbreviation"],
                "word_count": item["word_count"],
                "frequency": item["frequency"],
                "examples": " || ".join(item["examples"])
            })

        df = pd.DataFrame(data)

        # Финальная фильтрация после агрегации:
        #
        # 1) если кандидат состоит только из 2 слов,
        #    оставляем его лишь при частоте >= 2
        #    (иначе таких фраз слишком много и среди них много шума);
        #
        # 2) если слов 3 и больше — допускаем даже единичные вхождения,
        #    потому что длинные конструкции чаще являются реальными терминами.
        df = df[
            ((df["word_count"] == 2) & (df["frequency"] >= 2)) |
            (df["word_count"] >= 3)
        ].copy()

        # Финальная постобработка:
        # - убираем подфразы более длинных терминов;
        # - убираем процессный / организационный шум.
        df = self.postfilter_aggregated_candidates(df)

        # Сортировка:
        # сначала более частотные,
        # затем более длинные,
        # затем по алфавиту
        df = df.sort_values(
            by=["frequency", "word_count", "phrase_example"],
            ascending=[False, False, True]
        ).reset_index(drop=True)

        return df

    def postfilter_aggregated_candidates(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Финальная очистка агрегированных кандидатов.

        Что делает:
        1. Убирает процессные / организационные конструкции.
        2. Убирает подфразы, если они уже входят в более длинный термин.
        3. Возвращает очищенный DataFrame.

        Это особенно полезно для случаев вроде:
        - "Модуль анализа" при наличии "Модуль анализа текстовых документов"
        - "текстовых документов" как хвост более длинной фразы
        - "Рабочая группа" как организационная сущность
        """
        if df.empty:
            return df

        rows = df.to_dict("records")
        filtered_rows = []

        # Сортируем по длине фразы: длинные и более информативные выше.
        rows_sorted = sorted(
            rows,
            key=lambda x: (x["word_count"], len(str(x["phrase_example"]))),
            reverse=True
        )

        kept_phrases: List[str] = []

        for row in rows_sorted:
            phrase = str(row["phrase_example"]).strip().lower()
            normalized = str(row["normalized_phrase"]).strip().lower()
            lemmas = normalized.split()

            # 1. Убираем процессный шум на уровне агрегированной фразы.
            noise_count = sum(1 for lemma in lemmas if lemma in self.extra_noise_lemmas)
            if noise_count >= 2:
                continue

            # 2. Убираем подфразы уже сохранённых более длинных терминов.
            is_subphrase = False
            for kept in kept_phrases:
                if phrase != kept and phrase in kept:
                    is_subphrase = True
                    break

            if is_subphrase:
                continue

            kept_phrases.append(phrase)
            filtered_rows.append(row)

        result = pd.DataFrame(filtered_rows)

        if result.empty:
            return result

        result = result.sort_values(
            by=["frequency", "word_count", "phrase_example"],
            ascending=[False, False, True]
        ).reset_index(drop=True)

        return result

    # ========================================================
    # 12. Сохранение результатов
    # ========================================================

    def save_results(self, mentions: List[CandidateMention], output_dir: str | Path) -> Dict[str, Path]:
        """
        Сохраняет результаты анализа в файлы.

        Всегда сохраняет CSV.
        Дополнительно пытается сохранить XLSX.
        Если openpyxl не установлен, CSV всё равно сохранятся.
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        raw_df = self.mentions_to_dataframe(mentions)
        agg_df = self.aggregate_mentions(mentions)

        raw_csv = output_dir / "raw_mentions.csv"
        agg_csv = output_dir / "aggregated_candidates.csv"
        raw_xlsx = output_dir / "raw_mentions.xlsx"
        agg_xlsx = output_dir / "aggregated_candidates.xlsx"

        # CSV сохраняем всегда
        raw_df.to_csv(raw_csv, index=False, encoding="utf-8-sig")
        agg_df.to_csv(agg_csv, index=False, encoding="utf-8-sig")

        saved_files: Dict[str, Path] = {
            "raw_csv": raw_csv,
            "aggregated_csv": agg_csv,
        }

        # Попытка сохранить в Excel.
        # Если openpyxl не установлен, просто выводим предупреждение.
        try:
            raw_df.to_excel(raw_xlsx, index=False)
            agg_df.to_excel(agg_xlsx, index=False)

            saved_files["raw_xlsx"] = raw_xlsx
            saved_files["aggregated_xlsx"] = agg_xlsx

        except ModuleNotFoundError as exc:
            print("Внимание: не удалось сохранить XLSX-файлы.")
            print("Причина:", exc)
            print("CSV-файлы при этом успешно сохранены.")

        return saved_files


# ============================================================
# 13. Локальный запуск
# ============================================================

if __name__ == "__main__":
    """
    Пример запуска:
    1. Положите рядом с файлом .docx-документ.
    2. Укажите его имя в input_docx.
    3. Запустите файл.
    """

    # Имя входного документа Word.
    # При необходимости замените на своё.
    input_docx = "test_reduction_input.docx"

    # Папка, куда будут сохранены результаты
    output_dir = "result_stage1_v3"

    recognizer = ReducibleWordformRecognizerV3()

    # Шаг 1. Анализируем документ
    mentions = recognizer.analyze_document(input_docx)

    # Шаг 2. Сохраняем найденные кандидаты
    saved_files = recognizer.save_results(mentions, output_dir)

    # Шаг 3. Выводим краткую сводку в консоль
    print("=" * 60)
    print("ЭТАП 1 V3 ЗАВЕРШЁН: распознавание текста и поиск кандидатов")
    print("=" * 60)
    print(f"Всего найдено вхождений-кандидатов: {len(mentions)}")
    print("Сформированы файлы:")

    for name, path in saved_files.items():
        print(f"{name}: {path}")
