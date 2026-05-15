СБОРКА EXE ДЛЯ КОНЕЧНОГО ПОЛЬЗОВАТЕЛЯ
========================================

Что входит:
- end_user_app.py              -> интерфейс для конечного пользователя
- ui_backend_entry_enduser.py  -> backend для интерфейса
- ReductionAppGUI.spec         -> spec-файл PyInstaller
- build_end_user_exe.bat       -> bat-файл для сборки
- README_build_exe.txt         -> эта инструкция

Что умеет приложение:
1. Пользователь выбирает свой Word-файл.
2. Запускает анализ документа.
3. Видит варианты новых сокращений и сам выбирает, какие вводить.
4. Выбирает место вставки списка сокращений:
   - в конец документа;
   - по маркеру;
   - в существующий раздел;
   - в отдельный файл.
5. Выбирает папку сохранения и имя итогового документа.
6. При необходимости включает замену повторных объявлений на сокращения.
7. Получает готовый итоговый Word-файл.

Как встроить в проект:
1. Скопируйте файлы:
   - end_user_app.py
   - ui_backend_entry_enduser.py
   - ReductionAppGUI.spec
   - build_end_user_exe.bat
   в корень проекта, рядом с main.py.

2. Убедитесь, что в корне проекта есть ваши рабочие модули:
   - abbreviation_need_stage3.py
   - abbreviation_list_inserter.py
   - repeated_declaration_replacer.py
   - и все остальные зависимости проекта.

3. Убедитесь, что файл requirements.txt актуален.

Как собрать:
1. Откройте cmd или PowerShell в корне проекта.
2. Запустите:
   build_end_user_exe.bat

Где будет результат:
- dist\ReductionAppGUI\ReductionAppGUI.exe

Почему сборка сделана в режиме one-dir:
- так надёжнее для python-docx, pandas, openpyxl, pymorphy2 и словарей pymorphy2;
- меньше риск получить ошибку с отсутствующими словарями или ресурсами;
- проще отлаживать и передавать конечному пользователю.

Что передавать конечному пользователю:
- всю папку dist\ReductionAppGUI целиком

Что не нужно просить у конечного пользователя:
- путь к проекту
- путь к базе сокращений
- ручной запуск Python
- командную строку

Дополнительно:
- рабочие временные файлы интерфейс хранит в LOCALAPPDATA\ReductionApp
- это сделано, чтобы приложению не требовались права записи в папку с exe


ОБНОВЛЕНИЕ v2
-------------
- В интерфейсе добавлены вертикальная и горизонтальная прокрутка всего окна.
- В таблице сокращений добавлены собственные вертикальная и горизонтальная прокрутка.
- build_end_user_exe.bat заменён на ASCII-only версию без русских строк, чтобы не было проблем с кодировкой в cmd.exe.


ИСПРАВЛЕНИЕ v3
--------------
- Исправлена ошибка Tkinter 'cannot use geometry manager grid inside ...'
- Таблица сокращений теперь создаётся в отдельном контейнере tree_frame, где для неё и скроллов используется только grid.


ИСПРАВЛЕНИЕ v4
--------------
- Добавлен патч для pymorphy2 в frozen-среде PyInstaller.
- Если entry points недоступны внутри exe, путь к словарям ru теперь ищется напрямую в bundled папках.
- В spec-файл добавлено копирование metadata для pymorphy2 и pymorphy2-dicts-ru.


ИСПРАВЛЕНИЕ v5
--------------
- Исправлена ошибка сборки PyInstaller: metadata для python-docx теперь берётся по имени distribution "python-docx", а не по import-имени "docx".
- Добавлен безопасный helper safe_copy_metadata(), чтобы сборка не падала, если metadata какого-то пакета недоступна.


ИСПРАВЛЕНИЕ v6
--------------
- Интерфейс теперь умеет читать реальные файлы stage 3 вашего проекта:
  - abbreviation_recommendations.csv/xlsx
  - abbreviation_decisions.csv/xlsx
- Поддержаны колонки:
  - term
  - suggested_abbreviation
  - abbreviation_found_in_text
  - need_to_introduce
  - decision_score
  - reason
  - priority
- Файл abbreviation_recommendations теперь ищется раньше abbreviation_decisions.


ИСПРАВЛЕНИЕ v7
--------------
- Исправлено соответствие режимов AbbreviationListInserter.run():
  - append_existing_section -> append_existing_list
  - new_document -> separate_file
- Устранена ошибка "Неизвестный режим run()".
