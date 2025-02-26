# Пакетная печать документов

Программа для автоматизации процесса массовой печати документов. Она позволяет выбирать папку с файлами и отправлять их на печать одним нажатием кнопки, устраняя необходимость ручного открытия каждого файла.

## О проекте  

### Основные возможности:
- Поддержка различных форматов файлов (PDF, DOCX, TXT и др.).
- Простая настройка: выбор папки с файлами через графический интерфейс.
- Автоматическая отправка файлов на принтер по умолчанию.

### Установка

Для работы с программой необходимо:
- Установить Python 3.10 или выше.
- Установить необходимые библиотеки:
```bash
pip install pywin32
```
- Упаковать программу в .exe файл (опционально, если не хотите запускать через Python):
```bash
pyinstaller --onefile --icon=print.ico print.py
```

## Использование  

1. Запустите программу (или двойным щелчком по print.exe, или командой python print.py).
2. Выберите папку с файлами для печати через графический интерфейс.
3. Программа автоматически отправит все файлы из выбранной папки на принтер

## Поддержка и доработки

Если вы столкнулись с ошибками или хотите предложить новые функции, пожалуйста, создайте Issue в репозитории. Мы открыты для предложений и обратной связи!

## Лицензия

Проект предоставляется на условиях MIT License.

 
