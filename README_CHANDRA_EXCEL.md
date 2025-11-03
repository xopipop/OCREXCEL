# Chandra Excel Integration

Интеграция Chandra OCR с Excel через COM-объект Python для автоматической обработки PDF документов с использованием нейросетей.

## Возможности

- 📄 OCR для PDF и изображений через Chandra
- 🤖 Обработка данных через OpenRouter/ChatGPT
- 📊 Запись результатов в Excel
- 🎯 Готовые шаблоны промптов
- 🔌 Интеграция через VBA

## Архитектура

```
Excel VBA → COM Объект (Python) → Chandra OCR → OpenRouter API → Excel
```

## Требования

### Железо
- Windows 10/11
- 16+ ГБ ОЗУ
- NVIDIA GPU (опционально, для ускорения)

### Программное обеспечение
- Python 3.10+
- Microsoft Excel 2016+
- NVIDIA CUDA (опционально)

## Установка

### 1. Установка Python

Скачайте и установите Python с официального сайта:
```
https://www.python.org/downloads/
```

**Важно:** При установке отметьте "Add Python to PATH"

### 2. Установка зависимостей

Откройте командную строку в папке проекта и выполните:

```bash
pip install -r requirements.txt
```

Или установите вручную:

```bash
pip install chandra-ocr pywin32 openai python-dotenv
```

### 3. Настройка OpenRouter API

1. Зарегистрируйтесь на [OpenRouter.ai](https://openrouter.ai/)
2. Получите API ключ
3. Настройте переменные окружения:

**Windows:**
```cmd
set OPENROUTER_API_KEY=sk-or-v1-ваш-ключ
set OPENROUTER_MODEL=openai/gpt-4o
```

**Постоянно (Windows):**
1. Win+R → `sysdm.cpl` → Дополнительно → Переменные среды
2. Добавьте переменную `OPENROUTER_API_KEY`

### 4. Регистрация COM-сервера

Выполните в командной строке:

```bash
register_com.bat
```

Или вручную:

```bash
python chandra_excel_com.py --register
```

### 5. Настройка Excel

**Автоматическая установка (рекомендуется):**

1. Откройте Excel, нажмите `Alt+F11`
2. Insert → Module
3. Скопируйте содержимое файла `INSTALLER.vba` в модуль
4. Нажмите `Alt+F8`, выберите `InstallChandraOCR`, нажмите "Выполнить"
5. Дождитесь завершения установки
6. Сохраните файл как `ChandraOCR.xlsm` (с поддержкой макросов)

Все модули и формы созданы автоматически!

### 6. Назначение горячей клавиши

1. Alt+F8
2. Выберите `ShowChandraForm`
3. Нажмите "Параметры"
4. Выберите сочетание клавиш (например, Ctrl+Shift+C)

## Использование

### Базовое использование

1. Откройте Excel с макросами
2. Нажмите Ctrl+Shift+C (или ваше сочетание)
3. В форме:
   - Выберите PDF файл (кнопка "Обзор...")
   - Выберите шаблон промпта из списка
   - Укажите начальную ячейку (например, A5)
4. Нажмите "ОК"
5. Дождитесь обработки
6. Результаты появятся в Excel

### Доступные шаблоны

1. **Универсальный** - автоматическое извлечение структурированных данных
2. **Счет-фактура** - извлечение данных счетов-фактур
3. **Договор** - извлечение данных договоров
4. **Таблица** - извлечение табличных данных
5. **Контакты** - извлечение контактов и событий
6. **Акт выполненных работ** - извлечение актов

### Кастомный промпт

Если ни один шаблон не подходит:
1. Не выбирайте шаблон в ComboBox
2. Введите свой промпт в текстовое поле
3. **Важно:** GPT должен вернуть JSON

Пример кастомного промпта:

```
Извлеки следующую информацию в JSON:
{
  "название_поля": "значение",
  "другое_поле": "значение"
}
```

## Структура данных

### Формат результата

COM-объект возвращает JSON строку:

```json
{
  "success": true,
  "data": {
    "поле1": "значение1",
    "поле2": "значение2",
    ...
  },
  "stats": {
    "extracted_chars": 1500,
    "response_chars": 500
  }
}
```

### Ошибки

При ошибке:

```json
{
  "error": "Описание ошибки"
}
```

## Работа с шаблонами

### Добавление нового шаблона

Отредактируйте `templates_config.json`:

```json
{
  "templates": {
    "Мой шаблон": {
      "prompt": "Извлеки данные в JSON: { \"поле\": \"значение\" }",
      "fields": ["поле1", "поле2", "поле3"]
    }
  }
}
```

**Поля:**
- `prompt` - текст промпта для GPT
- `fields` - порядок записи в Excel (слева направо)

### Специальные значения fields

- `["all"]` - записать все поля из JSON в порядке возврата
- `["поле1", "поле2"]` - записать только указанные поля в заданном порядке

## Отладка

### Логи

Логи COM-сервера сохраняются в:
```
logs/chandra_com.log
```

### Тест подключения

В VBA редакторе:

```vba
Sub TestConnection()
    TestConnection  ' Вызов функции из модуля
End Sub
```

### Проверка COM-объекта

```vba
Sub TestCOM()
    Dim obj As Object
    Set obj = CreateObject("ChandraExcel.Processor")
    MsgBox obj.GetTemplates()
End Sub
```

## Устранение неполадок

### Ошибка: "COM-сервер не зарегистрирован"

Выполните:
```bash
register_com.bat
```

Если не помогло, запустите от имени администратора.

### Ошибка: "OPENROUTER_API_KEY не установлен"

Проверьте переменные окружения:
```cmd
echo %OPENROUTER_API_KEY%
```

Установите если нужно:
```cmd
set OPENROUTER_API_KEY=ваш-ключ
```

### Ошибка при загрузке модели Chandra

Убедитесь что установлена Chandra OCR:
```bash
pip install chandra-ocr
```

Для ускорения установите flash-attention:
```bash
pip install flash-attention
```

### Медленная обработка

Chandra OCR по умолчанию работает на CPU (метод `hf`).

Для ускорения:
1. Установите Docker Desktop
2. Запустите vLLM сервер (требуется NVIDIA GPU)
3. Измените в `chandra_excel_com.py`:
   ```python
   self.chandra_model = InferenceManager(method='vllm')
   ```

## Удаление

### Удаление COM-регистрации

```bash
python chandra_excel_com.py --unregister
```

### Удаление зависимостей

```bash
pip uninstall chandra-ocr pywin32 openai python-dotenv
```

## Файлы проекта

```
.
├── INSTALLER.vba                 # Единый установщик всех компонентов
├── chandra_excel_com.py          # COM-сервер Python
├── templates_config.json          # Конфигурация шаблонов
├── requirements.txt               # Зависимости Python
├── env.example                    # Пример конфигурации .env
├── register_com.bat               # Скрипт регистрации COM
├── setup_env.bat                  # Автоматическая установка
├── QUICKSTART.md                  # Быстрый старт
├── README_CHANDRA_EXCEL.md        # Полная документация
└── logs/                          # Логи
    └── chandra_com.log
```

## Лицензия

Этот проект использует:
- Chandra OCR - Apache 2.0 / OpenRAIL-M
- PyWin32 - Python Software Foundation License
- OpenAI Library - MIT License

## Поддержка

- GitHub Issues: [создать issue]
- Email: support@example.com

## Changelog

### v1.0.0 (2024-01-XX)
- Первый релиз
- Базовая интеграция Chandra + Excel
- Поддержка шаблонов промптов
- COM интерфейс для VBA
