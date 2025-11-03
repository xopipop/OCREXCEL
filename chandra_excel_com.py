"""
Chandra Excel COM Server
COM-сервер для интеграции Chandra OCR с Excel через VBA
"""

import json
import os
import sys
import logging
from pathlib import Path
from typing import Dict, Any
import traceback

# Импорт win32com должен быть в начале для правильной работы COM
import win32com.server.register

# Импорт Chandra OCR
from chandra.model import InferenceManager
from chandra.input import load_file
from chandra.model.schema import BatchInputItem

# Импорт OpenAI для OpenRouter
from openai import OpenAI
import pythoncom

# Загрузка переменных окружения из .env
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv опционален

# Настройка логирования
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
log_file = LOG_DIR / "chandra_com.log"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)


class ChandraExcelProcessor:
    """COM-сервер для обработки документов через Chandra OCR + OpenRouter"""
    
    _public_methods_ = ['ProcessDocument', 'GetTemplates', 'TestConnection']
    _reg_progid_ = "ChandraExcel.Processor"
    _reg_clsid_ = "{A1B2C3D4-E5F6-4A5B-8C9D-0E1F2A3B4C5D}"
    
    def __init__(self):
        """Инициализация процессора"""
        logger.info("Инициализация ChandraExcelProcessor")
        self.chandra_model = None
        self.openai_client = None
        self.templates = None
        
    def _get_chandra_model(self):
        """Ленивая загрузка модели Chandra OCR"""
        if self.chandra_model is None:
            logger.info("Загрузка модели Chandra OCR...")
            try:
                self.chandra_model = InferenceManager(method='hf')
                logger.info("Модель Chandra OCR загружена успешно")
            except Exception as e:
                logger.error(f"Ошибка загрузки модели Chandra: {e}")
                raise
        return self.chandra_model
    
    def _get_openai_client(self):
        """Инициализация клиента OpenRouter"""
        if self.openai_client is None:
            logger.info("Инициализация OpenRouter клиента...")
            
            # Загружаем API ключ из переменных окружения
            api_key = os.getenv('OPENROUTER_API_KEY')
            if not api_key:
                error_msg = "OPENROUTER_API_KEY не установлен в переменных окружения"
                logger.error(error_msg)
                raise ValueError(error_msg)
            
            # Base URL для OpenRouter
            base_url = os.getenv('OPENROUTER_BASE_URL', 'https://openrouter.ai/api/v1')
            
            # Имя модели
            model_name = os.getenv('OPENROUTER_MODEL', 'openai/gpt-4o')
            
            self.openai_client = OpenAI(
                base_url=base_url,
                api_key=api_key
            )
            self.model_name = model_name
            
            logger.info(f"OpenRouter клиент инициализирован, модель: {model_name}")
        
        return self.openai_client
    
    def _load_templates(self):
        """Загрузка шаблонов из конфигурационного файла"""
        if self.templates is None:
            config_path = Path(__file__).parent / "templates_config.json"
            
            if not config_path.exists():
                logger.warning(f"Файл шаблонов не найден: {config_path}")
                return {}
            
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.templates = config.get('templates', {})
                    logger.info(f"Загружено шаблонов: {len(self.templates)}")
            except Exception as e:
                logger.error(f"Ошибка загрузки шаблонов: {e}")
                self.templates = {}
        
        return self.templates
    
    def GetTemplates(self):
        """
        Получить список доступных шаблонов
        
        Returns:
            str: JSON строка со списком шаблонов
        """
        logger.info("Запрос списка шаблонов")
        templates = self._load_templates()
        
        template_names = list(templates.keys())
        result = {"templates": template_names}
        
        logger.info(f"Возвращено шаблонов: {len(template_names)}")
        return json.dumps(result, ensure_ascii=False, indent=2)
    
    def TestConnection(self):
        """
        Проверка подключения к сервисам
        
        Returns:
            str: JSON строка с результатом теста
        """
        logger.info("Тест подключения")
        result = {"chandra": False, "openrouter": False, "message": ""}
        
        try:
            # Тест Chandra
            try:
                self._get_chandra_model()
                result["chandra"] = True
                result["message"] = "Все сервисы доступны"
            except Exception as e:
                result["message"] = f"Ошибка Chandra: {str(e)}"
                logger.error(f"Тест Chandra провален: {e}")
            
            # Тест OpenRouter
            try:
                client = self._get_openai_client()
                result["openrouter"] = True
            except Exception as e:
                result["message"] = f"Ошибка OpenRouter: {str(e)}"
                logger.error(f"Тест OpenRouter провален: {e}")
        
        except Exception as e:
            result["message"] = f"Критическая ошибка: {str(e)}"
            logger.error(f"Ошибка теста подключения: {e}")
        
        return json.dumps(result, ensure_ascii=False)
    
    def ProcessDocument(self, file_path: str, prompt: str = "", template_name: str = ""):
        """
        Обработка документа через Chandra OCR и OpenRouter
        
        Args:
            file_path: Путь к PDF или изображению
            prompt: Пользовательский промпт (необязательно, если указан template_name)
            template_name: Имя шаблона из templates_config.json (необязательно)
        
        Returns:
            str: JSON строка с результатом обработки или ошибкой
        """
        logger.info(f"Обработка документа: {file_path}")
        
        try:
            # Валидация пути к файлу
            file_path_obj = Path(file_path)
            if not file_path_obj.exists():
                error_msg = f"Файл не найден: {file_path}"
                logger.error(error_msg)
                return json.dumps({"error": error_msg}, ensure_ascii=False)
            
            # Определение промпта
            if template_name:
                templates = self._load_templates()
                if template_name not in templates:
                    error_msg = f"Шаблон не найден: {template_name}"
                    logger.error(error_msg)
                    return json.dumps({"error": error_msg}, ensure_ascii=False)
                
                template = templates[template_name]
                final_prompt = template.get('prompt', '')
                logger.info(f"Используется шаблон: {template_name}")
            elif prompt:
                final_prompt = prompt
                logger.info("Используется пользовательский промпт")
            else:
                error_msg = "Не указан промпт или шаблон"
                logger.error(error_msg)
                return json.dumps({"error": error_msg}, ensure_ascii=False)
            
            # Шаг 1: OCR через Chandra
            logger.info("Начало OCR обработки...")
            extracted_text = self._perform_ocr(file_path)
            
            if not extracted_text:
                error_msg = "Не удалось извлечь текст из документа"
                logger.error(error_msg)
                return json.dumps({"error": error_msg}, ensure_ascii=False)
            
            logger.info(f"OCR завершен, извлечено символов: {len(extracted_text)}")
            
            # Шаг 2: Обработка через OpenRouter
            logger.info("Отправка данных в OpenRouter...")
            gpt_response = self._process_with_openrouter(extracted_text, final_prompt)
            
            if not gpt_response:
                error_msg = "Не получен ответ от OpenRouter"
                logger.error(error_msg)
                return json.dumps({"error": error_msg}, ensure_ascii=False)
            
            logger.info(f"Получен ответ от OpenRouter, символов: {len(gpt_response)}")
            
            # Шаг 3: Парсинг JSON
            try:
                parsed_data = self._parse_json_response(gpt_response)
                logger.info("JSON успешно распарсен")
            except Exception as e:
                logger.error(f"Ошибка парсинга JSON: {e}")
                return json.dumps({
                    "error": f"Ошибка парсинга JSON: {str(e)}",
                    "raw_response": gpt_response[:500]  # Первые 500 символов для отладки
                }, ensure_ascii=False)
            
            # Возврат результата
            result = {
                "success": True,
                "data": parsed_data,
                "stats": {
                    "extracted_chars": len(extracted_text),
                    "response_chars": len(gpt_response)
                }
            }
            
            logger.info("Обработка завершена успешно")
            return json.dumps(result, ensure_ascii=False, indent=2)
        
        except Exception as e:
            error_msg = f"Критическая ошибка: {str(e)}"
            logger.error(f"{error_msg}\n{traceback.format_exc()}")
            return json.dumps({"error": error_msg}, ensure_ascii=False)
    
    def _perform_ocr(self, file_path: str) -> str:
        """
        Выполнение OCR через Chandra
        
        Args:
            file_path: Путь к файлу
        
        Returns:
            str: Извлеченный текст в формате markdown
        """
        model = self._get_chandra_model()
        
        # Загрузка документа
        images = load_file(file_path, {})
        
        if not images:
            raise ValueError("Не удалось загрузить изображения из документа")
        
        logger.info(f"Загружено страниц: {len(images)}")
        
        # Обработка всех страниц
        all_markdown = []
        for i, img in enumerate(images):
            logger.info(f"Обработка страницы {i+1}/{len(images)}")
            batch = [BatchInputItem(image=img, prompt_type="ocr_layout")]
            result = model.generate(batch)[0]
            
            if result.error:
                logger.warning(f"Ошибка при обработке страницы {i+1}")
                continue
            
            all_markdown.append(result.markdown)
        
        # Объединение всех страниц
        full_text = "\n\n--- Страница ---\n\n".join(all_markdown)
        return full_text
    
    def _process_with_openrouter(self, extracted_text: str, prompt: str) -> str:
        """
        Обработка текста через OpenRouter API
        
        Args:
            extracted_text: Извлеченный текст из OCR
            prompt: Промпт для обработки
        
        Returns:
            str: Ответ от GPT
        """
        client = self._get_openai_client()
        
        # Формирование полного промпта
        full_prompt = f"""{prompt}

Содержимое документа:
{extracted_text}

Верни результат ТОЛЬКО в формате JSON без дополнительных объяснений."""
        
        # Вызов API
        response = client.chat.completions.create(
            model=self.model_name,
            messages=[
                {
                    "role": "system",
                    "content": "Ты помощник по извлечению структурированных данных из документов. Всегда возвращай результат ТОЛЬКО в формате JSON без дополнительного текста."
                },
                {
                    "role": "user",
                    "content": full_prompt
                }
            ],
            temperature=0.1,
            max_tokens=4000
        )
        
        result = response.choices[0].message.content
        return result
    
    def _parse_json_response(self, gpt_response: str) -> Dict[str, Any]:
        """
        Парсинг JSON ответа от GPT
        
        Args:
            gpt_response: Ответ от GPT
        
        Returns:
            dict: Распарсенные данные
        """
        # Удаляем markdown код-блоки если есть
        if "```json" in gpt_response:
            json_start = gpt_response.find("```json") + 7
            json_end = gpt_response.find("```", json_start)
            gpt_response = gpt_response[json_start:json_end].strip()
        elif "```" in gpt_response:
            json_start = gpt_response.find("```") + 3
            json_end = gpt_response.rfind("```")
            gpt_response = gpt_response[json_start:json_end].strip()
        
        # Извлекаем JSON
        json_start = gpt_response.find('{')
        json_end = gpt_response.rfind('}') + 1
        
        if json_start >= 0 and json_end > json_start:
            json_str = gpt_response[json_start:json_end]
            return json.loads(json_str)
        else:
            raise ValueError("Не найден JSON в ответе GPT")
    
    def _check_registry(self):
        """Проверка регистрации в реестре"""
        return


def main():
    """Регистрация COM-сервера"""
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--register":
        logger.info("Регистрация COM-сервера...")
        pythoncom.CoInitialize()
        win32com.server.register.UseCommandLine(ChandraExcelProcessor)
        logger.info("COM-сервер зарегистрирован успешно")
        print("COM-сервер зарегистрирован успешно!")
    elif len(sys.argv) > 1 and sys.argv[1] == "--unregister":
        logger.info("Удаление регистрации COM-сервера...")
        pythoncom.CoInitialize()
        win32com.server.register.UseCommandLine(ChandraExcelProcessor)
        logger.info("Регистрация COM-сервера удалена")
        print("Регистрация COM-сервера удалена!")
    else:
        # Запуск в режиме тестирования
        print("Для регистрации COM-сервера запустите:")
        print("  python chandra_excel_com.py --register")
        print("\nДля удаления регистрации:")
        print("  python chandra_excel_com.py --unregister")


if __name__ == "__main__":
    main()
