import json
import os

# Путь к файлу с настройками
CONFIG_FILE = "app_config.json"

# Стандартный список предлогов и союзов
DEFAULT_SHORT_WORDS = {
    "в", "с", "на", "по", "за", "к", "у", "о", "об", "под", "над", "из", "от",
    "для", "про", "между", "через", "перед", "при", "до", "без", "вне", "кроме",
    "вместо", "со", "ко", "во", "и", "а", "но"
}

def load_short_words():
    """Загружает список коротких слов из конфигурационного файла."""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return set(config.get('short_words', DEFAULT_SHORT_WORDS))
        return DEFAULT_SHORT_WORDS
    except Exception:
        # В случае ошибки возвращаем стандартный список
        return DEFAULT_SHORT_WORDS

def save_short_words(words_set):
    """Сохраняет список коротких слов в конфигурационный файл."""
    try:
        config = {'short_words': list(words_set)}
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False

# Загружаем список слов при импорте модуля
SHORT_WORDS = load_short_words()

# Настройки приложения
APP_NAME = "Обработка висячих предлогов"
VERSION = "1.0.0"