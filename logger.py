import logging
import os
from datetime import datetime
from logging.handlers import RotatingFileHandler


def setup_logging():
    """Настраивает систему логирования с удобным форматированием."""
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)

    # Имя файла с датой
    log_filename = os.path.join(log_dir, f"app_{datetime.now().strftime('%Y-%m-%d')}.log")

    # Создаем форматтер с более читаемым форматом
    formatter = logging.Formatter(
        '%(asctime)s | %(levelname)-8s | %(message)s',
        datefmt='%H:%M:%S'  # Только время без даты для компактности
    )

    # Ротация логов для ограничения размера файла
    file_handler = RotatingFileHandler(
        log_filename,
        maxBytes=5 * 1024 * 1024,  # 5 МБ максимальный размер
        backupCount=3,  # Хранить до 3 файлов истории
        encoding='utf-8'
    )
    file_handler.setFormatter(formatter)

    # Настройка вывода в консоль
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    # Настраиваем root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # Удаляем существующие обработчики, если они есть
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    # Добавляем новые обработчики
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    # Логируем начало новой сессии с разделителем для удобства чтения
    logging.info("=" * 50)
    logging.info("НАЧАЛО НОВОЙ СЕССИИ")
    logging.info("=" * 50)


def log_separator(message=None):
    """Добавляет разделитель в лог-файл с опциональным сообщением."""
    if message:
        logging.info("-" * 20 + f" {message} " + "-" * 20)
    else:
        logging.info("-" * 50)