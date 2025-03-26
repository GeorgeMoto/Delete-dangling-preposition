from ui import run_ui
from logger import setup_logging, log_separator  # Импортируем функции логирования
import logging

if __name__ == "__main__":

    setup_logging()

    logging.info("Запуск приложения")
    log_separator("Начало работы")

    try:
        run_ui()
    except Exception as e:
        logging.critical(f"Критическая ошибка: {e}", exc_info=True)
        log_separator("Завершение работы с ошибкой")
    else:
        log_separator("Завершение работы")
