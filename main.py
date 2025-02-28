import logging
from ui import run_ui, log_separator

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)

if __name__ == "__main__":
    logging.info("Запуск приложения")
    log_separator()
    run_ui()
