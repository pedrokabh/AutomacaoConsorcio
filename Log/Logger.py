import logging
from datetime import datetime

class ConsoleLogger:
    def __init__(self, execution_number, logger_name='ConsoleLogger', level=logging.DEBUG):
        self.logger = logging.getLogger(logger_name)
        self.logger.setLevel(level)
        self.execution_number = execution_number
        self.execution_start = datetime.now()
        
        # Formato do log
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # Log no console
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)
        
        # Log em arquivo
        file_handler = logging.FileHandler(f'Execution {str(execution_number)} - .log')
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        self.info(f"[START] DateTime Now [{str(datetime.now())}] Execution N°{execution_number}")
    
    def debug(self, message):
        self.logger.debug(message)
    
    def info(self, message):
        self.logger.info(message)
    
    def warning(self, message):
        self.logger.warning(message)
    
    def error(self, message):
        self.logger.error(message)
    
    def critical(self, message):
        self.logger.critical(message)
    
    def closeConsole(self):
        self.logger.debug(f"[END] Execution Finish at [{str(datetime.now())}]")
        self.logger.shutdown()