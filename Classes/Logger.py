import logging
from datetime import datetime

class Logger:
    def __init__(self, execution_number, write_log_directory, logger_name='ConsoleLogger', level=logging.DEBUG):
        self.logger = logging.getLogger(logger_name)
        self.logger.setLevel(level)
        self.execution_number = execution_number
        self.execution_start = datetime.now()
        
        # Formato do log
        formatter = logging.Formatter('[%(asctime)s - %(levelname)s - %(message)s]')
        
        # Log no console
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        self.logger.addHandler(console_handler)
        
        # Log em arquivo
        file_handler = logging.FileHandler(write_log_directory)
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
    
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
        logging.shutdown()