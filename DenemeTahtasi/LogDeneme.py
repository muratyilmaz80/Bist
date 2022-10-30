import logging
import sys
from logging.handlers import TimedRotatingFileHandler

from prettytable import PrettyTable

FORMATTER = logging.Formatter("%(asctime)s — %(name)s — %(levelname)s — %(message)s")
LOG_FILE = "my_app.log"

def get_console_handler():
   console_handler = logging.StreamHandler(sys.stdout)
   console_handler.setFormatter(FORMATTER)
   return console_handler

def get_file_handler():
   file_handler = TimedRotatingFileHandler(LOG_FILE)
   file_handler.setFormatter(FORMATTER)
   return file_handler

def get_logger(logger_name):
   logger = logging.getLogger(logger_name)
   logger.setLevel(logging.DEBUG) # better to have too much log than not enough
   logger.addHandler(get_console_handler())
   logger.addHandler(get_file_handler())
   # with this pattern, it's rarely necessary to propagate the error up to parent
   logger.propagate = False
   return logger

#my_logger = get_logger("my module name")
#my_logger.debug("a debug message")




# 2.KULLANIM

my_logger = logging.getLogger()
my_logger.setLevel(logging.DEBUG)

output_file_handler = logging.FileHandler("output.log")
#output_file_handler.setFormatter(logging.Formatter("%(asctime)s — %(name)s — %(levelname)s — %(message)s"))

stdout_handler = logging.StreamHandler(sys.stdout)
#stdout_handler.setFormatter(logging.Formatter("%(asctime)s — %(name)s — %(levelname)s — %(message)s"))

my_logger.addHandler(output_file_handler)
my_logger.addHandler(stdout_handler)

for i in range (1,4):
   my_logger.debug("This is line " + str(i))

#printx = "Bilanço Dönemi Faaliyet Kari Artisi: " + "{:.2%}".format(1.234567) + " >? 15% " + str(True)

#my_logger.info(printx)

my_logger.info("Bilanço Dönemi Özsermaye Karlılığı: %s", "{:.2%}".format(0.087678))
