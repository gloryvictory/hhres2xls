import os
from time import strftime  # Load just the strftime Module from Time

DATETIME_CURRENT = str(strftime("%Y-%m-%d-%H-%M-%S"))

FILE_LOG_NAME = 'hhres2xls'
FILE_LOG = DATETIME_CURRENT + '_' + FILE_LOG_NAME + '.log'
FILE_LOG_FORMAT = '%(asctime)s %(levelname)s %(message)s'

FOLDER_IN  = 'C:\\Glory\\Projects\\Python\\hhres2xls\\src\\data'

CSV_DELIMITER = ','

CSV_FILE = 'res_out.csv'
CSV_DICT = {'FIO': '',
                'EMAIL': '',
                'TEL': '',
                'CITY': '',
                'GENDER': '',
                'AGE': '',
                'OBR': '',
                'GR': '',
                'ZAN': '',
                'NAVIK': ''
                }
