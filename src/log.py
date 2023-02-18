import logging as log
import os

from src import cfg


def set_logger(file_name_log='log.log'):
    for handler in log.root.handlers[:]:  # Remove all handlers associated with the root logger object.
        log.root.removeHandler(handler)

    log_folder = cfg.FOLDER_IN
    if not os.path.exists(log_folder):
        os.makedirs(log_folder)

    file_log = str(os.path.join(os.getcwd(), log_folder, file_name_log))  # from cfg.file

    # if os.path.isfile(file_log):  # Если выходной LOG файл существует - удаляем его
    #     os.remove(file_log)
    log.basicConfig(filename=file_log, format=cfg.FILE_LOG_FORMAT, level=log.INFO,
                    filemode='w', encoding='utf8')
    # log.info(file_log)
    return log