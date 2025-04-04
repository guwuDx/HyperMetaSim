import logging
import time
import os

def setup_logging():
    # get current date & time
    current_time = time.strftime("%Y%m%d_%H%M%S", time.localtime())

    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file = os.path.join(log_dir, current_time + ".log")

    logging.basicConfig(
        level=logging.DEBUG,
        format="[%(asctime)s][%(levelname)s]: %(message)s",
        datefmt="%Y-%m-%d/%H:%M:%S",
        filename=log_file,   # 指定日志文件路径
        filemode="a"         # 追加写入
    )

    # 输出到控制台，添加 StreamHandler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter("[%(asctime)s][%(levelname)s]: %(message)s",
                                          datefmt="%Y-%m-%d/%H:%M:%S")
    console_handler.setFormatter(console_formatter)
    logging.getLogger().addHandler(console_handler)

# 当模块导入时自动加载日志配置
setup_logging()
