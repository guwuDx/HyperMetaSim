import logging
import time
import os
from tqdm import tqdm


class TqdmLoggingHandler(logging.Handler):
    """
    A custom logging handler that uses tqdm.write() so logs go above the progress bar.
    """
    def emit(self, record):
        # 1) Format the record
        msg = self.format(record)
        # 2) Use tqdm.write() instead of print() or sys.stdout,
        #    so that it doesn't ruin the tqdm progress bar
        tqdm.write(msg)


class ContextFormatter(logging.Formatter):
    """
    Custom formatter that uses %(context)s if present,
    otherwise prints empty.
    """
    def format(self, record):
        # If record has no 'context', set it to '[Thread/Main]'
        if not hasattr(record, 'context'):
            record.context = 'Thread/Main'
        return super().format(record)


def setup_logging_with_tqdm():
    """
    Set up logging so that:
      - Logs are written to a timestamped file
      - Console logs go through TqdmLoggingHandler,
        ensuring the tqdm progress bar stays at the bottom.
    """
    # 1) Determine log file path
    current_time = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file = os.path.join(log_dir, current_time + ".log")

    # 2) Configure basic logging to a file (for detailed logs)
    #    We'll not rely on the default console stream here.
    logging.basicConfig(
        level=logging.DEBUG,
        format="[%(asctime)s][%(levelname)s]: %(message)s",
        datefmt="%Y-%m-%d/%H:%M:%S",
        filename=log_file,
        filemode="a"
    )

    # 3) Remove the default StreamHandler (if any)
    root_logger = logging.getLogger()
    if root_logger.hasHandlers():
        # By default, basicConfig might create a StreamHandler
        # We remove them so we can replace them with TqdmLoggingHandler
        root_logger.handlers = [
            h for h in root_logger.handlers
            if not isinstance(h, logging.StreamHandler)
        ]

    # 4) Create our TqdmLoggingHandler for console output
    console_handler = TqdmLoggingHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = ContextFormatter(
        "[%(asctime)s][%(context)s][%(levelname)s]: %(message)s",
        datefmt="%Y-%m-%d/%H:%M:%S"
    )
    console_handler.setFormatter(console_formatter)
    root_logger.addHandler(console_handler)

    # Now root logger will write DEBUG+ to file, and INFO+ to console via TqdmLoggingHandler.
    logging.info(f"Logging is set up. File logs -> {log_file}")

# automatically set up logging when this module is imported
setup_logging_with_tqdm()