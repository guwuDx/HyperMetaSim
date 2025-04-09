import logging
from utils import log_conf

from fastapi import FastAPI

from utils import cst_handler
from utils import misc

import state

app = FastAPI()

@app.on_event("startup")
def startup_event():
    misc.print_logo()
    logging.info("FastAPI server started.")
    state.csth = cst_handler.CSTHandler()
    logging.info("CST handler initialized.")