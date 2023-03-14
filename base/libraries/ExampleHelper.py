from time import sleep

from robot.api import logger


def example_python_keyword():
    logger.warn("HOLAMUNDO DESDE PYTHON")
    logger.error("HOLAMUNDO DESDE PYTHON")

    for i in range(25):
        sleep(1)
        logger.info(f'Log {i}')
