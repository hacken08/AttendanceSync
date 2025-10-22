from loguru import logger
from pretty_json_loguru import setup_json_loguru
import sys

# Configure Loguru to serialize logs to JSON and send them to the console
logger.remove()
# logger.add(sys.stderr, serialize=True)
setup_json_loguru(level="DEBUG")