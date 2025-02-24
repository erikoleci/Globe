import logging
from app import app

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

if __name__ == "__main__":
    try:
        logger.info("Starting Flask server on port 5000...z")
        app.run(host="0.0.0.0", port=5000)
    except Exception as e:
        logger.error(f"Failed to start server: {str(e)}")
        raise
