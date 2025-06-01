"""
Main entry point for Mac PST Converter
"""

import sys
import logging
from pathlib import Path
from .gui.main_window import MainWindow


def setup_logging():
    """Set up logging configuration"""
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(Path.home() / 'macpst_converter.log')
        ]
    )


def main():
    """Main entry point"""
    try:
        setup_logging()
        logger = logging.getLogger(__name__)
        logger.info("Starting Mac PST Converter")
        
        app = MainWindow()
        app.run()
        
    except Exception as e:
        logging.error(f"Application error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()